/**
 * Microsoft Graph delegated auth.
 *
 * Interactive: OAuth 2.0 authorization-code flow (PKCE) against the Teams
 * public client via Kai's `openAuthWindow`, capturing the redirect to the
 * native-client URI. Reuses PIM's Azure AD form automation for auto-login/MFA.
 *
 * Silent: refresh_token grant. Because Teams and Office share a FOCI family,
 * the family refresh token obtained above is redeemed with the *Office*
 * client_id, whose Graph `.default` preauth includes Chat.ReadWrite.
 */

import { randomBytes, createHash } from 'crypto';
import {
  AAD_AUTHORIZE_URL,
  AAD_TOKEN_URL,
  AAD_NATIVE_REDIRECT,
  AUTH_CLIENT_ID,
  GRAPH_CLIENT_ID,
  GRAPH_SCOPE,
  AUTH_SCOPES,
  AUTH_TIMEOUT_MS,
  AUTH_PARTITION,
} from '../shared/constants.js';
import type { PluginAPI, GraphTokenData, MfaState, AuthWindowHelpers } from '../shared/types.js';
import * as tokenCache from './token-cache.js';
import { hasCredentials, getCredentials } from './credential-store.js';
import { performAutoLogin, isMicrosoftLoginHost } from './auto-login.js';
import type { AutoLoginHelpers, AutoLoginCallbacks } from './auto-login.js';
import { getLogger } from './logger-singleton.js';

// ── JWT helpers ──

function decodeJwtPayload(token: string): Record<string, unknown> {
  const parts = token.split('.');
  if (parts.length < 2) throw new Error('Invalid JWT');
  return JSON.parse(Buffer.from(parts[1], 'base64url').toString('utf-8'));
}

// ── MFA deferred (frontend supplies code via panel action) ──

let mfaCodeResolve: ((code: string) => void) | null = null;
let mfaCodeReject: ((err: Error) => void) | null = null;

export function submitMfaCode(code: string): void {
  if (mfaCodeResolve) { mfaCodeResolve(code); mfaCodeResolve = null; mfaCodeReject = null; }
}

export function cancelMfaCode(): void {
  if (mfaCodeReject) { mfaCodeReject(new Error('MFA cancelled by user')); mfaCodeResolve = null; mfaCodeReject = null; }
}

// ── Concurrency guard ──

let inFlight: Promise<GraphTokenData> | null = null;

// ── Token endpoint ──

interface TokenResponse {
  access_token: string;
  refresh_token?: string;
  id_token?: string;
  expires_in: number;
  scope?: string;
  token_type: string;
  error?: string;
  error_description?: string;
}

async function tokenRequest(api: PluginAPI, form: Record<string, string>): Promise<TokenResponse> {
  const resp = await api.fetch(AAD_TOKEN_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams(form).toString(),
  });
  const body = (await resp.json()) as TokenResponse;
  if (!resp.ok || body.error) {
    throw new Error(`Token endpoint ${resp.status}: ${body.error ?? ''} ${body.error_description ?? ''}`.trim());
  }
  return body;
}

function buildTokenData(tr: TokenResponse, prevRefreshToken?: string | null): GraphTokenData {
  const access = decodeJwtPayload(tr.access_token);
  let oid = String(access.oid ?? '');
  let tid = String(access.tid ?? '');
  let email = String(access.upn ?? access.unique_name ?? access.preferred_username ?? '');
  let displayName = (access.name as string | undefined) ?? null;

  if (tr.id_token) {
    try {
      const idc = decodeJwtPayload(tr.id_token);
      oid = String(idc.oid ?? oid);
      tid = String(idc.tid ?? tid);
      email = String(idc.preferred_username ?? idc.email ?? email);
      displayName = (idc.name as string | undefined) ?? displayName;
    } catch { /* ignore */ }
  }

  return {
    accessToken: tr.access_token,
    refreshToken: tr.refresh_token ?? prevRefreshToken ?? '',
    expiresAt: Date.now() + tr.expires_in * 1000,
    objectId: oid,
    tenantId: tid,
    email,
    displayName,
    scopes: String(access.scp ?? tr.scope ?? ''),
  };
}

// ── Silent refresh (refresh_token grant, FOCI-redeemed as Office client) ──

export async function acquireTokenSilent(api: PluginAPI): Promise<GraphTokenData> {
  const rt = tokenCache.getRefreshToken();
  if (!rt) throw new Error('No refresh token available');
  const session = tokenCache.currentSession();

  getLogger().info('acquireTokenSilent: redeeming refresh token via Office client (FOCI)');
  let tr: TokenResponse;
  try {
    tr = await tokenRequest(api, {
      client_id: GRAPH_CLIENT_ID,
      grant_type: 'refresh_token',
      refresh_token: rt,
      scope: GRAPH_SCOPE,
    });
  } catch (err) {
    // Fallback: if FOCI redemption is refused for any reason, redeem as the
    // original client. Loses Chat.Read* but keeps ChatMessage.Send etc.
    getLogger().warn(`FOCI refresh via Office client failed (${err}); falling back to auth client`);
    tr = await tokenRequest(api, {
      client_id: AUTH_CLIENT_ID,
      grant_type: 'refresh_token',
      refresh_token: rt,
      scope: GRAPH_SCOPE,
    });
  }

  const token = buildTokenData(tr, rt);
  if (session !== tokenCache.currentSession()) {
    throw new Error('Session invalidated during refresh');
  }
  tokenCache.set(token);
  tokenCache.persist(api);
  getLogger().info(`Silent refresh ok for ${token.email} (scp="${token.scopes.slice(0, 120)}…")`);
  return token;
}

// ── Interactive login ──

export interface AcquireTokenOptions {
  forceRefresh?: boolean;
}

export async function acquireTokenInteractive(
  api: PluginAPI,
  opts: AcquireTokenOptions = {},
): Promise<GraphTokenData> {
  if (!opts.forceRefresh) {
    const at = tokenCache.getValidAccessToken();
    if (at) return tokenCache.get()!;
    if (tokenCache.hasRefreshToken()) {
      try {
        return await acquireTokenSilent(api);
      } catch (err) {
        getLogger().warn(`Silent refresh before interactive failed: ${err}`);
      }
    }
  }

  if (inFlight) return inFlight;
  const session = tokenCache.currentSession();
  const run = doInteractive(api, session);
  inFlight = run;
  try {
    return await run;
  } finally {
    if (inFlight === run) inFlight = null;
  }
}

async function doInteractive(api: PluginAPI, session: number): Promise<GraphTokenData> {
  const shouldAutoLogin = hasCredentials(api);
  const state = randomBytes(16).toString('hex');
  const codeVerifier = randomBytes(32).toString('base64url');
  const codeChallenge = createHash('sha256').update(codeVerifier).digest('base64url');

  const authUrl = `${AAD_AUTHORIZE_URL}?` + new URLSearchParams({
    client_id: AUTH_CLIENT_ID,
    response_type: 'code',
    redirect_uri: AAD_NATIVE_REDIRECT,
    response_mode: 'query',
    scope: AUTH_SCOPES,
    state,
    code_challenge: codeChallenge,
    code_challenge_method: 'S256',
  }).toString();

  getLogger().info(`Interactive login (autoLogin=${shouldAutoLogin})`);
  api.state.set('auth.autoLoginStatus', shouldAutoLogin ? 'Opening sign-in (auto-login)…' : 'Opening sign-in…');
  api.state.set('error', null);

  let codeSettled = false;
  const timers: Array<ReturnType<typeof setInterval>> = [];
  let settleCode!: (code: string) => void;
  let settleErr!: (err: Error) => void;
  const codePromise = new Promise<string>((res, rej) => {
    const done = () => { codeSettled = true; for (const t of timers) clearInterval(t); };
    settleCode = (c) => { if (!codeSettled) { done(); res(c); } };
    settleErr = (e) => { if (!codeSettled) { done(); rej(e); } };
  });

  const isRedirect = (url: string): URL | null => {
    try {
      const u = new URL(url);
      if (u.origin + u.pathname === AAD_NATIVE_REDIRECT) return u;
    } catch { /* ignore */ }
    return null;
  };

  const authPromise = api.auth.openAuthWindow({
    url: authUrl,
    title: 'Microsoft Teams — Sign In',
    width: 520,
    height: 640,
    timeoutMs: AUTH_TIMEOUT_MS,
    showOnCreate: !shouldAutoLogin,
    showAfterMs: shouldAutoLogin ? 180_000 : undefined,
    partition: AUTH_PARTITION,
    customUserAgent: false,
    interceptUrls: [`${AAD_NATIVE_REDIRECT}*`],
    onReady: (helpers: AuthWindowHelpers) => {
      const handleUrl = (url: string) => {
        const u = isRedirect(url);
        if (!u) return;
        const err = u.searchParams.get('error');
        if (err) {
          settleErr(new Error(`${err}: ${u.searchParams.get('error_description') ?? ''}`));
        } else if (u.searchParams.get('state') !== state) {
          settleErr(new Error('OAuth state mismatch'));
        } else {
          const code = u.searchParams.get('code');
          if (code) settleCode(code);
          else settleErr(new Error('No authorization code in redirect'));
        }
        setTimeout(() => { try { helpers.close(); } catch { /* ignore */ } }, 300);
      };

      helpers.onDidNavigate(handleUrl);
      timers.push(setInterval(() => {
        try { handleUrl(helpers.getURL()); } catch { /* window may be closing */ }
      }, 500));

      if (!shouldAutoLogin) return;
      const credentials = getCredentials(api);
      if (!credentials) { helpers.show(); return; }

      const autoLoginHelpers: AutoLoginHelpers = {
        executeJavaScript: helpers.executeJavaScript,
        getURL: helpers.getURL,
        show: helpers.show,
      };
      const live = () => session === tokenCache.currentSession();
      const callbacks: AutoLoginCallbacks = {
        onMfaCodeNeeded: (type) => {
          if (!live()) return Promise.reject(new Error('Signed out'));
          api.state.set('mfa', { needed: true, type, approvalNumber: null } satisfies MfaState);
          return new Promise<string>((resolve, reject) => {
            mfaCodeResolve = resolve;
            mfaCodeReject = reject;
            const thisReject = reject;
            setTimeout(() => {
              if (mfaCodeReject === thisReject) {
                mfaCodeReject(new Error('MFA code entry timed out'));
                mfaCodeResolve = null; mfaCodeReject = null;
              }
            }, 120_000);
          });
        },
        onMfaApprovalNeeded: (approvalNumber) => {
          if (live()) api.state.set('mfa', { needed: true, type: 'push', approvalNumber: approvalNumber ?? null } satisfies MfaState);
        },
        onMfaApprovalComplete: () => {
          if (live()) api.state.set('mfa', { needed: false, type: null, approvalNumber: null } satisfies MfaState);
        },
        onFallback: (reason) => {
          getLogger().warn(`Auto-login fallback: ${reason}`);
          if (live()) api.state.set('auth.autoLoginStatus', `Fallback: ${reason}`);
        },
      };

      let autoLoginAttempted = false;
      const tryAutoLogin = (url: string) => {
        if (autoLoginAttempted || codeSettled || !live()) return;
        if (isRedirect(url)) return;
        if (!isMicrosoftLoginHost(url)) return;
        autoLoginAttempted = true;
        api.state.set('auth.autoLoginStatus', 'Auto-logging in…');
        performAutoLogin(autoLoginHelpers, credentials, callbacks)
          .then((r) => {
            if (live()) api.state.set('mfa', { needed: false, type: null, approvalNumber: null } satisfies MfaState);
            if (r.success) getLogger().info('Auto-login succeeded');
          })
          .catch((err) => {
            getLogger().error(`Auto-login error: ${err}`);
            if (live() && !codeSettled) helpers.show();
          });
      };
      helpers.onDidNavigate(tryAutoLogin);
      timers.push(setInterval(() => {
        try { tryAutoLogin(helpers.getURL()); } catch { /* ignore */ }
      }, 1000));
    },
  });

  // If the window closes/times out before we captured a code, reject so the
  // in-flight guard clears and the next Sign-in click can open a fresh window.
  authPromise.then(
    (r) => { if (!codeSettled) settleErr(new Error(r.error || 'Sign-in cancelled')); },
    (e) => { if (!codeSettled) settleErr(e instanceof Error ? e : new Error(String(e))); },
  );

  let code: string;
  try {
    code = await codePromise;
  } finally {
    if (session === tokenCache.currentSession()) {
      api.state.set('mfa', { needed: false, type: null, approvalNumber: null } satisfies MfaState);
      api.state.set('auth.autoLoginStatus', null);
    }
    // Ensure the window promise doesn't dangle unhandled.
    authPromise.catch(() => {});
  }

  // Exchange code → tokens (auth client), then immediately FOCI-refresh as Office
  // so the very first access token already carries Chat.ReadWrite.
  const initial = await tokenRequest(api, {
    client_id: AUTH_CLIENT_ID,
    grant_type: 'authorization_code',
    code,
    redirect_uri: AAD_NATIVE_REDIRECT,
    code_verifier: codeVerifier,
    scope: AUTH_SCOPES,
  });
  let token = buildTokenData(initial);
  if (!token.refreshToken) {
    throw new Error('No refresh_token returned; ensure offline_access scope was granted');
  }
  if (session !== tokenCache.currentSession()) {
    throw new Error('Sign-in cancelled');
  }
  tokenCache.set(token);
  tokenCache.persist(api);

  try {
    token = await acquireTokenSilent(api);
  } catch (err) {
    getLogger().warn(`Post-login FOCI redemption failed (${err}); continuing with auth-client token`);
  }
  if (session !== tokenCache.currentSession()) {
    throw new Error('Signed out');
  }

  getLogger().info(`Signed in as ${token.email} (expires ${new Date(token.expiresAt).toISOString()})`);
  return token;
}

// ── Public: ensure a valid Graph access token, refreshing automatically ──

export interface EnsureOptions {
  /** When false, never open an auth window; throw instead. Background paths must pass false. */
  allowInteractive?: boolean;
}

export async function ensureAccessToken(api: PluginAPI, opts: EnsureOptions = {}): Promise<string> {
  const at = tokenCache.getValidAccessToken();
  if (at) return at;
  const session = tokenCache.currentSession();
  if (tokenCache.hasRefreshToken()) {
    try {
      const t = await acquireTokenSilent(api);
      return t.accessToken;
    } catch (err) {
      getLogger().warn(`Silent refresh failed: ${err}`);
      if (opts.allowInteractive === false) {
        throw new Error(`Silent refresh failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  }
  if (session !== tokenCache.currentSession()) {
    throw new Error('Signed out');
  }
  if (opts.allowInteractive === false) {
    throw new Error('Not signed in. Please log in via the Teams panel.');
  }
  const t = await acquireTokenInteractive(api);
  return t.accessToken;
}

// ── Secondary FOCI tokens (per-client, in-memory only) ──

const fociTokens = new Map<string, { accessToken: string; expiresAt: number; session: number }>();

export function clearFociTokens(): void {
  fociTokens.clear();
}

/** Redeem the family refresh token as an arbitrary FOCI client (e.g. Outlook Mobile for Presence). */
export async function acquireFociAccessToken(api: PluginAPI, clientId: string): Promise<string> {
  const session = tokenCache.currentSession();
  const cached = fociTokens.get(clientId);
  if (cached && cached.session === session && cached.expiresAt > Date.now() + 60_000) {
    return cached.accessToken;
  }
  const rt = tokenCache.getRefreshToken();
  if (!rt) throw new Error('Not signed in');
  const tr = await tokenRequest(api, {
    client_id: clientId,
    grant_type: 'refresh_token',
    refresh_token: rt,
    scope: GRAPH_SCOPE,
  });
  if (session !== tokenCache.currentSession()) throw new Error('Signed out');
  fociTokens.set(clientId, {
    accessToken: tr.access_token,
    expiresAt: Date.now() + tr.expires_in * 1000,
    session,
  });
  return tr.access_token;
}

/** Force a refresh regardless of current expiry. Used on 401 retry. */
export async function forceRefresh(api: PluginAPI): Promise<string> {
  if (!tokenCache.hasRefreshToken()) {
    throw new Error('Not signed in. Please log in via the Teams panel.');
  }
  const t = await acquireTokenSilent(api);
  return t.accessToken;
}
