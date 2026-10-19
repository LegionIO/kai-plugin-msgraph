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

// One owner for the window, MFA callbacks, and token exchange. Repeated callers
// join it before any session invalidation or cache reset can occur.
interface SignInAttempt {
  session: number;
  recoveringSession?: number;
  controller: AbortController;
  promise: Promise<GraphTokenData>;
  helpers?: AuthWindowHelpers;
  rejectMfa?: (error: Error) => void;
  resolveMfa?: (code: string) => void;
}
let inFlight: SignInAttempt | null = null;
let silentFlight: { session: number; promise: Promise<GraphTokenData> } | null = null;
let automaticAttemptSession: number | null = null;
let lifecycle: {
  onStart?: () => void;
  onSuccess?: () => void;
} = {};

export function configureAuth(callbacks: typeof lifecycle): void {
  lifecycle = callbacks;
}

export function isSigningIn(): boolean { return inFlight !== null; }

export function showSignIn(): void {
  try { inFlight?.helpers?.show(); } catch { /* already closed */ }
}

export function submitMfaCode(code: string): void {
  const attempt = inFlight;
  if (!attempt?.resolveMfa || !code.trim()) return;
  attempt.resolveMfa(code.trim());
}

export function cancelSignIn(): void {
  const attempt = inFlight;
  if (!attempt) return;
  tokenCache.invalidateSession();
  automaticAttemptSession = tokenCache.currentSession();
  const token = tokenCache.get();
  if (token) tokenCache.set({ ...token, accessToken: '', expiresAt: 0 });
  attempt.controller.abort(new Error('Sign-in cancelled'));
  attempt.rejectMfa?.(new Error('Sign-in cancelled'));
  try { attempt.helpers?.close(); } catch { /* already closed */ }
  inFlight = null;
}

export const cancelMfaCode = cancelSignIn;

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

class TokenEndpointError extends Error {
  constructor(public readonly code: string, status: number, description: string) {
    super(`Token endpoint ${status}: ${code} ${description}`.trim());
  }
}

function needsReauthentication(error: unknown): boolean {
  return error instanceof TokenEndpointError &&
    ['invalid_grant', 'interaction_required', 'login_required', 'consent_required'].includes(error.code);
}

async function tokenRequest(api: PluginAPI, form: Record<string, string>): Promise<TokenResponse> {
  const resp = await api.fetch(AAD_TOKEN_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams(form).toString(),
  });
  const body = (await resp.json()) as TokenResponse;
  if (!resp.ok || body.error) {
    throw new TokenEndpointError(body.error ?? '', resp.status, body.error_description ?? '');
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

export function acquireTokenSilent(api: PluginAPI): Promise<GraphTokenData> {
  const session = tokenCache.currentSession();
  if (silentFlight?.session === session) return silentFlight.promise;
  const promise = redeemGraphToken(api).finally(() => {
    if (silentFlight?.promise === promise) silentFlight = null;
  });
  silentFlight = { session, promise };
  return promise;
}

async function redeemGraphToken(api: PluginAPI): Promise<GraphTokenData> {
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
    if (session !== tokenCache.currentSession()) throw new Error('Signed out');
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

export function acquireTokenInteractive(
  api: PluginAPI,
  opts: AcquireTokenOptions = {},
): Promise<GraphTokenData> {
  if (inFlight) return inFlight.promise;
  if (!opts.forceRefresh && tokenCache.getValidAccessToken()) return Promise.resolve(tokenCache.get()!);

  tokenCache.invalidateSession();
  const previous = tokenCache.get();
  if (previous) tokenCache.set({ ...previous, accessToken: '', expiresAt: 0 });
  clearFociTokens();
  const attempt: SignInAttempt = {
    session: tokenCache.currentSession(),
    controller: new AbortController(),
    promise: undefined!,
  };
  inFlight = attempt;
  automaticAttemptSession = attempt.session;
  api.state.set('auth.isSigningIn', true);
  api.state.set('auth.autoLoginStatus', 'Opening Microsoft sign-in…');
  api.state.set('auth.error', null);
  api.state.set('error', null);
  attempt.promise = Promise.resolve().then(async () => {
    if (inFlight !== attempt || attempt.controller.signal.aborted) throw new Error('Sign-in cancelled');
    lifecycle.onStart?.();
    const token = await doInteractive(api, attempt);
    if (inFlight !== attempt || attempt.session !== tokenCache.currentSession()) throw new Error('Sign-in cancelled');
    automaticAttemptSession = null;
    return token;
  }).then((token) => {
    if (inFlight === attempt) {
      inFlight = null;
      api.state.set('auth.isSigningIn', false);
      api.state.set('auth.autoLoginStatus', null);
      try { lifecycle.onSuccess?.(); } catch (error) { getLogger().error(`Unable to resume after sign-in: ${error}`); }
    }
    return token;
  }, (error) => {
    if (inFlight === attempt) {
      api.state.set('auth.error', error instanceof Error ? error.message : String(error));
    }
    throw error;
  }).finally(() => {
    attempt.controller.abort(new Error('Sign-in finished'));
    attempt.rejectMfa?.(new Error('Sign-in finished'));
    if (inFlight === attempt) inFlight = null;
    if (attempt.session === tokenCache.currentSession() && !inFlight) {
      api.state.set('mfa', { needed: false, type: null, approvalNumber: null } satisfies MfaState);
      api.state.set('auth.isSigningIn', false);
      api.state.set('auth.autoLoginStatus', null);
    }
  });
  return attempt.promise;
}

async function doInteractive(api: PluginAPI, attempt: SignInAttempt): Promise<GraphTokenData> {
  const { session, controller } = attempt;
  const signal = controller.signal;
  const live = () => inFlight === attempt && session === tokenCache.currentSession() && !signal.aborted;
  const assertLive = () => { if (!live()) throw new Error('Sign-in cancelled'); };
  assertLive();
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

  let poll: ReturnType<typeof setInterval> | undefined;
  let windowFinished = false;
  const authPromise = api.auth.openAuthWindow({
    url: authUrl,
    title: 'Microsoft Teams — Sign In',
    width: 520,
    height: 640,
    timeoutMs: AUTH_TIMEOUT_MS,
    showOnCreate: !shouldAutoLogin,
    // Reveal the same window early enough to complete an unfamiliar challenge.
    showAfterMs: shouldAutoLogin ? 30_000 : undefined,
    partition: AUTH_PARTITION,
    customUserAgent: false,
    callbackMatch: AAD_NATIVE_REDIRECT,
    extractParams: ['code', 'state', 'error', 'error_description'],
    onReady: (helpers: AuthWindowHelpers) => {
      if (!live() || windowFinished) { helpers.close(); return; }
      attempt.helpers = helpers;
      if (!shouldAutoLogin) return;
      const credentials = getCredentials(api);
      if (!credentials) { helpers.show(); return; }
      const active = () => live() && !windowFinished;
      const autoLoginHelpers: AutoLoginHelpers = {
        executeJavaScript: (code) => {
          if (!active()) return Promise.reject(new Error('Sign-in finished'));
          return helpers.executeJavaScript(code);
        },
        getURL: helpers.getURL,
        show: () => { if (active()) helpers.show(); },
      };
      const clearMfa = () => {
        if (active()) api.state.set('mfa', { needed: false, type: null, approvalNumber: null } satisfies MfaState);
      };
      const callbacks: AutoLoginCallbacks = {
        onMfaCodeNeeded: (type) => {
          if (!active()) return Promise.reject(new Error('Sign-in finished'));
          api.state.set('auth.autoLoginStatus', 'Enter your verification code to continue.');
          api.state.set('mfa', { needed: true, type, approvalNumber: null } satisfies MfaState);
          return new Promise<string>((resolve, reject) => {
            const cleanup = () => {
              clearTimeout(timer);
              attempt.resolveMfa = undefined;
              attempt.rejectMfa = undefined;
              clearMfa();
            };
            const timer = setTimeout(() => attempt.rejectMfa?.(new Error('MFA code entry timed out')), 120_000);
            attempt.resolveMfa = (code) => {
              cleanup();
              if (active()) api.state.set('auth.autoLoginStatus', 'Verifying code…');
              resolve(code);
            };
            attempt.rejectMfa = (error) => { cleanup(); reject(error); };
          });
        },
        onMfaApprovalNeeded: (approvalNumber) => {
          if (!active()) return;
          api.state.set('auth.autoLoginStatus', 'Approve the sign-in in Microsoft Authenticator.');
          api.state.set('mfa', { needed: true, type: 'push', approvalNumber: approvalNumber ?? null } satisfies MfaState);
        },
        onMfaApprovalComplete: clearMfa,
        onFallback: (reason) => {
          if (!active()) return;
          getLogger().warn(`Auto-login fallback: ${reason}`);
          clearMfa();
          api.state.set('auth.autoLoginStatus', 'Complete sign-in in the Microsoft window.');
        },
      };
      let autoLoginAttempted = false;
      const tryAutoLogin = async (url?: string) => {
        if (autoLoginAttempted || !active()) return;
        try {
          const currentUrl = url ?? await helpers.getURL();
          if (autoLoginAttempted || !active() || !isMicrosoftLoginHost(currentUrl) || currentUrl.startsWith(AAD_NATIVE_REDIRECT)) return;
          autoLoginAttempted = true;
          api.state.set('auth.autoLoginStatus', 'Signing in to Microsoft…');
          await performAutoLogin(autoLoginHelpers, credentials, callbacks, signal);
        } catch (error) {
          if (active()) {
            callbacks.onFallback(`Auto-login failed: ${error}`);
            helpers.show();
          }
        }
      };
      helpers.onDidNavigate((url) => { void tryAutoLogin(url); });
      poll = setInterval(() => { void tryAutoLogin(); }, 1000);
      void tryAutoLogin();
    },
  });

  let abort!: () => void;
  const cancelled = new Promise<never>((_, reject) => {
    abort = () => reject(signal.reason);
    signal.addEventListener('abort', abort, { once: true });
    if (signal.aborted) abort();
  });
  let result: Awaited<typeof authPromise>;
  try {
    result = await Promise.race([authPromise, cancelled]);
  } finally {
    windowFinished = true;
    if (poll) clearInterval(poll);
    signal.removeEventListener('abort', abort);
    attempt.rejectMfa?.(new Error('Sign-in window closed'));
    if (live()) api.state.set('mfa', { needed: false, type: null, approvalNumber: null } satisfies MfaState);
  }
  assertLive();
  // The host captures will-redirect before the renderer navigates (including
  // after a hidden window is revealed). Its successful result is authoritative.
  const params = result.params ?? {};
  if (!result.success) throw new Error(result.error || 'Sign-in cancelled');
  if (params.state !== state) throw new Error('OAuth state mismatch');
  if (params.error) throw new Error(`${params.error}: ${params.error_description ?? ''}`);
  if (!params.code) throw new Error('No authorization code in redirect');
  api.state.set('auth.autoLoginStatus', 'Finishing sign-in…');

  const initial = await tokenRequest(api, {
    client_id: AUTH_CLIENT_ID,
    grant_type: 'authorization_code',
    code: params.code,
    redirect_uri: AAD_NATIVE_REDIRECT,
    code_verifier: codeVerifier,
    scope: AUTH_SCOPES,
  });
  assertLive();
  let token = buildTokenData(initial);
  if (!token.refreshToken) throw new Error('No refresh_token returned; ensure offline_access scope was granted');
  tokenCache.set(token);
  tokenCache.persist(api);
  try {
    token = await acquireTokenSilent(api);
  } catch (err) {
    getLogger().warn(`Post-login FOCI redemption failed (${err}); continuing with auth-client token`);
  }
  assertLive();
  getLogger().info(`Signed in as ${token.email}`);
  return token;
}

// Background requests may recover an existing rejected session once. A failed
// or cancelled attempt requires an explicit click; polling must not prompt again.
async function recoverAuthentication(api: PluginAPI, error: unknown, session: number): Promise<GraphTokenData> {
  if (session !== tokenCache.currentSession()) {
    // Only join recovery of this session. An old request must not carry on
    // under a different account after an explicit sign-out and sign-in.
    if (inFlight?.recoveringSession === session) return inFlight.promise;
    throw new Error('Signed out');
  }
  if (!needsReauthentication(error) || !tokenCache.hasRefreshToken()) throw error;
  if (inFlight) return inFlight.promise;
  api.state.set('auth.isAuthenticated', false);
  if (automaticAttemptSession === session) throw new Error('Session expired. Sign in to reconnect.');
  const promise = acquireTokenInteractive(api, { forceRefresh: true });
  inFlight!.recoveringSession = session;
  return promise;
}

export interface EnsureOptions {
  /** Allow a first login. Existing rejected sessions still get one automatic recovery attempt. */
  allowInteractive?: boolean;
}

export async function ensureAccessToken(api: PluginAPI, opts: EnsureOptions = {}): Promise<string> {
  if (inFlight) return (await inFlight.promise).accessToken;
  const at = tokenCache.getValidAccessToken();
  if (at) return at;
  if (automaticAttemptSession === tokenCache.currentSession()) throw new Error('Session expired. Sign in to reconnect.');
  const session = tokenCache.currentSession();
  if (tokenCache.hasRefreshToken()) {
    try {
      return (await acquireTokenSilent(api)).accessToken;
    } catch (error) {
      return (await recoverAuthentication(api, error, session)).accessToken;
    }
  }
  if (opts.allowInteractive === false) throw new Error('Not signed in. Please log in via the Teams panel.');
  return (await acquireTokenInteractive(api)).accessToken;
}

// ── Secondary FOCI tokens (per client+scope, in-memory only) ──

const fociTokens = new Map<string, { accessToken: string; expiresAt: number; session: number }>();

export function clearFociTokens(): void {
  fociTokens.clear();
}

/**
 * Redeem the family refresh token as an arbitrary FOCI client for an arbitrary
 * resource scope (e.g. Outlook Mobile for Graph Presence, or Teams for the IC3
 * chat service). Cached per (clientId, scope) until near expiry.
 */
export async function acquireFociAccessToken(
  api: PluginAPI,
  clientId: string,
  scope: string = GRAPH_SCOPE,
  retried = false,
): Promise<string> {
  if (inFlight) await inFlight.promise;
  if (automaticAttemptSession === tokenCache.currentSession()) throw new Error('Session expired. Sign in to reconnect.');
  const session = tokenCache.currentSession();
  const key = `${clientId}|${scope}`;
  const cached = fociTokens.get(key);
  if (cached && cached.session === session && cached.expiresAt > Date.now() + 60_000) {
    return cached.accessToken;
  }
  const rt = tokenCache.getRefreshToken();
  if (!rt) throw new Error('Not signed in');
  let tr: TokenResponse;
  try {
    tr = await tokenRequest(api, {
      client_id: clientId,
      grant_type: 'refresh_token',
      refresh_token: rt,
      scope,
    });
  } catch (error) {
    if (retried) throw error;
    await recoverAuthentication(api, error, session);
    return acquireFociAccessToken(api, clientId, scope, true);
  }
  if (session !== tokenCache.currentSession()) throw new Error('Signed out');
  fociTokens.set(key, {
    accessToken: tr.access_token,
    expiresAt: Date.now() + tr.expires_in * 1000,
    session,
  });
  return tr.access_token;
}

/** Force a refresh regardless of current expiry. Used on 401 retry. */
export async function forceRefresh(api: PluginAPI): Promise<string> {
  if (inFlight) return (await inFlight.promise).accessToken;
  if (automaticAttemptSession === tokenCache.currentSession()) throw new Error('Session expired. Sign in to reconnect.');
  const session = tokenCache.currentSession();
  try {
    return (await acquireTokenSilent(api)).accessToken;
  } catch (error) {
    return (await recoverAuthentication(api, error, session)).accessToken;
  }
}
