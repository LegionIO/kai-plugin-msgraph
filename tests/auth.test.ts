import assert from 'node:assert/strict';
import { beforeEach, afterEach, test } from 'node:test';
import * as auth from '../src/backend/auth.ts';
import * as cache from '../src/backend/token-cache.ts';
import { setLogger } from '../src/backend/logger-singleton.ts';
import { GraphClient } from '../src/backend/graph-client.ts';
import { AAD_NATIVE_REDIRECT } from '../src/shared/constants.ts';
import { withRequestContext } from '../src/backend/request-scheduler.ts';
import type { PluginAPI } from '../src/shared/types.ts';

const jwt = `e30.${Buffer.from(JSON.stringify({ oid: 'user', tid: 'tenant', upn: 'user@example.test', scp: 'Chat.ReadWrite' })).toString('base64url')}.test`;
const tokenResponse = () => Response.json({ access_token: jwt, refresh_token: 'new-refresh', expires_in: 3600, token_type: 'Bearer' });
const invalidGrant = () => Response.json({ error: 'invalid_grant', error_description: 'Refresh token revoked' }, { status: 400 });
const deferred = <T,>() => { let resolve!: (value: T) => void; const promise = new Promise<T>((r) => { resolve = r; }); return { promise, resolve }; };
async function flush() { for (let i = 0; i < 80; i++) await Promise.resolve(); }
function seed() {
  cache.set({ accessToken: jwt, refreshToken: 'expired-refresh', expiresAt: 0, objectId: 'user', tenantId: 'tenant', email: 'user@example.test', displayName: null, scopes: 'Chat.ReadWrite' });
}
function harness(credentials = false) {
  const state: any = { auth: { isAuthenticated: false, isSigningIn: false }, mfa: { needed: false } };
  const data: any = credentials ? { encryptedCredentials: { method: 'safeStorage', username: 'user@example.test', encryptedPassword: 'test-password' } } : {};
  const calls: URLSearchParams[] = [];
  const windows: any[] = [];
  let responder = async (_form: URLSearchParams, _url: string) => tokenResponse();
  let page = { state: 'loading' } as any;
  const api = {
    pluginDir: '/nonexistent-msgraph-test',
    state: { get: () => state, set: (path: string, value: unknown) => {
      const keys = path.split('.'); let obj = state;
      for (const key of keys.slice(0, -1)) obj = obj[key] ??= {};
      obj[keys.at(-1)!] = value;
    } },
    config: { getPluginData: () => data, setPluginData: (key: string, value: unknown) => { data[key] = value; } },
    safeStorage: { isEncryptionAvailable: () => true, encryptString: (s: string) => `encrypted:${s}`, decryptString: (s: string) => s },
    fetch: async (url: string, options?: any) => {
      const form = new URLSearchParams(options?.body); calls.push(form); return responder(form, url);
    },
    auth: { openAuthWindow: (options: any) => {
      const result = deferred<any>();
      const win = { options, result, shown: 0, closed: 0 };
      windows.push(win);
      options.onReady({
        getURL: async () => 'https://login.microsoftonline.com/organizations/login',
        executeJavaScript: async () => page,
        onDidNavigate: () => {},
        show: () => { win.shown++; },
        close: () => { win.closed++; result.resolve({ success: false, error: 'Closed by plugin' }); },
      });
      return result.promise;
    } },
  } as unknown as PluginAPI;
  const complete = (index = windows.length - 1, extra = {}) => {
    const win = windows[index];
    assert.equal(win.options.callbackMatch, AAD_NATIVE_REDIRECT);
    const state = new URL(win.options.url).searchParams.get('state');
    win.result.resolve({ success: true, params: { state, code: 'test-authorization-code', ...extra } });
  };
  return { api, state, data, calls, windows, complete,
    respond: (r: typeof responder) => { responder = r; }, page: (p: any) => { page = p; } };
}

beforeEach(() => { cache.invalidateSession(); cache.clear(); auth.clearFociTokens(); auth.configureAuth({}); setLogger({ info() {}, warn() {}, error() {} }); });
afterEach(async () => { cache.invalidateSession(); auth.cancelSignIn(); auth.configureAuth({}); await flush(); });

test('repeated sign-in clicks share one session, window, and completion', async () => {
  const h = harness(); let starts = 0; let successes = 0;
  auth.configureAuth({ onStart: () => { starts++; }, onSuccess: () => { successes++; } });
  const before = cache.currentSession();
  const first = auth.acquireTokenInteractive(h.api, { forceRefresh: true });
  assert.equal(h.state.auth.isSigningIn, true);
  const second = auth.acquireTokenInteractive(h.api, { forceRefresh: true });
  assert.equal(first, second);
  assert.equal(cache.currentSession(), before + 1);
  await flush(); assert.equal(h.windows.length, 1);
  // No navigation callbacks: the host's redirect result must be sufficient.
  h.complete();
  await Promise.all([first, second]);
  assert.equal(starts, 1); assert.equal(successes, 1);
  assert.equal(h.state.auth.isSigningIn, false);
  assert.equal(h.state.auth.error, null);
  assert.equal(cache.getRefreshToken(), 'new-refresh');
  assert.equal(h.calls.filter((c) => c.get('grant_type') === 'authorization_code').length, 1);
});

test('expired access tokens share one silent refresh without opening a window', async () => {
  seed(); const h = harness();
  const results = await Promise.all(Array.from({ length: 8 }, () => auth.ensureAccessToken(h.api, { allowInteractive: false })));
  assert.ok(results.every((t) => t === jwt)); assert.equal(h.calls.length, 1); assert.equal(h.windows.length, 0);
});

test('revoked refresh tokens recover automatically across concurrent background requests', async () => {
  seed(); const h = harness();
  h.respond(async (form) => form.get('refresh_token') === 'expired-refresh' ? invalidGrant() : tokenResponse());
  const requests = Array.from({ length: 8 }, () => auth.ensureAccessToken(h.api, { allowInteractive: false }));
  await flush(); assert.equal(h.windows.length, 1); assert.equal(h.state.auth.isAuthenticated, false);
  const clicked = auth.acquireTokenInteractive(h.api, { forceRefresh: true });
  h.complete();
  await Promise.all([...requests, clicked]);
  assert.equal(h.windows.length, 1); assert.equal(cache.getRefreshToken(), 'new-refresh');
});

test('Graph 401 with an invalid refresh token reauthenticates and retries the request', async () => {
  seed(); cache.set({ ...cache.get()!, expiresAt: Date.now() + 3600000 });
  const h = harness(); let graphCalls = 0;
  h.respond(async (form, url) => {
    if (url.includes('graph.microsoft.com')) return ++graphCalls === 1 ? new Response(null, { status: 401 }) : Response.json({ id: 'user' });
    return form.get('refresh_token') === 'expired-refresh' ? invalidGrant() : tokenResponse();
  });
  const request = new GraphClient(h.api, false).getMe();
  await flush(); assert.equal(h.windows.length, 1);
  h.complete(); assert.equal((await request).id, 'user'); assert.equal(graphCalls, 2);
});

test('a logical action can continue after same-account silent recovery without losing its retry budget', async () => {
  seed(); const h = harness();
  h.respond(async (form, url) => url.includes('graph.microsoft.com') ? Response.json({ id: 'user' }) :
    form.get('refresh_token') === 'expired-refresh' ? invalidGrant() : tokenResponse());
  const action = withRequestContext(h.api, async () => {
    await auth.ensureAccessToken(h.api, { allowInteractive: false });
    return new GraphClient(h.api, false).getMe();
  });
  await flush(); assert.equal(h.windows.length, 1); h.complete();
  assert.equal((await action).id, 'user');
});

test('failed automatic sign-in does not prompt on every poll; explicit retry works', async () => {
  seed(); const h = harness(); h.respond(async (form) => form.get('refresh_token') === 'expired-refresh' ? invalidGrant() : tokenResponse());
  const attempt = auth.ensureAccessToken(h.api, { allowInteractive: false });
  const rejected = assert.rejects(attempt, /Auth window closed/);
  await flush(); h.windows[0].result.resolve({ success: false, error: 'Auth window closed' }); await rejected;
  const count = h.calls.length;
  for (let i = 0; i < 3; i++) await assert.rejects(auth.ensureAccessToken(h.api, { allowInteractive: false }), /Sign in to reconnect/);
  assert.equal(h.windows.length, 1); assert.equal(h.calls.length, count);
  const retry = auth.acquireTokenInteractive(h.api, { forceRefresh: true }); await flush(); h.complete(); await retry;
  assert.equal(h.windows.length, 2); assert.equal(h.state.auth.error, null);
});

test('network failures do not launch interactive auth or prevent a later refresh', async () => {
  seed(); const h = harness(); h.respond(async () => { throw new Error('Network unavailable'); });
  await assert.rejects(auth.ensureAccessToken(h.api, { allowInteractive: false }), /Network unavailable/);
  assert.equal(h.windows.length, 0);
  h.respond(async () => tokenResponse()); assert.equal(await auth.ensureAccessToken(h.api, { allowInteractive: false }), jwt);
});

test('sign-out during silent refresh cannot restore tokens or open a window', async () => {
  seed(); const h = harness(); const response = deferred<Response>(); h.respond(() => response.promise);
  const request = auth.ensureAccessToken(h.api, { allowInteractive: false });
  const rejected = assert.rejects(request, /Signed out/);
  cache.invalidateSession(); cache.clear(); response.resolve(tokenResponse()); await rejected;
  assert.equal(cache.get(), null); assert.equal(h.windows.length, 0);
});

test('cancelling an old window allows a new login without stale errors or MFA state', async () => {
  const h = harness(); const old = auth.acquireTokenInteractive(h.api, { forceRefresh: true });
  const rejected = assert.rejects(old, /cancelled|Closed by plugin/); await flush();
  auth.cancelSignIn(); const next = auth.acquireTokenInteractive(h.api, { forceRefresh: true });
  await flush(); assert.equal(h.windows.length, 2); assert.equal(h.state.auth.isSigningIn, true);
  h.complete(); await Promise.all([rejected, next]);
  assert.equal(h.state.auth.error, null); assert.equal(cache.getRefreshToken(), 'new-refresh');
});

test('sign-in remains busy through token exchange and cancellation prevents late token writes', async () => {
  const h = harness(); const response = deferred<Response>();
  h.respond((form) => form.get('grant_type') === 'authorization_code' ? response.promise : Promise.resolve(tokenResponse()));
  const attempt = auth.acquireTokenInteractive(h.api, { forceRefresh: true }); const rejected = assert.rejects(attempt, /cancelled/);
  await flush(); h.complete(); await flush();
  assert.equal(h.state.auth.autoLoginStatus, 'Finishing sign-in…');
  assert.equal(auth.acquireTokenInteractive(h.api, { forceRefresh: true }), attempt);
  auth.cancelSignIn(); cache.clear(); response.resolve(tokenResponse()); await rejected;
  assert.equal(cache.get(), null); assert.equal(h.data.graphToken, undefined);
});

for (const [name, params, pattern] of [
  ['wrong OAuth state', { state: 'wrong' }, /state mismatch/],
  ['missing code', { code: '' }, /No authorization code/],
  ['OAuth error', { error: 'access_denied', error_description: 'User declined' }, /access_denied/],
] as const) test(`rejects ${name} without exchanging a token`, async () => {
  const h = harness(); const attempt = auth.acquireTokenInteractive(h.api, { forceRefresh: true });
  const rejected = assert.rejects(attempt, pattern); await flush(); h.complete(0, params); await rejected;
  assert.equal(h.calls.length, 0); assert.equal(h.state.auth.isSigningIn, false);
});

test('secondary token recovery is bounded even if the resource keeps rejecting it', async () => {
  seed(); const h = harness(); h.respond(async (form) => form.get('client_id') === 'secondary' ? invalidGrant() : tokenResponse());
  const request = auth.acquireFociAccessToken(h.api, 'secondary'); const rejected = assert.rejects(request, /invalid_grant/);
  await flush(); h.complete(); await rejected; assert.equal(h.windows.length, 1);
});

test('direct MFA state is published immediately and cleared on cancellation', async () => {
  const h = harness(true); h.page({ state: 'mfa_authenticator_push', approvalNumber: '42' });
  const attempt = auth.acquireTokenInteractive(h.api, { forceRefresh: true }); const rejected = assert.rejects(attempt, /cancelled|Closed by plugin/);
  await flush(); assert.deepEqual(h.state.mfa, { needed: true, type: 'push', approvalNumber: '42' });
  auth.cancelSignIn(); await rejected; assert.equal(h.windows[0].closed, 1);
});

test('MFA code submission clears the challenge and never publishes the entered code', async () => {
  const h = harness(true); h.page({ state: 'mfa_totp_code' });
  const attempt = auth.acquireTokenInteractive(h.api, { forceRefresh: true }); await flush();
  assert.equal(h.state.mfa.type, 'totp'); h.page({ state: 'not_login_page' }); auth.submitMfaCode('123456');
  await flush(); assert.equal(h.state.mfa.needed, false); assert.ok(!JSON.stringify(h.state).includes('123456'));
  h.complete(); await attempt;
});

test('manual fallback uses the same window and accepts successful headful completion', async (t) => {
  t.mock.timers.enable({ apis: ['setTimeout', 'setInterval', 'Date'] });
  const h = harness(true); const attempt = auth.acquireTokenInteractive(h.api, { forceRefresh: true }); await flush();
  for (let i = 0; i < 32; i++) { t.mock.timers.tick(500); await flush(); }
  assert.ok(h.windows[0].shown > 0); assert.match(h.state.auth.autoLoginStatus, /Complete sign-in/);
  h.complete(); await attempt;
  assert.equal(h.windows.length, 1); assert.equal(h.state.auth.error, null); assert.equal(h.state.auth.isSigningIn, false);
});

test('an old failed request cannot join a login started after explicit sign-out', async () => {
  seed(); const h = harness(); const response = deferred<Response>();
  h.respond((form) => form.get('refresh_token') === 'expired-refresh' ? response.promise : Promise.resolve(tokenResponse()));
  const old = auth.ensureAccessToken(h.api, { allowInteractive: false }); const rejected = assert.rejects(old, /Signed out/);
  cache.invalidateSession(); cache.clear();
  const next = auth.acquireTokenInteractive(h.api, { forceRefresh: true }); await flush();
  response.resolve(invalidGrant()); await rejected;
  assert.equal(h.state.auth.isSigningIn, true); h.complete(); await next;
  assert.equal(h.state.auth.error, null);
});

test('a signed-out background request never initiates login', async () => {
  const h = harness(); await assert.rejects(auth.ensureAccessToken(h.api, { allowInteractive: false }), /Not signed in/);
  assert.equal(h.windows.length, 0); assert.equal(h.calls.length, 0);
});

test('MFA entry timeout hands off to the same window without cancelling OAuth', async (t) => {
  t.mock.timers.enable({ apis: ['setTimeout', 'setInterval', 'Date'] });
  const h = harness(true); h.page({ state: 'mfa_totp_code' });
  const attempt = auth.acquireTokenInteractive(h.api, { forceRefresh: true }); await flush();
  assert.equal(h.state.mfa.needed, true);
  t.mock.timers.tick(120_000); await flush();
  assert.equal(h.state.mfa.needed, false); assert.equal(h.state.auth.isSigningIn, true);
  assert.ok(h.windows[0].shown > 0); assert.match(h.state.auth.autoLoginStatus, /Complete sign-in/);
  h.complete(); await attempt; assert.equal(h.state.auth.error, null);
});
