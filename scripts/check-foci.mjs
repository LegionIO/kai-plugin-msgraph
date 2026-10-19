// Read-only, silent validation against an existing MSAL family-token cache.
// Usage: node scripts/check-foci.mjs /path/to/token_cache.json
// Credentials stay in memory. Only capability results are written to stdout.
import { build } from 'esbuild';
import { readFile, mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { pathToFileURL } from 'node:url';

const path = process.argv[2];
if (!path) throw new Error('Provide the path to an existing MSAL token cache.');
const cache = JSON.parse(await readFile(path, 'utf8'));
const accounts = Object.values(cache.Account ?? {});
if (accounts.length !== 1) throw new Error('Validation requires a cache containing exactly one account.');
const account = accounts[0];
const refresh = Object.values(cache.RefreshToken ?? {}).find(r => r.family_id && r.home_account_id === account.home_account_id);
if (!refresh) throw new Error('No matching family refresh token is available.');
const directory = await mkdtemp(join(tmpdir(), 'msgraph-capabilities-'));
try {
  const outfile = join(directory, 'probe.mjs');
  await build({ stdin: { contents: `
    export { acquireFociToken, TokenEndpointError, describeToken } from './src/backend/auth.ts';
    export { FOCI_CLIENTS, graphPolicy } from './src/backend/request-policy.ts';
    export { eligible } from './src/backend/request-transport.ts';
    export { serviceFetch, disposeScheduler, withNewRequestContext } from './src/backend/request-scheduler.ts';
    export * as cache from './src/backend/token-cache.ts';
    export { setLogger } from './src/backend/logger-singleton.ts';
  `, resolveDir: resolve('.') }, bundle: true, platform: 'node', format: 'esm', outfile });
  const m = await import(pathToFileURL(outfile).href);
  m.setLogger({ info() {}, warn() {}, error() {} });
  m.cache.set({ accessToken: '', refreshToken: refresh.secret, expiresAt: 0,
    objectId: account.local_account_id, tenantId: account.realm, email: '', displayName: null, scopes: '' });
  const api = { fetch: globalThis.fetch, config: { getPluginData: () => ({ preferences: { maxThrottleWaitSeconds: 30 } }) } };
  const results = [];
  let stopped = false;
  for (const client of m.FOCI_CLIENTS) {
    const row = { client: client.alias, token: 'unavailable', capabilities: {} };
    results.push(row);
    if (stopped) { row.token = 'skipped-after-auth-or-throttle-failure'; continue; }
    try {
      const token = await m.acquireFociToken(api, client.id, undefined, { candidate: true });
      row.token = token.verified ? 'verified' : 'unverifiable';
      const operations = [
        ['mail-read', 'GET', '/me/messages?$top=1&$select=id'],
        ['directory-read', 'GET', '/users?$top=1&$select=id'],
        ['people-read', 'GET', '/me/people?$top=1&$select=id'],
        ['chat-read', 'GET', '/me/chats?$top=1&$select=id'],
        ['presence-read', 'POST', '/communications/getPresencesByUserId'],
        ['mail-send', 'POST', '/me/sendMail'],
      ];
      for (const [name, method, path] of operations) {
        const url = `https://graph.microsoft.com/v1.0${path}`;
        const policy = m.graphPolicy(method, url);
        if (!m.eligible(token, policy)) { row.capabilities[name] = 'scope-not-granted'; continue; }
        if (name === 'mail-send') { row.capabilities[name] = 'token-qualified; write-not-probed'; continue; }
        const response = await m.serviceFetch(api, 'live-validation', client.id, url, {
          method, headers: { Authorization: `Bearer ${token.accessToken}`, 'Content-Type': 'application/json' },
          ...(name === 'presence-read' ? { body: JSON.stringify({ ids: [account.local_account_id] }) } : {}),
        }, true);
        row.capabilities[name] = response.ok ? 'endpoint-confirmed' : `http-${response.status}`;
        await response.body?.cancel();
      }
    } catch (error) {
      row.token = error instanceof m.TokenEndpointError ? error.code : `validation-failed-${error.statusCode ?? 'network'}`;
      // Do not multiply failures from a revoked session or a shared service throttle.
      stopped = !(error instanceof m.TokenEndpointError) || ['invalid_grant','interaction_required','login_required'].includes(error.code);
    }
  }
  m.disposeScheduler(api); m.cache.clear();
  process.stdout.write(JSON.stringify({ checkedAt: new Date().toISOString(), source: 'existing-local-msal-cache', results }, null, 2) + '\n');
} finally { await rm(directory, { recursive: true, force: true }); }
