import type { PluginAPI } from '../shared/types.js';
import { GraphApiError } from '../shared/types.js';
import { CLIENT_ID_OFFICE, GRAPH_SCOPE } from '../shared/constants.js';
import { acquireFociToken, describeToken, ensureAccessToken, forceRefresh, type FociToken } from './auth.js';
import * as tokens from './token-cache.js';
import { FOCI_CLIENTS, graphPolicy, type RequestPolicy } from './request-policy.js';
import { currentRequestContext, getScheduler, serviceFetch, withRequestContext } from './request-scheduler.js';

let denied = new WeakMap<PluginAPI, Set<string>>();
let reports = new WeakMap<PluginAPI, Map<string, { client: string; workload: string; eligible: boolean; reason: string }>>();
tokens.onSessionInvalidated(() => { denied = new WeakMap(); reports = new WeakMap(); });
export function capabilityReport(api: PluginAPI): unknown[] { return [...(reports.get(api)?.values() ?? [])]; }
export function eligible(token: FociToken, policy: RequestPolicy): boolean {
  return token.verified && token.session === tokens.currentSession() && token.expiresAt > Date.now() &&
    policy.scopes.some(set => set.every(scope => token.scopes.includes(scope)));
}
function denialKey(p: RequestPolicy, client: string): string { return `${tokens.currentSession()}|${client}|${p.workload}|${p.operation}|${p.resource}`; }
function report(api: PluginAPI, client: string, p: RequestPolicy, ok: boolean, reason: string): void {
  let map = reports.get(api); if (!map) { map = new Map(); reports.set(api, map); }
  map.set(`${client}|${p.workload}|${p.operation}`, { client, workload: p.workload, eligible: ok, reason });
}

export async function graphFetch(api: PluginAPI, input: string | URL | Request, init: RequestInit = {}, observed?: { response: Response; attempts: number }): Promise<Response> {
  const url = typeof input === 'string' ? input : input instanceof URL ? input.href : input.url;
  const target = new URL(url);
  if (target.origin !== 'https://graph.microsoft.com') {
    // Public CDN and preauthorized upload/download URLs never receive a Graph token.
    const headers = new Headers(init.headers);
    if (headers.has('authorization')) throw new GraphApiError('Refusing Graph credentials for a non-Graph destination', 0);
    return serviceFetch(api, 'files', 'upload', input, init);
  }
  return withRequestContext(api, async () => {
    const context = currentRequestContext()!;
    const headers = new Headers(init.headers);
    const initial = headers.get('authorization')?.replace(/^Bearer /i, '') ?? await ensureAccessToken(api);
    const descriptor = describeToken(initial, tokens.get()?.clientId ?? CLIENT_ID_OFFICE, GRAPH_SCOPE, tokens.get()?.expiresAt ?? Date.now() + 60_000);
    const preferred = descriptor.clientId;
    const handles = new Map<string, FociToken>([[preferred, descriptor]]);
    let body: unknown;
    if (typeof init.body === 'string') { try { body = JSON.parse(init.body); } catch { /* bytes */ } }
    const policy = graphPolicy(init.method ?? 'GET', url, body);
    const scheduler = getScheduler(api);
    let deny = denied.get(api); if (!deny) { deny = new Set(); denied.set(api, deny); }
    const choose = async (current: string, deadline: number): Promise<string | undefined> => {
      const choices: Array<{ id: string; at: number }> = [];
      for (const candidate of FOCI_CLIENTS) {
        if (candidate.id === current || !(candidate.workloads as readonly string[]).includes(policy.workload) || deny!.has(denialKey(policy, candidate.id))) continue;
        if (scheduler.admissionAt(policy, candidate.id) >= deadline) continue;
        try {
          const token = await acquireFociToken(api, candidate.id, GRAPH_SCOPE, { candidate: true });
          const ok = eligible(token, policy);
          report(api, candidate.alias, policy, ok, ok ? 'token-qualified' : 'scope-or-identity-mismatch');
          if (!ok) continue;
          handles.set(candidate.id, token);
          choices.push({ id: candidate.id, at: scheduler.admissionAt(policy, candidate.id) });
          // An immediately available client is optimal; registry order breaks ties.
          if (choices.at(-1)!.at <= Date.now()) break;
        } catch (error) {
          if (context.session !== tokens.currentSession() || context.controller.signal.aborted) throw error;
          report(api, candidate.alias, policy, false, 'token-unavailable');
          // Token throttles share one identity lane. Do not fan out to more clients.
          if (error instanceof GraphApiError) break;
        }
      }
      return choices.filter(c => c.at < deadline).sort((a,b) => a.at - b.at)[0]?.id;
    };

    // Replace blind channel 403 fallback: only a known missing scope can select another client.
    let selected = preferred;
    if (policy.workload === 'teams' && policy.scopes.length && descriptor.verified && !eligible(descriptor, policy)) {
      selected = await choose(preferred, Infinity) ?? preferred;
    }
    return scheduler.execute(policy, selected, context, async (client, signal) => {
      let token = handles.get(client)!;
      if (token.expiresAt <= Date.now() + 30_000) {
        token = await acquireFociToken(api, client, GRAPH_SCOPE, { candidate: client !== preferred });
        handles.set(client, token);
      }
      if (client !== preferred && (!eligible(token, policy) || deny!.has(denialKey(policy, client)))) throw new GraphApiError('Alternate client no longer qualifies for this operation', 403);
      const h = new Headers(headers); h.set('authorization', `Bearer ${token.accessToken}`);
      const response = await api.fetch(input, { ...init, headers: h, signal });
      if (response.ok && token.verified) report(api, client, policy, true, 'endpoint-confirmed');
      return response;
    }, {
      alternate: choose,
      denied: client => { deny!.add(denialKey(policy, client)); report(api, client, policy, false, 'endpoint-denied'); },
      refresh: async client => {
        const account = context.account;
        let refreshed: FociToken;
        if (client === preferred && initial === tokens.get()?.accessToken) {
          const access = await forceRefresh(api);
          refreshed = describeToken(access, client, GRAPH_SCOPE, tokens.get()?.expiresAt);
          // Authentication recovery may replace the session; never cross accounts.
          if (account !== `${tokens.getTenantId() ?? ''}|${tokens.getObjectId() ?? ''}`) throw new GraphApiError('Account changed during authentication', 0);
          context.session = tokens.currentSession(); context.controller = new AbortController();
        } else refreshed = await acquireFociToken(api, client, GRAPH_SCOPE, { force: true, candidate: client !== preferred });
        if (refreshed.clientId !== client) throw new GraphApiError('Token refresh changed the client identity', 401);
        handles.set(client, refreshed);
      },
    }, observed?.response, observed?.attempts ?? 1);
  });
}

/** Execute one batch item through the same budgets without replaying successful peers. */
export async function executeBatchItem(api: PluginAPI, url: string, method: string, body: unknown,
  initial: Response, initialAttempts = 1): Promise<Response> {
  // Results that already completed need no admission, token exchange, or network call.
  if (initial.status !== 429 && ![500,502,503,504].includes(initial.status)) return initial;
  return graphFetch(api, `https://graph.microsoft.com/v1.0${url}`, {
    method, headers: { Authorization: `Bearer ${await ensureAccessToken(api, { allowInteractive: false })}`, 'Content-Type': 'application/json' },
    ...(body === undefined ? {} : { body: JSON.stringify(body) }),
  }, { response: initial, attempts: initialAttempts });
}
