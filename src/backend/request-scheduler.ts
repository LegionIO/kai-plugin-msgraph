import { AsyncLocalStorage } from 'node:async_hooks';
import { GraphApiError, type PluginAPI } from '../shared/types.js';
import * as tokens from './token-cache.js';
import type { RequestPolicy } from './request-policy.js';
import { clientAlias } from './request-policy.js';
import { getLogger } from './logger-singleton.js';

export type ThrottleScope = 'application' | 'app-tenant' | 'tenant' | 'user' | 'resource' | 'service' | 'unknown';
export interface RequestContext {
  session: number; account: string; priority: 'interactive' | 'background'; remainingMs: number;
  adaptive: boolean; controller: AbortController;
}
type Scope = { kind: ThrottleScope; access: 'Read' | 'Write' | 'ReadWrite'; information?: string; resourceUnits?: number; requestId?: string };
type Cooldown = Scope & { until: number; client: string; policy: RequestPolicy };
export interface Routes {
  alternate: (current: string, deadline: number) => Promise<string | undefined>;
  refresh: (current: string) => Promise<void>;
  denied: (current: string) => void;
}
const local = new AsyncLocalStorage<RequestContext>();
const schedulers = new WeakMap<PluginAPI, RequestScheduler>();
export const responseAttempts = new WeakMap<Response, number>();
export const currentRequestContext = (): RequestContext | undefined => local.getStore();
export function throttleWaitSeconds(value: unknown): number {
  return typeof value === 'number' && Number.isFinite(value) ? Math.max(0, Math.min(300, Math.floor(value))) : 30;
}
export function newRequestContext(api: PluginAPI, priority: RequestContext['priority'] = 'interactive'): RequestContext {
  const prefs = api.config.getPluginData().preferences as Record<string, unknown> | undefined;
  return { session: tokens.currentSession(), account: `${tokens.getTenantId() ?? ''}|${tokens.getObjectId() ?? ''}`,
    priority, remainingMs: throttleWaitSeconds(prefs?.maxThrottleWaitSeconds) * 1000,
    adaptive: prefs?.adaptiveFociRouting !== false, controller: new AbortController() };
}
export function withRequestContext<T>(api: PluginAPI, fn: () => Promise<T>, priority: RequestContext['priority'] = 'interactive'): Promise<T> {
  return local.getStore() ? fn() : local.run(newRequestContext(api, priority), fn);
}
/** Only for a new auth attempt or a new background tick, never for replaying an old write. */
export function withNewRequestContext<T>(api: PluginAPI, fn: () => Promise<T>, priority: RequestContext['priority'] = 'interactive'): Promise<T> {
  return local.run(newRequestContext(api, priority), fn);
}
export function getScheduler(api: PluginAPI): RequestScheduler {
  let scheduler = schedulers.get(api);
  if (!scheduler) { scheduler = new RequestScheduler(); schedulers.set(api, scheduler); }
  return scheduler;
}
export function disposeScheduler(api: PluginAPI): void { schedulers.get(api)?.dispose(); schedulers.delete(api); }

/** Server delays are lower bounds. An absolute date uses the server's clock. */
export function retryAfterMs(headers: Headers, now = Date.now()): number | undefined {
  const delays: number[] = [];
  const add = (n: number) => { if (Number.isFinite(n) && n >= 0) delays.push(n); };
  for (const key of ['x-ms-retry-after-ms', 'retry-after-ms']) {
    const value = headers.get(key)?.trim();
    if (value && /^\d+(?:\.\d+)?$/.test(value)) add(Number(value));
  }
  const value = headers.get('retry-after')?.trim();
  if (value && /^\d+(?:\.\d+)?$/.test(value)) add(Number(value) * 1000);
  else if (value && /[a-z]/i.test(value)) {
    const time = Date.parse(value);
    const date = Date.parse(headers.get('date') ?? '');
    if (Number.isFinite(time)) add(Math.max(0, time - (Number.isFinite(date) ? date : now)));
  }
  return delays.length ? Math.max(...delays) : undefined;
}

export function throttleScope(headers: Headers, client: string, tenant: string): Scope {
  const units = Number(headers.get('x-ms-resource-unit'));
  const details = { information: headers.get('x-ms-throttle-information')?.slice(0, 120),
    resourceUnits: Number.isFinite(units) && units > 0 ? units : undefined,
    requestId: headers.get('request-id') ?? undefined };
  const [scope, limit, app, target] = (headers.get('x-ms-throttle-scope') ?? '').split('/');
  const kinds: Record<string, ThrottleScope> = { Application: 'application', Tenant_Application: 'app-tenant', Tenant: 'tenant', User: 'user', Resource: 'resource' };
  // A header for another application/tenant is not evidence of an independent budget.
  if (kinds[scope] && ['Read','Write','ReadWrite'].includes(limit) && app === client && target &&
      (!['Tenant','Tenant_Application','Application'].includes(scope) || target === tenant)) {
    return { ...details, kind: kinds[scope], access: limit as Scope['access'] };
  }
  return { ...details, kind: 'unknown', access: 'ReadWrite' };
}
export function preventsFallback(error: unknown): boolean {
  return error instanceof GraphApiError && (error.statusCode === 429 || error.statusCode >= 500 || error.statusCode === 0 || !!error.retry?.ambiguous);
}

export class RequestScheduler {
  private cooldowns: Cooldown[] = [];
  private active = new Map<string, { total: number; background: number }>();
  private pace = new Map<string, number>();
  private listeners = new Set<() => void>();
  private controllers = new Map<RequestContext, number>();
  private probes = new Map<string, { next: number; busy: boolean }>();
  private hints = new Map<string, { client: string; until: number }>();
  private closed = false;
  private unsubscribe = tokens.onSessionInvalidated(() => this.reset());
  constructor(private readonly now = () => Date.now(), private readonly random = () => Math.random()) {}
  private wake(): void { for (const fn of [...this.listeners]) fn(); }
  private reset(): void {
    for (const c of this.controllers.keys()) c.controller.abort(new Error('Session ended'));
    this.cooldowns = []; this.pace.clear(); this.probes.clear(); this.hints.clear(); this.wake();
  }
  dispose(): void { this.closed = true; this.reset(); this.unsubscribe(); }
  private check(ctx: RequestContext): void {
    if (this.closed || ctx.session !== tokens.currentSession() || ctx.controller.signal.aborted) throw new GraphApiError('Request cancelled: session ended', 0);
  }
  private lane(p: RequestPolicy): string { return `${p.workload}|${p.operation}|${p.resource}`; }
  private paceKeys(p: RequestPolicy): string[] {
    return p.workload === 'teams' && !p.readOnly ? [this.lane(p), 'teams|writes-user'] : [this.lane(p)];
  }
  private retain(ctx: RequestContext): void { this.controllers.set(ctx, (this.controllers.get(ctx) ?? 0) + 1); }
  private releaseContext(ctx: RequestContext): void {
    const count = (this.controllers.get(ctx) ?? 1) - 1;
    if (count) this.controllers.set(ctx, count); else this.controllers.delete(ctx);
  }
  private matches(c: Cooldown, p: RequestPolicy, client: string): boolean {
    if (c.policy.workload !== p.workload) return false;
    if (c.access !== 'ReadWrite' && (p.readOnly ? 'Read' : 'Write') !== c.access) return false;
    if (c.kind === 'application' || c.kind === 'app-tenant') return c.client === client;
    if (c.kind === 'tenant' || c.kind === 'user') return true;
    if (c.kind === 'resource') return c.policy.resource === p.resource;
    const batchOverlap = p.mailbox && c.policy.mailbox && p.resource === c.policy.resource && (p.batch || c.policy.batch);
    return (this.lane(c.policy) === this.lane(p) || !!batchOverlap) && (c.kind === 'service' || c.client === client);
  }
  admissionAt(p: RequestPolicy, client: string): number {
    const now = this.now();
    this.cooldowns = this.cooldowns.filter(c => c.until > now);
    return Math.max(now, ...this.cooldowns.filter(c => this.matches(c, p, client)).map(c => c.until), ...this.paceKeys(p).map(k => this.pace.get(k) ?? 0));
  }
  sharedAt(p: RequestPolicy): number {
    return Math.max(this.now(), ...this.cooldowns.filter(c => !['application','app-tenant','unknown'].includes(c.kind) && this.matches(c, p, '')).map(c => c.until), ...this.paceKeys(p).map(k => this.pace.get(k) ?? 0));
  }
  private block(p: RequestPolicy, client: string, scope: Scope, until: number): void {
    const existing = this.cooldowns.find(c => c.client === client && c.kind === scope.kind && c.access === scope.access && this.lane(c.policy) === this.lane(p));
    if (existing) existing.until = Math.max(existing.until, until);
    else this.cooldowns.push({ ...scope, until, client, policy: p });
    this.wake();
  }
  private later(until: number, scope = 'unknown', attempts = 0, details?: Scope): GraphApiError {
    return new GraphApiError(`Microsoft is busy. Retry in ${Math.max(1, Math.ceil((until - this.now()) / 1000))} seconds.`, 429, undefined,
      { retryAt: until, scope, attempts, information: details?.information, resourceUnits: details?.resourceUnits, requestId: details?.requestId });
  }
  async delay(ms: number, ctx: RequestContext, wakeable = false): Promise<void> {
    this.check(ctx);
    if (ms <= 0) return;
    if (ms > ctx.remainingMs) throw this.later(this.now() + ms);
    const start = this.now();
    this.retain(ctx);
    try {
      await new Promise<void>((resolve, reject) => {
        const finish = (error?: unknown) => {
          clearTimeout(timer); this.listeners.delete(wake); ctx.controller.signal.removeEventListener('abort', abort);
          if (error) reject(error); else resolve();
        };
        const wake = () => finish();
        const abort = () => finish(new GraphApiError('Request cancelled: session ended', 0));
        const timer = setTimeout(wake, Math.min(ms, 2_147_483_647));
        if (wakeable) this.listeners.add(wake);
        ctx.controller.signal.addEventListener('abort', abort, { once: true });
      });
    } finally {
      ctx.remainingMs = Math.max(0, ctx.remainingMs - (this.now() - start));
      this.releaseContext(ctx);
    }
    this.check(ctx);
  }
  /** Bound waiting for a serialized candidate operation, including cancellation. */
  async awaitTurn(promise: Promise<void>, ctx: RequestContext): Promise<void> {
    this.check(ctx); this.retain(ctx);
    const start = this.now();
    try {
      await new Promise<void>((resolve, reject) => {
        const finish = (error?: unknown) => { clearTimeout(timer); ctx.controller.signal.removeEventListener('abort', abort); error ? reject(error) : resolve(); };
        const abort = () => finish(new GraphApiError('Request cancelled: session ended', 0));
        const timer = setTimeout(() => finish(this.later(this.now() + 1000)), ctx.remainingMs);
        ctx.controller.signal.addEventListener('abort', abort, { once: true });
        promise.then(() => finish(), finish);
      });
      this.check(ctx);
    } finally { ctx.remainingMs = Math.max(0, ctx.remainingMs - (this.now() - start)); this.releaseContext(ctx); }
  }
  private async acquire(p: RequestPolicy, client: string, ctx: RequestContext): Promise<() => void> {
    const key = `${ctx.session}|${p.workload}|${client}|${p.mailbox ? p.resource : ''}`;
    const weight = p.batch ? 4 : 1;
    for (;;) {
      this.check(ctx);
      const wait = this.admissionAt(p, client) - this.now();
      if (wait > ctx.remainingMs) {
        const cause = this.cooldowns.filter(c => this.matches(c,p,client)).sort((a,b)=>b.until-a.until)[0];
        throw this.later(this.now() + wait, cause?.kind, 0, cause);
      }
      if (wait > 0) { await this.delay(wait, ctx, true); continue; }
      const a = this.active.get(key) ?? { total: 0, background: 0 };
      const limit = p.mailbox ? 4 : 8;
      if (a.total + weight > limit || (p.mailbox && !p.batch && ctx.priority === 'background' && a.background >= 3)) {
        if (ctx.remainingMs <= 0) throw this.later(this.now() + 1000);
        await this.delay(Math.min(1000, ctx.remainingMs), ctx, true); continue;
      }
      a.total += weight; if (ctx.priority === 'background') a.background += weight;
      this.active.set(key, a);
      if (p.spacing) for (const pace of this.paceKeys(p)) this.pace.set(pace, this.now() + p.spacing);
      return () => { a.total -= weight; if (ctx.priority === 'background') a.background -= weight; if (!a.total) this.active.delete(key); this.wake(); };
    }
  }

  async execute(p: RequestPolicy, preferred: string, ctx: RequestContext,
    send: (client: string, signal: AbortSignal) => Promise<Response>, routes?: Routes, initialResponse?: Response, initialAttempts = 1): Promise<Response> {
    if (p.mailbox && p.resource === tokens.getObjectId()) p = { ...p, resource: 'me' };
    this.check(ctx); this.retain(ctx);
    let client = preferred, switched = false, refreshed = false, probe = false, originalUntil = 0;
    const lane = this.lane(p);
    try {
      const hint = this.hints.get(lane);
      const appBlocked = this.cooldowns.some(c => ['application','app-tenant'].includes(c.kind) && c.until > this.now() && this.matches(c,p,preferred));
      if (!initialResponse && ctx.adaptive && p.rotation && routes && (appBlocked || (hint && hint.until > this.now())) && this.admissionAt(p, preferred) > this.sharedAt(p)) {
        // Eligibility is rechecked by the routing callback; hints never authorize a token.
        const alternate = await routes.alternate(preferred, this.admissionAt(p, preferred));
        if (alternate) {
          client = alternate; switched = true;
          if (!appBlocked && hint) { probe = true; originalUntil = hint.until; }
        }
      }
      for (let attempt = initialAttempts; attempt <= 4; attempt++) {
        const observed = attempt === initialAttempts ? initialResponse : undefined;
        const release = observed ? () => {} : await this.acquire(p, client, ctx);
        let resp: Response;
        try { resp = observed ?? await send(client, ctx.controller.signal); }
        catch (error) {
          this.check(ctx);
          if (error instanceof GraphApiError) throw error;
          if (!p.readOnly || attempt === 4) throw new GraphApiError('Network unavailable; completion could not be confirmed.', 0, undefined, { attempts: attempt, ambiguous: !p.readOnly });
          await this.blockTransient(p, client, attempt);
          continue;
        } finally { release(); }
        this.check(ctx);
        if (resp.status === 401 && routes && !refreshed && attempt < 4) {
          refreshed = true; await resp.body?.cancel(); await routes.refresh(client); continue;
        }
        if (resp.status === 403) routes?.denied(client);
        const transient = [500,502,503,504].includes(resp.status);
        if (resp.status !== 429 && !transient) {
          if (resp.ok && probe) this.hints.set(lane, { client, until: originalUntil });
          responseAttempts.set(resp, attempt);
          return resp;
        }
        const scope = throttleScope(resp.headers, client, tokens.getTenantId() ?? '');
        const delay = retryAfterMs(resp.headers, this.now()) ?? this.backoff(attempt);
        const until = this.now() + Math.max(1, delay);
        this.block(p, client, scope, until);
        getLogger().warn('Microsoft throttle', { workload: p.workload, operation: p.operation, client: clientAlias(client), scope: scope.kind,
          delayMs: delay, attempt, requestId: scope.requestId, information: scope.information, resourceUnits: scope.resourceUnits });
        if (probe) {
          this.block(p, '', { kind: 'service', access: 'ReadWrite' }, Math.max(originalUntil, until));
          this.hints.delete(lane);
        }
        await resp.body?.cancel();
        if (transient && !p.readOnly) throw new GraphApiError('Microsoft returned a server error; the operation may have completed and was not repeated.', resp.status, undefined, { retryAt: until, scope: scope.kind, attempts: attempt, ambiguous: true });
        if (attempt === 4) throw this.later(until, scope.kind, attempt, scope);
        const appScope = scope.kind === 'application' || scope.kind === 'app-tenant';
        const state = this.probes.get(lane);
        const mayProbe = scope.kind === 'unknown' && p.readOnly && (!state || (!state.busy && state.next <= this.now()));
        if (resp.status === 429 && ctx.adaptive && routes && p.rotation && !switched && (appScope || mayProbe) && this.sharedAt(p) < until) {
          if (mayProbe) this.probes.set(lane, { busy: true, next: this.now() + 300_000 });
          try {
            const alternate = await routes.alternate(client, until);
            if (alternate && this.admissionAt(p, alternate) < until) {
              originalUntil = until; client = alternate; switched = true; probe = mayProbe;
              getLogger().info('FOCI routing', { workload: p.workload, operation: p.operation, client: clientAlias(client), reason: appScope ? scope.kind : 'read-probe' });
            }
          } finally { const state = this.probes.get(lane); if (mayProbe && !probe && state) state.busy = false; }
        }
      }
      throw new Error('Unreachable retry state');
    } finally {
      if (probe) { const state = this.probes.get(lane); if (state) state.busy = false; }
      this.releaseContext(ctx);
    }
  }
  private backoff(attempt: number): number { return Math.min(30_000, 1000 * 2 ** (attempt - 1) * (1 + this.random() * 0.2)); }
  private async blockTransient(p: RequestPolicy, client: string, attempt: number): Promise<void> {
    this.block(p, client, { kind: 'unknown', access: 'ReadWrite' }, this.now() + this.backoff(attempt));
  }
}

const backgrounds = new WeakMap<PluginAPI, Map<string, Promise<unknown>>>();
/** One refresh (including a delayed retry) per key. Never enqueue user writes here. */
export function coalesceRefresh<T>(api: PluginAPI, key: string, fn: () => Promise<T>): Promise<T> {
  let tasks = backgrounds.get(api);
  if (!tasks) { tasks = new Map(); backgrounds.set(api, tasks); }
  key = `${tokens.currentSession()}|${key}`;
  const existing = tasks.get(key);
  if (existing) return existing as Promise<T>;
  const session = tokens.currentSession();
  const background = !local.getStore() || local.getStore()!.priority === 'background';
  const run = async (): Promise<T> => {
    try { return await fn(); }
    catch (error) {
      if (!background || !(error instanceof GraphApiError) || !error.retry?.retryAt || session !== tokens.currentSession()) throw error;
      // A deferred refresh is safe; user actions never enter this branch.
      const wait = newRequestContext(api, 'background'); wait.remainingMs = Infinity;
      await getScheduler(api).delay(Math.max(1, error.retry.retryAt - Date.now()), wait);
      return withNewRequestContext(api, run, 'background');
    }
  };
  const promise = withRequestContext(api, run, 'background').finally(() => { if (tasks!.get(key) === promise) tasks!.delete(key); });
  tasks.set(key, promise);
  return promise;
}

export function serviceFetch(api: PluginAPI, workload: string, preferred: string, url: string | URL | Request, init?: RequestInit, readOnly?: boolean): Promise<Response> {
  return withRequestContext(api, () => getScheduler(api).execute({ workload, operation: 'service', resource: '',
    readOnly: readOnly ?? (!init?.method || init.method === 'GET'), rotation: false, scopes: [] }, preferred, local.getStore()!,
  (_client, signal) => api.fetch(url, { ...init, signal })));
}
