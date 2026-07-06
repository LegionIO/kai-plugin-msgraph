import type { PluginAPI } from '../shared/types.js';
import { GraphApiError } from '../shared/types.js';
import { GraphClient } from './graph-client.js';
import { getBotIcon } from './ic3-client.js';
import * as tokenCache from './token-cache.js';
import { DiskCache } from './disk-cache.js';
import * as mediaServer from './media-server.js';
import { getLogger } from './logger-singleton.js';

/** userId or appId → data URL, or null when confirmed no-photo. */
const cache = new Map<string, string | null>();
const inFlight = new Set<string>();
const CONCURRENCY = 4;

const HARD_TTL_MS = 14 * 24 * 60 * 60_000;
const SOFT_TTL_MS = 24 * 60 * 60_000;

let disk: DiskCache<string | null> | null = null;
/** Entries older than SOFT_TTL loaded from disk — served immediately, revalidated lazily. */
const staleAt = new Map<string, number>();

let publishTimer: ReturnType<typeof setTimeout> | null = null;
function schedulePublish(api: PluginAPI): void {
  if (publishTimer) return;
  publishTimer = setTimeout(() => {
    publishTimer = null;
    api.state.set('photos', published());
  }, 120);
}

export function init(api: PluginAPI): void {
  mediaServer.register('photo', (id) => cache.get(id) ?? undefined);
  disk = new DiskCache<string | null>(
    api.pluginName,
    'photos',
    { hardTtlMs: HARD_TTL_MS, maxEntries: 2000 },
    () => {
      const now = Date.now();
      for (const [k, e] of disk!.entries()) {
        if (!cache.has(k)) {
          cache.set(k, e.v);
          if (now - e.at > SOFT_TTL_MS) staleAt.set(k, e.at);
        }
      }
      getLogger().info(`photo-cache: hydrated ${cache.size} entries from disk (${staleAt.size} stale)`);
      schedulePublish(api);
    },
  );
}

/** State-facing view: local-server URLs when available, else raw data-URLs. */
export function published(): Record<string, string | null> {
  if (!mediaServer.baseUrl()) return Object.fromEntries(cache);
  const out: Record<string, string | null> = {};
  for (const [k, v] of cache) out[k] = v === null ? null : mediaServer.urlFor('photo', k);
  return out;
}

export function get(): Record<string, string | null> {
  return Object.fromEntries(cache);
}

export function clear(): void {
  cache.clear();
  inFlight.clear();
  staleAt.clear();
  disk?.clear();
  if (publishTimer) { clearTimeout(publishTimer); publishTimer = null; }
}

export function flush(): void {
  if (publishTimer) { clearTimeout(publishTimer); publishTimer = null; }
  disk?.dispose();
}

/** Fetch any userIds not already cached; publishes to state.photos as results arrive. */
export function ensure(api: PluginAPI, client: GraphClient, userIds: Iterable<string>): void {
  ensureWith(api, userIds, (id) => client.getUserPhoto(id));
}

/** Fetch bot/app icons into the same photos map, keyed by application id. */
export function ensureApps(api: PluginAPI, client: GraphClient, appIds: Iterable<string>): void {
  ensureWith(api, appIds, async (id) => (await getBotIcon(api, id)) ?? (await client.getAppIcon(id)));
}

function ensureWith(
  api: PluginAPI,
  ids: Iterable<string>,
  fetchOne: (id: string) => Promise<string | null>,
): void {
  const all = [...new Set(ids)].filter(Boolean);
  // Need a network fetch if: not cached at all, OR cached-but-stale (SWR revalidate).
  const missing = all.filter((id) => !inFlight.has(id) && (!cache.has(id) || staleAt.has(id)));
  if (missing.length === 0) return;
  for (const id of missing) inFlight.add(id);
  const session = tokenCache.currentSession();

  void (async () => {
    let idx = 0;
    const workers = Array.from({ length: Math.min(CONCURRENCY, missing.length) }, async () => {
      while (idx < missing.length) {
        if (session !== tokenCache.currentSession()) return;
        const id = missing[idx++];
        try {
          const url = await fetchOne(id);
          if (session === tokenCache.currentSession()) {
            cache.set(id, url);
            staleAt.delete(id);
            disk?.set(id, url);
          }
        } catch (err) {
          const status = err instanceof GraphApiError ? err.statusCode : 0;
          if (status >= 400 && status < 500 && status !== 429 && session === tokenCache.currentSession()) {
            // Permanent (403 no-permission, 400 bad id, etc.) — cache as no-photo.
            cache.set(id, null);
            staleAt.delete(id);
            disk?.set(id, null);
          }
          // 429 / 5xx / network: keep whatever we had (stale or nothing); retry next ensure().
          getLogger().warn(`photo fetch failed for ${id}: ${err}`);
        } finally {
          inFlight.delete(id);
        }
        if (session === tokenCache.currentSession()) schedulePublish(api);
      }
    });
    await Promise.all(workers);
  })();
}
