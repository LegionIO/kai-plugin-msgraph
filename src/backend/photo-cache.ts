import type { PluginAPI } from '../shared/types.js';
import { GraphApiError } from '../shared/types.js';
import { GraphClient } from './graph-client.js';
import * as tokenCache from './token-cache.js';
import { getLogger } from './logger-singleton.js';

/** userId → data URL, or null when confirmed no-photo. */
const cache = new Map<string, string | null>();
const inFlight = new Set<string>();
const CONCURRENCY = 4;

export function get(): Record<string, string | null> {
  return Object.fromEntries(cache);
}

export function clear(): void {
  cache.clear();
  inFlight.clear();
}

/** Fetch any userIds not already cached; publishes to state.photos as results arrive. */
export function ensure(api: PluginAPI, client: GraphClient, userIds: Iterable<string>): void {
  const missing = [...new Set(userIds)].filter((id) => id && !cache.has(id) && !inFlight.has(id));
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
          const url = await client.getUserPhoto(id);
          if (session === tokenCache.currentSession()) cache.set(id, url);
        } catch (err) {
          const status = err instanceof GraphApiError ? err.statusCode : 0;
          if (status >= 400 && status < 500 && status !== 429 && session === tokenCache.currentSession()) {
            // Permanent (403 no-permission, 400 bad id, etc.) — cache as no-photo.
            cache.set(id, null);
          }
          // 429 / 5xx / network: leave uncached so a later ensure() retries.
          getLogger().warn(`photo fetch failed for ${id}: ${err}`);
        } finally {
          inFlight.delete(id);
        }
        if (session === tokenCache.currentSession()) api.state.set('photos', get());
      }
    });
    await Promise.all(workers);
  })();
}
