import type { PluginAPI } from '../shared/types.js';
import { GraphClient } from './graph-client.js';
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

  void (async () => {
    let idx = 0;
    const workers = Array.from({ length: Math.min(CONCURRENCY, missing.length) }, async () => {
      while (idx < missing.length) {
        const id = missing[idx++];
        try {
          const url = await client.getUserPhoto(id);
          cache.set(id, url);
        } catch (err) {
          getLogger().warn(`photo fetch failed for ${id}: ${err}`);
          cache.set(id, null);
        } finally {
          inFlight.delete(id);
        }
        api.state.set('photos', get());
      }
    });
    await Promise.all(workers);
  })();
}
