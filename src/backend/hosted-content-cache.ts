import type { PluginAPI } from '../shared/types.js';
import { GraphApiError } from '../shared/types.js';
import { GraphClient } from './graph-client.js';
import * as tokenCache from './token-cache.js';
import { getLogger } from './logger-singleton.js';

const cache = new Map<string, string | null>();
const inFlight = new Set<string>();
const CONCURRENCY = 3;

let publishTimer: ReturnType<typeof setTimeout> | null = null;
function schedulePublish(api: PluginAPI): void {
  if (publishTimer) return;
  publishTimer = setTimeout(() => {
    publishTimer = null;
    api.state.set('hostedContents', get());
  }, 120);
}

export function get(): Record<string, string | null> {
  return Object.fromEntries(cache);
}

export function clear(): void {
  cache.clear();
  inFlight.clear();
  if (publishTimer) { clearTimeout(publishTimer); publishTimer = null; }
}

export function ensure(api: PluginAPI, client: GraphClient, urls: Iterable<string>): void {
  const missing = [...new Set(urls)].filter((u) => u && !cache.has(u) && !inFlight.has(u));
  if (missing.length === 0) return;
  for (const u of missing) inFlight.add(u);
  const session = tokenCache.currentSession();

  void (async () => {
    let idx = 0;
    const workers = Array.from({ length: Math.min(CONCURRENCY, missing.length) }, async () => {
      while (idx < missing.length) {
        if (session !== tokenCache.currentSession()) return;
        const u = missing[idx++];
        try {
          const data = await client.getHostedContent(u);
          if (session === tokenCache.currentSession()) cache.set(u, data);
        } catch (err) {
          const status = err instanceof GraphApiError ? err.statusCode : 0;
          if (status >= 400 && status < 500 && status !== 429 && session === tokenCache.currentSession()) {
            cache.set(u, null);
          }
          getLogger().warn(`hostedContent fetch failed: ${err}`);
        } finally {
          inFlight.delete(u);
        }
        if (session === tokenCache.currentSession()) schedulePublish(api);
      }
    });
    await Promise.all(workers);
  })();
}
