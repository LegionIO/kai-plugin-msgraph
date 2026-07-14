import type { PluginAPI } from '../shared/types.js';
import { GraphApiError } from '../shared/types.js';
import { GraphClient } from './graph-client.js';
import * as tokenCache from './token-cache.js';
import * as mediaServer from './media-server.js';
import { DiskCache } from './disk-cache.js';
import { getLogger } from './logger-singleton.js';
import { DEFAULT_IMAGE_CACHE_MAX_ENTRIES } from '../shared/constants.js';

/** hostedContents URL → data URL ("data:<mime>;base64,<bytes>"), or null when confirmed unfetchable. */
const cache = new Map<string, string | null>();
const inFlight = new Set<string>();
const CONCURRENCY = 3;

const HARD_TTL_MS = 14 * 24 * 60 * 60_000;
const SOFT_TTL_MS = 24 * 60 * 60_000;

let disk: DiskCache<string | null> | null = null;
/** Entries older than SOFT_TTL loaded from disk — served immediately, revalidated lazily. */
const staleAt = new Map<string, number>();

mediaServer.register('hosted', (u) => cache.get(u) ?? undefined);

let publishTimer: ReturnType<typeof setTimeout> | null = null;
function schedulePublish(api: PluginAPI): void {
  if (publishTimer) return;
  publishTimer = setTimeout(() => {
    publishTimer = null;
    api.state.set('hostedContents', published());
  }, 120);
}

/** Wire up disk persistence. Call once at startup. maxEntries caps the on-disk image cache. */
export function init(api: PluginAPI, maxEntries = DEFAULT_IMAGE_CACHE_MAX_ENTRIES): void {
  disk = new DiskCache<string | null>(
    api.pluginName,
    'hosted-content',
    { hardTtlMs: HARD_TTL_MS, maxEntries: Math.max(1, maxEntries) },
    () => {
      const now = Date.now();
      for (const [k, e] of disk!.entries()) {
        if (!cache.has(k)) {
          cache.set(k, e.v);
          if (now - e.at > SOFT_TTL_MS) staleAt.set(k, e.at);
        }
      }
      getLogger().info(`hosted-content-cache: hydrated ${cache.size} entries from disk (${staleAt.size} stale)`);
      schedulePublish(api);
    },
  );
}

export function published(): Record<string, string | null> {
  if (!mediaServer.baseUrl()) return Object.fromEntries(cache);
  const out: Record<string, string | null> = {};
  for (const [k, v] of cache) out[k] = v === null ? null : mediaServer.urlFor('hosted', k);
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

/** Update the on-disk entry cap live (e.g. when the user changes the setting). */
export function setMaxEntries(maxEntries: number): void {
  disk?.setMaxEntries(Math.max(1, maxEntries));
}

/** Store a freshly-fetched data URL in both memory and disk. */
function store(url: string, data: string | null): void {
  cache.set(url, data);
  staleAt.delete(url);
  disk?.set(url, data);
}

/** Parse a "data:<mime>;base64,<bytes>" URL into raw base64 + media type. */
function splitDataUrl(dataUrl: string): { base64: string; mediaType: string } {
  const comma = dataUrl.indexOf(',');
  const header = comma >= 0 ? dataUrl.slice(5, comma) : ''; // strip "data:"
  const base64 = comma >= 0 ? dataUrl.slice(comma + 1) : '';
  const mediaType = header.split(';')[0] || 'image/png';
  return { base64, mediaType };
}

/**
 * Get one hosted content as raw base64 + media type, serving from cache when
 * present and fetching from Graph only on a miss. This is the rate-limit-safe
 * path for tooling: repeated requests for the same image never re-hit Graph.
 * Returns null when the content is confirmed unfetchable (permanent 4xx).
 */
export async function getOne(
  api: PluginAPI,
  client: GraphClient,
  url: string,
): Promise<{ base64: string; mediaType: string } | null> {
  if (!url) return null;
  const cached = cache.get(url);
  if (cached !== undefined && !staleAt.has(url)) {
    return cached === null ? null : splitDataUrl(cached);
  }
  // Coalesce with any in-flight fetch (background ensure() or a concurrent
  // getOne) for the same URL: wait for it rather than issuing a duplicate call.
  if (inFlight.has(url)) {
    for (let i = 0; i < 100 && inFlight.has(url); i++) {
      await new Promise((r) => setTimeout(r, 100));
    }
    const after = cache.get(url);
    if (after !== undefined && !staleAt.has(url)) {
      return after === null ? null : splitDataUrl(after);
    }
    // Fell through (timeout, or still stale): fetch below only if nobody else is.
    if (inFlight.has(url)) {
      const anyCopy = cache.get(url);
      return anyCopy ? splitDataUrl(anyCopy) : null;
    }
  }
  inFlight.add(url);
  const session = tokenCache.currentSession();
  try {
    const raw = await client.getHostedContentRaw(url);
    if (session === tokenCache.currentSession()) {
      store(url, `data:${raw.mediaType};base64,${raw.base64}`);
      schedulePublish(api);
    }
    return raw;
  } catch (err) {
    const status = err instanceof GraphApiError ? err.statusCode : 0;
    if (status >= 400 && status < 500 && status !== 429 && session === tokenCache.currentSession()) {
      store(url, null); // permanent failure — remember as unfetchable
      return null;
    }
    // 429 / 5xx / network: if we have a stale copy, fall back to it rather than error.
    const stale = cache.get(url);
    if (stale) return splitDataUrl(stale);
    throw err;
  } finally {
    inFlight.delete(url);
  }
}

/** Background prefetch for the UI: fills any missing/stale URLs, publishing as results arrive. */
export function ensure(api: PluginAPI, client: GraphClient, urls: Iterable<string>): void {
  const missing = [...new Set(urls)].filter(
    (u) => u && !inFlight.has(u) && (!cache.has(u) || staleAt.has(u)),
  );
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
          if (session === tokenCache.currentSession()) store(u, data);
        } catch (err) {
          const status = err instanceof GraphApiError ? err.statusCode : 0;
          if (status >= 400 && status < 500 && status !== 429 && session === tokenCache.currentSession()) {
            store(u, null);
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
