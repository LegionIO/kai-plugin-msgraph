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

/** Only Teams hostedContents URLs are surfaced to the UI state map; mail/ref
 *  keys share the same byte cache but the renderer never looks them up.
 *  Image file-attachment URLs are opted in explicitly via publishedKeys. */
const publishedKeys = new Set<string>();
function isHostedUrl(key: string): boolean {
  return key.startsWith('https://') || publishedKeys.has(key);
}

export function published(): Record<string, string | null> {
  const out: Record<string, string | null> = {};
  const base = mediaServer.baseUrl();
  for (const [k, v] of cache) {
    if (!isHostedUrl(k)) continue;
    out[k] = v === null ? null : base ? mediaServer.urlFor('hosted', k) : v;
  }
  return out;
}

export function get(): Record<string, string | null> {
  return Object.fromEntries(cache);
}

export function clear(): void {
  cache.clear();
  inFlight.clear();
  staleAt.clear();
  publishedKeys.clear();
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
 * Generic get-or-fetch against the shared attachment-bytes cache, keyed by an
 * arbitrary string. Serves cached bytes with zero network, coalesces concurrent
 * in-flight fetches for the same key, and on a transient failure (429/5xx/
 * network) falls back to a stale cached copy rather than erroring. Returns null
 * when the fetch is a permanent failure (4xx) or yields no bytes.
 *
 * `key` must uniquely and stably identify the content (e.g. the hostedContents
 * URL, or `mail:<messageId>:<attachmentId>`, or `ref:<contentUrl>`).
 */
export async function getOneVia(
  api: PluginAPI,
  key: string,
  fetcher: () => Promise<{ base64: string; mediaType: string }>,
): Promise<{ base64: string; mediaType: string } | null> {
  if (!key) return null;
  const cached = cache.get(key);
  if (cached !== undefined && !staleAt.has(key)) {
    return cached === null ? null : splitDataUrl(cached);
  }
  // Coalesce with any in-flight fetch (background ensure() or a concurrent
  // getOne*) for the same key: wait for it rather than issuing a duplicate call.
  if (inFlight.has(key)) {
    for (let i = 0; i < 100 && inFlight.has(key); i++) {
      await new Promise((r) => setTimeout(r, 100));
    }
    const after = cache.get(key);
    if (after !== undefined && !staleAt.has(key)) {
      return after === null ? null : splitDataUrl(after);
    }
    if (inFlight.has(key)) {
      const anyCopy = cache.get(key);
      return anyCopy ? splitDataUrl(anyCopy) : null;
    }
  }
  inFlight.add(key);
  const session = tokenCache.currentSession();
  try {
    const raw = await fetcher();
    if (!raw.base64) {
      if (session === tokenCache.currentSession()) store(key, null);
      return null;
    }
    if (session === tokenCache.currentSession()) {
      store(key, `data:${raw.mediaType};base64,${raw.base64}`);
      schedulePublish(api);
    }
    return raw;
  } catch (err) {
    const status = err instanceof GraphApiError ? err.statusCode : 0;
    if (status >= 400 && status < 500 && status !== 429 && session === tokenCache.currentSession()) {
      store(key, null); // permanent failure — remember as unfetchable
      return null;
    }
    // 429 / 5xx / network: if we have a stale copy, fall back to it rather than error.
    const stale = cache.get(key);
    if (stale) return splitDataUrl(stale);
    throw err;
  } finally {
    inFlight.delete(key);
  }
}

/**
 * Get one Teams hostedContents URL as raw base64 + media type, cached. This is
 * the rate-limit-safe path for tooling: repeated requests never re-hit Graph.
 * The cache key is the URL itself, matching the UI's ensure() prefetch path.
 */
export function getOne(
  api: PluginAPI,
  client: GraphClient,
  url: string,
): Promise<{ base64: string; mediaType: string } | null> {
  return getOneVia(api, url, () => client.getHostedContentRaw(url));
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

/**
 * Prefetch image file-attachments (SharePoint references or graph-hosted files)
 * for the UI, keyed by their raw URL and published to state.hostedContents so
 * the panel can render them inline. Shares the same byte cache as the tools, so
 * a UI-loaded image file costs the AI zero extra Graph calls (and vice-versa).
 */
export function ensureImageFiles(api: PluginAPI, client: GraphClient, urls: Iterable<string>): void {
  const missing = [...new Set(urls)].filter(
    (u) => u && !inFlight.has(u) && (!cache.has(u) || staleAt.has(u)),
  );
  if (missing.length === 0) return;
  for (const u of missing) { inFlight.add(u); publishedKeys.add(u); }
  const session = tokenCache.currentSession();

  void (async () => {
    let idx = 0;
    const workers = Array.from({ length: Math.min(CONCURRENCY, missing.length) }, async () => {
      while (idx < missing.length) {
        if (session !== tokenCache.currentSession()) return;
        const u = missing[idx++];
        try {
          const raw = u.startsWith('https://graph.microsoft.com/')
            ? await client.getHostedContentRaw(u)
            : await client.downloadReferenceAttachment(u);
          if (session === tokenCache.currentSession()) {
            store(u, raw.base64 ? `data:${raw.mediaType};base64,${raw.base64}` : null);
          }
        } catch (err) {
          const status = err instanceof GraphApiError ? err.statusCode : 0;
          if (status >= 400 && status < 500 && status !== 429 && session === tokenCache.currentSession()) {
            store(u, null);
          }
          getLogger().warn(`image-file fetch failed: ${err}`);
        } finally {
          inFlight.delete(u);
        }
        if (session === tokenCache.currentSession()) schedulePublish(api);
      }
    });
    await Promise.all(workers);
  })();
}
