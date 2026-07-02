import type { PluginAPI, Presence } from '../shared/types.js';
import { GraphClient } from './graph-client.js';
import * as tokenCache from './token-cache.js';
import { getLogger } from './logger-singleton.js';

const cache = new Map<string, Presence>();
let disabled = false;

export function get(): Record<string, Presence> {
  return Object.fromEntries(cache);
}

export function clear(): void {
  cache.clear();
  disabled = false;
}

/** Refresh presence for the given ids (batched, max 650 per Graph call). */
export function refresh(api: PluginAPI, client: GraphClient, userIds: Iterable<string>): void {
  if (disabled) return;
  const ids = [...new Set(userIds)].filter(Boolean);
  if (ids.length === 0) return;
  const session = tokenCache.currentSession();

  void (async () => {
    try {
      for (let i = 0; i < ids.length; i += 650) {
        const chunk = ids.slice(i, i + 650);
        const res = await client.getPresences(chunk);
        if (session !== tokenCache.currentSession()) return;
        for (const [id, p] of Object.entries(res)) cache.set(id, p);
      }
      if (session === tokenCache.currentSession()) api.state.set('presence', get());
    } catch (err) {
      const status = (err as { statusCode?: number }).statusCode ?? 0;
      if (status === 403 || status === 401) {
        getLogger().warn(`Presence disabled (token lacks Presence.Read.All): ${err}`);
        disabled = true;
      } else {
        getLogger().warn(`Presence refresh failed: ${err}`);
      }
    }
  })();
}
