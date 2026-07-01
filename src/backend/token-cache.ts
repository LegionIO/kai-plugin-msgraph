import { TOKEN_REFRESH_BUFFER_MS } from '../shared/constants.js';
import type { GraphTokenData, PluginAPI } from '../shared/types.js';

let cached: GraphTokenData | null = null;

export function get(): GraphTokenData | null {
  return cached;
}

/** Returns the access token only if it is still valid past the refresh buffer. */
export function getValidAccessToken(): string | null {
  if (!cached) return null;
  if (cached.expiresAt <= Date.now() + TOKEN_REFRESH_BUFFER_MS) return null;
  return cached.accessToken;
}

export function getRefreshToken(): string | null {
  return cached?.refreshToken ?? null;
}

export function set(token: GraphTokenData): void {
  cached = token;
}

export function clear(): void {
  cached = null;
}

export function isTokenValid(): boolean {
  return getValidAccessToken() !== null;
}

export function hasRefreshToken(): boolean {
  return !!cached?.refreshToken;
}

export function minutesRemaining(): number | null {
  if (!cached) return null;
  const remaining = cached.expiresAt - Date.now();
  if (remaining <= 0) return 0;
  return Math.floor(remaining / 60_000);
}

export function getEmail(): string | null {
  return cached?.email ?? null;
}

export function getDisplayName(): string | null {
  return cached?.displayName ?? null;
}

export function getObjectId(): string | null {
  return cached?.objectId ?? null;
}

export function persist(api: PluginAPI): void {
  api.config.setPluginData('graphToken', cached);
}

export function loadPersisted(api: PluginAPI): void {
  const data = api.config.getPluginData();
  const stored = data.graphToken as GraphTokenData | null | undefined;
  if (stored?.refreshToken) {
    // Keep even if the access token is expired — the refresh token lets us mint a new one.
    cached = stored;
  }
}
