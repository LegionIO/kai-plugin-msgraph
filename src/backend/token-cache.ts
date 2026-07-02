import { TOKEN_REFRESH_BUFFER_MS } from '../shared/constants.js';
import type { GraphTokenData, PluginAPI } from '../shared/types.js';
import { getLogger } from './logger-singleton.js';

let cached: GraphTokenData | null = null;
/** Bumped on logout; async writers check this before committing. */
let sessionGen = 0;

export function currentSession(): number {
  return sessionGen;
}

export function invalidateSession(): void {
  sessionGen++;
}

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

type PersistedToken = Omit<GraphTokenData, 'accessToken' | 'refreshToken'> & {
  refreshTokenEnc?: string;
};

export function persist(api: PluginAPI): void {
  if (!cached) {
    api.config.setPluginData('graphToken', null);
    return;
  }
  const { accessToken: _at, refreshToken, ...rest } = cached;
  void _at;
  const out: PersistedToken = { ...rest };
  try {
    if (!api.safeStorage.isEncryptionAvailable()) throw new Error('safeStorage unavailable');
    out.refreshTokenEnc = api.safeStorage.encryptString(refreshToken);
  } catch (err) {
    // Fail closed: keep the refresh token in memory only for this session.
    getLogger().warn(`Not persisting refresh token (encryption unavailable): ${err}`);
    api.config.setPluginData('graphToken', null);
    return;
  }
  api.config.setPluginData('graphToken', out);
}

export function loadPersisted(api: PluginAPI): void {
  const data = api.config.getPluginData();
  const stored = data.graphToken as
    | (PersistedToken & { refreshToken?: string; refreshTokenPlain?: string })
    | null
    | undefined;
  if (!stored) return;
  // Drop any legacy plaintext at rest.
  if (stored.refreshToken || stored.refreshTokenPlain) {
    api.config.setPluginData('graphToken', null);
    return;
  }
  if (!stored.refreshTokenEnc) return;
  let refreshToken: string;
  try {
    refreshToken = api.safeStorage.decryptString(stored.refreshTokenEnc);
  } catch {
    return;
  }
  cached = {
    accessToken: '',
    refreshToken,
    expiresAt: stored.expiresAt ?? 0,
    objectId: stored.objectId ?? '',
    email: stored.email ?? '',
    displayName: stored.displayName ?? null,
    scopes: stored.scopes ?? '',
  };
}
