// ── Error Types ──
export class TokenExpiredError extends Error {
  name = 'TokenExpiredError' as const;
}

export class GraphApiError extends Error {
  name = 'GraphApiError' as const;
  constructor(message: string, public readonly statusCode: number, public readonly body?: unknown) {
    super(message);
  }
}

// ── Token / Auth ──
export interface GraphTokenData {
  accessToken: string;
  /** Family refresh token (FOCI). Long-lived; used to silently mint new access tokens. */
  refreshToken: string;
  /** Epoch ms. */
  expiresAt: number;
  /** AAD object id (from id_token oid claim). */
  objectId: string;
  email: string;
  displayName: string | null;
  /** Space-separated scp claim from the access token, for diagnostics. */
  scopes: string;
}

export interface AuthStatus {
  isAuthenticated: boolean;
  email: string | null;
  displayName: string | null;
  minutesRemaining: number | null;
  autoLoginStatus: string | null;
}

export interface CredentialStatus {
  hasCredentials: boolean;
  username: string | null;
  encryptionMethod: 'os-keychain' | 'aes256gcm' | 'none';
}

export interface MfaState {
  needed: boolean;
  type: 'sms' | 'totp' | 'push' | null;
  approvalNumber: string | null;
}

// ── Graph API Response Types (subset used) ──
export interface GraphUser {
  id: string;
  displayName?: string | null;
  userPrincipalName?: string | null;
  mail?: string | null;
}

export interface GraphChatMember {
  '@odata.type'?: string;
  id?: string;
  userId?: string | null;
  displayName?: string | null;
  email?: string | null;
  roles?: string[];
}

export type GraphChatType = 'oneOnOne' | 'group' | 'meeting' | 'unknownFutureValue';

export interface GraphChat {
  id: string;
  topic?: string | null;
  chatType: GraphChatType;
  createdDateTime?: string;
  lastUpdatedDateTime?: string;
  webUrl?: string | null;
  members?: GraphChatMember[];
  lastMessagePreview?: {
    id?: string;
    createdDateTime?: string;
    body?: { content?: string; contentType?: string };
    from?: { user?: { id?: string; displayName?: string } | null } | null;
  } | null;
}

export interface GraphReaction {
  reactionType: string;
  createdDateTime?: string;
  user?: { user?: { id?: string; displayName?: string } | null } | null;
}

export interface GraphMessage {
  id: string;
  chatId?: string;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  deletedDateTime?: string | null;
  messageType?: string;
  from?: { user?: { id?: string; displayName?: string } | null; application?: unknown } | null;
  body?: { contentType?: 'text' | 'html'; content?: string } | null;
  attachments?: Array<{ id?: string; contentType?: string; name?: string; contentUrl?: string }>;
  mentions?: unknown[];
  reactions?: GraphReaction[];
}

// ── Normalized Types (frontend / tool output) ──
export interface NormalizedChat {
  id: string;
  type: GraphChatType;
  topic: string | null;
  members: Array<{ id: string; displayName: string; email: string | null }>;
  lastUpdated: string | null;
  lastMessagePreview: string | null;
  lastMessageFrom: string | null;
  webUrl: string | null;
}

export interface NormalizedReaction {
  emoji: string;
  type: string;
  count: number;
  users: string[];
}

export interface NormalizedMessage {
  id: string;
  chatId: string;
  createdDateTime: string | null;
  fromId: string | null;
  fromName: string | null;
  fromMe: boolean;
  contentType: 'text' | 'html';
  text: string;
  attachments: Array<{ name: string | null; contentType: string | null; url: string | null }>;
  reactions: NormalizedReaction[];
}

// ── Plugin State ──
export interface MsgraphPluginState {
  auth: AuthStatus;
  mfa: MfaState;
  credentials: CredentialStatus;
  chats: NormalizedChat[];
  activeChatId: string | null;
  activeChatMessages: NormalizedMessage[];
  /** userId → data-URL (or null when the user has no photo, so we stop retrying). */
  photos: Record<string, string | null>;
  loadingChats: boolean;
  loadingMessages: boolean;
  sendingMessage: boolean;
  error: string | null;
}

// ── Preferences / Config ──
export interface UserPreferences {
  notifications: boolean;
  pollIntervalSeconds: number;
  debugLogging: boolean;
}

export interface ToolPermissions {
  authStatus: boolean;
  findUser: boolean;
  listChats: boolean;
  getChatMessages: boolean;
  searchMessages: boolean;
  sendMessage: boolean;
  sendDm: boolean;
  createGroupChat: boolean;
  reactToMessage: boolean;
}

export const DEFAULT_TOOL_PERMISSIONS: ToolPermissions = {
  authStatus: true,
  findUser: true,
  listChats: true,
  getChatMessages: true,
  searchMessages: true,
  sendMessage: true,
  sendDm: true,
  createGroupChat: true,
  reactToMessage: true,
};

// ── Plugin API (subset used by this plugin) ──
export interface AuthWindowHelpers {
  executeJavaScript: (code: string) => Promise<unknown>;
  getURL: () => string;
  show: () => void;
  close: () => void;
  onDidNavigate: (cb: (url: string) => void) => void;
}

export interface PluginAPI {
  pluginName: string;
  pluginDir: string;
  config: {
    get: () => Record<string, unknown>;
    set: (path: string, value: unknown) => void;
    getPluginData: () => Record<string, unknown>;
    setPluginData: (path: string, value: unknown) => void;
    onChanged: (callback: (config: Record<string, unknown>) => void) => () => void;
  };
  state: {
    get: () => Record<string, unknown>;
    replace: (next: Record<string, unknown>) => void;
    set: (path: string, value: unknown) => void;
    emitEvent: (eventName: string, data?: unknown) => void;
  };
  tools: {
    register: (tools: unknown[]) => void;
    unregister: (toolNames: string[]) => void;
  };
  ui: {
    showBanner: (descriptor: Record<string, unknown>) => void;
    hideBanner?: (id: string) => void;
    registerSettingsView: (descriptor: Record<string, unknown>) => void;
    registerPanelView: (descriptor: Record<string, unknown>) => void;
    registerNavigationItem: (descriptor: Record<string, unknown>) => void;
  };
  notifications: {
    show: (descriptor: Record<string, unknown>) => void;
    dismiss: (id: string) => void;
  };
  navigation: {
    open: (target: Record<string, unknown>) => void;
  };
  auth: {
    openAuthWindow: (options: Record<string, unknown>) => Promise<{ success: boolean; params?: Record<string, string>; error?: string }>;
  };
  shell: {
    openExternal: (url: string) => Promise<void>;
  };
  safeStorage: {
    isEncryptionAvailable: () => boolean;
    encryptString: (plaintext: string) => string;
    decryptString: (base64Cipher: string) => string;
  };
  log: {
    info: (...args: unknown[]) => void;
    warn: (...args: unknown[]) => void;
    error: (...args: unknown[]) => void;
  };
  onAction: (targetId: string, handler: (action: string, data?: unknown) => void | Promise<void>) => void;
  fetch: typeof globalThis.fetch;
}
