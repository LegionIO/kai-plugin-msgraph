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
  tenantId: string;
  email: string;
  displayName: string | null;
  /** Space-separated scp claim from the access token, for diagnostics. */
  scopes: string;
}

export interface AuthStatus {
  isAuthenticated: boolean;
  email: string | null;
  displayName: string | null;
  objectId: string | null;
  jobTitle: string | null;
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
  jobTitle?: string | null;
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
  viewpoint?: { isHidden?: boolean; lastMessageReadDateTime?: string | null } | null;
  members?: GraphChatMember[];
  lastMessagePreview?: {
    id?: string;
    createdDateTime?: string;
    messageType?: string;
    body?: { content?: string; contentType?: string };
    from?: {
      user?: { id?: string; displayName?: string } | null;
      application?: { id?: string; displayName?: string } | null;
    } | null;
    eventDetail?: { '@odata.type'?: string; [k: string]: unknown } | null;
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
  from?: {
    user?: { id?: string; displayName?: string } | null;
    application?: { id?: string; displayName?: string } | null;
  } | null;
  body?: { contentType?: 'text' | 'html'; content?: string } | null;
  attachments?: Array<{ id?: string; contentType?: string; name?: string; contentUrl?: string; content?: string }>;
  mentions?: unknown[];
  reactions?: GraphReaction[];
  eventDetail?: { '@odata.type'?: string; [k: string]: unknown } | null;
}

// ── Normalized Types (frontend / tool output) ──
export interface NormalizedChat {
  id: string;
  type: GraphChatType;
  topic: string | null;
  members: Array<{ id: string; displayName: string; email: string | null; isBot?: boolean }>;
  lastUpdated: string | null;
  lastMessagePreview: string | null;
  lastMessageFrom: string | null;
  unread: boolean;
  webUrl: string | null;
}

export interface NormalizedReaction {
  emoji: string;
  type: string;
  count: number;
  users: string[];
  mine: boolean;
}

export type BodySegment =
  | { type: 'text'; text: string }
  | { type: 'br' }
  | { type: 'mention'; userId: string | null; displayName: string }
  | { type: 'code'; code: string }
  | { type: 'codeblock'; code: string; lang: string | null }
  | { type: 'blockquote'; segments: BodySegment[] }
  | { type: 'heading'; level: 1 | 2 | 3; segments: BodySegment[] }
  | { type: 'hr' }
  | { type: 'link'; href: string; text: string }
  | { type: 'image'; url: string }
  | { type: 'table'; header: string[] | null; rows: string[][] };

export interface NormalizedMessage {
  id: string;
  chatId: string;
  createdDateTime: string | null;
  fromId: string | null;
  fromName: string | null;
  fromApp: boolean;
  fromMe: boolean;
  contentType: 'text' | 'html';
  text: string;
  segments: BodySegment[];
  /** Auth-protected Graph hostedContents URLs extracted from inline <img> tags. */
  hostedImages: string[];
  replyTo: { id: string | null; senderName: string | null; text: string | null } | null;
  forwarded: { senderName: string | null; text: string | null; originalDate: string | null } | null;
  files: Array<{ name: string; url: string | null; contentType: string | null }>;
  /** Image-typed file attachments (by extension) that can be rendered inline; bytes fetched into the hosted-content cache and published to state.hostedContents keyed by url. */
  imageFiles: Array<{ name: string; url: string; contentType: string | null }>;
  cards: Array<{ id: string | null; name: string | null; contentJson: string }>;
  attachments: Array<{ name: string | null; contentType: string | null; url: string | null }>;
  systemEvent: string | null;
  reactions: NormalizedReaction[];
  deleted: boolean;
}

// ── Mail (Outlook via Graph) ──

export interface MailAddress {
  name: string | null;
  address: string;
}

export interface MailFolder {
  id: string;
  displayName: string;
  wellKnownName: string | null;
  unreadItemCount: number;
  totalItemCount: number;
  childFolderCount: number;
  parentId: string | null;
  depth: number;
}

export interface NormalizedMailSummary {
  id: string;
  conversationId: string | null;
  subject: string;
  from: MailAddress | null;
  toRecipients: MailAddress[];
  receivedDateTime: string | null;
  isRead: boolean;
  isDraft: boolean;
  hasAttachments: boolean;
  flagged: boolean;
  importance: 'low' | 'normal' | 'high';
  bodyPreview: string;
  webLink: string | null;
}

export interface MailAttachmentMeta {
  id: string;
  name: string;
  contentType: string | null;
  size: number;
  isInline: boolean;
  contentId: string | null;
}

export interface NormalizedMail extends NormalizedMailSummary {
  ccRecipients: MailAddress[];
  bccRecipients: MailAddress[];
  bodyHtml: string;
  attachments: MailAttachmentMeta[];
}

export interface OutgoingMail {
  to: MailAddress[];
  cc?: MailAddress[];
  bcc?: MailAddress[];
  subject: string;
  bodyHtml: string;
  /** Base64 file attachments. Set contentId + isInline for images referenced via cid: in bodyHtml. */
  attachments?: Array<{
    name: string;
    contentType: string;
    contentBytes: string;
    contentId?: string;
    isInline?: boolean;
  }>;
}

export interface MailSignature {
  html: string;
  autoAddOnNew: boolean;
  autoAddOnReply: boolean;
  /** 'owa' = fetched from Outlook web settings; 'config' = user-entered override in plugin settings. */
  source: 'owa' | 'config';
}

export interface MailComposeState {
  mode: 'new' | 'reply' | 'replyAll' | 'forward';
  /** Source message when replying/forwarding. */
  sourceId: string | null;
  to: MailAddress[];
  cc: MailAddress[];
  subject: string;
  /** Quoted original (HTML) appended below the user's body when replying/forwarding. */
  quotedHtml: string | null;
  /** Signature HTML appended between the user's body and the quoted original. */
  signatureHtml: string | null;
}

export interface Presence {
  availability: string;
  activity: string;
  statusMessage?: string | null;
}

export interface TypingState {
  chatId: string;
  userId: string;
  displayName: string | null;
  /** Epoch ms after which the indicator should be hidden. */
  until: number;
}

// ── Card action invocation (via Teams IC3; see backend/ic3-client.ts) ──
export type CardActionKind = 'messageback' | 'task/fetch' | 'task/submit' | 'execute';

export interface CardActionPayload {
  kind: CardActionKind;
  chatId: string;
  messageId: string;
  botId: string;
  /** Merged action.data + collected Input.* values, minus the msteams routing block. */
  data: Record<string, unknown>;
  /** Action.Execute only. */
  verb?: string | null;
  /** UI label for the pending overlay. */
  title?: string | null;
}

export interface TaskModuleState {
  botId: string;
  chatId: string;
  messageId: string;
  title: string | null;
  /** Adaptive Card JSON to render in the dialog; null while loading. */
  card: unknown;
  /** When the bot returns a web-view task module instead of a card. */
  url?: string | null;
  width?: number | string;
  height?: number | string;
  submitting: boolean;
  error: string | null;
  /** Dynamic ChoiceSet typeahead round-trip. */
  choiceSearch?: {
    reqId: number;
    query: string;
    loading: boolean;
    results: Array<{ title: string; value: string }>;
    error?: string | null;
  } | null;
}

// ── Plugin State ──
export interface MsgraphPluginState {
  /** Which sub-view the panel router shows. Source of truth over Kai's panel selection. */
  activeView: 'teams' | 'mail';
  auth: AuthStatus;
  mfa: MfaState;
  credentials: CredentialStatus;
  chats: NormalizedChat[];
  chatsNextLink: string | null;
  chatsFullyLoaded: boolean;
  loadingMoreChats: boolean;
  remoteSearch: { query: string; loading: boolean; results: NormalizedChat[]; error?: string | null } | null;
  peopleSearch: {
    query: string;
    loading: boolean;
    results: Array<{ id: string; displayName: string; email: string | null }>;
    error?: string | null;
  } | null;
  /** Populated by the 'load-user-card' action. */
  userCard: {
    userId: string;
    loading: boolean;
    displayName: string | null;
    email: string | null;
    jobTitle: string | null;
    error?: string | null;
  } | null;
  composerReplyTo: { messageId: string; senderName: string | null; text: string | null } | null;
  composerEditing: {
    chatId: string;
    messageId: string;
    segments: BodySegment[];
    /** Original attachments (cards, files, message refs) — preserved verbatim on save. */
    attachments: Array<{ id?: string; contentType?: string; name?: string; contentUrl?: string; content?: string }>;
  } | null;
  forwardTarget: { chatId: string; messageId: string; senderName: string | null; text: string | null } | null;
  taskModule: TaskModuleState | null;
  /** Transient: message-id currently awaiting a card-action round-trip (drives per-card spinner). */
  cardActionPending: string | null;
  activeChatId: string | null;
  activeChatMessages: NormalizedMessage[];
  activeChatMessagesNextLink: string | null;
  loadingOlderMessages: boolean;
  /** Real-time push connection status. */
  realtime: 'connecting' | 'connected' | 'disconnected' | 'disabled';
  realtimeError: string | null;
  /** Per-chat typing indicators keyed by chatId. */
  typing: Record<string, TypingState>;
  /** chatId → userId → last-read message id (arrival timestamp string). */
  readReceipts: Record<string, Record<string, string>>;
  /** userId → data-URL (or null when the user has no photo, so we stop retrying). */
  photos: Record<string, string | null>;
  /** userId → presence (availability/activity). */
  presence: Record<string, Presence>;
  /** hostedContent URL → data-URL (or null when fetch permanently failed). */
  hostedContents: Record<string, string | null>;
  // ── Mail ──
  mailFolders: MailFolder[];
  /** Folder ids whose children have been fetched and are shown. */
  mailFoldersExpanded: string[];
  activeMailFolder: string;
  mailList: NormalizedMailSummary[];
  mailListNextLink: string | null;
  mailSearch: { query: string; loading: boolean; results: NormalizedMailSummary[]; error?: string | null } | null;
  /** email address (lowercased) → AAD user id (or null when external / not found). */
  mailSenderIds: Record<string, string | null>;
  activeMailId: string | null;
  activeMail: NormalizedMail | null;
  /** attachmentId → data-URL (inline images). */
  mailInlineAttachments: Record<string, string | null>;
  composingMail: MailComposeState | null;
  mailSignature: MailSignature | null;
  loadingMailFolders: boolean;
  loadingMailList: boolean;
  loadingMail: boolean;
  sendingMail: boolean;
  mailError: string | null;
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
  /** Max attachment files (Teams images, mail & SharePoint attachments) kept in the on-disk bytes cache (oldest evicted). */
  imageCacheMaxEntries?: number;
  /** Manual signature override (HTML). Used when Outlook's roaming signatures aren't fetchable. */
  mailSignatureHtml?: string;
  mailSignatureAutoNew?: boolean;
  mailSignatureAutoReply?: boolean;
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
  editMessage: boolean;
  deleteMessage: boolean;
  forwardMessage: boolean;
  markChatRead: boolean;
  getPresence: boolean;
  setPresence: boolean;
  setStatusMessage: boolean;
  invokeCardAction: boolean;
  getTeamsImage: boolean;
  getAttachment: boolean;
  listFolders: boolean;
  listMail: boolean;
  getMail: boolean;
  searchMail: boolean;
  sendMail: boolean;
  replyToMail: boolean;
  markMail: boolean;
  archiveMail: boolean;
  deleteMail: boolean;
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
  editMessage: true,
  deleteMessage: true,
  forwardMessage: true,
  markChatRead: true,
  getPresence: true,
  setPresence: true,
  setStatusMessage: true,
  invokeCardAction: false,
  getTeamsImage: true,
  getAttachment: true,
  listFolders: true,
  listMail: true,
  getMail: true,
  searchMail: true,
  sendMail: true,
  replyToMail: true,
  markMail: true,
  archiveMail: true,
  deleteMail: false,
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
  /** Automation event bus (Kai ≥ 1.0.121). Optional so older hosts still load us. */
  events?: {
    declare: (decl: {
      events?: Array<{ event: string; title: string; description?: string; payloadSchema?: Record<string, unknown> }>;
      actions?: Array<{ targetId: string; title: string; description?: string; inputSchema?: Record<string, unknown> }>;
    }) => void;
    emit: (event: string, payload?: unknown) => void;
    on: (key: string, handler: (event: { key: string; source: string; event: string; payload: unknown; ts: number }) => void) => () => void;
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
