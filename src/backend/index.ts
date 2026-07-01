import {
  acquireTokenInteractive,
  acquireTokenSilent,
  ensureAccessToken,
  submitMfaCode,
  cancelMfaCode,
} from './auth.js';
import * as tokenCache from './token-cache.js';
import * as credentialStore from './credential-store.js';
import * as photoCache from './photo-cache.js';
import { GraphClient, normalizeChat, normalizeMessage } from './graph-client.js';
import { buildMsgraphTools, ALL_TOOL_NAMES } from './tools.js';
import { setLogger, getLogger } from './logger-singleton.js';
import {
  PANEL_ID,
  NAV_ID,
  SETTINGS_ID,
  DEFAULT_POLL_INTERVAL_SECONDS,
  TOKEN_REFRESH_BUFFER_MS,
} from '../shared/constants.js';
import type {
  PluginAPI,
  MsgraphPluginState,
  UserPreferences,
  ToolPermissions,
  MfaState,
} from '../shared/types.js';
import { DEFAULT_TOOL_PERMISSIONS } from '../shared/types.js';

let pollTimer: ReturnType<typeof setInterval> | null = null;
let refreshTimer: ReturnType<typeof setInterval> | null = null;
let unsubConfig: (() => void) | null = null;
let hadValidTokenSinceLogout = false;

// ── Config helpers ──

function getPreferences(api: PluginAPI): UserPreferences {
  const data = api.config.getPluginData();
  const prefs = (data.preferences ?? {}) as Partial<UserPreferences>;
  return {
    notifications: prefs.notifications ?? true,
    pollIntervalSeconds: prefs.pollIntervalSeconds ?? DEFAULT_POLL_INTERVAL_SECONDS,
    debugLogging: prefs.debugLogging ?? false,
  };
}

function getToolPermissions(api: PluginAPI): ToolPermissions {
  const data = api.config.getPluginData();
  const perms = (data.toolPermissions ?? {}) as Partial<ToolPermissions>;
  return { ...DEFAULT_TOOL_PERMISSIONS, ...perms };
}

// ── State ──

function initialState(): MsgraphPluginState {
  return {
    auth: {
      isAuthenticated: false,
      email: null,
      displayName: null,
      minutesRemaining: null,
      autoLoginStatus: null,
    },
    mfa: { needed: false, type: null, approvalNumber: null },
    credentials: { hasCredentials: false, username: null, encryptionMethod: 'none' },
    photos: {},
    chats: [],
    activeChatId: null,
    activeChatMessages: [],
    loadingChats: false,
    loadingMessages: false,
    sendingMessage: false,
    error: null,
  };
}

function publishAuthState(api: PluginAPI): void {
  api.state.set('auth', {
    isAuthenticated: tokenCache.isTokenValid() || tokenCache.hasRefreshToken(),
    email: tokenCache.getEmail(),
    displayName: tokenCache.getDisplayName(),
    minutesRemaining: tokenCache.minutesRemaining(),
    autoLoginStatus: null,
  });
}

function publishCredentialState(api: PluginAPI): void {
  api.state.set('credentials', credentialStore.getCredentialStatus(api));
}

// ── Client / auth wiring ──

async function ensureAuthenticated(api: PluginAPI): Promise<GraphClient> {
  await ensureAccessToken(api);
  hadValidTokenSinceLogout = true;
  publishAuthState(api);
  return new GraphClient(api);
}

// ── Tools ──

function registerEnabledTools(api: PluginAPI): void {
  const perms = getToolPermissions(api);
  const allTools = buildMsgraphTools({
    api,
    ensureAuthenticated: () => ensureAuthenticated(api),
  });

  const permMap: Record<string, keyof ToolPermissions> = {
    'auth-status': 'authStatus',
    'find-user': 'findUser',
    'list-chats': 'listChats',
    'get-chat-messages': 'getChatMessages',
    'search-messages': 'searchMessages',
    'send-message': 'sendMessage',
    'send-dm': 'sendDm',
    'react-to-message': 'reactToMessage',
    'create-group-chat': 'createGroupChat',
  };

  const enabled = allTools.filter((t) => perms[permMap[t.name]] !== false);

  api.tools.unregister(ALL_TOOL_NAMES);
  if (enabled.length > 0) api.tools.register(enabled);
  getLogger().info(
    `Registered ${enabled.length}/${allTools.length} msgraph tools (${enabled.map((t) => t.name).join(', ') || 'none'})`,
  );
}

// ── Data loaders ──

async function loadChats(api: PluginAPI): Promise<void> {
  if (!tokenCache.hasRefreshToken() && !tokenCache.isTokenValid()) return;
  api.state.set('loadingChats', true);
  api.state.set('error', null);
  try {
    const client = await ensureAuthenticated(api);
    const raw = await client.listChats({});
    const myId = tokenCache.getObjectId();
    const chats = raw.map((c) => normalizeChat(c, myId));
    api.state.set('chats', chats);
    const ids = new Set<string>();
    if (myId) ids.add(myId);
    for (const c of chats) for (const m of c.members) ids.add(m.id);
    photoCache.ensure(api, client, ids);
  } catch (err) {
    api.state.set('error', err instanceof Error ? err.message : String(err));
  } finally {
    api.state.set('loadingChats', false);
  }
}

async function loadMessages(api: PluginAPI, chatId: string): Promise<void> {
  api.state.set('activeChatId', chatId);
  api.state.set('loadingMessages', true);
  api.state.set('error', null);
  try {
    const client = await ensureAuthenticated(api);
    const msgs = await client.getChatMessages(chatId, 40);
    const myId = tokenCache.getObjectId();
    const normalized = msgs
      .filter((m) => m.messageType === 'message' || m.messageType == null)
      .map((m) => normalizeMessage(m, myId))
      .reverse();
    api.state.set('activeChatMessages', normalized);
    const ids = new Set<string>();
    for (const m of normalized) if (m.fromId) ids.add(m.fromId);
    photoCache.ensure(api, client, ids);
  } catch (err) {
    api.state.set('error', err instanceof Error ? err.message : String(err));
  } finally {
    api.state.set('loadingMessages', false);
  }
}

// ── Actions ──

async function handlePanelAction(api: PluginAPI, action: string, data?: unknown): Promise<void> {
  const log = getLogger();
  try {
    switch (action) {
      case 'login': {
        api.state.set('error', null);
        await acquireTokenInteractive(api, { forceRefresh: true });
        hadValidTokenSinceLogout = true;
        publishAuthState(api);
        await loadChats(api);
        break;
      }
      case 'logout': {
        tokenCache.clear();
        tokenCache.persist(api);
        photoCache.clear();
        hadValidTokenSinceLogout = false;
        api.state.replace(initialState() as unknown as Record<string, unknown>);
        publishCredentialState(api);
        break;
      }
      case 'refresh-chats': {
        await loadChats(api);
        break;
      }
      case 'select-chat': {
        const { chatId } = data as { chatId: string };
        await loadMessages(api, chatId);
        break;
      }
      case 'send-message': {
        const { chatId, text } = data as { chatId: string; text: string };
        api.state.set('sendingMessage', true);
        try {
          const client = await ensureAuthenticated(api);
          await client.sendMessage(chatId, text);
          await loadMessages(api, chatId);
        } finally {
          api.state.set('sendingMessage', false);
        }
        break;
      }
      case 'react-to-message': {
        const { chatId, messageId, reactionType, remove } = data as {
          chatId: string; messageId: string; reactionType: string; remove?: boolean;
        };
        const client = await ensureAuthenticated(api);
        if (remove) await client.unsetReaction(chatId, messageId, reactionType);
        else await client.setReaction(chatId, messageId, reactionType);
        await loadMessages(api, chatId);
        break;
      }
      case 'submit-mfa-code': {
        const { code } = data as { code: string };
        submitMfaCode(code);
        break;
      }
      case 'cancel-mfa': {
        cancelMfaCode();
        api.state.set('mfa', { needed: false, type: null, approvalNumber: null } satisfies MfaState);
        break;
      }
      case 'open-in-teams': {
        const { url } = data as { url: string };
        await api.shell.openExternal(url);
        break;
      }
      default:
        log.warn(`Unknown panel action: ${action}`);
    }
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    log.error(`Panel action '${action}' failed: ${message}`);
    api.state.set('error', message);
  }
}

async function handleSettingsAction(api: PluginAPI, action: string, data?: unknown): Promise<void> {
  const log = getLogger();
  try {
    switch (action) {
      case 'login':
      case 'logout':
        return handlePanelAction(api, action, data);
      case 'save-credentials': {
        const { username, password } = data as { username: string; password: string };
        credentialStore.saveCredentials(api, username, password);
        publishCredentialState(api);
        break;
      }
      case 'clear-credentials': {
        credentialStore.clearCredentials(api);
        publishCredentialState(api);
        break;
      }
      case 'set-tool-permission': {
        const { key, value } = data as { key: keyof ToolPermissions; value: boolean };
        const current = getToolPermissions(api);
        api.config.setPluginData('toolPermissions', { ...current, [key]: value });
        registerEnabledTools(api);
        break;
      }
      case 'set-preference': {
        const { key, value } = data as { key: keyof UserPreferences; value: unknown };
        const current = getPreferences(api);
        api.config.setPluginData('preferences', { ...current, [key]: value });
        break;
      }
      default:
        log.warn(`Unknown settings action: ${action}`);
    }
  } catch (err) {
    log.error(`Settings action '${action}' failed: ${err instanceof Error ? err.message : String(err)}`);
  }
}

// ── Background: proactive refresh + poll ──

async function refreshTick(api: PluginAPI): Promise<void> {
  const t = tokenCache.get();
  if (!t?.refreshToken || !hadValidTokenSinceLogout) return;
  if (t.expiresAt - Date.now() > TOKEN_REFRESH_BUFFER_MS) {
    publishAuthState(api);
    return;
  }
  try {
    await acquireTokenSilent(api);
    publishAuthState(api);
  } catch (err) {
    getLogger().warn(`Proactive silent refresh failed: ${err}`);
  }
}

// ── Lifecycle ──

export async function activate(api: PluginAPI): Promise<void> {
  setLogger(api.log);
  const log = getLogger();
  log.info('msgraph plugin activating');

  tokenCache.loadPersisted(api);
  hadValidTokenSinceLogout = tokenCache.hasRefreshToken();

  api.state.replace(initialState() as unknown as Record<string, unknown>);
  publishAuthState(api);
  publishCredentialState(api);

  api.ui.registerPanelView({
    id: PANEL_ID,
    title: 'Teams',
    visible: true,
  });
  api.ui.registerNavigationItem({
    id: NAV_ID,
    visible: true,
    target: { type: 'panel', panelId: PANEL_ID },
  });
  api.ui.registerSettingsView({
    id: SETTINGS_ID,
    label: 'Teams',
  });

  api.onAction(`panel:${PANEL_ID}`, (action, data) => handlePanelAction(api, action, data));
  api.onAction('settings:SettingsView', (action, data) => handleSettingsAction(api, action, data));

  registerEnabledTools(api);

  unsubConfig = api.config.onChanged(() => {
    registerEnabledTools(api);
    publishCredentialState(api);
  });

  // Proactive refresh_token loop — keeps access token fresh without user interaction.
  refreshTimer = setInterval(() => void refreshTick(api), 60_000);

  // If we already have a refresh token, warm the access token and load chats.
  if (tokenCache.hasRefreshToken()) {
    void (async () => {
      try {
        await acquireTokenSilent(api);
        publishAuthState(api);
        await loadChats(api);
      } catch (err) {
        log.warn(`Startup silent refresh failed: ${err}`);
        publishAuthState(api);
      }
    })();
  }

  const prefs = getPreferences(api);
  pollTimer = setInterval(() => {
    if (tokenCache.hasRefreshToken()) void loadChats(api);
  }, Math.max(15, prefs.pollIntervalSeconds) * 1000);

  log.info('msgraph plugin activated');
}

export async function deactivate(): Promise<void> {
  if (pollTimer) clearInterval(pollTimer);
  if (refreshTimer) clearInterval(refreshTimer);
  if (unsubConfig) unsubConfig();
  pollTimer = null;
  refreshTimer = null;
  unsubConfig = null;
}
