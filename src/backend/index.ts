import {
  acquireTokenInteractive,
  acquireTokenSilent,
  ensureAccessToken,
  submitMfaCode,
  cancelMfaCode,
  clearFociTokens,
} from './auth.js';
import * as tokenCache from './token-cache.js';
import * as credentialStore from './credential-store.js';
import * as photoCache from './photo-cache.js';
import * as presenceCache from './presence-cache.js';
import * as hostedContentCache from './hosted-content-cache.js';
import * as mediaServer from './media-server.js';
import {
  mailInitialState,
  initMail,
  loadMailFolders,
  loadMailList,
  startMailPoll,
  stopMailPoll,
  disposeMail,
  handleMailAction,
  updateMailNavBadge,
} from './mail-backend.js';
import { GraphClient, normalizeChat, normalizeMessage } from './graph-client.js';
import {
  invokeMessageback,
  invokeTask,
  invokeExecute,
  invokeSearch,
  clearIC3State,
  getConsumptionHorizons,
  setForcedAvailability,
  setStatusNote,
  sendTyping,
  sendClearTyping,
  clearTypingThrottle,
  type UpsAvailability,
} from './ic3-client.js';
import { TrouterListener, type TrouterEvent } from './trouter.js';
import { buildMessageBody, withMessageRef, type PendingImage } from '../shared/markdown.js';
import { DiskCache } from './disk-cache.js';
import { buildMsgraphTools, ALL_TOOL_NAMES } from './tools.js';
import { setLogger, getLogger } from './logger-singleton.js';
import {
  PANEL_ID,
  NAV_ID,
  MAIL_PANEL_ID,
  MAIL_NAV_ID,
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
  CardActionPayload,
  TaskModuleState,
} from '../shared/types.js';
import { DEFAULT_TOOL_PERMISSIONS } from '../shared/types.js';

let pollTimer: ReturnType<typeof setInterval> | null = null;
let refreshTimer: ReturnType<typeof setInterval> | null = null;
let unsubConfig: (() => void) | null = null;
let trouter: TrouterListener | null = null;
const typingTimers = new Map<string, ReturnType<typeof setTimeout>>();
const chatRefreshDebounce = new Map<string, ReturnType<typeof setTimeout>>();
const TYPING_TTL_MS = 8000;
let hadValidTokenSinceLogout = false;
let messageLoadSeq = 0;
let remoteSearchSeq = 0;
let meJobTitle: string | null = null;
let paginationInFlight = false;
const MAX_CHAT_PAGES = 20;

type PeopleResults = Array<{ id: string; displayName: string; email: string | null }>;
const peopleSearchCache = new Map<string, { at: number; results: PeopleResults }>();
let peopleSearchDisk: DiskCache<PeopleResults> | null = null;
let threadCache: DiskCache<MsgraphPluginState['activeChatMessages']> | null = null;
const PEOPLE_CACHE_TTL_MS = 60 * 60_000;
const PEOPLE_CACHE_MAX = 400;

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
      objectId: null,
      jobTitle: null,
      minutesRemaining: null,
      autoLoginStatus: null,
    },
    mfa: { needed: false, type: null, approvalNumber: null },
    credentials: { hasCredentials: false, username: null, encryptionMethod: 'none' },
    photos: {},
    presence: {},
    hostedContents: {},
    chats: [],
    chatsNextLink: null,
    chatsFullyLoaded: false,
    loadingMoreChats: false,
    remoteSearch: null,
    peopleSearch: null,
    userCard: null,
    composerReplyTo: null,
    composerEditing: null,
    forwardTarget: null,
    taskModule: null,
    cardActionPending: null,
    activeChatId: null,
    activeChatMessages: [],
    activeChatMessagesNextLink: null,
    loadingOlderMessages: false,
    realtime: 'disabled',
    realtimeError: null,
    typing: {},
    readReceipts: {},
    loadingChats: false,
    loadingMessages: false,
    sendingMessage: false,
    error: null,
    ...mailInitialState(),
  };
}

function publishAuthState(api: PluginAPI): void {
  api.state.set('auth', {
    isAuthenticated: tokenCache.isTokenValid() || tokenCache.hasRefreshToken(),
    email: tokenCache.getEmail(),
    displayName: tokenCache.getDisplayName(),
    objectId: tokenCache.getObjectId(),
    jobTitle: meJobTitle,
    minutesRemaining: tokenCache.minutesRemaining(),
    autoLoginStatus: null,
  });
}

function publishCredentialState(api: PluginAPI): void {
  api.state.set('credentials', credentialStore.getCredentialStatus(api));
}

function updateNavBadge(api: PluginAPI): void {
  const chats = ((api.state.get() as Partial<MsgraphPluginState>).chats ?? []);
  const unread = chats.reduce((n, c) => n + (c.unread ? 1 : 0), 0);
  api.ui.registerNavigationItem({
    id: NAV_ID,
    label: 'Teams',
    icon: { lucide: 'message-square-more' },
    visible: true,
    priority: 0,
    badge: unread > 0 ? unread : undefined,
    target: { type: 'panel', panelId: PANEL_ID },
  });
}

// ── Client / auth wiring ──

async function ensureAuthenticated(api: PluginAPI, allowInteractive = true): Promise<GraphClient> {
  await ensureAccessToken(api, { allowInteractive });
  hadValidTokenSinceLogout = true;
  publishAuthState(api);
  return new GraphClient(api, allowInteractive);
}

// ── Tools ──

function registerEnabledTools(api: PluginAPI): void {
  const perms = getToolPermissions(api);
  const allTools = buildMsgraphTools({
    api,
    ensureAuthenticated: () => ensureAuthenticated(api, false),
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
    'edit-message': 'editMessage',
    'delete-message': 'deleteMessage',
    'forward-message': 'forwardMessage',
    'mark-chat-read': 'markChatRead',
    'get-presence': 'getPresence',
    'set-presence': 'setPresence',
    'set-status-message': 'setStatusMessage',
    'invoke-card-action': 'invokeCardAction',
    'list-mail': 'listMail',
    'get-mail': 'getMail',
    'search-mail': 'searchMail',
    'send-mail': 'sendMail',
    'reply-to-mail': 'replyToMail',
    'mark-mail': 'markMail',
    'archive-mail': 'archiveMail',
    'delete-mail': 'deleteMail',
    'create-group-chat': 'createGroupChat',
  };

  const enabled = allTools.filter((t) => perms[permMap[t.name]] !== false);

  api.tools.unregister(ALL_TOOL_NAMES);
  if (enabled.length > 0) api.tools.register(enabled);
  getLogger().info(
    `Registered ${enabled.length}/${allTools.length} msgraph tools (${enabled.map((t) => t.name).join(', ') || 'none'})`,
  );
}

async function refreshMeProfile(api: PluginAPI): Promise<void> {
  try {
    const client = await ensureAuthenticated(api, false);
    const me = await client.getMe();
    meJobTitle = me.jobTitle ?? null;
    publishAuthState(api);
  } catch (err) {
    getLogger().warn(`refreshMeProfile failed: ${err}`);
  }
}

// ── Data loaders ──

async function loadChats(api: PluginAPI, allowInteractive = false): Promise<void> {
  if (!tokenCache.hasRefreshToken() && !tokenCache.isTokenValid()) return;
  const session = tokenCache.currentSession();
  api.state.set('loadingChats', true);
  api.state.set('error', null);
  try {
    const client = await ensureAuthenticated(api, allowInteractive);
    const { chats: raw, nextLink } = await client.listChats({});
    if (session !== tokenCache.currentSession()) return;
    const myId = tokenCache.getObjectId();
    const page1 = raw.map((c) => normalizeChat(c, myId));
    const prev = ((api.state.get() as Partial<MsgraphPluginState>).chats ?? []);
    // Preserve already-loaded tail so a periodic refresh of page 1 doesn't shrink the list.
    const p1Ids = new Set(page1.map((c) => c.id));
    const merged = [...page1, ...prev.filter((c) => !p1Ids.has(c.id))];
    api.state.set('chats', merged);
    api.state.set('chatsNextLink', nextLink);
    updateNavBadge(api);
    const ids = new Set<string>();
    if (myId) ids.add(myId);
    for (const c of page1) for (const m of c.members) ids.add(m.id);
    photoCache.ensure(api, client, ids);
    presenceCache.refresh(api, client, ids);
    trouter?.subscribePresence(ids);

    // Background: page to the end (once) so member-based search is complete.
    if (
      nextLink &&
      !paginationInFlight &&
      !(api.state.get() as Partial<MsgraphPluginState>).chatsFullyLoaded
    ) {
      paginationInFlight = true;
      void (async () => {
        try {
          let link: string | null = nextLink;
          let pages = 1;
          while (link && pages < MAX_CHAT_PAGES) {
            if (session !== tokenCache.currentSession()) return;
            await new Promise((r) => setTimeout(r, 60));
            const { chats: more, nextLink: nl } = await client.listChatsPage(link);
            if (session !== tokenCache.currentSession()) return;
            const cur = ((api.state.get() as Partial<MsgraphPluginState>).chats ?? []);
            const seen = new Set(cur.map((c) => c.id));
            const add = more.map((c) => normalizeChat(c, myId)).filter((c) => !seen.has(c.id));
            if (add.length) api.state.set('chats', [...cur, ...add]);
            api.state.set('chatsNextLink', nl);
            link = nl;
            pages++;
          }
          if (session === tokenCache.currentSession()) {
            api.state.set('chatsFullyLoaded', !link);
            updateNavBadge(api);
          }
        } catch (err) {
          getLogger().warn(`background chat pagination stopped: ${err}`);
        } finally {
          paginationInFlight = false;
        }
      })();
    } else if (!nextLink) {
      api.state.set('chatsFullyLoaded', true);
    }
  } catch (err) {
    if (session === tokenCache.currentSession()) {
      api.state.set('error', err instanceof Error ? err.message : String(err));
    }
  } finally {
    if (session === tokenCache.currentSession()) api.state.set('loadingChats', false);
  }
}

async function loadMessages(api: PluginAPI, chatId: string): Promise<void> {
  const seq = ++messageLoadSeq;
  const current = (api.state.get() as Partial<MsgraphPluginState>).activeChatId;
  api.state.set('activeChatId', chatId);
  if (current !== chatId) {
    const snap = threadCache?.get(chatId);
    api.state.set('activeChatMessages', snap?.v ?? []);
  }
  api.state.set('loadingMessages', true);
  api.state.set('error', null);
  try {
    const client = await ensureAuthenticated(api);
    const { messages: msgs, nextLink } = await client.getChatMessages(chatId, 40);
    if (seq !== messageLoadSeq) return;
    const myId = tokenCache.getObjectId();
    const normalized = msgs
      .map((m) => normalizeMessage(m, myId))
      .filter((m) => !m.deleted)
      .reverse();
    api.state.set('activeChatMessages', normalized);
    api.state.set('activeChatMessagesNextLink', nextLink);
    threadCache?.set(chatId, normalized);
    const userIds = new Set<string>();
    const appIds = new Set<string>();
    const hosted = new Set<string>();
    for (const m of normalized) {
      if (m.fromId) (m.fromApp ? appIds : userIds).add(m.fromId);
      for (const u of m.hostedImages) hosted.add(u);
    }
    photoCache.ensure(api, client, userIds);
    photoCache.ensureApps(api, client, appIds);
    presenceCache.refresh(api, client, userIds);
    trouter?.subscribePresence(userIds);
    hostedContentCache.ensure(api, client, hosted);
    void loadReadReceipts(api, chatId);
  } catch (err) {
    if (seq !== messageLoadSeq) return;
    api.state.set('error', err instanceof Error ? err.message : String(err));
  } finally {
    if (seq === messageLoadSeq) api.state.set('loadingMessages', false);
  }
}

async function loadReadReceipts(api: PluginAPI, chatId: string): Promise<void> {
  try {
    const horizons = await getConsumptionHorizons(api, chatId);
    const cur = ((api.state.get() as Partial<MsgraphPluginState>).readReceipts ?? {});
    api.state.set('readReceipts', { ...cur, [chatId]: horizons });
  } catch (err) {
    getLogger().warn(`consumptionhorizons(${chatId}) failed: ${err}`);
  }
}

// ── Real-time (Trouter) ──

function clearTyping(api: PluginAPI, chatId: string): void {
  const t = typingTimers.get(chatId);
  if (t) { clearTimeout(t); typingTimers.delete(chatId); }
  const cur = ((api.state.get() as Partial<MsgraphPluginState>).typing ?? {});
  if (cur[chatId]) {
    const { [chatId]: _drop, ...rest } = cur;
    api.state.set('typing', rest);
  }
}

/** Fetch a single message and merge it into activeChatMessages (much cheaper than a full page reload). */
async function mergeMessage(api: PluginAPI, chatId: string, messageId: string): Promise<void> {
  const st = () => api.state.get() as Partial<MsgraphPluginState>;
  if (st().activeChatId !== chatId || !messageId) return;
  try {
    const client = await ensureAuthenticated(api, false);
    const raw = await client.getMessage(chatId, messageId);
    if (st().activeChatId !== chatId) return;
    const myId = tokenCache.getObjectId();
    const norm = normalizeMessage(raw, myId);
    const cur = st().activeChatMessages ?? [];
    let next: typeof cur;
    const idx = cur.findIndex((m) => m.id === norm.id);
    if (norm.deleted) {
      next = idx >= 0 ? [...cur.slice(0, idx), ...cur.slice(idx + 1)] : cur;
    } else if (idx >= 0) {
      next = [...cur]; next[idx] = norm;
    } else {
      // Insert in chronological order (list is oldest→newest).
      const nId = Number(norm.id);
      let pos = cur.length;
      while (pos > 0 && Number(cur[pos - 1].id) > nId) pos--;
      next = [...cur.slice(0, pos), norm, ...cur.slice(pos)];
    }
    api.state.set('activeChatMessages', next);
    threadCache?.set(chatId, next);
    if (norm.fromId) {
      if (norm.fromApp) photoCache.ensureApps(api, client, [norm.fromId]);
      else {
        photoCache.ensure(api, client, [norm.fromId]);
        trouter?.subscribePresence([norm.fromId]);
      }
    }
    if (norm.hostedImages.length) hostedContentCache.ensure(api, client, norm.hostedImages);
  } catch (err) {
    getLogger().warn(`mergeMessage(${chatId},${messageId}) failed: ${err}; falling back to full reload`);
    if (st().activeChatId === chatId) void loadMessages(api, chatId);
  }
}

function scheduleMessageMerge(api: PluginAPI, chatId: string, messageId: string): void {
  const key = `${chatId}|${messageId}`;
  if (chatRefreshDebounce.has(key)) return;
  chatRefreshDebounce.set(key, setTimeout(() => {
    chatRefreshDebounce.delete(key);
    void mergeMessage(api, chatId, messageId);
  }, 150));
}

function chatDisplayName(api: PluginAPI, chatId: string): string | null {
  const c = ((api.state.get() as Partial<MsgraphPluginState>).chats ?? []).find((x) => x.id === chatId);
  if (!c) return null;
  return c.topic ?? (c.members.map((m) => m.displayName).join(', ') || null);
}

function busEmit(api: PluginAPI, event: string, payload: unknown): void {
  try { api.events?.emit(event, payload); } catch (err) { getLogger().warn(`events.emit(${event}) failed: ${err}`); }
}

function handleTrouterEvent(api: PluginAPI, ev: TrouterEvent): void {
  const st = api.state.get() as Partial<MsgraphPluginState>;
  const myId = tokenCache.getObjectId();
  switch (ev.kind) {
    case 'connected':
      api.state.set('realtime', 'connected');
      api.state.set('realtimeError', null);
      return;
    case 'disconnected':
      api.state.set('realtime', ev.willRetry ? 'connecting' : 'disconnected');
      return;
    case 'error':
      api.state.set('realtimeError', ev.message);
      return;
    case 'typing': {
      const t = typingTimers.get(ev.chatId);
      if (t) clearTimeout(t);
      const next = {
        ...(st.typing ?? {}),
        [ev.chatId]: {
          chatId: ev.chatId,
          userId: ev.fromUserId,
          displayName: ev.fromName,
          until: Date.now() + TYPING_TTL_MS,
        },
      };
      api.state.set('typing', next);
      typingTimers.set(ev.chatId, setTimeout(() => clearTyping(api, ev.chatId), TYPING_TTL_MS));
      busEmit(api, 'typing-started', {
        chatId: ev.chatId,
        chatTitle: chatDisplayName(api, ev.chatId),
        from: { id: ev.fromUserId, displayName: ev.fromName },
      });
      return;
    }
    case 'clearTyping': {
      clearTyping(api, ev.chatId);
      return;
    }
    case 'readReceipt': {
      const cur = st.readReceipts ?? {};
      const forChat = { ...(cur[ev.chatId] ?? {}), [ev.userId]: ev.lastReadMessageId };
      api.state.set('readReceipts', { ...cur, [ev.chatId]: forChat });
      if (ev.userId !== myId) {
        busEmit(api, 'message-read', {
          chatId: ev.chatId,
          chatTitle: chatDisplayName(api, ev.chatId),
          reader: {
            id: ev.userId,
            displayName: (st.chats ?? []).find((c) => c.id === ev.chatId)?.members.find((m) => m.id === ev.userId)?.displayName ?? null,
          },
          lastReadMessageId: ev.lastReadMessageId,
        });
      }
      return;
    }
    case 'message': {
      clearTyping(api, ev.chatId);
      if (!ev.own) {
        busEmit(api, 'message-received', {
          chatId: ev.chatId,
          messageId: ev.messageId,
          chatTitle: chatDisplayName(api, ev.chatId),
          from: { id: ev.fromUserId, displayName: ev.fromName },
          preview: ev.preview,
          isActiveChat: st.activeChatId === ev.chatId,
        });
      }
      if (st.activeChatId === ev.chatId) {
        scheduleMessageMerge(api, ev.chatId, ev.messageId);
      } else if (!ev.own) {
        const chats = (st.chats ?? []).map((c) =>
          c.id === ev.chatId ? { ...c, unread: true } : c,
        );
        const idx = chats.findIndex((c) => c.id === ev.chatId);
        if (idx > 0) chats.unshift(...chats.splice(idx, 1));
        api.state.set('chats', chats);
        updateNavBadge(api);
        if (idx < 0) void loadChats(api, false);
      }
      return;
    }
    case 'messageUpdate': {
      if (st.activeChatId === ev.chatId) scheduleMessageMerge(api, ev.chatId, ev.messageId);
      if (ev.reaction && ev.reaction.userId !== myId) {
        busEmit(api, 'reaction-added', {
          chatId: ev.chatId,
          messageId: ev.messageId,
          chatTitle: chatDisplayName(api, ev.chatId),
          reaction: ev.reaction.type,
          from: {
            id: ev.reaction.userId,
            displayName: (st.chats ?? []).find((c) => c.id === ev.chatId)?.members.find((m) => m.id === ev.reaction!.userId)?.displayName ?? null,
          },
        });
      }
      return;
    }
    case 'conversationUpdate':
      // Sidebar preview / lastUpdated changed — cheap to just refresh page-1.
      void loadChats(api, false);
      return;
    case 'presence': {
      const cur = st.presence ?? {};
      const prev = cur[ev.userId];
      api.state.set('presence', {
        ...cur,
        [ev.userId]: { ...(prev ?? {}), availability: ev.availability, activity: ev.activity },
      });
      if (!prev || prev.availability !== ev.availability || prev.activity !== ev.activity) {
        busEmit(api, 'presence-changed', {
          userId: ev.userId,
          availability: ev.availability,
          activity: ev.activity,
          previous: prev ? { availability: prev.availability, activity: prev.activity } : null,
        });
      }
      return;
    }
  }
}

function startRealtime(api: PluginAPI): void {
  if (trouter) return;
  api.state.set('realtime', 'connecting');
  trouter = new TrouterListener(api, tokenCache.getObjectId(), (ev) => handleTrouterEvent(api, ev));
  const seed = Object.keys((api.state.get() as Partial<MsgraphPluginState>).presence ?? {});
  if (seed.length) trouter.subscribePresence(seed);
  trouter.start();
}

function stopRealtime(api: PluginAPI): void {
  trouter?.stop();
  trouter = null;
  for (const t of typingTimers.values()) clearTimeout(t);
  typingTimers.clear();
  for (const t of chatRefreshDebounce.values()) clearTimeout(t);
  chatRefreshDebounce.clear();
  api.state.set('realtime', 'disabled');
  api.state.set('typing', {});
}

// ── Actions ──

async function handlePanelAction(api: PluginAPI, action: string, data?: unknown): Promise<void> {
  const log = getLogger();
  try {
    switch (action) {
      case 'login': {
        api.state.set('error', null);
        // A forced login may switch accounts. Rotate the session so all
        // secondary caches (FOCI tokens, IC3 region, trouter) key off the new
        // identity instead of leaking the previous account's tokens.
        stopRealtime(api);
        stopMailPoll();
        clearFociTokens();
        clearIC3State();
        clearTypingThrottle();
        tokenCache.invalidateSession();
        await acquireTokenInteractive(api, { forceRefresh: true });
        hadValidTokenSinceLogout = true;
        publishAuthState(api);
        startRealtime(api);
        startMailPoll(api);
        await loadChats(api, true);
        void loadMailFolders(api).then(() => loadMailList(api, 'inbox'));
        void refreshMeProfile(api);
        break;
      }
      case 'logout': {
        tokenCache.invalidateSession();
        messageLoadSeq++;
        remoteSearchSeq++;
        meJobTitle = null;
        cancelMfaCode();
        stopRealtime(api);
        stopMailPoll();
        clearFociTokens();
        clearIC3State();
        clearTypingThrottle();
        tokenCache.clear();
        tokenCache.persist(api);
        photoCache.clear();
        presenceCache.clear();
        hostedContentCache.clear();
        peopleSearchCache.clear();
        peopleSearchDisk?.clear();
        hadValidTokenSinceLogout = false;
        api.state.replace(initialState() as unknown as Record<string, unknown>);
        publishCredentialState(api);
        updateNavBadge(api);
        updateMailNavBadge(api);
        break;
      }
      case 'refresh-chats': {
        await loadChats(api, true);
        break;
      }
      case 'load-more-chats': {
        const st = api.state.get() as Partial<MsgraphPluginState>;
        const link = st.chatsNextLink;
        if (!link || st.loadingMoreChats) break;
        api.state.set('loadingMoreChats', true);
        try {
          const client = await ensureAuthenticated(api, false);
          const session = tokenCache.currentSession();
          const { chats: raw, nextLink } = await client.listChatsPage(link);
          if (session !== tokenCache.currentSession()) break;
          const myId = tokenCache.getObjectId();
          const existing = (api.state.get() as Partial<MsgraphPluginState>).chats ?? [];
          const seen = new Set(existing.map((c) => c.id));
          const more = raw.map((c) => normalizeChat(c, myId)).filter((c) => !seen.has(c.id));
          api.state.set('chats', [...existing, ...more]);
          api.state.set('chatsNextLink', nextLink);
          updateNavBadge(api);
          const ids = new Set<string>();
          for (const c of more) for (const m of c.members) ids.add(m.id);
          photoCache.ensure(api, client, ids);
          presenceCache.refresh(api, client, ids);
    trouter?.subscribePresence(ids);
        } finally {
          api.state.set('loadingMoreChats', false);
        }
        break;
      }
      case 'search-chats': {
        const { query, mode = 'people' } = data as { query: string; mode?: 'people' | 'content' };
        const q = query.trim();
        const seq = ++remoteSearchSeq;
        if (q.length < 2) {
          api.state.set('remoteSearch', null);
          break;
        }
        api.state.set('remoteSearch', { query: q, loading: true, results: [] });
        try {
          const client = await ensureAuthenticated(api, false);
          const myId = tokenCache.getObjectId();
          const isChatId = (id: string) => /@(unq\.gbl\.spaces|thread\.v2)$/i.test(id);
          const chatIdsToFetch = async (kql: string, top: number) => {
            const hits = await client.searchMessages(kql, top);
            return [...new Set(hits.map((h) => h.chatId).filter((id): id is string => !!id && isChatId(id)))];
          };

          let rawChats: import('../shared/types.js').GraphChat[] = [];

          if (mode === 'content') {
            const ids = (await chatIdsToFetch(q, 50).catch((e) => {
              log.warn(`content search failed: ${e}`);
              return [] as string[];
            })).slice(0, 12);
            const fetched = await Promise.allSettled(ids.map((id) => client.getChat(id)));
            rawChats = fetched.flatMap((p) => (p.status === 'fulfilled' ? [p.value] : []));
          } else {
            // People mode: resolve people → probe existing 1:1s. Group-chat membership
            // matches come from the local filter over the (background-)fully-loaded list.
            let users: import('../shared/types.js').GraphUser[] = [];
            if (q.includes('@')) {
              const u = await client.getUserByEmail(q);
              if (u) users = [u];
            }
            if (users.length === 0) {
              try {
                users = await client.searchPeople(q, 10);
              } catch (err) {
                log.warn(`searchPeople failed (${err}); falling back to /users`);
              }
            }
            if (users.length === 0) users = await client.findUsers(q, 25);
            if (seq !== remoteSearchSeq) return;
            const rs = await Promise.allSettled(
              users.filter((u) => u.id && u.id !== myId).map((u) => client.probeOneOnOne(u.id)),
            );
            rawChats = rs.flatMap((p) => (p.status === 'fulfilled' && p.value ? [p.value] : []));
          }

          if (seq !== remoteSearchSeq) return;
          const seen = new Set<string>();
          const results = rawChats
            .map((c) => normalizeChat(c, myId))
            .filter((c) => (seen.has(c.id) ? false : (seen.add(c.id), true)))
            .sort((a, b) => (b.lastUpdated ?? '').localeCompare(a.lastUpdated ?? ''));
          api.state.set('remoteSearch', { query: q, loading: false, results });
          const ids = new Set<string>();
          for (const c of results) for (const m of c.members) ids.add(m.id);
          photoCache.ensure(api, client, ids);
          presenceCache.refresh(api, client, ids);
    trouter?.subscribePresence(ids);
        } catch (err) {
          if (seq === remoteSearchSeq) {
            const msg = err instanceof Error ? err.message : String(err);
            api.state.set('remoteSearch', { query: q, loading: false, results: [], error: msg });
            log.warn(`search-chats failed: ${msg}`);
          }
        }
        break;
      }
      case 'clear-search': {
        remoteSearchSeq++;
        api.state.set('remoteSearch', null);
        break;
      }
      case 'load-older-messages': {
        const st0 = api.state.get() as Partial<MsgraphPluginState>;
        const link = st0.activeChatMessagesNextLink;
        const chatId = st0.activeChatId;
        if (!link || !chatId || st0.loadingOlderMessages) break;
        const seq = messageLoadSeq;
        api.state.set('loadingOlderMessages', true);
        try {
          const client = await ensureAuthenticated(api, false);
          const { messages: page, nextLink } = await client.getChatMessagesPage(link);
          if (seq !== messageLoadSeq) break;
          const myId = tokenCache.getObjectId();
          const cur = (api.state.get() as Partial<MsgraphPluginState>).activeChatMessages ?? [];
          const have = new Set(cur.map((m) => m.id));
          const older = page
            .map((m) => normalizeMessage(m, myId))
            .filter((m) => !m.deleted && !have.has(m.id))
            .reverse();
          const merged = [...older, ...cur];
          api.state.set('activeChatMessages', merged);
          api.state.set('activeChatMessagesNextLink', nextLink);
          threadCache?.set(chatId, merged);
          const userIds = new Set<string>();
          const appIds = new Set<string>();
          const hosted = new Set<string>();
          for (const m of older) {
            if (m.fromId) (m.fromApp ? appIds : userIds).add(m.fromId);
            for (const u of m.hostedImages) hosted.add(u);
          }
          photoCache.ensure(api, client, userIds);
          photoCache.ensureApps(api, client, appIds);
          hostedContentCache.ensure(api, client, hosted);
        } finally {
          if (seq === messageLoadSeq) api.state.set('loadingOlderMessages', false);
        }
        break;
      }
      case 'select-chat': {
        const { chatId } = data as { chatId: string };
        api.state.set('composerReplyTo', null);
        api.state.set('composerEditing', null);
        await loadMessages(api, chatId);
        // Optimistically clear the unread dot, then tell Graph.
        const chats = ((api.state.get() as Partial<MsgraphPluginState>).chats ?? []).map((c) =>
          c.id === chatId ? { ...c, unread: false } : c,
        );
        api.state.set('chats', chats);
        updateNavBadge(api);
        try {
          const client = await ensureAuthenticated(api, false);
          await client.markChatRead(chatId);
        } catch (err) {
          log.warn(`markChatRead failed: ${err}`);
        }
        break;
      }
      case 'send-message': {
        const { chatId, text, images, payload } = data as {
          chatId: string;
          text?: string;
          images?: PendingImage[];
          payload?: Record<string, unknown>;
        };
        api.state.set('sendingMessage', true);
        try {
          const client = await ensureAuthenticated(api);
          let body = (payload ?? buildMessageBody(text ?? '', images ?? [])) as {
            body: { contentType: 'text' | 'html'; content: string };
            attachments?: unknown[];
          };
          const rt = (api.state.get() as Partial<MsgraphPluginState>).composerReplyTo;
          if (rt) {
            body = withMessageRef(body, {
              contentType: 'messageReference',
              id: rt.messageId,
              messageId: rt.messageId,
              messagePreview: rt.text,
              messageSender: rt.senderName ? { user: { displayName: rt.senderName } } : null,
            });
            api.state.set('composerReplyTo', null);
          }
          await client.sendMessageRaw(chatId, body);
          void sendClearTyping(api, chatId).catch(() => {});
          if ((api.state.get() as Partial<MsgraphPluginState>).activeChatId === chatId) {
            await loadMessages(api, chatId);
          }
        } finally {
          api.state.set('sendingMessage', false);
        }
        break;
      }
      case 'compose-new-chat': {
        const { recipients, topic, text, images } = data as {
          recipients: Array<{ id: string; displayName?: string }>;
          topic?: string;
          text?: string;
          images?: PendingImage[];
        };
        const client = await ensureAuthenticated(api);
        const myId = tokenCache.getObjectId();
        const ids = [...new Set(recipients.map((r) => r.id).filter((id) => id && id !== myId))];
        if (ids.length === 0) throw new Error('Select at least one recipient');

        let chatId: string;
        if (ids.length === 1) {
          const chat = await client.getOrCreateOneOnOne(ids[0]);
          chatId = chat.id;
        } else {
          // Reuse an existing untitled group with exactly these members if one is already loaded.
          const want = new Set(ids);
          const existing = ((api.state.get() as Partial<MsgraphPluginState>).chats ?? []).find(
            (c) =>
              c.type === 'group' &&
              !c.topic &&
              c.members.length === want.size &&
              c.members.every((m) => want.has(m.id)),
          );
          if (existing && !topic) {
            chatId = existing.id;
          } else {
            const chat = await client.createGroupChat(topic || null, ids);
            chatId = chat.id;
          }
        }

        if ((text && text.trim()) || (images && images.length)) {
          const payload = buildMessageBody(text ?? '', images ?? []);
          await client.sendMessageRaw(chatId, payload);
        }
        await loadChats(api, false);
        await loadMessages(api, chatId);
        break;
      }
      case 'load-user-card': {
        const { userId } = data as { userId: string };
        api.state.set('userCard', { userId, loading: true, displayName: null, email: null, jobTitle: null });
        try {
          const client = await ensureAuthenticated(api, false);
          const u = await client.getUser(userId);
          if ((api.state.get() as Partial<MsgraphPluginState>).userCard?.userId !== userId) break;
          api.state.set('userCard', {
            userId,
            loading: false,
            displayName: u.displayName ?? null,
            email: u.mail ?? u.userPrincipalName ?? null,
            jobTitle: u.jobTitle ?? null,
          });
          photoCache.ensure(api, client, [userId]);
          presenceCache.refresh(api, client, [userId]);
        } catch (err) {
          if ((api.state.get() as Partial<MsgraphPluginState>).userCard?.userId === userId) {
            api.state.set('userCard', {
              userId, loading: false, displayName: null, email: null, jobTitle: null,
              error: err instanceof Error ? err.message : String(err),
            });
          }
        }
        break;
      }
      case 'close-user-card': {
        api.state.set('userCard', null);
        break;
      }
      case 'open-chat-with': {
        const { userId } = data as { userId: string };
        const client = await ensureAuthenticated(api);
        const chat = await client.getOrCreateOneOnOne(userId);
        api.state.set('userCard', null);
        await loadMessages(api, chat.id);
        void loadChats(api, false);
        break;
      }
      case 'search-people': {
        const { query } = data as { query: string };
        const q = query.trim();
        if (q.length < 1) {
          api.state.set('peopleSearch', { query: q, loading: false, results: [] });
          break;
        }
        const cacheKey = q.toLowerCase();
        const cached = peopleSearchCache.get(cacheKey);
        if (cached && Date.now() - cached.at < PEOPLE_CACHE_TTL_MS) {
          api.state.set('peopleSearch', { query: q, loading: false, results: cached.results });
          break;
        }
        api.state.set('peopleSearch', { query: q, loading: true, results: [] });
        try {
          const client = await ensureAuthenticated(api, false);
          let results = q.includes('@')
            ? await client.getUserByEmail(q).then((u) => (u ? [u] : []))
            : await client.searchPeople(q, 8);
          if (results.length === 0) results = await client.findUsers(q, 8);
          const mapped = results.map((u) => ({
            id: u.id,
            displayName: u.displayName ?? u.userPrincipalName ?? u.id,
            email: u.mail ?? u.userPrincipalName ?? null,
          }));
          if (peopleSearchCache.size >= PEOPLE_CACHE_MAX) {
            const oldest = peopleSearchCache.keys().next().value;
            if (oldest) peopleSearchCache.delete(oldest);
          }
          peopleSearchCache.set(cacheKey, { at: Date.now(), results: mapped });
          peopleSearchDisk?.set(cacheKey, mapped);
          api.state.set('peopleSearch', { query: q, loading: false, results: mapped });
          photoCache.ensure(api, client, results.map((u) => u.id));
        } catch (err) {
          api.state.set('peopleSearch', {
            query: q,
            loading: false,
            results: [],
            error: err instanceof Error ? err.message : String(err),
          });
        }
        break;
      }
      case 'set-reply-to': {
        api.state.set('composerReplyTo', data as MsgraphPluginState['composerReplyTo']);
        api.state.set('composerEditing', null);
        break;
      }
      case 'start-edit': {
        const { messageId } = data as { messageId: string };
        const st = api.state.get() as Partial<MsgraphPluginState>;
        const chatId = st.activeChatId;
        if (!chatId) break;
        const client = await ensureAuthenticated(api, false);
        // Fetch fresh so we have the raw attachments array to preserve on save.
        const raw = await client.getMessage(chatId, messageId);
        const norm = normalizeMessage(raw, tokenCache.getObjectId());
        api.state.set('composerReplyTo', null);
        api.state.set('composerEditing', {
          chatId,
          messageId,
          segments: norm.segments,
          attachments: raw.attachments ?? [],
        });
        // Make sure any hosted images are in the cache so the editor can display them.
        hostedContentCache.ensure(api, client, norm.hostedImages);
        break;
      }
      case 'cancel-edit': {
        api.state.set('composerEditing', null);
        break;
      }
      case 'save-edit': {
        const { payload } = data as {
          payload: {
            body: { contentType: 'text' | 'html'; content: string };
            hostedContents?: unknown[];
            mentions?: unknown[];
          };
        };
        const ed = (api.state.get() as Partial<MsgraphPluginState>).composerEditing;
        if (!ed) break;
        const client = await ensureAuthenticated(api);
        // Re-attach original attachments (cards, files, message refs) with their <attachment> markers.
        const markers = (ed.attachments ?? [])
          .filter((a) => a.id)
          .map((a) => `<attachment id="${a.id}"></attachment>`)
          .join('');
        let content = payload.body.content;
        let contentType: 'text' | 'html' = payload.body.contentType;
        if (markers) {
          if (contentType === 'text') {
            content = content.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
            contentType = 'html';
          }
          content += markers;
        }
        const patch: Record<string, unknown> = {
          body: { contentType, content },
        };
        if (ed.attachments?.length) patch.attachments = ed.attachments;
        if (payload.mentions) patch.mentions = payload.mentions;
        if (payload.hostedContents) patch.hostedContents = payload.hostedContents;
        await client.editMessage(ed.chatId, ed.messageId, patch as { body: { contentType: 'text' | 'html'; content: string } });
        api.state.set('composerEditing', null);
        if ((api.state.get() as Partial<MsgraphPluginState>).activeChatId === ed.chatId) {
          await loadMessages(api, ed.chatId);
        }
        break;
      }
      case 'set-forward-target': {
        api.state.set('forwardTarget', data as MsgraphPluginState['forwardTarget']);
        break;
      }
      case 'edit-message': {
        const { chatId, messageId, text } = data as { chatId: string; messageId: string; text: string };
        const client = await ensureAuthenticated(api);
        const p = buildMessageBody(text);
        await client.editMessage(chatId, messageId, { body: p.body });
        if ((api.state.get() as Partial<MsgraphPluginState>).activeChatId === chatId) {
          await loadMessages(api, chatId);
        }
        break;
      }
      case 'delete-message': {
        const { chatId, messageId } = data as { chatId: string; messageId: string };
        const client = await ensureAuthenticated(api);
        await client.deleteMessage(chatId, messageId);
        if ((api.state.get() as Partial<MsgraphPluginState>).activeChatId === chatId) {
          await loadMessages(api, chatId);
        }
        break;
      }
      case 'forward-message': {
        const { source, recipients } = data as {
          source: NonNullable<MsgraphPluginState['forwardTarget']>;
          recipients: Array<{ id: string; displayName?: string }>;
        };
        const client = await ensureAuthenticated(api);
        const myId = tokenCache.getObjectId();
        const ids = [...new Set(recipients.map((r) => r.id).filter((id) => id && id !== myId))];
        if (ids.length === 0) throw new Error('Select at least one recipient');
        let targetChatId: string;
        if (ids.length === 1) {
          targetChatId = (await client.getOrCreateOneOnOne(ids[0])).id;
        } else {
          targetChatId = (await client.createGroupChat(null, ids)).id;
        }
        const body = withMessageRef(
          { body: { contentType: 'html', content: '' } },
          {
            contentType: 'forwardedMessageReference',
            id: source.messageId,
            messageId: source.messageId,
            messagePreview: source.text,
            messageSender: source.senderName ? { user: { displayName: source.senderName } } : null,
          },
        );
        await client.sendMessageRaw(targetChatId, body);
        api.state.set('forwardTarget', null);
        await loadChats(api, false);
        await loadMessages(api, targetChatId);
        break;
      }
      case 'react-to-message': {
        const { chatId, messageId, reactionType, remove } = data as {
          chatId: string; messageId: string; reactionType: string; remove?: boolean;
        };
        const client = await ensureAuthenticated(api);
        if (remove) await client.unsetReaction(chatId, messageId, reactionType);
        else await client.setReaction(chatId, messageId, reactionType);
        if ((api.state.get() as Partial<MsgraphPluginState>).activeChatId === chatId) {
          await loadMessages(api, chatId);
        }
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
      case 'invoke-card-action': {
        const p = data as CardActionPayload;
        const ctx = { botId: p.botId, chatId: p.chatId, messageId: p.messageId };
        api.state.set('cardActionPending', p.messageId);
        api.state.set('error', null);
        try {
          await ensureAuthenticated(api, false);
          if (p.kind === 'task/fetch') {
            const r = await invokeTask(api, ctx, 'task/fetch', { ...p.data, type: 'task/fetch' });
            if (r?.card || r?.url) {
              api.state.set('taskModule', {
                botId: p.botId, chatId: p.chatId, messageId: p.messageId,
                title: r.title ?? p.title ?? null,
                card: r.card ?? null, url: r.url ?? null,
                width: r.width, height: r.height, submitting: false, error: null,
                choiceSearch: null,
              } satisfies TaskModuleState);
            } else {
              throw new Error('Bot returned an empty dialog');
            }
          } else if (p.kind === 'execute') {
            await invokeExecute(api, ctx, p.verb ?? null, p.data);
            if ((api.state.get() as Partial<MsgraphPluginState>).activeChatId === p.chatId) {
              await loadMessages(api, p.chatId);
            }
          } else {
            await invokeMessageback(api, ctx, p.data);
            // Bot replies asynchronously; give it a beat, then refresh the thread.
            await new Promise((r) => setTimeout(r, 900));
            if ((api.state.get() as Partial<MsgraphPluginState>).activeChatId === p.chatId) {
              await loadMessages(api, p.chatId);
            }
          }
        } catch (err) {
          throw new Error(
            `Card action failed (${err instanceof Error ? err.message : String(err)}). ` +
            `You can open this message in Teams instead.`,
          );
        } finally {
          if ((api.state.get() as Partial<MsgraphPluginState>).cardActionPending === p.messageId) {
            api.state.set('cardActionPending', null);
          }
        }
        break;
      }
      case 'submit-task-module': {
        const { data: form } = data as { data: Record<string, unknown> };
        const tm = (api.state.get() as Partial<MsgraphPluginState>).taskModule;
        if (!tm) break;
        api.state.set('taskModule', { ...tm, submitting: true, error: null });
        try {
          const r = await invokeTask(
            api,
            { botId: tm.botId, chatId: tm.chatId, messageId: tm.messageId },
            'task/submit',
            form,
          );
          if (r?.type === 'continue' && (r.card || r.url)) {
            api.state.set('taskModule', {
              ...tm, title: r.title ?? tm.title,
              card: r.card ?? null, url: r.url ?? null,
              width: r.width ?? tm.width, height: r.height ?? tm.height,
              submitting: false, error: null, choiceSearch: null,
            } satisfies TaskModuleState);
          } else {
            api.state.set('taskModule', null);
            if (r?.url) await handlePanelAction(api, 'open-external', { url: r.url });
            if ((api.state.get() as Partial<MsgraphPluginState>).activeChatId === tm.chatId) {
              await loadMessages(api, tm.chatId);
            }
          }
        } catch (err) {
          const cur = (api.state.get() as Partial<MsgraphPluginState>).taskModule;
          if (cur) {
            api.state.set('taskModule', {
              ...cur, submitting: false,
              error: err instanceof Error ? err.message : String(err),
            });
          }
        }
        break;
      }
      case 'close-task-module': {
        api.state.set('taskModule', null);
        break;
      }
      case 'search-task-choices': {
        const { reqId, dataset, query } = data as { reqId: number; dataset: string; query: string };
        const tm = (api.state.get() as Partial<MsgraphPluginState>).taskModule;
        if (!tm) break;
        api.state.set('taskModule', {
          ...tm,
          choiceSearch: { reqId, query, loading: true, results: [] },
        });
        try {
          const results = await invokeSearch(
            api,
            { botId: tm.botId, chatId: tm.chatId, messageId: tm.messageId },
            dataset,
            query,
          );
          const cur = (api.state.get() as Partial<MsgraphPluginState>).taskModule;
          if (cur && cur.choiceSearch?.reqId === reqId) {
            api.state.set('taskModule', { ...cur, choiceSearch: { reqId, query, loading: false, results } });
          }
        } catch (err) {
          const cur = (api.state.get() as Partial<MsgraphPluginState>).taskModule;
          if (cur && cur.choiceSearch?.reqId === reqId) {
            api.state.set('taskModule', {
              ...cur,
              choiceSearch: {
                reqId, query, loading: false, results: [],
                error: err instanceof Error ? err.message : String(err),
              },
            });
          }
        }
        break;
      }
      case 'set-presence': {
        const { availability } = data as { availability: UpsAvailability | null };
        await setForcedAvailability(api, availability);
        // Reflect immediately for the self dot; presenceCache will catch up on next poll.
        const me = tokenCache.getObjectId();
        if (me && availability) {
          const cur = ((api.state.get() as Partial<MsgraphPluginState>).presence ?? {});
          api.state.set('presence', {
            ...cur,
            [me]: { ...(cur[me] ?? { activity: availability }), availability, activity: availability },
          });
        }
        break;
      }
      case 'set-status-message': {
        const { message, pinned } = data as { message: string; pinned?: boolean };
        await setStatusNote(api, message, { pinned });
        const me = tokenCache.getObjectId();
        if (me) {
          const cur = ((api.state.get() as Partial<MsgraphPluginState>).presence ?? {});
          api.state.set('presence', {
            ...cur,
            [me]: { ...(cur[me] ?? { availability: 'Available', activity: 'Available' }), statusMessage: message },
          });
        }
        break;
      }
      case 'typing': {
        const { chatId } = data as { chatId: string };
        void sendTyping(api, chatId).catch((e) => log.warn(`sendTyping: ${e}`));
        break;
      }
      case 'navigate-panel': {
        const { view } = data as { view: 'teams' | 'mail' };
        api.navigation.open({ type: 'panel', panelId: view === 'mail' ? MAIL_PANEL_ID : PANEL_ID });
        break;
      }
      case 'open-in-teams': {
        const { url } = data as { url: string };
        let ok = false;
        try {
          const u = new URL(url);
          ok = u.protocol === 'https:' && (u.hostname === 'teams.microsoft.com' || u.hostname.endsWith('.teams.microsoft.com'));
        } catch { /* invalid */ }
        if (!ok) throw new Error('Refusing to open non-Teams URL');
        await api.shell.openExternal(url);
        break;
      }
      case 'open-external': {
        const { url } = data as { url: string };
        let ok = false;
        try {
          const u = new URL(url);
          ok = u.protocol === 'https:' || u.protocol === 'http:' || u.protocol === 'mailto:';
        } catch { /* invalid */ }
        if (!ok) throw new Error(`Refusing to open URL with unsupported scheme: ${url}`);
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
      case 'clear-cache': {
        photoCache.clear();
        hostedContentCache.clear();
        presenceCache.clear();
        peopleSearchCache.clear();
        peopleSearchDisk?.clear();
        threadCache?.clear();
        api.state.set('photos', {});
        api.state.set('presence', {});
        api.state.set('hostedContents', {});
        log.info('caches cleared');
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

  mediaServer.start();
  photoCache.init(api);
  initMail(api, (allowInteractive = false) => ensureAuthenticated(api, allowInteractive));
  peopleSearchDisk = new DiskCache<PeopleResults>(
    api.pluginName,
    'people-search',
    { hardTtlMs: 24 * 60 * 60_000, maxEntries: PEOPLE_CACHE_MAX },
    () => {
      for (const [k, e] of peopleSearchDisk!.entries()) {
        if (!peopleSearchCache.has(k)) peopleSearchCache.set(k, { at: e.at, results: e.v });
      }
    },
  );
  threadCache = new DiskCache(api.pluginName, 'threads', {
    hardTtlMs: 7 * 24 * 60 * 60_000,
    maxEntries: 100,
  });
  publishAuthState(api);
  publishCredentialState(api);

  api.ui.registerPanelView({
    id: PANEL_ID,
    title: 'Teams',
    visible: true,
    props: { view: 'teams' },
  });
  api.ui.registerPanelView({
    id: MAIL_PANEL_ID,
    title: 'Outlook',
    visible: true,
    width: 'wide',
    props: { view: 'mail' },
  });
  api.ui.registerNavigationItem({
    id: NAV_ID,
    label: 'Teams',
    icon: { lucide: 'message-square-more' },
    visible: true,
    priority: 0,
    target: { type: 'panel', panelId: PANEL_ID },
  });
  api.ui.registerNavigationItem({
    id: MAIL_NAV_ID,
    label: 'Outlook',
    icon: { lucide: 'mail' },
    visible: true,
    priority: 1,
    target: { type: 'panel', panelId: MAIL_PANEL_ID },
  });
  api.ui.registerSettingsView({
    id: SETTINGS_ID,
    label: 'Teams & Outlook',
  });

  api.onAction(`panel:${PANEL_ID}`, (action, data) => handlePanelAction(api, action, data));
  const SHARED_ACTIONS = new Set(['login', 'logout', 'open-external', 'search-people', 'navigate-panel']);
  api.onAction(`panel:${MAIL_PANEL_ID}`, (action, data) =>
    SHARED_ACTIONS.has(action) ? handlePanelAction(api, action, data) : handleMailAction(api, action, data),
  );
  api.onAction('settings:SettingsView', (action, data) => handleSettingsAction(api, action, data));

  try {
    api.events?.declare({
      events: [
        {
          event: 'message-received',
          title: 'Message received',
          description: 'A new Teams chat message arrived (not sent by you).',
          payloadSchema: {
            type: 'object',
            properties: {
              chatId: { type: 'string' },
              messageId: { type: 'string' },
              chatTitle: { type: 'string' },
              from: { type: 'object', properties: { id: { type: 'string' }, displayName: { type: 'string' } } },
              preview: { type: 'string' },
              isActiveChat: { type: 'boolean' },
            },
          },
        },
        {
          event: 'reaction-added',
          title: 'Reaction added',
          description: 'Someone reacted to a message in one of your chats.',
          payloadSchema: {
            type: 'object',
            properties: {
              chatId: { type: 'string' },
              messageId: { type: 'string' },
              chatTitle: { type: 'string' },
              reaction: { type: 'string' },
              from: { type: 'object', properties: { id: { type: 'string' }, displayName: { type: 'string' } } },
            },
          },
        },
        {
          event: 'message-read',
          title: 'Message read',
          description: 'A chat member advanced their read receipt (viewed your messages).',
          payloadSchema: {
            type: 'object',
            properties: {
              chatId: { type: 'string' },
              chatTitle: { type: 'string' },
              reader: { type: 'object', properties: { id: { type: 'string' }, displayName: { type: 'string' } } },
              lastReadMessageId: { type: 'string' },
            },
          },
        },
        {
          event: 'typing-started',
          title: 'Typing started',
          description: 'Someone started typing in a chat.',
          payloadSchema: {
            type: 'object',
            properties: {
              chatId: { type: 'string' },
              chatTitle: { type: 'string' },
              from: { type: 'object', properties: { id: { type: 'string' }, displayName: { type: 'string' } } },
            },
          },
        },
        {
          event: 'mail-received',
          title: 'Mail received',
          description: 'A new inbox message arrived.',
          payloadSchema: {
            type: 'object',
            properties: {
              messageId: { type: 'string' },
              subject: { type: 'string' },
              from: { type: 'object', properties: { name: { type: 'string' }, address: { type: 'string' } } },
              preview: { type: 'string' },
              receivedDateTime: { type: 'string' },
              hasAttachments: { type: 'boolean' },
            },
          },
        },
        {
          event: 'presence-changed',
          title: 'Presence changed',
          description: "A subscribed user's Teams availability changed.",
          payloadSchema: {
            type: 'object',
            properties: {
              userId: { type: 'string' },
              availability: { type: 'string' },
              activity: { type: 'string' },
              previous: { type: 'object' },
            },
          },
        },
      ],
      actions: [
        {
          targetId: `panel:${MAIL_PANEL_ID}`,
          title: 'Outlook panel action',
          description:
            'Dispatch a mail action. Set the action verb to one of: send-mail {mail:{to,subject,bodyHtml}}, ' +
            'mark-mail {messageId,isRead|flag}, archive-mail {messageId}, delete-mail {messageId}, ' +
            'select-folder {folderId}, select-mail {messageId}.',
          inputSchema: { type: 'object', additionalProperties: true },
        },
        {
          targetId: `panel:${PANEL_ID}`,
          title: 'Teams panel action',
          description:
            'Dispatch a panel action. Set the action verb to one of: send-message {chatId,text}, ' +
            'react-to-message {chatId,messageId,reactionType}, mark-chat-read {chatId}, ' +
            'set-presence {availability}, set-status-message {message,pinned}, select-chat {chatId}.',
          inputSchema: { type: 'object', additionalProperties: true },
        },
      ],
    });
  } catch (err) {
    log.warn(`events.declare unavailable: ${err}`);
  }

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
        startRealtime(api);
        startMailPoll(api);
        await loadChats(api);
        void loadMailFolders(api).then(() => loadMailList(api, 'inbox'));
        await refreshMeProfile(api);
      } catch (err) {
        log.warn(`Startup silent refresh failed: ${err}`);
        publishAuthState(api);
      }
    })();
  }

  // Poll remains as a safety net (trouter can drop events); it's cheap and skips
  // the active-thread reload when the push connection is live.
  const prefs = getPreferences(api);
  pollTimer = setInterval(() => {
    if (!tokenCache.hasRefreshToken()) return;
    void loadChats(api);
    const st = api.state.get() as Partial<MsgraphPluginState>;
    if (st.activeChatId && st.realtime !== 'connected') void loadMessages(api, st.activeChatId);
  }, Math.max(15, prefs.pollIntervalSeconds) * 1000);

  log.info('msgraph plugin activated');
}

export async function deactivate(): Promise<void> {
  if (pollTimer) clearInterval(pollTimer);
  if (refreshTimer) clearInterval(refreshTimer);
  if (unsubConfig) unsubConfig();
  trouter?.stop();
  trouter = null;
  for (const t of typingTimers.values()) clearTimeout(t);
  typingTimers.clear();
  for (const t of chatRefreshDebounce.values()) clearTimeout(t);
  chatRefreshDebounce.clear();
  photoCache.flush();
  peopleSearchDisk?.dispose();
  threadCache?.dispose();
  disposeMail();
  mediaServer.stop();
  pollTimer = null;
  refreshTimer = null;
  unsubConfig = null;
}
