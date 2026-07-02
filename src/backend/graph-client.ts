import { GRAPH_BASE_URL, DEFAULT_CHAT_LIST_TOP, DEFAULT_MESSAGE_TOP, CLIENT_ID_OUTLOOK_MOBILE, CLIENT_ID_TEAMS } from '../shared/constants.js';
import type {
  PluginAPI,
  GraphUser,
  GraphChat,
  GraphChatType,
  GraphMessage,
  GraphReaction,
  NormalizedChat,
  NormalizedMessage,
  NormalizedReaction,
  Presence,
} from '../shared/types.js';
import { GraphApiError, TokenExpiredError } from '../shared/types.js';
import { ensureAccessToken, forceRefresh, acquireFociAccessToken } from './auth.js';
import { parseHtmlBody } from './html-segments.js';
import * as tokenCache from './token-cache.js';
import { getLogger } from './logger-singleton.js';

type Fetch = typeof globalThis.fetch;

interface GraphList<T> {
  value: T[];
  '@odata.nextLink'?: string;
}

export class GraphClient {
  constructor(
    private readonly api: PluginAPI,
    private readonly allowInteractive = true,
  ) {}

  private get fetch(): Fetch {
    return this.api.fetch;
  }

  private async request<T>(
    method: string,
    path: string,
    opts: { query?: Record<string, string>; headers?: Record<string, string>; body?: unknown } = {},
    retried = false,
  ): Promise<T> {
    const token = await ensureAccessToken(this.api, { allowInteractive: this.allowInteractive });
    let url: string;
    if (/^https?:/i.test(path)) {
      if (!path.startsWith('https://graph.microsoft.com/')) {
        throw new GraphApiError(`Refusing to send Graph token to non-Graph URL: ${path}`, 0);
      }
      url = path;
    } else {
      url = `${GRAPH_BASE_URL}${path}${opts.query ? '?' + new URLSearchParams(opts.query).toString() : ''}`;
    }

    const resp = await this.fetch(url, {
      method,
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: 'application/json',
        ...(opts.body ? { 'Content-Type': 'application/json' } : {}),
        ...(opts.headers ?? {}),
      },
      body: opts.body ? JSON.stringify(opts.body) : undefined,
    });

    if (resp.status === 401 && !retried) {
      getLogger().warn(`Graph 401 on ${method} ${path}; refreshing and retrying once`);
      try {
        await forceRefresh(this.api);
      } catch (err) {
        throw new TokenExpiredError(`Token refresh failed: ${err instanceof Error ? err.message : String(err)}`);
      }
      return this.request<T>(method, path, opts, true);
    }

    if (resp.status === 204) return undefined as T;

    let body: unknown;
    try { body = await resp.json(); } catch { body = undefined; }

    if (!resp.ok) {
      const msg = (body as { error?: { code?: string; message?: string } } | undefined)?.error;
      throw new GraphApiError(
        `${method} ${path} → ${resp.status} ${msg?.code ?? ''} ${msg?.message ?? resp.statusText}`.trim(),
        resp.status,
        body,
      );
    }
    return body as T;
  }

  /**
   * Token for /users and /me/people. The Office client shares a tenant-wide
   * aadgraph throttle bucket that 429s in busy tenants; the Teams client has
   * User.ReadBasic.All + People.Read on a separate bucket.
   */
  private directoryToken(): Promise<string> {
    return acquireFociAccessToken(this.api, CLIENT_ID_TEAMS);
  }

  private async directoryGet<T>(path: string, opts: { query?: Record<string, string>; headers?: Record<string, string> } = {}): Promise<T> {
    const token = await this.directoryToken();
    const url = `${GRAPH_BASE_URL}${path}${opts.query ? '?' + new URLSearchParams(opts.query).toString() : ''}`;
    const resp = await this.fetch(url, {
      headers: { Authorization: `Bearer ${token}`, Accept: 'application/json', ...(opts.headers ?? {}) },
    });
    if (!resp.ok) {
      let body: unknown;
      try { body = await resp.json(); } catch { body = undefined; }
      const msg = (body as { error?: { code?: string; message?: string } } | undefined)?.error;
      throw new GraphApiError(`GET ${path} → ${resp.status} ${msg?.code ?? ''} ${msg?.message ?? resp.statusText}`.trim(), resp.status, body);
    }
    return (await resp.json()) as T;
  }

  // ── Identity ──

  async getMe(): Promise<GraphUser> {
    return this.request<GraphUser>('GET', '/me', {
      query: { $select: 'id,displayName,userPrincipalName,mail,jobTitle' },
    });
  }

  async getUser(userId: string): Promise<GraphUser> {
    return this.directoryGet<GraphUser>(`/users/${encodeURIComponent(userId)}`, {
      query: { $select: 'id,displayName,userPrincipalName,mail,jobTitle' },
    });
  }

  async findUsers(query: string, top = 10): Promise<GraphUser[]> {
    const q = query.trim();
    if (!q) return [];
    const escaped = q.replace(/"/g, '\\"');
    const r = await this.directoryGet<GraphList<GraphUser>>('/users', {
      headers: { ConsistencyLevel: 'eventual' },
      query: {
        $search: `"displayName:${escaped}" OR "mail:${escaped}" OR "userPrincipalName:${escaped}"`,
        $select: 'id,displayName,userPrincipalName,mail',
        $top: String(top),
        $count: 'true',
      },
    });
    return r.value;
  }

  /** Relevance-ranked people search (/me/people) — orders by the signed-in user's interaction history. */
  async searchPeople(query: string, top = 10): Promise<GraphUser[]> {
    const q = query.trim();
    if (!q) return [];
    const r = await this.directoryGet<
      GraphList<{ id: string; displayName?: string; userPrincipalName?: string; scoredEmailAddresses?: Array<{ address?: string }> }>
    >('/me/people', {
      query: {
        $search: `"${q.replace(/"/g, '\\"')}"`,
        $select: 'id,displayName,userPrincipalName,scoredEmailAddresses',
        $top: String(top),
      },
    });
    return r.value.map((p) => ({
      id: p.id,
      displayName: p.displayName ?? null,
      userPrincipalName: p.userPrincipalName ?? null,
      mail: p.scoredEmailAddresses?.[0]?.address ?? null,
    }));
  }

  /** Returns a data:image/* URL for the user's 48×48 profile photo, or null when the user has none. Throws on transient errors so callers can retry later. */
  async getUserPhoto(userId: string, retried = false): Promise<string | null> {
    const token = await this.directoryToken();
    const path = userId === tokenCache.getObjectId() ? '/me' : `/users/${encodeURIComponent(userId)}`;
    const resp = await this.fetch(`${GRAPH_BASE_URL}${path}/photos/48x48/$value`, {
      headers: { Authorization: `Bearer ${token}` },
    });
    void retried;
    if (resp.status === 404) return null;
    if (!resp.ok) {
      throw new GraphApiError(`GET ${path}/photos → ${resp.status}`, resp.status);
    }
    const buf = Buffer.from(await resp.arrayBuffer());
    const ct = resp.headers.get('content-type') || 'image/jpeg';
    return `data:${ct};base64,${buf.toString('base64')}`;
  }

  async getUserByEmail(email: string): Promise<GraphUser | null> {
    try {
      return await this.directoryGet<GraphUser>(`/users/${encodeURIComponent(email)}`, {
        query: { $select: 'id,displayName,userPrincipalName,mail' },
      });
    } catch (err) {
      if (err instanceof GraphApiError && err.statusCode === 404) return null;
      throw err;
    }
  }

  // ── Chats ──

  async listChats(
    opts: { chatType?: GraphChatType; top?: number } = {},
  ): Promise<{ chats: GraphChat[]; nextLink: string | null }> {
    // Graph caps /me/chats $top at 50 and members-expansion truncates large groups.
    const top = Math.min(Math.max(opts.top ?? DEFAULT_CHAT_LIST_TOP, 1), 50);
    const query: Record<string, string> = {
      $expand: 'members,lastMessagePreview',
      $orderby: 'lastMessagePreview/createdDateTime desc',
      $top: String(top),
    };
    if (opts.chatType) query.$filter = `chatType eq '${opts.chatType}'`;
    const r = await this.request<GraphList<GraphChat>>('GET', '/me/chats', { query });
    return { chats: r.value, nextLink: r['@odata.nextLink'] ?? null };
  }

  async listChatsPage(nextLink: string): Promise<{ chats: GraphChat[]; nextLink: string | null }> {
    const r = await this.request<GraphList<GraphChat>>('GET', nextLink);
    return { chats: r.value, nextLink: r['@odata.nextLink'] ?? null };
  }

  async getChat(chatId: string): Promise<GraphChat> {
    return this.request<GraphChat>('GET', `/chats/${encodeURIComponent(chatId)}`, {
      query: { $expand: 'members,lastMessagePreview' },
    });
  }

  async getChatMessages(chatId: string, top = DEFAULT_MESSAGE_TOP): Promise<GraphMessage[]> {
    const r = await this.request<GraphList<GraphMessage>>(
      'GET',
      `/chats/${encodeURIComponent(chatId)}/messages`,
      { query: { $top: String(top), $orderby: 'createdDateTime desc' } },
    );
    return r.value;
  }

  /** Fetch an auth-protected Graph hostedContents/$value URL and return a data URL. */
  async getHostedContent(url: string): Promise<string> {
    if (!url.startsWith('https://graph.microsoft.com/')) {
      throw new GraphApiError(`Refusing non-Graph hosted content URL`, 0);
    }
    const token = await ensureAccessToken(this.api, { allowInteractive: false });
    const resp = await this.fetch(url, { headers: { Authorization: `Bearer ${token}` } });
    if (!resp.ok) throw new GraphApiError(`GET hostedContent → ${resp.status}`, resp.status);
    const buf = Buffer.from(await resp.arrayBuffer());
    const ct = resp.headers.get('content-type') || 'image/png';
    return `data:${ct};base64,${buf.toString('base64')}`;
  }

  async sendMessage(
    chatId: string,
    text: string,
    contentType: 'text' | 'html' = 'text',
  ): Promise<GraphMessage> {
    return this.request<GraphMessage>('POST', `/chats/${encodeURIComponent(chatId)}/messages`, {
      body: { body: { contentType, content: text } },
    });
  }

  async sendMessageRaw(chatId: string, payload: Record<string, unknown>): Promise<GraphMessage> {
    return this.request<GraphMessage>('POST', `/chats/${encodeURIComponent(chatId)}/messages`, {
      body: payload,
    });
  }

  async setReaction(chatId: string, messageId: string, reactionType: string): Promise<void> {
    await this.request<void>(
      'POST',
      `/chats/${encodeURIComponent(chatId)}/messages/${encodeURIComponent(messageId)}/setReaction`,
      { body: { reactionType: toReactionGlyph(reactionType) } },
    );
  }

  async unsetReaction(chatId: string, messageId: string, reactionType: string): Promise<void> {
    await this.request<void>(
      'POST',
      `/chats/${encodeURIComponent(chatId)}/messages/${encodeURIComponent(messageId)}/unsetReaction`,
      { body: { reactionType: toReactionGlyph(reactionType) } },
    );
  }

  async markChatRead(chatId: string): Promise<void> {
    const id = tokenCache.getObjectId();
    const tenantId = tokenCache.getTenantId();
    if (!id || !tenantId) throw new Error('Not signed in');
    await this.request<void>('POST', `/chats/${encodeURIComponent(chatId)}/markChatReadForUser`, {
      body: { user: { id, tenantId } },
    });
  }

  /** Batch presence lookup. Uses the Outlook Mobile FOCI client (only one with Presence.Read.All). */
  async getPresences(userIds: string[]): Promise<Record<string, Presence>> {
    if (userIds.length === 0) return {};
    const token = await acquireFociAccessToken(this.api, CLIENT_ID_OUTLOOK_MOBILE);
    const resp = await this.fetch(`${GRAPH_BASE_URL}/communications/getPresencesByUserId`, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
        Accept: 'application/json',
      },
      body: JSON.stringify({ ids: userIds }),
    });
    if (!resp.ok) {
      throw new GraphApiError(`getPresencesByUserId → ${resp.status}`, resp.status);
    }
    const body = (await resp.json()) as {
      value?: Array<{
        id: string;
        availability: string;
        activity: string;
        statusMessage?: { message?: { content?: string } | null } | null;
      }>;
    };
    const out: Record<string, Presence> = {};
    for (const p of body.value ?? []) {
      const rawStatus = p.statusMessage?.message?.content ?? null;
      out[p.id] = {
        availability: p.availability,
        activity: p.activity,
        statusMessage: rawStatus ? stripHtml(rawStatus) || null : null,
      };
    }
    return out;
  }

  /** Construct the deterministic 1:1 chat id and fetch it if it exists (does not create). */
  async probeOneOnOne(otherUserId: string): Promise<GraphChat | null> {
    const meId = tokenCache.getObjectId();
    if (!meId) throw new Error('Not signed in');
    const [a, b] = [meId, otherUserId].sort();
    const chatId = `19:${a}_${b}@unq.gbl.spaces`;
    try {
      return await this.getChat(chatId);
    } catch (err) {
      if (err instanceof GraphApiError && (err.statusCode === 404 || err.statusCode === 403)) return null;
      throw err;
    }
  }

  /** Find or create the 1:1 chat with `otherUserId`. */
  async getOrCreateOneOnOne(otherUserId: string): Promise<GraphChat> {
    const meId = tokenCache.getObjectId();
    if (!meId) throw new Error('Not signed in');
    const body = {
      chatType: 'oneOnOne',
      members: [meId, otherUserId].map((id) => ({
        '@odata.type': '#microsoft.graph.aadUserConversationMember',
        roles: ['owner'],
        'user@odata.bind': `https://graph.microsoft.com/v1.0/users('${id}')`,
      })),
    };
    return this.request<GraphChat>('POST', '/chats', { body });
  }

  async createGroupChat(topic: string | null, memberUserIds: string[]): Promise<GraphChat> {
    const meId = tokenCache.getObjectId();
    if (!meId) throw new Error('Not signed in');
    const ids = Array.from(new Set([meId, ...memberUserIds]));
    if (ids.length < 3) {
      throw new Error('Group chats require at least 2 other members (3 total including you)');
    }
    const body: Record<string, unknown> = {
      chatType: 'group',
      members: ids.map((id) => ({
        '@odata.type': '#microsoft.graph.aadUserConversationMember',
        roles: ['owner'],
        'user@odata.bind': `https://graph.microsoft.com/v1.0/users('${id}')`,
      })),
    };
    if (topic) body.topic = topic;
    return this.request<GraphChat>('POST', '/chats', { body });
  }

  /** Search across all chat messages the signed-in user can access. */
  async searchMessages(query: string, top = 15): Promise<Array<{
    chatId: string | null;
    summary: string;
    from: string | null;
    createdDateTime: string | null;
    webUrl: string | null;
  }>> {
    const body = {
      requests: [{
        entityTypes: ['chatMessage'],
        query: { queryString: query },
        from: 0,
        size: top,
      }],
    };
    const r = await this.request<{
      value: Array<{
        hitsContainers: Array<{
          hits?: Array<{
            summary?: string;
            resource?: {
              chatId?: string;
              createdDateTime?: string;
              from?: { user?: { displayName?: string } };
              webUrl?: string;
              body?: { content?: string };
            };
          }>;
        }>;
      }>;
    }>('POST', '/search/query', { body });

    const hits = r.value?.[0]?.hitsContainers?.[0]?.hits ?? [];
    return hits.map((h) => ({
      chatId: h.resource?.chatId ?? null,
      summary: h.summary ?? h.resource?.body?.content ?? '',
      from: h.resource?.from?.user?.displayName ?? null,
      createdDateTime: h.resource?.createdDateTime ?? null,
      webUrl: h.resource?.webUrl ?? null,
    }));
  }
}

// ── Normalization helpers ──

const HOSTED_IMG_RE = /<img\b[^>]*\bsrc\s*=\s*"(https:\/\/graph\.microsoft\.com\/[^"]*?\/hostedContents\/[^"]+?\/\$value)"/gi;

function extractHostedImages(html: string): string[] {
  const out: string[] = [];
  for (const m of html.matchAll(HOSTED_IMG_RE)) out.push(m[1]);
  return out;
}

function stripHtml(html: string): string {
  return html
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/p>/gi, '\n')
    .replace(/<[^>]+>/g, '')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

export function normalizeChat(c: GraphChat, myId: string | null): NormalizedChat {
  const members = (c.members ?? [])
    .filter((m) => m.userId && m.userId !== myId)
    .map((m) => ({
      id: m.userId!,
      displayName: m.displayName ?? m.email ?? m.userId!,
      email: m.email ?? null,
    }))
    .sort((a, b) => a.displayName.localeCompare(b.displayName) || a.id.localeCompare(b.id));
  const previewBody = c.lastMessagePreview?.body?.content ?? null;
  const lastMsgAt = c.lastMessagePreview?.createdDateTime ?? null;
  const readAt = c.viewpoint?.lastMessageReadDateTime ?? null;
  const lastFromMe = c.lastMessagePreview?.from?.user?.id === myId;
  const unread = !!lastMsgAt && !lastFromMe && (!readAt || Date.parse(readAt) < Date.parse(lastMsgAt));
  return {
    id: c.id,
    type: c.chatType,
    topic: c.topic ?? (c.chatType === 'oneOnOne' ? members[0]?.displayName ?? null : null),
    members,
    lastUpdated: lastMsgAt ?? c.lastUpdatedDateTime ?? null,
    lastMessagePreview: previewBody ? stripHtml(previewBody).slice(0, 200) : null,
    lastMessageFrom: c.lastMessagePreview?.from?.user?.displayName ?? null,
    unread,
    webUrl: c.webUrl ?? null,
  };
}

const REACTION_EMOJI: Record<string, string> = {
  like: '👍',
  heart: '❤️',
  laugh: '😆',
  surprised: '😮',
  sad: '😢',
  angry: '😡',
};

/**
 * Graph returns reactionType as a name ("like") when reading, but setReaction /
 * unsetReaction require the Unicode emoji glyph. Accept either and emit the glyph.
 */
export function toReactionGlyph(input: string): string {
  const key = input.trim().toLowerCase();
  if (REACTION_EMOJI[key]) return REACTION_EMOJI[key];
  return input.trim();
}

function normalizeReactions(reactions: GraphReaction[] | undefined): NormalizedReaction[] {
  if (!reactions?.length) return [];
  const grouped = new Map<string, NormalizedReaction>();
  for (const r of reactions) {
    const type = r.reactionType ?? 'like';
    const emoji = REACTION_EMOJI[type] ?? type;
    let g = grouped.get(emoji);
    if (!g) {
      g = { emoji, type, count: 0, users: [] };
      grouped.set(emoji, g);
    }
    g.count++;
    const name = r.user?.user?.displayName;
    if (name) g.users.push(name);
  }
  return [...grouped.values()];
}

function parseMessageReference(content: string | undefined): NormalizedMessage['replyTo'] {
  if (!content) return null;
  try {
    const j = JSON.parse(content) as {
      messageId?: string;
      messagePreview?: string;
      messageSender?: { user?: { displayName?: string; id?: string } };
    };
    return {
      id: j.messageId ?? null,
      senderName: j.messageSender?.user?.displayName ?? null,
      text: j.messagePreview ? stripHtml(j.messagePreview) : null,
    };
  } catch {
    return null;
  }
}

export function normalizeMessage(m: GraphMessage, myId: string | null): NormalizedMessage {
  const fromId = m.from?.user?.id ?? null;
  const contentType = m.body?.contentType === 'html' ? 'html' : 'text';
  const raw = m.body?.content ?? '';
  const hostedImages = contentType === 'html' ? extractHostedImages(raw) : [];
  const segments = contentType === 'html'
    ? parseHtmlBody(raw, m.mentions)
    : (raw ? [{ type: 'text' as const, text: raw }] : []);
  const refAttachment = (m.attachments ?? []).find((a) => a.contentType === 'messageReference');
  const replyTo = refAttachment ? parseMessageReference(refAttachment.content) : null;
  const attachments = (m.attachments ?? []).filter((a) => a.contentType !== 'messageReference');
  return {
    id: m.id,
    chatId: m.chatId ?? '',
    createdDateTime: m.createdDateTime ?? null,
    fromId,
    fromName: m.from?.user?.displayName ?? null,
    fromMe: !!myId && fromId === myId,
    contentType,
    text: contentType === 'html' ? stripHtml(raw) : raw,
    segments,
    hostedImages,
    replyTo,
    attachments: attachments.map((a) => ({
      name: a.name ?? null,
      contentType: a.contentType ?? null,
      url: a.contentUrl ?? null,
    })),
    reactions: normalizeReactions(m.reactions),
    deleted: !!m.deletedDateTime,
  };
}
