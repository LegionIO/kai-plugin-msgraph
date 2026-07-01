import { GRAPH_BASE_URL, DEFAULT_CHAT_LIST_TOP, DEFAULT_MESSAGE_TOP } from '../shared/constants.js';
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
} from '../shared/types.js';
import { GraphApiError, TokenExpiredError } from '../shared/types.js';
import { ensureAccessToken, forceRefresh } from './auth.js';
import * as tokenCache from './token-cache.js';
import { getLogger } from './logger-singleton.js';

type Fetch = typeof globalThis.fetch;

interface GraphList<T> {
  value: T[];
  '@odata.nextLink'?: string;
}

export class GraphClient {
  constructor(private readonly api: PluginAPI) {}

  private get fetch(): Fetch {
    return this.api.fetch;
  }

  private async request<T>(
    method: string,
    path: string,
    opts: { query?: Record<string, string>; headers?: Record<string, string>; body?: unknown } = {},
    retried = false,
  ): Promise<T> {
    const token = await ensureAccessToken(this.api);
    const url = path.startsWith('http')
      ? path
      : `${GRAPH_BASE_URL}${path}${opts.query ? '?' + new URLSearchParams(opts.query).toString() : ''}`;

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

  // ── Identity ──

  async getMe(): Promise<GraphUser> {
    return this.request<GraphUser>('GET', '/me', {
      query: { $select: 'id,displayName,userPrincipalName,mail' },
    });
  }

  async findUsers(query: string, top = 10): Promise<GraphUser[]> {
    const q = query.trim();
    if (!q) return [];
    // $search covers displayName tokens regardless of "Last, First" ordering.
    const escaped = q.replace(/"/g, '\\"');
    const r = await this.request<GraphList<GraphUser>>('GET', '/users', {
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

  /** Returns a data:image/* URL for the user's 48×48 profile photo, or null if none. */
  async getUserPhoto(userId: string): Promise<string | null> {
    const token = await ensureAccessToken(this.api);
    const path = userId === tokenCache.getObjectId() ? '/me' : `/users/${encodeURIComponent(userId)}`;
    const resp = await this.fetch(`${GRAPH_BASE_URL}${path}/photos/48x48/$value`, {
      headers: { Authorization: `Bearer ${token}` },
    });
    if (resp.status === 404) return null;
    if (!resp.ok) {
      getLogger().warn(`getUserPhoto(${userId}) → ${resp.status}`);
      return null;
    }
    const buf = Buffer.from(await resp.arrayBuffer());
    const ct = resp.headers.get('content-type') || 'image/jpeg';
    return `data:${ct};base64,${buf.toString('base64')}`;
  }

  async getUserByEmail(email: string): Promise<GraphUser | null> {
    try {
      return await this.request<GraphUser>('GET', `/users/${encodeURIComponent(email)}`, {
        query: { $select: 'id,displayName,userPrincipalName,mail' },
      });
    } catch (err) {
      if (err instanceof GraphApiError && err.statusCode === 404) return null;
      throw err;
    }
  }

  // ── Chats ──

  async listChats(opts: { chatType?: GraphChatType; top?: number } = {}): Promise<GraphChat[]> {
    const query: Record<string, string> = {
      $expand: 'members,lastMessagePreview',
      $orderby: 'lastMessagePreview/createdDateTime desc',
      $top: String(opts.top ?? DEFAULT_CHAT_LIST_TOP),
    };
    if (opts.chatType) query.$filter = `chatType eq '${opts.chatType}'`;
    const r = await this.request<GraphList<GraphChat>>('GET', '/me/chats', { query });
    return r.value;
  }

  async getChat(chatId: string): Promise<GraphChat> {
    return this.request<GraphChat>('GET', `/chats/${encodeURIComponent(chatId)}`, {
      query: { $expand: 'members' },
    });
  }

  async getChatMessages(chatId: string, top = DEFAULT_MESSAGE_TOP): Promise<GraphMessage[]> {
    const r = await this.request<GraphList<GraphMessage>>(
      'GET',
      `/chats/${encodeURIComponent(chatId)}/messages`,
      { query: { $top: String(top) } },
    );
    return r.value;
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
  return {
    id: c.id,
    type: c.chatType,
    topic: c.topic ?? (c.chatType === 'oneOnOne' ? members[0]?.displayName ?? null : null),
    members,
    lastUpdated: c.lastMessagePreview?.createdDateTime ?? c.lastUpdatedDateTime ?? null,
    lastMessagePreview: previewBody ? stripHtml(previewBody).slice(0, 200) : null,
    lastMessageFrom: c.lastMessagePreview?.from?.user?.displayName ?? null,
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

export function normalizeMessage(m: GraphMessage, myId: string | null): NormalizedMessage {
  const fromId = m.from?.user?.id ?? null;
  const contentType = m.body?.contentType === 'html' ? 'html' : 'text';
  const raw = m.body?.content ?? '';
  return {
    id: m.id,
    chatId: m.chatId ?? '',
    createdDateTime: m.createdDateTime ?? null,
    fromId,
    fromName: m.from?.user?.displayName ?? null,
    fromMe: !!myId && fromId === myId,
    contentType,
    text: contentType === 'html' ? stripHtml(raw) : raw,
    attachments: (m.attachments ?? []).map((a) => ({
      name: a.name ?? null,
      contentType: a.contentType ?? null,
      url: a.contentUrl ?? null,
    })),
    reactions: normalizeReactions(m.reactions),
  };
}
