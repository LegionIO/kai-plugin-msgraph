import { GRAPH_BASE_URL, DEFAULT_CHAT_LIST_TOP, DEFAULT_MESSAGE_TOP, MAIL_LIST_TOP, CLIENT_ID_OUTLOOK_MOBILE, CLIENT_ID_TEAMS } from '../shared/constants.js';
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
  MailFolder,
  MailAddress,
  NormalizedMailSummary,
  NormalizedMail,
  MailAttachmentMeta,
  OutgoingMail,
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

    const safePath = path.replace(/\?.*$/, '');

    if (resp.status === 401 && !retried) {
      getLogger().warn(`Graph 401 on ${method} ${safePath}; refreshing and retrying once`);
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
        `${method} ${safePath} → ${resp.status} ${msg?.code ?? ''} ${msg?.message ?? resp.statusText}`.trim(),
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

  /** Resolve a bot's Teams-app icon. `appId` is message.from.application.id (== teamsApp id). */
  async getAppIcon(appId: string): Promise<string | null> {
    const token = await this.directoryToken();
    const H = { Authorization: `Bearer ${token}` };
    // 1. Look up the app + latest definition (v1.0; no $top supported here).
    const listUrl = `${GRAPH_BASE_URL}/appCatalogs/teamsApps?$filter=id eq '${appId}'&$expand=appDefinitions`;
    const lr = await this.fetch(listUrl, { headers: H });
    if (!lr.ok) throw new GraphApiError(`appCatalogs lookup → ${lr.status}`, lr.status);
    const list = (await lr.json()) as {
      value?: Array<{ id: string; appDefinitions?: Array<{ id: string }> }>;
    };
    const app = list.value?.[0];
    const def = app?.appDefinitions?.[app.appDefinitions.length - 1];
    if (!app || !def) return null;
    // 2. colorIcon is beta-only; store apps expose a public CDN webUrl.
    const metaUrl = `https://graph.microsoft.com/beta/appCatalogs/teamsApps/${app.id}/appDefinitions/${encodeURIComponent(def.id)}/colorIcon`;
    const mr = await this.fetch(metaUrl, { headers: H });
    if (!mr.ok) return null;
    const meta = (await mr.json()) as { webUrl?: string };
    // 3. Prefer webUrl (public CDN); fall back to hostedContent bytes for tenant-uploaded apps.
    if (meta.webUrl) {
      const ir = await this.fetch(meta.webUrl);
      if (!ir.ok) return null;
      const buf = Buffer.from(await ir.arrayBuffer());
      const ct = ir.headers.get('content-type') || 'image/png';
      return `data:${ct};base64,${buf.toString('base64')}`;
    }
    const hcUrl = `${metaUrl}/hostedContent/$value`;
    const hr = await this.fetch(hcUrl, { headers: H });
    if (!hr.ok) return null;
    const buf = Buffer.from(await hr.arrayBuffer());
    const ct = hr.headers.get('content-type') || 'image/png';
    return `data:${ct};base64,${buf.toString('base64')}`;
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
    // /users/{email} only matches UPN; many senders' `mail` differs from UPN
    // (e.g. @uhg.com vs @optum.com), so filter on both.
    const q = email.replace(/'/g, "''");
    try {
      const r = await this.directoryGet<GraphList<GraphUser>>('/users', {
        query: {
          $filter: `mail eq '${q}' or userPrincipalName eq '${q}'`,
          $select: 'id,displayName,userPrincipalName,mail',
          $top: '1',
        },
      });
      if (r.value[0]) return r.value[0];
    } catch (err) {
      if (!(err instanceof GraphApiError) || (err.statusCode !== 404 && err.statusCode !== 400)) throw err;
    }
    // Fallback for guest/external where filter isn't indexed: try direct.
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

  async getChatMessages(
    chatId: string,
    top = DEFAULT_MESSAGE_TOP,
  ): Promise<{ messages: GraphMessage[]; nextLink: string | null }> {
    const r = await this.request<GraphList<GraphMessage>>(
      'GET',
      `/chats/${encodeURIComponent(chatId)}/messages`,
      { query: { $top: String(top), $orderby: 'createdDateTime desc' } },
    );
    return { messages: r.value, nextLink: r['@odata.nextLink'] ?? null };
  }

  async getChatMessagesPage(nextLink: string): Promise<{ messages: GraphMessage[]; nextLink: string | null }> {
    const r = await this.request<GraphList<GraphMessage>>('GET', nextLink);
    return { messages: r.value, nextLink: r['@odata.nextLink'] ?? null };
  }

  /** Fetch an auth-protected Graph hostedContents/$value URL as raw base64 + media type. */
  async getHostedContentRaw(url: string): Promise<{ base64: string; mediaType: string }> {
    if (!url.startsWith('https://graph.microsoft.com/')) {
      throw new GraphApiError(`Refusing non-Graph hosted content URL`, 0);
    }
    const token = await ensureAccessToken(this.api, { allowInteractive: false });
    const resp = await this.fetch(url, { headers: { Authorization: `Bearer ${token}` } });
    if (!resp.ok) throw new GraphApiError(`GET hostedContent → ${resp.status}`, resp.status);
    const buf = Buffer.from(await resp.arrayBuffer());
    const mediaType = resp.headers.get('content-type') || 'image/png';
    return { base64: buf.toString('base64'), mediaType };
  }

  /** Fetch an auth-protected Graph hostedContents/$value URL and return a data URL. */
  async getHostedContent(url: string): Promise<string> {
    const { base64, mediaType } = await this.getHostedContentRaw(url);
    return `data:${mediaType};base64,${base64}`;
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

  async editMessage(
    chatId: string,
    messageId: string,
    payload: {
      body: { contentType: 'text' | 'html'; content: string };
      mentions?: unknown[];
      hostedContents?: unknown[];
      attachments?: unknown[];
    },
  ): Promise<void> {
    await this.request<void>('PATCH', `/chats/${encodeURIComponent(chatId)}/messages/${encodeURIComponent(messageId)}`, {
      body: payload,
    });
  }

  async deleteMessage(chatId: string, messageId: string): Promise<void> {
    // softDelete is only routed under /me (or /users/{id}); bare /chats/... 405s.
    await this.request<void>('POST', `/me/chats/${encodeURIComponent(chatId)}/messages/${encodeURIComponent(messageId)}/softDelete`, {
      body: {},
    });
  }

  async undeleteMessage(chatId: string, messageId: string): Promise<void> {
    await this.request<void>('POST', `/me/chats/${encodeURIComponent(chatId)}/messages/${encodeURIComponent(messageId)}/undoSoftDelete`, {
      body: {},
    });
  }

  async getMessage(chatId: string, messageId: string): Promise<GraphMessage> {
    return this.request<GraphMessage>('GET', `/chats/${encodeURIComponent(chatId)}/messages/${encodeURIComponent(messageId)}`);
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

  // ── Mail (Outlook) ──

  async listMailFolders(): Promise<{ folders: MailFolder[]; truncated: boolean }> {
    const [first, ...wk] = await Promise.all([
      this.request<GraphList<RawMailFolder>>('GET', '/me/mailFolders', {
        query: { $top: '100', $select: MAIL_FOLDER_SELECT },
      }),
      ...(['inbox', 'drafts', 'sentitems', 'archive', 'deleteditems', 'junkemail'] as const).map((n) =>
        this.request<{ id: string }>('GET', `/me/mailFolders/${n}`, { query: { $select: 'id' } })
          .then((r) => ({ name: n, id: r.id }))
          .catch(() => null),
      ),
    ]);
    const raw = [...first.value];
    let next = first['@odata.nextLink'] ?? null;
    let i = 0;
    for (; next && i < 20; i++) {
      const page = await this.request<GraphList<RawMailFolder>>('GET', next);
      raw.push(...page.value);
      next = page['@odata.nextLink'] ?? null;
    }
    const idToWk = new Map<string, string>();
    for (const x of wk) if (x) idToWk.set(x.id, x.name);
    return {
      folders: raw.map((f) => normalizeMailFolder(f, idToWk.get(f.id) ?? null, null, 0)),
      truncated: next != null,
    };
  }

  /**
   * All mail folders, descending up to `maxDepth` levels of children. Child-listing
   * errors propagate (fail closed) so name-resolution ambiguity checks see the full
   * set rather than silently missing a folder that shares a display name. `truncated`
   * is true when the depth cap left unexplored children — callers must not treat a
   * name match as authoritative in that case.
   */
  async listMailFoldersDeep(maxDepth = 3): Promise<{ folders: MailFolder[]; truncated: boolean }> {
    const top = await this.listMailFolders();
    const out = [...top.folders];
    let truncated = top.truncated;
    let frontier = out.filter((f) => f.childFolderCount > 0);
    for (let depth = 1; depth <= maxDepth && frontier.length; depth++) {
      const pages = await mapLimit(frontier, 5, (f) => this.listChildFoldersPaged(f.id, depth));
      const kids = pages.flatMap((p) => p.folders);
      if (pages.some((p) => p.truncated)) truncated = true;
      out.push(...kids);
      if (depth < maxDepth) {
        frontier = kids.filter((f) => f.childFolderCount > 0);
      } else {
        if (kids.some((f) => f.childFolderCount > 0)) truncated = true;
        frontier = [];
      }
    }
    return { folders: out, truncated };
  }

  /**
   * Resolve a well-known name, folder id, or (case-insensitive) displayName to a
   * folder id usable in Graph URLs. Custom folders (e.g. "External Senders") can
   * be passed by name; Graph itself only accepts well-known names or base64 ids.
   */
  async resolveFolderId(nameOrId: string): Promise<string> {
    const q = nameOrId.trim();
    if (!q) throw new Error('Folder name or id is required');
    const lower = q.toLowerCase();
    const wellKnown = new Set([
      'inbox', 'drafts', 'sentitems', 'archive', 'deleteditems', 'junkemail', 'outbox',
      'clutter', 'conflicts', 'conversationhistory', 'localfailures', 'msgfolderroot',
      'recoverableitemsdeletions', 'scheduled', 'searchfolders', 'serverfailures', 'syncissues',
    ]);
    if (wellKnown.has(lower)) return lower;
    // Id-shaped input (long, opaque base64-ish): probe it directly so a valid folder
    // id keeps working even if the full-tree enumeration below would throttle/fail.
    if (q.length > 80 && /^[A-Za-z0-9+/=_-]+$/.test(q)) {
      try {
        await this.request<{ id: string }>('GET', `/me/mailFolders/${encodeURIComponent(q)}`, { query: { $select: 'id' } });
        return q;
      } catch (err) {
        if (!(err instanceof GraphApiError) || (err.statusCode !== 404 && err.statusCode !== 400)) throw err;
        // Not a folder id after all — fall through to name resolution.
      }
    }
    const { folders, truncated } = await this.listMailFoldersDeep();
    // Exact id match wins (Graph ids are opaque; don't guess by shape).
    if (folders.some((f) => f.id === q)) return q;
    const exact = folders.filter((f) => f.displayName.toLowerCase() === lower);
    if (exact.length === 1 && !truncated) return exact[0].id;
    if (exact.length > 1) {
      throw new Error(`Multiple folders named "${nameOrId}" (in different parents). Use the folder id from list-folders.`);
    }
    const partial = folders.filter((f) => f.displayName.toLowerCase().includes(lower));
    if (partial.length === 1 && !truncated) return partial[0].id;
    if (partial.length > 1) {
      throw new Error(`Ambiguous folder "${nameOrId}" — matches: ${partial.map((f) => f.displayName).join(', ')}. Use an exact name.`);
    }
    // No unique name match (or the folder tree was too deep to enumerate fully):
    // fall back to treating the input as an opaque folder id.
    if (/^[A-Za-z0-9+/=_-]+$/.test(q)) return q;
    throw new Error(`No mail folder named "${nameOrId}". Use list-folders to see available folders.`);
  }

  private async batch(
    reqs: Array<{ method: string; url: string; body?: unknown }>,
  ): Promise<Array<{ id: string; status: number; body?: unknown }>> {
    const out: Array<{ id: string; status: number; body?: unknown }> = [];
    for (let i = 0; i < reqs.length; i += 20) {
      const chunk = reqs.slice(i, i + 20).map((r, j) => ({
        id: String(i + j),
        method: r.method,
        url: r.url,
        ...(r.body !== undefined ? { body: r.body, headers: { 'Content-Type': 'application/json' } } : {}),
      }));
      const res = await this.request<{ responses: Array<{ id: string; status: number; body?: unknown }> }>(
        'POST', '/$batch', { body: { requests: chunk } },
      );
      out.push(...res.responses);
    }
    return out;
  }

  async markFolderRead(folderId: string): Promise<number> {
    let done = 0;
    let url: string | null =
      `${GRAPH_BASE_URL}/me/mailFolders/${encodeURIComponent(folderId)}/messages?` +
      new URLSearchParams({ $filter: 'isRead eq false', $select: 'id', $top: '100' }).toString();
    while (url) {
      const page: GraphList<{ id: string }> = await this.request('GET', url);
      if (page.value.length) {
        await this.batch(
          page.value.map((m) => ({
            method: 'PATCH',
            url: `/me/messages/${m.id}`,
            body: { isRead: true },
          })),
        );
        done += page.value.length;
      }
      url = page['@odata.nextLink'] ?? null;
      if (done >= 1000) break;
    }
    return done;
  }

  async emptyFolder(folderId: string, hardDelete: boolean): Promise<number> {
    let done = 0;
    let url: string | null =
      `${GRAPH_BASE_URL}/me/mailFolders/${encodeURIComponent(folderId)}/messages?` +
      new URLSearchParams({ $select: 'id', $top: '100' }).toString();
    while (url) {
      const page: GraphList<{ id: string }> = await this.request('GET', url);
      if (!page.value.length) break;
      await this.batch(
        page.value.map((m) =>
          hardDelete
            ? { method: 'DELETE', url: `/me/messages/${m.id}` }
            : { method: 'POST', url: `/me/messages/${m.id}/move`, body: { destinationId: 'deleteditems' } },
        ),
      );
      done += page.value.length;
      // Re-fetch first page each time (moving/deleting shifts the collection).
      url =
        `${GRAPH_BASE_URL}/me/mailFolders/${encodeURIComponent(folderId)}/messages?` +
        new URLSearchParams({ $select: 'id', $top: '100' }).toString();
      if (done >= 2000) break;
    }
    return done;
  }

  async createChildFolder(parentId: string, displayName: string): Promise<MailFolder> {
    const r = await this.request<RawMailFolder>(
      'POST',
      `/me/mailFolders/${encodeURIComponent(parentId)}/childFolders`,
      { body: { displayName } },
    );
    return normalizeMailFolder(r, null, parentId, 0);
  }

  async renameFolder(folderId: string, displayName: string): Promise<void> {
    await this.request('PATCH', `/me/mailFolders/${encodeURIComponent(folderId)}`, { body: { displayName } });
  }

  async deleteFolder(folderId: string): Promise<void> {
    await this.request('DELETE', `/me/mailFolders/${encodeURIComponent(folderId)}`);
  }

  async listChildFolders(parentId: string, depth: number): Promise<MailFolder[]> {
    return (await this.listChildFoldersPaged(parentId, depth)).folders;
  }

  private async listChildFoldersPaged(parentId: string, depth: number): Promise<{ folders: MailFolder[]; truncated: boolean }> {
    const first = await this.request<GraphList<RawMailFolder>>(
      'GET',
      `/me/mailFolders/${encodeURIComponent(parentId)}/childFolders`,
      { query: { $top: '100', $select: MAIL_FOLDER_SELECT } },
    );
    const raw = [...first.value];
    let next = first['@odata.nextLink'] ?? null;
    for (let i = 0; next && i < 20; i++) {
      const page = await this.request<GraphList<RawMailFolder>>('GET', next);
      raw.push(...page.value);
      next = page['@odata.nextLink'] ?? null;
    }
    return { folders: raw.map((f) => normalizeMailFolder(f, null, parentId, depth)), truncated: next != null };
  }

  async listMail(
    folderId: string,
    top = MAIL_LIST_TOP,
  ): Promise<{ messages: NormalizedMailSummary[]; nextLink: string | null }> {
    const r = await this.request<GraphList<RawMail>>(
      'GET',
      `/me/mailFolders/${encodeURIComponent(folderId)}/messages`,
      {
        query: {
          $top: String(Math.min(Math.max(top, 1), 100)),
          $select: MAIL_SUMMARY_SELECT,
          $orderby: 'receivedDateTime desc',
        },
      },
    );
    return { messages: r.value.map(normalizeMailSummary), nextLink: r['@odata.nextLink'] ?? null };
  }

  async listMailPage(nextLink: string): Promise<{ messages: NormalizedMailSummary[]; nextLink: string | null }> {
    const r = await this.request<GraphList<RawMail>>('GET', nextLink);
    return { messages: r.value.map(normalizeMailSummary), nextLink: r['@odata.nextLink'] ?? null };
  }

  /**
   * Delta on a folder. Pass null to start; store the returned deltaLink and pass
   * it next time to get only changes (new/updated/@removed).
   */
  async mailDelta(
    folderId: string,
    deltaLink: string | null,
  ): Promise<{ changes: Array<NormalizedMailSummary | { removedId: string }>; deltaLink: string | null }> {
    let url =
      deltaLink ??
      `${GRAPH_BASE_URL}/me/mailFolders/${encodeURIComponent(folderId)}/messages/delta?` +
        new URLSearchParams({ $select: MAIL_SUMMARY_SELECT }).toString();
    const changes: Array<NormalizedMailSummary | { removedId: string }> = [];
    let nextDelta: string | null = null;
    for (let i = 0; i < 10; i++) {
      const r = await this.request<
        GraphList<RawMail & { '@removed'?: { reason?: string } }> & { '@odata.deltaLink'?: string }
      >('GET', url, deltaLink && i === 0 ? {} : { headers: { Prefer: 'odata.maxpagesize=50' } });
      for (const m of r.value) {
        if (m['@removed']) changes.push({ removedId: m.id ?? '' });
        else changes.push(normalizeMailSummary(m));
      }
      if (r['@odata.deltaLink']) { nextDelta = r['@odata.deltaLink']; break; }
      if (!r['@odata.nextLink']) break;
      url = r['@odata.nextLink'];
    }
    return { changes, deltaLink: nextDelta };
  }

  async getMail(messageId: string): Promise<NormalizedMail> {
    const m = await this.request<RawMail>('GET', `/me/messages/${encodeURIComponent(messageId)}`, {
      query: { $select: `${MAIL_SUMMARY_SELECT},body,ccRecipients,bccRecipients` },
      headers: { Prefer: 'outlook.body-content-type="html"' },
    });
    // Graph's `hasAttachments` excludes inline-only attachments, so check the
    // body for cid: references too.
    const hasInlineRefs = /\bsrc\s*=\s*["']?cid:/i.test(m.body?.content ?? '');
    let attachments: MailAttachmentMeta[] = [];
    if (m.hasAttachments || hasInlineRefs) {
      const ar = await this.request<GraphList<RawAttachment>>(
        'GET',
        `/me/messages/${encodeURIComponent(messageId)}/attachments`,
        { query: { $select: 'id,name,contentType,size,isInline' } },
      );
      attachments = ar.value.map((a) => ({
        id: a.id ?? '',
        name: a.name ?? 'attachment',
        contentType: a.contentType ?? null,
        size: a.size ?? 0,
        isInline: !!a.isInline,
        contentId: null,
      }));
    }
    return {
      ...normalizeMailSummary(m),
      ccRecipients: (m.ccRecipients ?? []).map(addr),
      bccRecipients: (m.bccRecipients ?? []).map(addr),
      bodyHtml: m.body?.content ?? '',
      attachments,
    };
  }

  /** Fetch a single mail attachment as raw base64 (contentBytes for fileAttachment). */
  async getMailAttachmentRaw(
    messageId: string,
    attachmentId: string,
  ): Promise<{ contentId: string | null; base64: string; mediaType: string; name: string }> {
    const a = await this.request<RawAttachment & { contentBytes?: string }>(
      'GET',
      `/me/messages/${encodeURIComponent(messageId)}/attachments/${encodeURIComponent(attachmentId)}`,
    );
    return {
      contentId: a.contentId ?? null,
      base64: a.contentBytes ?? '',
      mediaType: a.contentType || 'application/octet-stream',
      name: a.name ?? 'attachment',
    };
  }

  /** Fetch a single attachment fully (contentId + contentBytes for fileAttachment). */
  async getMailAttachment(
    messageId: string,
    attachmentId: string,
  ): Promise<{ contentId: string | null; dataUrl: string; name: string }> {
    const { contentId, base64, mediaType, name } = await this.getMailAttachmentRaw(messageId, attachmentId);
    return { contentId, dataUrl: `data:${mediaType};base64,${base64}`, name };
  }

  /**
   * Download a Teams "reference" attachment (a SharePoint/OneDrive file) by its
   * contentUrl, via the Graph shares API. Returns raw base64 + media type.
   * NOTE: verified opportunistically — falls back through a couple of shapes.
   */
  async downloadReferenceAttachment(
    contentUrl: string,
  ): Promise<{ base64: string; mediaType: string }> {
    // Graph "shares" encoding: base64url of the full URL, prefixed with "u!".
    const b64 = Buffer.from(contentUrl, 'utf8')
      .toString('base64')
      .replace(/=+$/, '')
      .replace(/\//g, '_')
      .replace(/\+/g, '-');
    const shareId = `u!${b64}`;
    const token = await ensureAccessToken(this.api, { allowInteractive: false });
    const url = `${GRAPH_BASE_URL}/shares/${encodeURIComponent(shareId)}/driveItem/content`;
    const resp = await this.fetch(url, { headers: { Authorization: `Bearer ${token}` }, redirect: 'follow' });
    if (!resp.ok) throw new GraphApiError(`GET shared driveItem → ${resp.status}`, resp.status);
    const buf = Buffer.from(await resp.arrayBuffer());
    const mediaType = resp.headers.get('content-type') || 'application/octet-stream';
    return { base64: buf.toString('base64'), mediaType };
  }

  async patchMail(messageId: string, patch: { isRead?: boolean; flag?: 'flagged' | 'complete' | 'notFlagged' }): Promise<void> {
    const body: Record<string, unknown> = {};
    if (patch.isRead !== undefined) body.isRead = patch.isRead;
    if (patch.flag) body.flag = { flagStatus: patch.flag };
    await this.request('PATCH', `/me/messages/${encodeURIComponent(messageId)}`, { body });
  }

  async moveMail(messageId: string, destinationFolderId: string): Promise<string> {
    const r = await this.request<{ id: string }>('POST', `/me/messages/${encodeURIComponent(messageId)}/move`, {
      body: { destinationId: destinationFolderId },
    });
    return r.id;
  }

  async deleteMail(messageId: string): Promise<void> {
    await this.request('DELETE', `/me/messages/${encodeURIComponent(messageId)}`);
  }

  async sendMail(mail: OutgoingMail): Promise<void> {
    await this.request('POST', '/me/sendMail', {
      body: {
        message: toGraphMail(mail),
        saveToSentItems: true,
      },
    });
  }

  async replyMail(
    messageId: string,
    mode: 'reply' | 'replyAll' | 'forward',
    mail: Partial<OutgoingMail> & { comment?: string },
  ): Promise<void> {
    const path =
      mode === 'reply' ? 'reply' : mode === 'replyAll' ? 'replyAll' : 'forward';
    const body: Record<string, unknown> = {};
    // /reply and /replyAll accept `comment` (plain) or a full `message`; /forward requires toRecipients.
    const msg = toGraphMail({
      to: mail.to ?? [],
      cc: mail.cc,
      bcc: mail.bcc,
      subject: mail.subject ?? '',
      bodyHtml: mail.bodyHtml ?? '',
      attachments: mail.attachments,
    });
    if (mail.bodyHtml || mail.attachments?.length || mail.cc?.length || mail.bcc?.length) {
      body.message = msg;
      // Graph rejects an empty subject on message override for reply; omit it.
      if (!mail.subject) delete (body.message as Record<string, unknown>).subject;
    } else if (mail.comment) {
      body.comment = mail.comment;
    }
    if (mode === 'forward') {
      body.toRecipients = (mail.to ?? []).map((a) => ({ emailAddress: { address: a.address, name: a.name ?? undefined } }));
    }
    await this.request('POST', `/me/messages/${encodeURIComponent(messageId)}/${path}`, { body });
  }

  async searchMail(
    query: string,
    top = 25,
    opts: { from?: string; since?: string; until?: string; unreadOnly?: boolean; hasAttachments?: boolean } = {},
  ): Promise<NormalizedMailSummary[]> {
    const date = (label: string, v?: string) => {
      if (!v) return null;
      if (!/^\d{4}-\d{2}-\d{2}$/.test(v)) throw new Error(`${label} must be a YYYY-MM-DD date, got "${v}"`);
      return v;
    };
    const since = date('since', opts.since);
    const until = date('until', opts.until);
    const pageSize = Math.min(top, 100);

    // Only free-text and from: require $search (KQL). Everything else (isRead,
    // received dates, hasAttachments) is expressible as an OData $filter, which
    // filters server-side — so prefer $filter unless a real $search term is present.
    const needsSearch = !!(query && query.trim()) || !!opts.from;
    if (!needsSearch) {
      const filters: string[] = [];
      if (opts.unreadOnly) filters.push('isRead eq false');
      if (since) filters.push(`receivedDateTime ge ${since}T00:00:00Z`);
      if (until) filters.push(`receivedDateTime le ${until}T23:59:59Z`);
      if (opts.hasAttachments) filters.push('hasAttachments eq true');
      const q: Record<string, string> = {
        $top: String(pageSize),
        $select: MAIL_SUMMARY_SELECT,
      };
      // Graph rejects $orderby combined with a $filter unless the ordered field leads
      // the filter, so only order when there's no filter.
      if (filters.length) q.$filter = filters.join(' and ');
      else q.$orderby = 'receivedDateTime desc';
      const r = await this.request<GraphList<RawMail>>('GET', '/me/messages', { query: q });
      return r.value.map(normalizeMailSummary).slice(0, top);
    }

    // $search + $filter cannot be combined on /messages, so express refiners as KQL inside $search.
    const kqlStr = (v: string) => `"${v.replace(/["\\]/g, '\\$&')}"`;
    const terms: string[] = [];
    if (query && query.trim()) terms.push(kqlStr(query));
    if (opts.from) terms.push(`from:${kqlStr(opts.from)}`);
    if (opts.hasAttachments) terms.push('hasAttachments:true');
    if (since) terms.push(`received>=${since}`);
    if (until) terms.push(`received<=${until}`);

    // isRead is not a searchable $search property; post-filter instead of injecting it into KQL.
    const search = terms.join(' AND ');
    // Unread filtering is client-side, so fetch a larger candidate page (independent of `top`)
    // to avoid a small `top` starving the filter.
    const candidatePage = opts.unreadOnly ? Math.min(Math.max(top, 50), 100) : pageSize;
    let url =
      `${GRAPH_BASE_URL}/me/messages?` +
      new URLSearchParams({ $search: search, $top: String(candidatePage), $select: MAIL_SUMMARY_SELECT }).toString();
    const out: NormalizedMailSummary[] = [];
    // With unreadOnly we post-filter, so page until we have `top` unread (bounded).
    for (let i = 0; i < (opts.unreadOnly ? 5 : 1) && url; i++) {
      const r: GraphList<RawMail> = await this.request('GET', url, {
        headers: { ConsistencyLevel: 'eventual' },
      });
      for (const m of r.value.map(normalizeMailSummary)) {
        if (opts.unreadOnly && m.isRead) continue;
        out.push(m);
      }
      if (out.length >= top) break;
      url = r['@odata.nextLink'] ?? '';
    }
    return out.slice(0, top);
  }
}

// ── Mail normalization ──

const MAIL_SUMMARY_SELECT =
  'id,conversationId,subject,from,toRecipients,receivedDateTime,isRead,isDraft,hasAttachments,flag,importance,bodyPreview,webLink';

const MAIL_FOLDER_SELECT = 'id,displayName,unreadItemCount,totalItemCount,childFolderCount';

interface RawMailFolder {
  id: string;
  displayName?: string;
  unreadItemCount?: number;
  totalItemCount?: number;
  childFolderCount?: number;
}
interface RawEmailAddress { emailAddress?: { name?: string; address?: string } }
interface RawAttachment {
  id?: string; name?: string; contentType?: string; size?: number; isInline?: boolean; contentId?: string;
}
interface RawMail {
  id?: string;
  conversationId?: string;
  subject?: string;
  from?: RawEmailAddress;
  toRecipients?: RawEmailAddress[];
  ccRecipients?: RawEmailAddress[];
  bccRecipients?: RawEmailAddress[];
  receivedDateTime?: string;
  isRead?: boolean;
  isDraft?: boolean;
  hasAttachments?: boolean;
  flag?: { flagStatus?: string };
  importance?: string;
  bodyPreview?: string;
  webLink?: string;
  body?: { contentType?: string; content?: string };
}

/** Map with bounded concurrency so large folder trees don't fan out unbounded parallel Graph calls. */
async function mapLimit<T, R>(items: T[], limit: number, fn: (item: T) => Promise<R>): Promise<R[]> {
  const out = new Array<R>(items.length);
  let next = 0;
  const workers = Array.from({ length: Math.min(limit, items.length) }, async () => {
    while (next < items.length) {
      const i = next++;
      out[i] = await fn(items[i]);
    }
  });
  await Promise.all(workers);
  return out;
}

function addr(r: RawEmailAddress): MailAddress {
  return { name: r.emailAddress?.name ?? null, address: r.emailAddress?.address ?? '' };
}

function normalizeMailFolder(
  f: RawMailFolder,
  wellKnownName: string | null,
  parentId: string | null,
  depth: number,
): MailFolder {
  return {
    id: f.id,
    displayName: f.displayName ?? f.id,
    wellKnownName,
    unreadItemCount: f.unreadItemCount ?? 0,
    totalItemCount: f.totalItemCount ?? 0,
    childFolderCount: f.childFolderCount ?? 0,
    parentId,
    depth,
  };
}

export function normalizeMailSummary(m: RawMail): NormalizedMailSummary {
  return {
    id: m.id ?? '',
    conversationId: m.conversationId ?? null,
    subject: m.subject ?? '(no subject)',
    from: m.from ? addr(m.from) : null,
    toRecipients: (m.toRecipients ?? []).map(addr),
    receivedDateTime: m.receivedDateTime ?? null,
    isRead: m.isRead ?? true,
    isDraft: m.isDraft ?? false,
    hasAttachments: m.hasAttachments ?? false,
    flagged: (m.flag?.flagStatus ?? 'notFlagged') === 'flagged',
    importance: (m.importance as 'low' | 'normal' | 'high') ?? 'normal',
    bodyPreview: (m.bodyPreview ?? '').replace(/\s+/g, ' ').trim().slice(0, 200),
    webLink: m.webLink ?? null,
  };
}

function toGraphMail(mail: OutgoingMail): Record<string, unknown> {
  const rcpt = (list?: MailAddress[]) =>
    (list ?? []).map((a) => ({ emailAddress: { address: a.address, name: a.name ?? undefined } }));
  const out: Record<string, unknown> = {
    subject: mail.subject,
    body: { contentType: 'html', content: mail.bodyHtml },
    toRecipients: rcpt(mail.to),
  };
  if (mail.cc?.length) out.ccRecipients = rcpt(mail.cc);
  if (mail.bcc?.length) out.bccRecipients = rcpt(mail.bcc);
  if (mail.attachments?.length) {
    out.attachments = mail.attachments.map((a) => ({
      '@odata.type': '#microsoft.graph.fileAttachment',
      name: a.name,
      contentType: a.contentType,
      contentBytes: a.contentBytes,
      ...(a.contentId ? { contentId: a.contentId, isInline: a.isInline ?? true } : {}),
    }));
  }
  return out;
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
  const members: NormalizedChat['members'] = (c.members ?? [])
    .filter((m) => m.userId && m.userId !== myId)
    .map((m) => ({
      id: m.userId!,
      displayName: m.displayName ?? m.email ?? m.userId!,
      email: m.email ?? null,
    }))
    .sort((a, b) => a.displayName.localeCompare(b.displayName) || a.id.localeCompare(b.id));
  // Bot 1:1 chats: Graph omits the bot from `members` entirely (only the human
  // user is listed). Synthesize a bot member from lastMessagePreview.from so the
  // sidebar/header have something to render; the friendly name & icon are
  // resolved later via Teams MT (see enrichBotChats in index.ts).
  const lpApp = c.lastMessagePreview?.from?.application;
  if (members.length === 0 && c.chatType === 'oneOnOne' && lpApp?.id) {
    members.push({ id: lpApp.id, displayName: lpApp.displayName ?? 'Bot', email: null, isBot: true });
  }
  const lp = c.lastMessagePreview;
  const rawBody = lp?.body?.content ?? '';
  const lastMsgAt = lp?.createdDateTime ?? null;
  const readAt = c.viewpoint?.lastMessageReadDateTime ?? null;
  const lastFromMe = lp?.from?.user?.id === myId;
  const unread = !!lastMsgAt && !lastFromMe && (!readAt || Date.parse(readAt) < Date.parse(lastMsgAt));

  let preview = stripHtml(rawBody).slice(0, 200);
  if (!preview) {
    if (lp?.eventDetail) preview = describeEvent(lp.eventDetail) ?? '';
    else if (/<attachment\b/i.test(rawBody)) preview = 'sent a card';
    else if (/<img\b/i.test(rawBody)) preview = 'sent an image';
  }

  return {
    id: c.id,
    type: c.chatType,
    topic: c.topic ?? (c.chatType === 'oneOnOne' ? members[0]?.displayName ?? null : null),
    members,
    lastUpdated: lastMsgAt ?? c.lastUpdatedDateTime ?? null,
    lastMessagePreview: preview || null,
    lastMessageFrom:
      lp?.from?.user?.displayName ?? lp?.from?.application?.displayName ?? null,
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

function normalizeReactions(reactions: GraphReaction[] | undefined, myId: string | null): NormalizedReaction[] {
  if (!reactions?.length) return [];
  const grouped = new Map<string, NormalizedReaction>();
  for (const r of reactions) {
    const type = r.reactionType ?? 'like';
    const emoji = REACTION_EMOJI[type] ?? type;
    let g = grouped.get(emoji);
    if (!g) {
      g = { emoji, type, count: 0, users: [], mine: false };
      grouped.set(emoji, g);
    }
    g.count++;
    const uid = r.user?.user?.id;
    const name = r.user?.user?.displayName;
    // Graph returns displayName:null for reactors; store the id so the frontend
    // can resolve it against the chat's member list.
    if (uid) g.users.push(name || uid);
    else if (name) g.users.push(name);
    if (myId && uid === myId) g.mine = true;
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

function parseForwarded(content: string | undefined): NormalizedMessage['forwarded'] {
  if (!content) return null;
  try {
    const j = JSON.parse(content) as {
      originalMessage?: {
        body?: { content?: string };
        from?: { user?: { displayName?: string } };
        createdDateTime?: string;
      };
      messagePreview?: string;
      messageSender?: { user?: { displayName?: string } };
      originalCreatedDateTime?: string;
    };
    const om = j.originalMessage;
    return {
      senderName: om?.from?.user?.displayName ?? j.messageSender?.user?.displayName ?? null,
      text: om?.body?.content ? stripHtml(om.body.content) : j.messagePreview ? stripHtml(j.messagePreview) : null,
      originalDate: om?.createdDateTime ?? j.originalCreatedDateTime ?? null,
    };
  } catch {
    return null;
  }
}

function describeEvent(ed: GraphMessage['eventDetail']): string | null {
  if (!ed) return null;
  const t = String(ed['@odata.type'] ?? '').replace('#microsoft.graph.', '');
  const names = (arr: unknown): string =>
    Array.isArray(arr)
      ? (arr as Array<{ displayName?: string; user?: { displayName?: string } }>)
          .map((m) => m.displayName ?? m.user?.displayName)
          .filter(Boolean)
          .join(', ')
      : '';
  const initiator =
    (ed as { initiator?: { user?: { displayName?: string } } }).initiator?.user?.displayName ?? null;
  switch (t) {
    case 'membersAddedEventMessageDetail':
      return `${initiator ? `${initiator} added ` : 'Added '}${names((ed as { members?: unknown }).members)}`;
    case 'membersDeletedEventMessageDetail':
      return `${initiator ? `${initiator} removed ` : 'Removed '}${names((ed as { members?: unknown }).members)}`;
    case 'membersJoinedEventMessageDetail':
      return `${names((ed as { members?: unknown }).members) || 'Someone'} joined`;
    case 'membersLeftEventMessageDetail':
      return `${names((ed as { members?: unknown }).members) || 'Someone'} left`;
    case 'chatRenamedEventMessageDetail':
      return `${initiator ?? 'Chat'} renamed the chat to “${(ed as { chatDisplayName?: string }).chatDisplayName ?? ''}”`;
    case 'callStartedEventMessageDetail':
      return `${initiator ?? 'Call'} started a call`;
    case 'callEndedEventMessageDetail': {
      const dur = (ed as { callDuration?: string }).callDuration;
      return `Call ended${dur ? ` · ${dur.replace(/^PT/i, '').toLowerCase()}` : ''}`;
    }
    case 'callRecordingEventMessageDetail':
      return 'Recording available';
    case 'callTranscriptEventMessageDetail':
      return 'Transcript available';
    case 'messagePinnedEventMessageDetail':
      return `${initiator ?? 'Someone'} pinned a message`;
    case 'meetingPolicyUpdatedEventMessageDetail':
      return 'Meeting options updated';
    default:
      return t.replace(/EventMessageDetail$/, '').replace(/([A-Z])/g, ' $1').trim() || 'System event';
  }
}

export function normalizeMessage(m: GraphMessage, myId: string | null): NormalizedMessage {
  const fromUser = m.from?.user ?? null;
  const fromAppObj = m.from?.application ?? null;
  const fromId = fromUser?.id ?? fromAppObj?.id ?? null;
  const contentType = m.body?.contentType === 'html' ? 'html' : 'text';
  const raw = m.body?.content ?? '';
  const hostedImages = contentType === 'html' ? extractHostedImages(raw) : [];
  const segments = contentType === 'html'
    ? parseHtmlBody(raw, m.mentions)
    : (raw ? [{ type: 'text' as const, text: raw }] : []);
  const all = m.attachments ?? [];
  const refAttachment = all.find((a) => a.contentType === 'messageReference');
  const fwdAttachment = all.find((a) => a.contentType === 'forwardedMessageReference');
  const replyTo = refAttachment ? parseMessageReference(refAttachment.content) : null;
  const forwarded = fwdAttachment ? parseForwarded(fwdAttachment.content) : null;
  const files = all
    .filter((a) => a.contentType === 'reference')
    .map((a) => ({ name: a.name ?? 'file', url: a.contentUrl ?? null, contentType: a.contentType ?? null }));
  const cards = all
    .filter((a) => (a.contentType ?? '').startsWith('application/vnd.microsoft.card.') && a.content)
    .map((a) => ({ id: a.id ?? null, name: a.name ?? null, contentJson: a.content! }));
  const handled = new Set(['messageReference', 'forwardedMessageReference', 'reference']);
  const attachments = all.filter(
    (a) => !handled.has(a.contentType ?? '') && !(a.contentType ?? '').startsWith('application/vnd.microsoft.card.'),
  );
  const systemEvent = m.messageType && m.messageType !== 'message' ? describeEvent(m.eventDetail) : null;
  return {
    id: m.id,
    chatId: m.chatId ?? '',
    createdDateTime: m.createdDateTime ?? null,
    fromId,
    fromName: fromUser?.displayName ?? fromAppObj?.displayName ?? null,
    fromApp: !!fromAppObj && !fromUser,
    fromMe: !!myId && fromUser?.id === myId,
    contentType,
    text: contentType === 'html' ? stripHtml(raw) : raw,
    segments,
    hostedImages,
    replyTo,
    forwarded,
    files,
    cards,
    attachments: attachments.map((a) => ({
      name: a.name ?? null,
      contentType: a.contentType ?? null,
      url: a.contentUrl ?? null,
    })),
    systemEvent,
    reactions: normalizeReactions(m.reactions, myId),
    deleted: !!m.deletedDateTime,
  };
}
