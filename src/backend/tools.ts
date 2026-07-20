import type { PluginAPI, GraphUser, GraphChatType, GraphTeam, GraphChannel, GraphMessage, MailAddress } from '../shared/types.js';
import { GraphClient, normalizeChat, normalizeMessage } from './graph-client.js';
import * as tokenCache from './token-cache.js';
import { buildMessageBody, withMessageRef, withFileReferences, type PendingImage } from '../shared/markdown.js';
import {
  setForcedAvailability,
  setStatusNote,
  getMyPresence,
  invokeMessageback,
  invokeTask,
  invokeExecute,
  type UpsAvailability,
} from './ic3-client.js';
import { getLogger } from './logger-singleton.js';
import * as hostedContentCache from './hosted-content-cache.js';
import { readFile } from 'fs/promises';
import { homedir } from 'os';
import { resolve, extname, basename } from 'path';
import { randomUUID } from 'crypto';

const AVAIL_VALUES = ['Available', 'Busy', 'DoNotDisturb', 'BeRightBack', 'Away', 'Offline'] as const;

/** Extension → image MIME type, used to sniff images for the hostedContents path. */
const IMAGE_MIME_BY_EXT: Record<string, string> = {
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.gif': 'image/gif',
  '.webp': 'image/webp',
  '.bmp': 'image/bmp',
  '.svg': 'image/svg+xml',
  '.tif': 'image/tiff',
  '.tiff': 'image/tiff',
  '.ico': 'image/x-icon',
  '.heic': 'image/heic',
  '.heif': 'image/heif',
};

/** Extension → MIME for common non-image file attachments. Falls back to
 * application/octet-stream; the value is only advisory for the upload PUT. */
const FILE_MIME_BY_EXT: Record<string, string> = {
  '.pdf': 'application/pdf',
  '.txt': 'text/plain',
  '.md': 'text/markdown',
  '.csv': 'text/csv',
  '.json': 'application/json',
  '.xml': 'application/xml',
  '.html': 'text/html',
  '.zip': 'application/zip',
  '.gz': 'application/gzip',
  '.tar': 'application/x-tar',
  '.doc': 'application/msword',
  '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  '.xls': 'application/vnd.ms-excel',
  '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  '.ppt': 'application/vnd.ms-powerpoint',
  '.pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  '.mp4': 'video/mp4',
  '.mov': 'video/quicktime',
  '.mp3': 'audio/mpeg',
  '.wav': 'audio/wav',
  '.dmg': 'application/x-apple-diskimage',
};

/** Expand a leading `~` / `~/` to the user's home directory, then make absolute. */
function expandPath(p: string): string {
  let out = p;
  if (out === '~') out = homedir();
  else if (out.startsWith('~/')) out = resolve(homedir(), out.slice(2));
  return resolve(out);
}

/** Read a local file with ~ expansion and clear ENOENT/EISDIR errors. */
async function readLocalFile(raw: string): Promise<{ abs: string; buf: Buffer; name: string; ext: string }> {
  const abs = expandPath(raw);
  let buf: Buffer;
  try {
    buf = await readFile(abs);
  } catch (e) {
    const code = (e as NodeJS.ErrnoException).code;
    throw new Error(
      code === 'ENOENT'
        ? `File not found: ${raw} (resolved to ${abs})`
        : code === 'EISDIR'
          ? `Path is a directory, not a file: ${raw}`
          : `Could not read ${raw}: ${(e as Error).message}`,
    );
  }
  return { abs, buf, name: basename(abs), ext: extname(abs) };
}

/** Sniff the image MIME type from magic bytes, falling back to the extension. */
function sniffImageMime(buf: Buffer, ext: string): string | null {
  if (buf.length >= 8 && buf[0] === 0x89 && buf[1] === 0x50 && buf[2] === 0x4e && buf[3] === 0x47) return 'image/png';
  if (buf.length >= 3 && buf[0] === 0xff && buf[1] === 0xd8 && buf[2] === 0xff) return 'image/jpeg';
  if (buf.length >= 6 && buf[0] === 0x47 && buf[1] === 0x49 && buf[2] === 0x46) return 'image/gif';
  if (buf.length >= 12 && buf.toString('ascii', 0, 4) === 'RIFF' && buf.toString('ascii', 8, 12) === 'WEBP') return 'image/webp';
  if (buf.length >= 2 && buf[0] === 0x42 && buf[1] === 0x4d) return 'image/bmp';
  return IMAGE_MIME_BY_EXT[ext.toLowerCase()] ?? null;
}

/** Best-effort MIME for any file: image sniff first, then extension map, else octet-stream. */
function guessFileMime(buf: Buffer, ext: string): string {
  return sniffImageMime(buf, ext) ?? FILE_MIME_BY_EXT[ext.toLowerCase()] ?? 'application/octet-stream';
}

/** Read image files from disk and turn them into PendingImage entries for the
 * hostedContents pipeline. Throws with a clear message on missing/non-image files. */
async function readImagePaths(paths: string[], startIndex: number): Promise<PendingImage[]> {
  const out: PendingImage[] = [];
  for (let i = 0; i < paths.length; i++) {
    const raw = paths[i];
    const { buf, name, ext } = await readLocalFile(raw);
    const mime = sniffImageMime(buf, ext);
    if (!mime) {
      throw new Error(
        `${raw} is not a recognized image. Use filePaths to attach non-image files.`,
      );
    }
    out.push({
      id: `img${startIndex + i}`,
      contentType: mime,
      contentBytes: buf.toString('base64'),
      name,
    });
  }
  return out;
}

/** A file uploaded to OneDrive and ready to attach as a Teams `reference` card. */
interface FileAttachment {
  id: string;
  contentUrl: string;
  name: string;
}

/** Upload each local file to OneDrive, create an org-view share link, and return
 * the reference-attachment descriptors. Renders as file cards in Teams. */
async function prepareFileAttachments(client: GraphClient, paths: string[]): Promise<FileAttachment[]> {
  const out: FileAttachment[] = [];
  for (const raw of paths) {
    const { buf, name, ext } = await readLocalFile(raw);
    const mime = guessFileMime(buf, ext);
    const item = await client.uploadDriveFile(name, buf, mime);
    const contentUrl = await client.createShareLink(item.id);
    out.push({ id: randomUUID(), contentUrl, name: item.name });
  }
  return out;
}

/** Teams caps the whole chat-message payload at 4 MiB. Inline images ride in the
 * POST body as base64 hostedContents, so their combined base64 size (plus the body
 * and JSON envelope) must stay under it. Budget conservatively to leave headroom. */
const INLINE_IMAGE_BUDGET_BYTES = 3_900_000;

/** Map an image MIME type to a file extension for a synthesized filename. */
function imageExtForMime(mime: string): string {
  const m: Record<string, string> = {
    'image/png': 'png', 'image/jpeg': 'jpg', 'image/gif': 'gif', 'image/webp': 'webp',
    'image/bmp': 'bmp', 'image/svg+xml': 'svg', 'image/tiff': 'tif', 'image/heic': 'heic', 'image/heif': 'heif',
  };
  return m[mime] ?? 'png';
}

/**
 * Keep inline images under Teams' 4 MiB per-message payload cap. If their combined
 * base64 size exceeds the budget, offload ALL of them to OneDrive and return them as
 * file-card attachments instead so the message still sends. (All-or-nothing keeps the
 * result predictable rather than silently sending some inline and some as cards.)
 */
async function fitInlineImages(
  client: GraphClient,
  images: PendingImage[],
): Promise<{ images: PendingImage[]; fileAttachments: FileAttachment[] }> {
  const total = images.reduce((n, img) => n + img.contentBytes.length, 0);
  if (images.length === 0 || total <= INLINE_IMAGE_BUDGET_BYTES) {
    return { images, fileAttachments: [] };
  }
  const fileAttachments: FileAttachment[] = [];
  for (const img of images) {
    const bytes = Buffer.from(img.contentBytes, 'base64');
    const name = img.name && img.name.trim() ? img.name : `image-${randomUUID().slice(0, 8)}.${imageExtForMime(img.contentType)}`;
    const item = await client.uploadDriveFile(name, bytes, img.contentType);
    const contentUrl = await client.createShareLink(item.id);
    fileAttachments.push({ id: randomUUID(), contentUrl, name: item.name });
  }
  return { images: [], fileAttachments };
}

interface CardAction {
  id: string | null;
  title: string | null;
  type: string;
  data?: Record<string, unknown>;
  verb?: string;
  url?: string;
}

function walkActions(node: unknown, out: CardAction[]): void {
  if (Array.isArray(node)) { for (const c of node) walkActions(c, out); return; }
  if (!node || typeof node !== 'object') return;
  const n = node as Record<string, unknown>;
  if (typeof n.type === 'string' && n.type.startsWith('Action.')) {
    out.push({
      id: typeof n.id === 'string' ? n.id : null,
      title: typeof n.title === 'string' ? n.title : null,
      type: n.type,
      data: (n.data && typeof n.data === 'object') ? (n.data as Record<string, unknown>) : undefined,
      verb: typeof n.verb === 'string' ? n.verb : undefined,
      url: typeof n.url === 'string' ? n.url : undefined,
    });
  }
  for (const v of Object.values(n)) walkActions(v, out);
}

export type ToolDefinition = {
  name: string;
  description: string;
  inputSchema: unknown;
  execute: (input: unknown, context?: unknown) => Promise<unknown>;
};

export interface ToolDeps {
  api: PluginAPI;
  ensureAuthenticated: () => Promise<GraphClient>;
}

function errResult(err: unknown): { error: string } {
  const message = err instanceof Error ? err.message : String(err);
  getLogger().error(`Tool error: ${message}`);
  return { error: message };
}

function clampTop(v: unknown, def: number, max: number): number {
  const n = Number.isFinite(v as number) ? Math.floor(v as number) : def;
  return Math.min(Math.max(n, 1), max);
}

/** Default true unless explicitly set to false. */
function wantConcise(input: unknown): boolean {
  const c = (input as { concise?: unknown } | null)?.concise;
  return c !== false;
}

/** A model-visible content part carried on a tool result (see host tool-model-content). */
type ModelContentPart =
  | { type: 'text'; text: string }
  | { type: 'image'; data: string; mediaType: string }
  | { type: 'file'; data: string; mediaType: string; filename?: string };

/** Classify a media type into the model-content part kind the model can consume. */
function partForMedia(base64: string, mediaType: string, filename?: string): ModelContentPart {
  if (mediaType.startsWith('image/')) return { type: 'image', data: base64, mediaType };
  return { type: 'file', data: base64, mediaType, ...(filename ? { filename } : {}) };
}

type MailSummary = {
  id: string; conversationId: string | null; subject: string; from: MailAddress | null; toRecipients: MailAddress[];
  receivedDateTime: string | null; isRead: boolean; hasAttachments: boolean;
  flagged: boolean; importance: string; bodyPreview: string;
};

function conciseMail(m: MailSummary) {
  return {
    id: m.id,
    conversationId: m.conversationId,
    subject: m.subject,
    from: m.from ? `${m.from.name ?? ''} <${m.from.address}>`.trim() : null,
    to: m.toRecipients.length,
    received: m.receivedDateTime,
    unread: !m.isRead || undefined,
    hasAttachments: m.hasAttachments || undefined,
    flagged: m.flagged || undefined,
    importance: m.importance !== 'normal' ? m.importance : undefined,
    preview: (m.bodyPreview ?? '').slice(0, 150),
  };
}

function conciseChat(c: ReturnType<typeof normalizeChat>) {
  const names = c.members.map((m) => m.displayName);
  const shown = names.slice(0, 4).join(', ');
  const members = names.length > 4 ? `${shown}, +${names.length - 4}` : shown;
  return {
    id: c.id,
    type: c.type,
    topic: c.topic,
    members,
    lastUpdated: c.lastUpdated,
    lastMessagePreview: c.lastMessagePreview,
    unread: c.unread || undefined,
  };
}

function conciseMessage(m: ReturnType<typeof normalizeMessage>) {
  return {
    id: m.id,
    chatId: m.chatId,
    from: m.fromName,
    fromMe: m.fromMe || undefined,
    at: m.createdDateTime,
    text: m.text,
    replyTo: m.replyTo ? { from: m.replyTo.senderName, preview: m.replyTo.text } : undefined,
    hasAttachments:
      (m.attachments.length + m.files.length + m.cards.length + m.hostedImages.length) > 0 || undefined,
    reactions: m.reactions.length ? m.reactions : undefined,
    systemEvent: m.systemEvent ?? undefined,
    deleted: m.deleted || undefined,
  };
}

interface ChannelMessageRef {
  teamId: string;
  channelId: string;
  messageId: string;
  parentMessageId: string;
  webUrl?: string;
}

/** Parse the canonical Teams /l/message deep link used by channel posts and replies. */
function parseTeamsMessageUrl(raw: string): ChannelMessageRef {
  let url: URL;
  try {
    url = new URL(raw);
  } catch {
    throw new Error('Invalid Teams message URL. Expected https://teams.microsoft.com/l/message/...');
  }
  const parts = url.pathname.split('/').filter(Boolean);
  const messageIdx = parts.findIndex((p, i) => p === 'message' && parts[i - 1] === 'l');
  if (messageIdx < 0 || !parts[messageIdx + 1] || !parts[messageIdx + 2]) {
    throw new Error('This is not a Teams channel-message link (/l/message/{channelId}/{messageId}).');
  }
  const teamId = url.searchParams.get('groupId')?.trim();
  if (!teamId) throw new Error('The Teams message link is missing its groupId (team id).');
  const channelId = decodeURIComponent(parts[messageIdx + 1]);
  const messageId = decodeURIComponent(parts[messageIdx + 2]);
  const parentMessageId = url.searchParams.get('parentMessageId')?.trim() || messageId;
  return { teamId, channelId, messageId, parentMessageId, webUrl: url.toString() };
}

function resolveChannelMessageRef(input: {
  webUrl?: string;
  teamId?: string;
  channelId?: string;
  messageId?: string;
  parentMessageId?: string;
}): ChannelMessageRef {
  if (input.webUrl?.trim()) return parseTeamsMessageUrl(input.webUrl.trim());
  const teamId = input.teamId?.trim();
  const channelId = input.channelId?.trim();
  const messageId = input.messageId?.trim();
  if (!teamId || !channelId || !messageId) {
    throw new Error('Pass webUrl, or pass teamId + channelId + messageId.');
  }
  return {
    teamId,
    channelId,
    messageId,
    parentMessageId: input.parentMessageId?.trim() || messageId,
  };
}

function channelMessageOutput(
  m: GraphMessage,
  ref: Pick<ChannelMessageRef, 'teamId' | 'channelId' | 'parentMessageId'>,
  myId: string | null,
  concise: boolean,
) {
  const n = normalizeMessage(m, myId);
  const base = {
    id: m.id,
    parentMessageId: m.replyToId || ref.parentMessageId || m.id,
    teamId: ref.teamId,
    channelId: ref.channelId,
    from: n.fromName,
    fromId: n.fromId,
    fromMe: n.fromMe || undefined,
    createdDateTime: n.createdDateTime,
    lastModifiedDateTime: m.lastModifiedDateTime ?? null,
    subject: m.subject ?? null,
    text: n.text,
    webUrl: m.webUrl ?? null,
    deleted: n.deleted || undefined,
    reactions: n.reactions.length ? n.reactions : undefined,
  };
  if (concise) return base;
  return {
    ...base,
    contentType: n.contentType,
    bodyContent: m.body?.content ?? '',
    segments: n.segments,
    mentions: m.mentions ?? [],
    attachments: m.attachments ?? [],
    hostedImages: n.hostedImages,
    files: n.files,
    cards: n.cards,
  };
}

function channelErrResult(err: unknown): { error: string } {
  const base = errResult(err);
  if (/\b403\b/.test(base.error)) {
    base.error += ' Channel operations require the signed-in Graph token to include the relevant Team.ReadBasic.All, Channel.ReadBasic.All, ChannelMessage.Read.All, or ChannelMessage.ReadWrite delegated permission. Use auth-status to inspect scopes.';
  }
  return base;
}

function matchesQuery(values: Array<string | null | undefined>, query?: string): boolean {
  if (!query?.trim()) return true;
  const q = query.trim().toLowerCase();
  return values.some((v) => (v ?? '').toLowerCase().includes(q));
}

function pickUnique<T>(items: T[], ref: string, label: (item: T) => string, kind: string): T {
  const q = ref.trim().toLowerCase();
  const exact = items.filter((item) => label(item).toLowerCase() === q);
  if (exact.length === 1) return exact[0];
  const partial = items.filter((item) => label(item).toLowerCase().includes(q));
  if (partial.length === 1) return partial[0];
  const choices = (exact.length ? exact : partial).map(label);
  if (choices.length > 1) throw new Error(`Ambiguous ${kind} "${ref}" — matches: ${choices.join(', ')}`);
  throw new Error(`No ${kind} found matching "${ref}".`);
}

async function resolveTeam(client: GraphClient, ref: string): Promise<GraphTeam> {
  const q = ref.trim();
  if (/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(q)) return client.getTeam(q);
  return pickUnique(await client.listJoinedTeams(), q, (t) => t.displayName ?? t.id, 'joined team');
}

async function mapConcurrent<T, R>(items: T[], limit: number, fn: (item: T) => Promise<R>): Promise<R[]> {
  const out = new Array<R>(items.length);
  let next = 0;
  await Promise.all(Array.from({ length: Math.min(limit, items.length) }, async () => {
    while (true) {
      const i = next++;
      if (i >= items.length) return;
      out[i] = await fn(items[i]);
    }
  }));
  return out;
}

async function resolveChannel(
  client: GraphClient,
  channelRef: string,
  teamRef?: string,
): Promise<{ team: GraphTeam; channel: GraphChannel }> {
  const teams = teamRef ? [await resolveTeam(client, teamRef)] : await client.listJoinedTeams();
  if (channelRef.startsWith('19:')) {
    if (teams.length !== 1) throw new Error('A team id or name is required when channel is supplied as an id.');
    const channels = await client.listChannels(teams[0].id);
    const channel = channels.find((c) => c.id === channelRef);
    if (!channel) throw new Error(`Channel ${channelRef} is not visible in team ${teams[0].displayName ?? teams[0].id}.`);
    return { team: teams[0], channel };
  }
  const nested = await mapConcurrent(teams, 4, async (team) => ({ team, channels: await client.listChannels(team.id) }));
  const candidates = nested.flatMap(({ team, channels }) => channels.map((channel) => ({ team, channel })));
  const q = channelRef.trim().toLowerCase();
  const exact = candidates.filter((x) => (x.channel.displayName ?? x.channel.id).toLowerCase() === q);
  const partial = candidates.filter((x) => (x.channel.displayName ?? x.channel.id).toLowerCase().includes(q));
  const matches = exact.length ? exact : partial;
  if (matches.length === 1) return matches[0];
  if (matches.length > 1) {
    throw new Error(
      `Ambiguous channel "${channelRef}" — matches: ${matches.map((x) =>
        `${x.channel.displayName ?? x.channel.id} (${x.team.displayName ?? x.team.id})`).join(', ')}`,
    );
  }
  throw new Error(`No channel found matching "${channelRef}".`);
}

function namesFromTeamsUrl(webUrl: string | null): { teamName: string | null; channelName: string | null } {
  if (!webUrl) return { teamName: null, channelName: null };
  try {
    const u = new URL(webUrl);
    return { teamName: u.searchParams.get('teamName'), channelName: u.searchParams.get('channelName') };
  } catch {
    return { teamName: null, channelName: null };
  }
}

const CHANNEL_MESSAGE_REF_SCHEMA = {
  webUrl: { type: 'string', description: 'A Teams /l/message deep link. When supplied, IDs are parsed from it.' },
  teamId: { type: 'string', description: 'Team id (the groupId in a Teams message link).' },
  channelId: { type: 'string', description: 'Channel id, usually 19:...@thread.tacv2.' },
  messageId: { type: 'string', description: 'Channel post or reply id.' },
  parentMessageId: { type: 'string', description: 'Root post id when messageId identifies a reply. Omit for a root post.' },
};

async function buildOutgoing(
  client: GraphClient,
  chatId: string,
  input: {
    text: string;
    contentType?: 'text' | 'html';
    replyToMessageId?: string;
    images?: PendingImage[];
    fileAttachments?: FileAttachment[];
  },
): Promise<Record<string, unknown>> {
  let body = input.contentType
    ? { body: { contentType: input.contentType, content: input.text } }
    : (buildMessageBody(input.text, input.images ?? []) as {
        body: { contentType: 'text' | 'html'; content: string };
        attachments?: unknown[];
      });
  if (input.replyToMessageId) {
    let preview: string | null = null;
    let sender: string | null = null;
    try {
      const ref = await client.getMessage(chatId, input.replyToMessageId);
      preview = (ref.body?.content ?? '').replace(/<[^>]+>/g, '').slice(0, 200) || null;
      sender = ref.from?.user?.displayName ?? ref.from?.application?.displayName ?? null;
    } catch { /* best-effort */ }
    body = withMessageRef(body, {
      contentType: 'messageReference',
      id: input.replyToMessageId,
      messageId: input.replyToMessageId,
      messagePreview: preview,
      messageSender: sender ? { user: { displayName: sender } } : null,
    });
  }
  if (input.fileAttachments?.length) {
    body = withFileReferences(body, input.fileAttachments);
  }
  return body;
}

async function resolveUser(client: GraphClient, ref: string): Promise<GraphUser> {
  const q = ref.trim();
  if (/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(q)) {
    return { id: q };
  }
  if (q.includes('@')) {
    const u = await client.getUserByEmail(q);
    if (u) return u;
  }
  const matches = await client.findUsers(q, 5);
  if (matches.length === 0) throw new Error(`No user found matching "${q}"`);
  if (matches.length > 1) {
    const list = matches.map((m) => `${m.displayName} <${m.userPrincipalName ?? m.mail}>`).join('; ');
    throw new Error(`Ambiguous user "${q}" — matches: ${list}. Use an exact email/UPN.`);
  }
  return matches[0];
}

export function buildMsgraphTools(deps: ToolDeps): ToolDefinition[] {
  const { api, ensureAuthenticated } = deps;

  return [
    {
      name: 'auth-status',
      description:
        'Check Microsoft Graph / Teams authentication status. Returns whether the user is signed in, their email, display name, minutes until the access token expires, and whether a refresh token is available for silent renewal.',
      inputSchema: { type: 'object', properties: {}, additionalProperties: false },
      execute: async () => {
        try {
          await ensureAuthenticated();
          return {
            isAuthenticated: true,
            email: tokenCache.getEmail(),
            displayName: tokenCache.getDisplayName(),
            objectId: tokenCache.getObjectId(),
            minutesRemaining: tokenCache.minutesRemaining(),
            hasRefreshToken: tokenCache.hasRefreshToken(),
            scopes: tokenCache.get()?.scopes ?? null,
          };
        } catch (err) {
          return { isAuthenticated: false, error: err instanceof Error ? err.message : String(err) };
        }
      },
    },

    {
      name: 'find-user',
      description:
        'Search the tenant directory for people by name, email, or UPN. Use this to resolve a colleague to their AAD id before sending a DM or creating a group chat.',
      inputSchema: {
        type: 'object',
        properties: {
          query: { type: 'string', description: 'Name fragment, email, or UPN to search for.' },
          top: { type: 'number', description: 'Max results (default 10, max 25).' },
        },
        required: ['query'],
        additionalProperties: false,
      },
      execute: async (input) => {
        try {
          const client = await ensureAuthenticated();
          const { query, top } = input as { query: string; top?: number };
          const users = await client.findUsers(query, clampTop(top, 10, 25));
          return {
            success: true,
            count: users.length,
            users: users.map((u) => ({
              id: u.id,
              displayName: u.displayName ?? null,
              email: u.mail ?? u.userPrincipalName ?? null,
              upn: u.userPrincipalName ?? null,
            })),
          };
        } catch (err) {
          return errResult(err);
        }
      },
    },

    {
      name: 'list-teams',
      description:
        "List or find Microsoft Teams the signed-in user has joined. Use this when a request names a team but does not provide its id.",
      inputSchema: {
        type: 'object',
        properties: {
          query: { type: 'string', description: 'Optional case-insensitive substring matched against team name and description.' },
        },
        additionalProperties: false,
      },
      execute: async (input) => {
        try {
          const client = await ensureAuthenticated();
          const query = (input as { query?: string } | null)?.query;
          const teams = (await client.listJoinedTeams()).filter((t) => matchesQuery([t.displayName, t.description], query));
          return {
            success: true,
            count: teams.length,
            teams: teams.map((t) => ({
              id: t.id,
              name: t.displayName ?? t.id,
              description: t.description ?? null,
              webUrl: t.webUrl ?? null,
              archived: t.isArchived || undefined,
            })),
          };
        } catch (err) {
          return channelErrResult(err);
        }
      },
    },

    {
      name: 'list-channels',
      description:
        'List or find channels visible to the signed-in user. Pass a team id/name to restrict the search, or omit it to search channels across all joined teams. Returns actionable teamId and channelId values.',
      inputSchema: {
        type: 'object',
        properties: {
          team: { type: 'string', description: 'Optional team id or unambiguous team-name substring.' },
          query: { type: 'string', description: 'Optional case-insensitive substring matched against channel name and description.' },
        },
        additionalProperties: false,
      },
      execute: async (input) => {
        try {
          const client = await ensureAuthenticated();
          const { team: teamRef, query } = (input ?? {}) as { team?: string; query?: string };
          const teams = teamRef ? [await resolveTeam(client, teamRef)] : await client.listJoinedTeams();
          const nested = await mapConcurrent(teams, 4, async (team) => ({
            team,
            channels: await client.listChannels(team.id),
          }));
          const channels = nested.flatMap(({ team, channels: teamChannels }) =>
            teamChannels
              .filter((c) => matchesQuery([c.displayName, c.description], query))
              .map((c) => ({
                teamId: team.id,
                teamName: team.displayName ?? team.id,
                channelId: c.id,
                channelName: c.displayName ?? c.id,
                description: c.description ?? null,
                membershipType: c.membershipType ?? null,
                webUrl: c.webUrl ?? null,
              })),
          );
          return { success: true, count: channels.length, channels };
        } catch (err) {
          return channelErrResult(err);
        }
      },
    },

    {
      name: 'list-channel-messages',
      description:
        "Read recent root posts from a Teams channel, optionally including replies and/or only the signed-in user's messages. Team and channel may be ids or unambiguous name fragments. Useful when the exact search terms are unknown.",
      inputSchema: {
        type: 'object',
        properties: {
          team: { type: 'string', description: 'Team id or unambiguous team-name substring.' },
          channel: { type: 'string', description: 'Channel id or unambiguous channel-name substring.' },
          top: { type: 'number', description: 'Number of root posts to fetch (default 25, max 50).' },
          includeReplies: { type: 'boolean', description: 'Include replies nested under those root posts (default false).' },
          fromMe: { type: 'boolean', description: 'Only return messages authored by the signed-in user.' },
          concise: { type: 'boolean', description: 'Compact output by default. Set false to include exact HTML body content, segments, mentions, and attachments.' },
        },
        required: ['team', 'channel'],
        additionalProperties: false,
      },
      execute: async (input) => {
        try {
          const client = await ensureAuthenticated();
          const { team: teamRef, channel: channelRef, top, includeReplies, fromMe } = input as {
            team: string; channel: string; top?: number; includeReplies?: boolean; fromMe?: boolean;
          };
          const { team, channel } = await resolveChannel(client, channelRef, teamRef);
          const { messages: roots } = await client.listChannelMessages(
            team.id,
            channel.id,
            clampTop(top, 25, 50),
            !!includeReplies,
          );
          const myId = tokenCache.getObjectId();
          const messages: ReturnType<typeof channelMessageOutput>[] = [];
          for (const root of roots) {
            const rootRef = { teamId: team.id, channelId: channel.id, parentMessageId: root.id };
            const normalizedRoot = normalizeMessage(root, myId);
            if (!fromMe || normalizedRoot.fromMe) messages.push(channelMessageOutput(root, rootRef, myId, wantConcise(input)));
            if (includeReplies) {
              for (const reply of root.replies ?? []) {
                if (!fromMe || normalizeMessage(reply, myId).fromMe) {
                  messages.push(channelMessageOutput(reply, rootRef, myId, wantConcise(input)));
                }
              }
            }
          }
          return {
            success: true,
            team: { id: team.id, name: team.displayName ?? team.id },
            channel: { id: channel.id, name: channel.displayName ?? channel.id },
            rootCount: roots.length,
            count: messages.length,
            messages,
          };
        } catch (err) {
          return channelErrResult(err);
        }
      },
    },

    {
      name: 'get-channel-message',
      description:
        'Read an exact Teams channel post or reply from a pasted /l/message link or explicit ids. Root-post links include the complete reply thread by default. Exact HTML is returned by default so formatting can be inspected and repaired.',
      inputSchema: {
        type: 'object',
        properties: {
          ...CHANNEL_MESSAGE_REF_SCHEMA,
          includeReplies: { type: 'boolean', description: 'For a root post, include every reply (default true).' },
          concise: { type: 'boolean', description: 'Set true for compact plain-text output. Default false preserves exact HTML.' },
        },
        additionalProperties: false,
      },
      execute: async (input) => {
        try {
          const client = await ensureAuthenticated();
          const opts = input as Parameters<typeof resolveChannelMessageRef>[0] & { includeReplies?: boolean; concise?: boolean };
          const ref = resolveChannelMessageRef(opts);
          const message = await client.getChannelMessage(ref.teamId, ref.channelId, ref.messageId, ref.parentMessageId);
          const myId = tokenCache.getObjectId();
          const concise = opts.concise === true;
          const isRoot = ref.parentMessageId === ref.messageId;
          const replies = isRoot && opts.includeReplies !== false
            ? await client.getChannelReplies(ref.teamId, ref.channelId, ref.messageId)
            : [];
          return {
            success: true,
            reference: ref,
            message: channelMessageOutput(message, ref, myId, concise),
            replies: replies.map((r) => channelMessageOutput(r, ref, myId, concise)),
          };
        } catch (err) {
          return channelErrResult(err);
        }
      },
    },

    {
      name: 'search-channel-messages',
      description:
        "Full-text search across Teams channel posts and replies visible to the signed-in user. Optionally scope by team/channel name or id and to the user's own messages. Results contain all ids needed by get-channel-message and edit-channel-message.",
      inputSchema: {
        type: 'object',
        properties: {
          query: { type: 'string', description: 'Search text or a Microsoft Search KQL query.' },
          team: { type: 'string', description: 'Optional team id or unambiguous team-name substring.' },
          channel: { type: 'string', description: 'Optional channel id or unambiguous channel-name substring.' },
          fromMe: { type: 'boolean', description: 'Only return messages authored by the signed-in user.' },
          top: { type: 'number', description: 'Maximum matching results (default 15, max 50).' },
        },
        required: ['query'],
        additionalProperties: false,
      },
      execute: async (input) => {
        try {
          const client = await ensureAuthenticated();
          const { query, team: teamRef, channel: channelRef, fromMe, top } = input as {
            query: string; team?: string; channel?: string; fromMe?: boolean; top?: number;
          };
          const max = clampTop(top, 15, 50);
          let teamId: string | null = null;
          let channelId: string | null = null;
          let resolvedNames: { teamName?: string; channelName?: string } = {};
          if (channelRef) {
            const resolved = await resolveChannel(client, channelRef, teamRef);
            teamId = resolved.team.id;
            channelId = resolved.channel.id;
            resolvedNames = {
              teamName: resolved.team.displayName ?? resolved.team.id,
              channelName: resolved.channel.displayName ?? resolved.channel.id,
            };
          } else if (teamRef) {
            const team = await resolveTeam(client, teamRef);
            teamId = team.id;
            resolvedNames.teamName = team.displayName ?? team.id;
          }

          const myId = tokenCache.getObjectId();
          const myName = tokenCache.getDisplayName();
          const results: Array<Record<string, unknown>> = [];
          for (let offset = 0; offset < 100 && results.length < max; offset += 25) {
            const page = await client.searchMessages(query, 25, offset);
            for (const hit of page) {
              let hitTeamId = hit.teamId;
              let hitChannelId = hit.channelId;
              let hitMessageId = hit.messageId;
              let hitParentId = hit.replyToId;
              if ((!hitTeamId || !hitChannelId || !hitMessageId) && hit.webUrl) {
                try {
                  const parsed = parseTeamsMessageUrl(hit.webUrl);
                  hitTeamId ||= parsed.teamId;
                  hitChannelId ||= parsed.channelId;
                  hitMessageId ||= parsed.messageId;
                  hitParentId ||= parsed.parentMessageId;
                } catch { /* a chat search result, not a channel deep link */ }
              }
              if (!hitTeamId || !hitChannelId || !hitMessageId) continue;
              if (teamId && hitTeamId !== teamId) continue;
              if (channelId && hitChannelId !== channelId) continue;
              const authoredByMe = !!myId && hit.fromId === myId || (!hit.fromId && !!myName && hit.from === myName);
              if (fromMe && !authoredByMe) continue;
              const urlNames = namesFromTeamsUrl(hit.webUrl);
              results.push({
                messageId: hitMessageId,
                parentMessageId: hitParentId || hitMessageId,
                teamId: hitTeamId,
                teamName: resolvedNames.teamName ?? urlNames.teamName,
                channelId: hitChannelId,
                channelName: resolvedNames.channelName ?? urlNames.channelName,
                from: hit.from,
                fromId: hit.fromId,
                fromMe: authoredByMe || undefined,
                createdDateTime: hit.createdDateTime,
                summary: hit.summary.replace(/<\/?c\d+>/gi, ''),
                webUrl: hit.webUrl,
              });
              if (results.length >= max) break;
            }
            if (page.length < 25) break;
          }
          return { success: true, count: results.length, results };
        } catch (err) {
          return channelErrResult(err);
        }
      },
    },

    {
      name: 'edit-channel-message',
      description:
        "Edit one of the signed-in user's own Teams channel posts or replies. Accepts a pasted Teams /l/message link or explicit ids. Markdown and @[Name](aad:<id>) mentions are converted to Teams HTML.",
      inputSchema: {
        type: 'object',
        properties: {
          ...CHANNEL_MESSAGE_REF_SCHEMA,
          text: { type: 'string', description: 'Replacement message body.' },
          contentType: { type: 'string', enum: ['text', 'html'], description: 'Force raw text or HTML. Omit to convert Markdown automatically.' },
        },
        required: ['text'],
        additionalProperties: false,
      },
      execute: async (input) => {
        try {
          const client = await ensureAuthenticated();
          const opts = input as Parameters<typeof resolveChannelMessageRef>[0] & { text: string; contentType?: 'text' | 'html' };
          const ref = resolveChannelMessageRef(opts);
          const current = await client.getChannelMessage(ref.teamId, ref.channelId, ref.messageId, ref.parentMessageId);
          const myId = tokenCache.getObjectId();
          if (!myId || current.from?.user?.id !== myId) {
            throw new Error(`Cannot edit channel message ${ref.messageId}: it was not authored by the signed-in user.`);
          }
          const p = opts.contentType
            ? { body: { contentType: opts.contentType, content: opts.text } as { contentType: 'text' | 'html'; content: string } }
            : buildMessageBody(opts.text);
          await client.editChannelMessage(ref.teamId, ref.channelId, ref.messageId, ref.parentMessageId, {
            body: p.body,
            ...('mentions' in p && p.mentions ? { mentions: p.mentions } : {}),
            ...('hostedContents' in p && p.hostedContents ? { hostedContents: p.hostedContents } : {}),
          });
          return { success: true, ...ref };
        } catch (err) {
          return channelErrResult(err);
        }
      },
    },

    {
      name: 'delete-channel-message',
      description:
        "Soft-delete one of the signed-in user's own Teams channel posts or replies. Accepts a pasted Teams /l/message link or explicit ids.",
      inputSchema: {
        type: 'object',
        properties: CHANNEL_MESSAGE_REF_SCHEMA,
        additionalProperties: false,
      },
      execute: async (input) => {
        try {
          const client = await ensureAuthenticated();
          const ref = resolveChannelMessageRef(input as Parameters<typeof resolveChannelMessageRef>[0]);
          const current = await client.getChannelMessage(ref.teamId, ref.channelId, ref.messageId, ref.parentMessageId);
          const myId = tokenCache.getObjectId();
          if (!myId || current.from?.user?.id !== myId) {
            throw new Error(`Cannot delete channel message ${ref.messageId}: it was not authored by the signed-in user.`);
          }
          await client.deleteChannelMessage(ref.teamId, ref.channelId, ref.messageId, ref.parentMessageId);
          return { success: true, ...ref, softDeleted: true };
        } catch (err) {
          return channelErrResult(err);
        }
      },
    },

    {
      name: 'list-chats',
      description:
        "List the signed-in user's Teams chats (both DMs and group chats), most-recently-active first. Optionally filter by chat type or by a member name/email substring. Returns chat id, type, topic/title, members, and a preview of the last message.",
      inputSchema: {
        type: 'object',
        properties: {
          chatType: {
            type: 'string',
            enum: ['oneOnOne', 'group', 'meeting'],
            description: 'Restrict to a single chat type. Omit to list all types.',
          },
          filter: {
            type: 'string',
            description:
              'Case-insensitive substring matched against chat topic and member display names/emails after fetching.',
          },
          top: { type: 'number', description: 'Max chats to fetch from Graph before filtering (default 50).' },
          concise: { type: 'boolean', description: 'Compact output (default true): trims each chat to id, type, topic, a short member summary, last update, preview and unread. Set false for full members[] and webUrl.' },
        },
        additionalProperties: false,
      },
      execute: async (input) => {
        try {
          const client = await ensureAuthenticated();
          const { chatType, filter, top } = (input ?? {}) as {
            chatType?: GraphChatType;
            filter?: string;
            top?: number;
          };
          const { chats: raw } = await client.listChats({ chatType, top: clampTop(top, 50, 50) });
          const myId = tokenCache.getObjectId();
          let chats = raw.map((c) => normalizeChat(c, myId));
          if (filter) {
            const f = filter.toLowerCase();
            chats = chats.filter(
              (c) =>
                (c.topic ?? '').toLowerCase().includes(f) ||
                c.members.some(
                  (m) =>
                    m.displayName.toLowerCase().includes(f) ||
                    (m.email ?? '').toLowerCase().includes(f),
                ),
            );
          }
          const out = wantConcise(input) ? chats.map(conciseChat) : chats;
          return { success: true, count: chats.length, chats: out };
        } catch (err) {
          return errResult(err);
        }
      },
    },

    {
      name: 'get-chat-messages',
      description:
        'Read recent messages from a Teams chat by chat id (newest first). HTML bodies are stripped to plain text. Use list-chats or send-dm first to obtain a chat id.',
      inputSchema: {
        type: 'object',
        properties: {
          chatId: { type: 'string', description: 'The chat id, e.g. 19:...@unq.gbl.spaces or 19:...@thread.v2' },
          top: { type: 'number', description: 'Max messages to return (default 25, max 50).' },
          concise: { type: 'boolean', description: 'Compact output (default true): id, from, at, text, and flags for replies/attachments/reactions. Set false for full segments[], hostedImages[], card JSON and attachment details.' },
        },
        required: ['chatId'],
        additionalProperties: false,
      },
      execute: async (input) => {
        try {
          const client = await ensureAuthenticated();
          const { chatId, top } = input as { chatId: string; top?: number };
          const { messages: msgs } = await client.getChatMessages(chatId, clampTop(top, 25, 50));
          const myId = tokenCache.getObjectId();
          const messages = msgs
            .filter((m) => m.messageType === 'message' || m.messageType == null)
            .map((m) => normalizeMessage(m, myId));
          const out = wantConcise(input) ? messages.map(conciseMessage) : messages;
          return { success: true, chatId, count: messages.length, messages: out };
        } catch (err) {
          return errResult(err);
        }
      },
    },

    {
      name: 'search-messages',
      description:
        "Full-text search across all Teams chat messages the signed-in user can access. Returns hit summaries with chat ids so you can follow up with get-chat-messages.",
      inputSchema: {
        type: 'object',
        properties: {
          query: { type: 'string', description: 'Search query (KQL supported).' },
          top: { type: 'number', description: 'Max hits (default 15, max 25).' },
        },
        required: ['query'],
        additionalProperties: false,
      },
      execute: async (input) => {
        try {
          const client = await ensureAuthenticated();
          const { query, top } = input as { query: string; top?: number };
          const hits = await client.searchMessages(query, clampTop(top, 15, 25));
          return { success: true, count: hits.length, hits };
        } catch (err) {
          return errResult(err);
        }
      },
    },

    {
      name: 'send-message',
      description:
        'Send a message to an existing Teams chat by chat id. For a person you have not yet resolved to a chat id, prefer send-dm instead.',
      inputSchema: {
        type: 'object',
        properties: {
          chatId: { type: 'string' },
          text: {
            type: 'string',
            description:
              'Message body. Markdown (**bold**, *italic*, ~~strike~~, `code`, ```fenced blocks```) is auto-converted to Teams HTML unless contentType is set. To @-mention someone, embed the token @[Display Name](aad:<their-AAD-object-id>) in the text (use find-user to get the id).',
          },
          contentType: {
            type: 'string',
            enum: ['text', 'html'],
            description: 'Force a specific body content type. Omit to auto-detect markdown / mentions.',
          },
          replyToMessageId: {
            type: 'string',
            description: 'Quote-reply to an existing message in the same chat (attaches a messageReference).',
          },
          images: {
            type: 'array',
            description: 'Inline images to embed as hostedContents. If the combined image size would exceed Teams\' ~4MB per-message limit, they are automatically uploaded and sent as file-card attachments instead.',
            items: {
              type: 'object',
              properties: {
                contentType: { type: 'string', description: 'e.g. image/png' },
                contentBytes: { type: 'string', description: 'Base64-encoded image bytes (no data: prefix).' },
                name: { type: 'string' },
              },
              required: ['contentType', 'contentBytes'],
            },
          },
          imagePaths: {
            type: 'array',
            description:
              'Local image file paths to attach inline (embedded as hostedContents). Preferred over `images` when the image is already a file on disk — avoids base64 round-tripping. Supports ~ expansion; MIME type is sniffed automatically. For non-image files use filePaths.',
            items: { type: 'string' },
          },
          filePaths: {
            type: 'array',
            description:
              'Local file paths to attach as Teams file cards (any type, any size — PDFs, docs, archives, etc.). Supports ~ expansion. Each file is uploaded to your OneDrive "Microsoft Teams Chat Files" folder, shared with an organization view link, and attached so recipients can open it. For inline images prefer imagePaths.',
            items: { type: 'string' },
          },
        },
        required: ['chatId', 'text'],
        additionalProperties: false,
      },
      execute: async (input) => {
        try {
          const client = await ensureAuthenticated();
          const { chatId, text, contentType, replyToMessageId, images, imagePaths, filePaths } = input as {
            chatId: string;
            text: string;
            contentType?: 'text' | 'html';
            replyToMessageId?: string;
            images?: Array<{ contentType: string; contentBytes: string; name?: string }>;
            imagePaths?: string[];
            filePaths?: string[];
          };
          const imgs: PendingImage[] = (images ?? []).map((i, idx) => ({
            id: `img${idx}`,
            contentType: i.contentType,
            contentBytes: i.contentBytes,
            name: i.name,
          }));
          if (imagePaths?.length) {
            imgs.push(...(await readImagePaths(imagePaths, imgs.length)));
          }
          const explicitFiles = filePaths?.length ? await prepareFileAttachments(client, filePaths) : [];
          // Offload inline images to file cards if they'd blow Teams' 4MiB message cap.
          const fitted = await fitInlineImages(client, imgs);
          const fileAttachments = [...explicitFiles, ...fitted.fileAttachments];
          const body = await buildOutgoing(client, chatId, {
            text,
            contentType,
            replyToMessageId,
            images: fitted.images.length ? fitted.images : undefined,
            fileAttachments: fileAttachments.length ? fileAttachments : undefined,
          });
          const m = await client.sendMessageRaw(chatId, body);
          return {
            success: true,
            chatId,
            messageId: m.id,
            imagesOffloaded: fitted.fileAttachments.length || undefined,
            createdDateTime: m.createdDateTime ?? null,
          };
        } catch (err) {
          return errResult(err);
        }
      },
    },

    {
      name: 'send-dm',
      description:
        'Send a direct (1:1) Teams message to a person. Resolves the recipient by email/UPN/AAD id (or unambiguous name), finds or creates the 1:1 chat, and posts the message. Returns the chat id for follow-up reads.',
      inputSchema: {
        type: 'object',
        properties: {
          to: {
            type: 'string',
            description:
              'Recipient email, UPN, AAD object id, or unambiguous display name. Use find-user first if the name may be ambiguous.',
          },
          text: {
            type: 'string',
            description:
              'Message body. Supports markdown and @[Name](aad:<id>) mention tokens (see send-message).',
          },
          contentType: { type: 'string', enum: ['text', 'html'] },
          replyToMessageId: { type: 'string', description: 'Quote-reply to a message in the resulting 1:1 chat.' },
          images: {
            type: 'array',
            description: 'Inline images to embed as hostedContents. If the combined image size would exceed Teams\' ~4MB per-message limit, they are automatically uploaded and sent as file-card attachments instead.',
            items: {
              type: 'object',
              properties: {
                contentType: { type: 'string', description: 'e.g. image/png' },
                contentBytes: { type: 'string', description: 'Base64-encoded image bytes (no data: prefix).' },
                name: { type: 'string' },
              },
              required: ['contentType', 'contentBytes'],
            },
          },
          imagePaths: {
            type: 'array',
            description:
              'Local image file paths to attach inline (embedded as hostedContents). Preferred over `images` when the image is already a file on disk — avoids base64 round-tripping. Supports ~ expansion; MIME type is sniffed automatically. For non-image files use filePaths.',
            items: { type: 'string' },
          },
          filePaths: {
            type: 'array',
            description:
              'Local file paths to attach as Teams file cards (any type, any size — PDFs, docs, archives, etc.). Supports ~ expansion. Each file is uploaded to your OneDrive "Microsoft Teams Chat Files" folder, shared with an organization view link, and attached so the recipient can open it. For inline images prefer imagePaths.',
            items: { type: 'string' },
          },
        },
        required: ['to', 'text'],
        additionalProperties: false,
      },
      execute: async (input) => {
        try {
          const client = await ensureAuthenticated();
          const { to, text, contentType, replyToMessageId, images, imagePaths, filePaths } = input as {
            to: string;
            text: string;
            contentType?: 'text' | 'html';
            replyToMessageId?: string;
            images?: Array<{ contentType: string; contentBytes: string; name?: string }>;
            imagePaths?: string[];
            filePaths?: string[];
          };
          const user = await resolveUser(client, to);
          const chat = await client.getOrCreateOneOnOne(user.id);
          const imgs: PendingImage[] = (images ?? []).map((i, idx) => ({
            id: `img${idx}`,
            contentType: i.contentType,
            contentBytes: i.contentBytes,
            name: i.name,
          }));
          if (imagePaths?.length) {
            imgs.push(...(await readImagePaths(imagePaths, imgs.length)));
          }
          const explicitFiles = filePaths?.length ? await prepareFileAttachments(client, filePaths) : [];
          // Offload inline images to file cards if they'd blow Teams' 4MiB message cap.
          const fitted = await fitInlineImages(client, imgs);
          const fileAttachments = [...explicitFiles, ...fitted.fileAttachments];
          const body = await buildOutgoing(client, chat.id, {
            text,
            contentType,
            replyToMessageId,
            images: fitted.images.length ? fitted.images : undefined,
            fileAttachments: fileAttachments.length ? fileAttachments : undefined,
          });
          const m = await client.sendMessageRaw(chat.id, body);
          return {
            success: true,
            recipient: {
              id: user.id,
              displayName: user.displayName ?? null,
              email: user.mail ?? user.userPrincipalName ?? null,
            },
            chatId: chat.id,
            messageId: m.id,
            imagesOffloaded: fitted.fileAttachments.length || undefined,
            createdDateTime: m.createdDateTime ?? null,
          };
        } catch (err) {
          return errResult(err);
        }
      },
    },

    {
      name: 'react-to-message',
      description:
        'Add or remove a reaction on a Teams chat message. Standard Teams reactions are like, heart, laugh, surprised, sad, angry; any single emoji is also accepted. Use get-chat-messages first to obtain the messageId.',
      inputSchema: {
        type: 'object',
        properties: {
          chatId: { type: 'string' },
          messageId: { type: 'string' },
          reactionType: {
            type: 'string',
            description:
              'One of: like, heart, laugh, surprised, sad, angry — or a single emoji character for a custom reaction.',
          },
          remove: {
            type: 'boolean',
            description: 'Set true to remove the reaction instead of adding it. Default false.',
          },
        },
        required: ['chatId', 'messageId', 'reactionType'],
        additionalProperties: false,
      },
      execute: async (input) => {
        try {
          const client = await ensureAuthenticated();
          const { chatId, messageId, reactionType, remove } = input as {
            chatId: string; messageId: string; reactionType: string; remove?: boolean;
          };
          if (remove) await client.unsetReaction(chatId, messageId, reactionType);
          else await client.setReaction(chatId, messageId, reactionType);
          return { success: true, chatId, messageId, reactionType, removed: !!remove };
        } catch (err) {
          return errResult(err);
        }
      },
    },

    {
      name: 'edit-message',
      description:
        "Edit one of the signed-in user's own messages. Markdown/mention syntax is supported (same as send-message). Pass contentType: 'html' to send a raw HTML body verbatim (for links, <br>, etc.).",
      inputSchema: {
        type: 'object',
        properties: {
          chatId: { type: 'string' },
          messageId: { type: 'string' },
          text: { type: 'string' },
          contentType: {
            type: 'string',
            enum: ['text', 'html'],
            description: 'Force a specific body content type. Omit to auto-detect markdown / mentions.',
          },
        },
        required: ['chatId', 'messageId', 'text'],
        additionalProperties: false,
      },
      execute: async (input) => {
        try {
          const client = await ensureAuthenticated();
          const { chatId, messageId, text, contentType } = input as {
            chatId: string;
            messageId: string;
            text: string;
            contentType?: 'text' | 'html';
          };
          const p = contentType
            ? { body: { contentType, content: text } as { contentType: 'text' | 'html'; content: string }, mentions: undefined, hostedContents: undefined }
            : buildMessageBody(text);
          await client.editMessage(chatId, messageId, {
            body: p.body,
            ...(p.mentions ? { mentions: p.mentions } : {}),
            ...(p.hostedContents ? { hostedContents: p.hostedContents } : {}),
          });
          return { success: true, chatId, messageId };
        } catch (err) {
          return errResult(err);
        }
      },
    },

    {
      name: 'delete-message',
      description: "Soft-delete one of the signed-in user's own messages.",
      inputSchema: {
        type: 'object',
        properties: { chatId: { type: 'string' }, messageId: { type: 'string' } },
        required: ['chatId', 'messageId'],
        additionalProperties: false,
      },
      execute: async (input) => {
        try {
          const client = await ensureAuthenticated();
          const { chatId, messageId } = input as { chatId: string; messageId: string };
          await client.deleteMessage(chatId, messageId);
          return { success: true, chatId, messageId };
        } catch (err) {
          return errResult(err);
        }
      },
    },

    {
      name: 'forward-message',
      description:
        'Forward an existing message to another person or chat. `to` may be an email/UPN/AAD-id (opens or creates a 1:1) or a chatId.',
      inputSchema: {
        type: 'object',
        properties: {
          sourceChatId: { type: 'string' },
          messageId: { type: 'string' },
          to: { type: 'string', description: 'Recipient email/UPN/AAD id, or an existing chatId (19:...).' },
        },
        required: ['sourceChatId', 'messageId', 'to'],
        additionalProperties: false,
      },
      execute: async (input) => {
        try {
          const client = await ensureAuthenticated();
          const { sourceChatId, messageId, to } = input as { sourceChatId: string; messageId: string; to: string };
          const src = await client.getMessage(sourceChatId, messageId);
          const preview = (src.body?.content ?? '').replace(/<[^>]+>/g, '').slice(0, 500);
          const sender = src.from?.user?.displayName ?? src.from?.application?.displayName ?? null;
          let targetChatId: string;
          if (/^19:.+@(unq\.gbl\.spaces|thread\.v2)$/i.test(to)) {
            targetChatId = to;
          } else {
            const u = await resolveUser(client, to);
            targetChatId = (await client.getOrCreateOneOnOne(u.id)).id;
          }
          const body = withMessageRef(
            { body: { contentType: 'html', content: '' } },
            {
              contentType: 'forwardedMessageReference',
              id: messageId,
              messageId,
              messagePreview: preview || null,
              messageSender: sender ? { user: { displayName: sender } } : null,
            },
          );
          const m = await client.sendMessageRaw(targetChatId, body);
          return { success: true, targetChatId, messageId: m.id };
        } catch (err) {
          return errResult(err);
        }
      },
    },

    {
      name: 'mark-chat-read',
      description: "Mark a chat as read for the signed-in user (clears its unread indicator).",
      inputSchema: {
        type: 'object',
        properties: { chatId: { type: 'string' } },
        required: ['chatId'],
        additionalProperties: false,
      },
      execute: async (input) => {
        try {
          const client = await ensureAuthenticated();
          const { chatId } = input as { chatId: string };
          await client.markChatRead(chatId);
          return { success: true, chatId };
        } catch (err) {
          return errResult(err);
        }
      },
    },

    {
      name: 'get-presence',
      description:
        'Get Teams presence (availability, activity, status message) for one or more users by AAD id, email, or UPN.',
      inputSchema: {
        type: 'object',
        properties: {
          users: { type: 'array', items: { type: 'string' }, minItems: 1, maxItems: 50 },
        },
        required: ['users'],
        additionalProperties: false,
      },
      execute: async (input) => {
        try {
          const client = await ensureAuthenticated();
          const { users } = input as { users: string[] };
          const ids: string[] = [];
          for (const ref of users) {
            if (/^[0-9a-f-]{36}$/i.test(ref)) ids.push(ref);
            else {
              const u = await client.getUserByEmail(ref);
              if (u) ids.push(u.id);
            }
          }
          const presence = await client.getPresences(ids);
          return { success: true, presence };
        } catch (err) {
          return errResult(err);
        }
      },
    },

    {
      name: 'invoke-card-action',
      description:
        'Click a button on an Adaptive Card message sent by a bot. Omit `action` to list available actions on the card. ' +
        'Plain submit buttons cause the bot to post a reply into the chat; task-module buttons return the dialog card content. ' +
        'Only works on messages where the sender is a bot/app.',
      inputSchema: {
        type: 'object',
        properties: {
          chatId: { type: 'string', description: 'The chat containing the card message.' },
          messageId: { type: 'string', description: 'The message id (arrival timestamp).' },
          action: {
            type: 'string',
            description: 'Button title or id to click. Omit to list all actions on the card.',
          },
          inputs: {
            type: 'object',
            additionalProperties: true,
            description: 'Optional Input.* values to include with the submit (form fields).',
          },
        },
        required: ['chatId', 'messageId'],
        additionalProperties: false,
      },
      execute: async (input) => {
        try {
          const client = await ensureAuthenticated();
          const { chatId, messageId, action, inputs } = input as {
            chatId: string; messageId: string; action?: string; inputs?: Record<string, unknown>;
          };
          const msg = await client.getMessage(chatId, messageId);
          const botId = msg.from?.application?.id;
          const actions: CardAction[] = [];
          for (const att of msg.attachments ?? []) {
            if (att.contentType === 'application/vnd.microsoft.card.adaptive' && att.content) {
              try { walkActions(JSON.parse(att.content), actions); } catch { /* skip */ }
            }
          }
          const listed = actions.map((a) => ({ id: a.id, title: a.title, type: a.type, url: a.url }));
          if (!action) {
            return { success: true, botId: botId ?? null, actions: listed };
          }
          if (!botId) {
            return { error: 'Message was not sent by a bot; card actions cannot be invoked.', actions: listed };
          }
          const q = action.toLowerCase();
          const hit = actions.find(
            (a) => a.id?.toLowerCase() === q || a.title?.toLowerCase() === q,
          ) ?? actions.find(
            (a) => (a.title ?? '').toLowerCase().includes(q) || (a.id ?? '').toLowerCase().includes(q),
          );
          if (!hit) {
            return { error: `No action matching "${action}".`, actions: listed };
          }
          const ctx = { botId, chatId, messageId };
          if (hit.type === 'Action.OpenUrl') {
            return { success: true, action: hit.title ?? hit.id, type: hit.type, url: hit.url };
          }
          if (hit.type === 'Action.ToggleVisibility' || hit.type === 'Action.ShowCard') {
            return { error: `${hit.type} is a client-side render toggle; nothing to invoke.` };
          }
          const data = { ...(hit.data ?? {}), ...(inputs ?? {}) };
          const { msteams, ...rest } = data as Record<string, unknown> & { msteams?: { type?: string; value?: { type?: string } } };
          if (hit.type === 'Action.Execute') {
            const res = await invokeExecute(api, ctx, hit.verb ?? null, rest);
            return { success: true, action: hit.title ?? hit.id, type: hit.type, response: res };
          }
          const mt = msteams?.type?.toLowerCase();
          if (mt === 'task/fetch' || (mt === 'invoke' && msteams?.value?.type === 'task/fetch')) {
            const r = await invokeTask(api, ctx, 'task/fetch', { ...rest, type: 'task/fetch' });
            return {
              success: true, action: hit.title ?? hit.id, type: 'task/fetch',
              dialog: r ? { title: r.title, card: r.card, url: r.url } : null,
              note: 'This opened a multi-step dialog; further submission is not supported via this tool — use the plugin UI.',
            };
          }
          await invokeMessageback(api, ctx, rest);
          return {
            success: true, action: hit.title ?? hit.id, type: 'messageback',
            note: 'Bot will post its reply into the chat asynchronously; call get-chat-messages to see it.',
          };
        } catch (err) {
          return errResult(err);
        }
      },
    },

    {
      name: 'set-presence',
      description:
        "Set the signed-in user's Teams presence (forced availability). Pass reset=true to clear the override and return to automatic. This is a user-visible change others will see immediately.",
      inputSchema: {
        type: 'object',
        properties: {
          availability: {
            type: 'string',
            enum: AVAIL_VALUES as unknown as string[],
            description: 'Desired presence. Ignored when reset=true.',
          },
          reset: { type: 'boolean', description: 'Clear the forced override (reverts to Available).' },
        },
        additionalProperties: false,
      },
      execute: async (input) => {
        try {
          await ensureAuthenticated();
          const { availability, reset } = input as { availability?: UpsAvailability; reset?: boolean };
          if (reset) {
            await setForcedAvailability(api, null);
          } else {
            if (!availability || !(AVAIL_VALUES as readonly string[]).includes(availability)) {
              return { error: `availability must be one of: ${AVAIL_VALUES.join(', ')} (or set reset=true)` };
            }
            await setForcedAvailability(api, availability);
          }
          const now = await getMyPresence(api).catch(() => null);
          return { success: true, availability: now?.availability ?? availability ?? 'Available', activity: now?.activity ?? null };
        } catch (err) {
          return errResult(err);
        }
      },
    },

    {
      name: 'set-status-message',
      description:
        "Set or clear the signed-in user's Teams status message (the note shown on your profile / when people message you). Pass an empty string to clear.",
      inputSchema: {
        type: 'object',
        properties: {
          message: { type: 'string', description: 'Status text. Empty string clears the note.' },
          pinned: {
            type: 'boolean',
            description: 'When true, Teams shows the note to people who message you. Default true.',
          },
          expiry: {
            type: 'string',
            description: 'Optional ISO-8601 expiry timestamp. Omit for no expiry.',
          },
        },
        required: ['message'],
        additionalProperties: false,
      },
      execute: async (input) => {
        try {
          await ensureAuthenticated();
          const { message, pinned, expiry } = input as { message: string; pinned?: boolean; expiry?: string };
          await setStatusNote(api, message, { pinned: pinned ?? true, expiry });
          return { success: true, message, pinned: pinned ?? true, expiry: expiry ?? null };
        } catch (err) {
          return errResult(err);
        }
      },
    },

    // ── Mail (Outlook) ──
    {
      name: 'list-folders',
      description:
        'List mailbox folders (Inbox plus custom folders like "External Senders") with id, name, and unread/total counts. Use this to discover a folder id or its exact name before calling list-mail on a custom folder.',
      inputSchema: { type: 'object', properties: {}, additionalProperties: false },
      execute: async () => {
        try {
          const client = await ensureAuthenticated();
          const { folders, truncated } = await client.listMailFoldersDeep();
          return {
            success: true,
            count: folders.length,
            ...(truncated ? { truncated: true, note: 'Folder tree is deeper than the enumeration limit; some nested folders are omitted. Use a folder id for those.' } : {}),
            folders: folders.map((f) => ({ id: f.id, name: f.displayName, unread: f.unreadItemCount, total: f.totalItemCount })),
          };
        } catch (err) { return errResult(err); }
      },
    },
    {
      name: 'list-mail',
      description: 'List messages in a mail folder (default: inbox). Accepts a well-known name, a custom folder name (e.g. "External Senders"), or a folder id. Returns summaries only; use get-mail for the full body.',
      inputSchema: {
        type: 'object',
        properties: {
          folder: { type: 'string', description: 'Well-known name (inbox, sentitems, archive, drafts, deleteditems, junkemail), a custom folder display name, or a folder id.', default: 'inbox' },
          top: { type: 'number', default: 20 },
          concise: { type: 'boolean', description: 'Compact output (default true): trims each message to id, subject, from, recipient count, received, and a short preview. Set false for full toRecipients[], webLink and conversationId.' },
        },
        additionalProperties: false,
      },
      execute: async (input) => {
        try {
          const client = await ensureAuthenticated();
          const { folder, top } = input as { folder?: string; top?: number };
          const folderId = await client.resolveFolderId(folder ?? 'inbox');
          const { messages } = await client.listMail(folderId, clampTop(top, 20, 50));
          const out = wantConcise(input) ? messages.map(conciseMail) : messages;
          return { success: true, folder: folder ?? 'inbox', count: messages.length, messages: out };
        } catch (err) { return errResult(err); }
      },
    },
    {
      name: 'get-mail',
      description: 'Fetch a single mail message by id, including HTML body, all recipients, and attachment metadata.',
      inputSchema: { type: 'object', properties: { messageId: { type: 'string' } }, required: ['messageId'], additionalProperties: false },
      execute: async (input) => {
        try {
          const client = await ensureAuthenticated();
          const { messageId } = input as { messageId: string };
          const m = await client.getMail(messageId);
          return { success: true, message: m };
        } catch (err) { return errResult(err); }
      },
    },
    {
      name: 'search-mail',
      description: 'Full-text search across the mailbox (subject, body, participants). Supports KQL-style filters in the query (from:, to:, hasAttachments:true) plus the structured from/since/until/unreadOnly/hasAttachments params below.',
      inputSchema: {
        type: 'object',
        properties: {
          query: { type: 'string', description: 'Free-text query. May be empty when using the structured filters below.' },
          top: { type: 'number', default: 15 },
          from: { type: 'string', description: 'Restrict to a sender (name fragment or email address).' },
          since: { type: 'string', description: 'Only messages received on/after this date (YYYY-MM-DD).' },
          until: { type: 'string', description: 'Only messages received on/before this date (YYYY-MM-DD).' },
          unreadOnly: { type: 'boolean', description: 'Only unread messages (applied client-side to the search results).' },
          hasAttachments: { type: 'boolean', description: 'Only messages with attachments.' },
          concise: { type: 'boolean', description: 'Compact output (default true). Set false for full toRecipients[], webLink and conversationId.' },
        },
        required: ['query'],
        additionalProperties: false,
      },
      execute: async (input) => {
        try {
          const client = await ensureAuthenticated();
          const { query, top, from, since, until, unreadOnly, hasAttachments } = input as {
            query: string; top?: number; from?: string; since?: string; until?: string;
            unreadOnly?: boolean; hasAttachments?: boolean;
          };
          const results = await client.searchMail(query, clampTop(top, 15, 50), { from, since, until, unreadOnly, hasAttachments });
          const out = wantConcise(input) ? results.map(conciseMail) : results;
          return { success: true, query, count: results.length, results: out };
        } catch (err) { return errResult(err); }
      },
    },
    {
      name: 'send-mail',
      description:
        'Send a new email as the signed-in user. Recipients may be plain email addresses. Body accepts markdown ' +
        '(bold/italic/code/links) or raw HTML when contentType="html".',
      inputSchema: {
        type: 'object',
        properties: {
          to: { type: 'array', items: { type: 'string' }, minItems: 1 },
          cc: { type: 'array', items: { type: 'string' } },
          bcc: { type: 'array', items: { type: 'string' } },
          subject: { type: 'string' },
          body: { type: 'string' },
          contentType: { type: 'string', enum: ['markdown', 'html'], default: 'markdown' },
          attachments: {
            type: 'array',
            items: {
              type: 'object',
              properties: {
                name: { type: 'string' },
                contentType: { type: 'string' },
                contentBytes: { type: 'string', description: 'Base64-encoded file contents.' },
              },
              required: ['name', 'contentBytes'],
            },
          },
        },
        required: ['to', 'subject', 'body'],
        additionalProperties: false,
      },
      execute: async (input) => {
        try {
          const client = await ensureAuthenticated();
          const { to, cc, bcc, subject, body, contentType, attachments } = input as {
            to: string[]; cc?: string[]; bcc?: string[]; subject: string; body: string;
            contentType?: 'markdown' | 'html';
            attachments?: Array<{ name: string; contentType?: string; contentBytes: string }>;
          };
          const addrs = (list?: string[]): MailAddress[] => (list ?? []).map((a) => ({ name: null, address: a }));
          const md = buildMessageBody(body);
          const bodyHtml = contentType === 'html' ? body : (md.body.contentType === 'html' ? md.body.content : `<p>${body.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/\n/g, '<br>')}</p>`);
          await client.sendMail({
            to: addrs(to), cc: addrs(cc), bcc: addrs(bcc), subject, bodyHtml,
            attachments: attachments?.map((a) => ({ name: a.name, contentType: a.contentType ?? 'application/octet-stream', contentBytes: a.contentBytes })),
          });
          return { success: true, to, subject };
        } catch (err) { return errResult(err); }
      },
    },
    {
      name: 'reply-to-mail',
      description: 'Reply, reply-all, or forward an existing message. For forward, provide `to`.',
      inputSchema: {
        type: 'object',
        properties: {
          messageId: { type: 'string' },
          mode: { type: 'string', enum: ['reply', 'replyAll', 'forward'], default: 'reply' },
          body: { type: 'string' },
          contentType: { type: 'string', enum: ['markdown', 'html'], default: 'markdown' },
          to: { type: 'array', items: { type: 'string' }, description: 'Required for forward.' },
        },
        required: ['messageId', 'body'],
        additionalProperties: false,
      },
      execute: async (input) => {
        try {
          const client = await ensureAuthenticated();
          const { messageId, mode, body, contentType, to } = input as {
            messageId: string; mode?: 'reply' | 'replyAll' | 'forward'; body: string;
            contentType?: 'markdown' | 'html'; to?: string[];
          };
          const md = buildMessageBody(body);
          const bodyHtml = contentType === 'html' ? body : (md.body.contentType === 'html' ? md.body.content : `<p>${body.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/\n/g, '<br>')}</p>`);
          await client.replyMail(messageId, mode ?? 'reply', {
            bodyHtml,
            to: (to ?? []).map((a) => ({ name: null, address: a })),
          });
          return { success: true, messageId, mode: mode ?? 'reply' };
        } catch (err) { return errResult(err); }
      },
    },
    {
      name: 'mark-mail',
      description: 'Mark a mail message as read/unread and/or flagged/unflagged.',
      inputSchema: {
        type: 'object',
        properties: {
          messageId: { type: 'string' },
          isRead: { type: 'boolean' },
          flagged: { type: 'boolean' },
        },
        required: ['messageId'],
        additionalProperties: false,
      },
      execute: async (input) => {
        try {
          const client = await ensureAuthenticated();
          const { messageId, isRead, flagged } = input as { messageId: string; isRead?: boolean; flagged?: boolean };
          await client.patchMail(messageId, {
            ...(isRead !== undefined ? { isRead } : {}),
            ...(flagged !== undefined ? { flag: flagged ? 'flagged' : 'notFlagged' } : {}),
          });
          return { success: true, messageId, isRead, flagged };
        } catch (err) { return errResult(err); }
      },
    },
    {
      name: 'archive-mail',
      description: 'Move a message to the Archive folder.',
      inputSchema: { type: 'object', properties: { messageId: { type: 'string' } }, required: ['messageId'], additionalProperties: false },
      execute: async (input) => {
        try {
          const client = await ensureAuthenticated();
          await client.moveMail((input as { messageId: string }).messageId, 'archive');
          return { success: true };
        } catch (err) { return errResult(err); }
      },
    },
    {
      name: 'delete-mail',
      description: 'Move a message to Deleted Items.',
      inputSchema: { type: 'object', properties: { messageId: { type: 'string' } }, required: ['messageId'], additionalProperties: false },
      execute: async (input) => {
        try {
          const client = await ensureAuthenticated();
          await client.moveMail((input as { messageId: string }).messageId, 'deleteditems');
          return { success: true };
        } catch (err) { return errResult(err); }
      },
    },

    {
      name: 'get-teams-image',
      description:
        'Fetch the inline (pasted/screenshot) images from a Teams chat message and return them so you can actually see them — e.g. to read an error screenshot someone sent. ' +
        'Use get-chat-messages first to find the message id; messages with images show hasAttachments. ' +
        'By default returns every inline image on the message; pass index to fetch just one.',
      inputSchema: {
        type: 'object',
        properties: {
          chatId: { type: 'string', description: 'The chat id containing the message.' },
          messageId: { type: 'string', description: 'The message id (from get-chat-messages).' },
          index: { type: 'number', description: 'Zero-based index to fetch a single inline image instead of all.' },
        },
        required: ['chatId', 'messageId'],
        additionalProperties: false,
      },
      execute: async (input) => {
        try {
          const client = await ensureAuthenticated();
          const { chatId, messageId, index } = input as { chatId: string; messageId: string; index?: number };
          const raw = await client.getMessage(chatId, messageId);
          const msg = normalizeMessage(raw, tokenCache.getObjectId());
          let urls = msg.hostedImages;
          if (urls.length === 0) {
            return { error: 'That message has no inline images. (Files/SharePoint attachments use get-attachment.)' };
          }
          if (typeof index === 'number') {
            if (index < 0 || index >= urls.length) {
              return { error: `index ${index} out of range; message has ${urls.length} inline image(s).` };
            }
            urls = [urls[index]];
          }
          const modelContent: ModelContentPart[] = [];
          const fetched: Array<{ mediaType: string; bytes: number }> = [];
          for (const u of urls) {
            // Served from the shared hosted-content cache (memory→disk) so
            // repeated questions about the same image don't re-hit Graph (429s).
            const got = await hostedContentCache.getOne(api, client, u);
            if (!got) continue; // permanently unfetchable
            modelContent.push(partForMedia(got.base64, got.mediaType));
            fetched.push({ mediaType: got.mediaType, bytes: Math.floor((got.base64.length * 3) / 4) });
          }
          if (modelContent.length === 0) {
            return { error: 'The inline image(s) could not be fetched (they may have expired or been removed).' };
          }
          return { success: true, chatId, messageId, count: fetched.length, images: fetched, _modelContent: modelContent };
        } catch (err) {
          return errResult(err);
        }
      },
    },

    {
      name: 'get-attachment',
      description:
        'Fetch a file/attachment from a Teams message or an Outlook mail and return its contents so you can read it — images (screenshots, photos), PDFs, and other documents. ' +
        'For Teams, pass source="teams" with chatId + messageId; the attachment is picked by name, index, or attachmentId (omit all to grab the first file). ' +
        'For mail, pass source="mail" with messageId; use get-mail first to list attachment ids/names. ' +
        'Name matching is tolerant of whitespace/case; when a name is ambiguous or fails, the error lists available attachments with [index] — retry with that index. ' +
        'For a pasted/inline screenshot use get-teams-image instead.',
      inputSchema: {
        type: 'object',
        properties: {
          source: { type: 'string', enum: ['teams', 'mail'], description: 'Where the attachment lives.' },
          messageId: { type: 'string', description: 'Teams message id or mail message id.' },
          chatId: { type: 'string', description: 'Required when source="teams".' },
          attachmentId: { type: 'string', description: 'Mail attachment id (from get-mail), or Teams attachment id.' },
          name: { type: 'string', description: 'Match an attachment by file name (tolerant of case/whitespace, falls back to substring).' },
          index: { type: 'number', description: 'Zero-based index into the message\'s attachment list (see the [index] hints in an error). Most reliable selector.' },
        },
        required: ['source', 'messageId'],
        additionalProperties: false,
      },
      execute: async (input) => {
        try {
          const client = await ensureAuthenticated();
          const { source, messageId, chatId, attachmentId, name, index } = input as {
            source: 'teams' | 'mail'; messageId: string; chatId?: string; attachmentId?: string; name?: string; index?: number;
          };
          // Filenames from Graph can carry stray whitespace / zero-width chars,
          // so match tolerantly: normalize (trim, strip zero-width, collapse
          // whitespace, casefold) and compare exact-then-substring rather than
          // requiring a byte-for-byte equal name.
          const norm = (s: string | null | undefined): string =>
            (s ?? '')
              .replace(/[\u200B-\u200D\uFEFF]/g, '') // strip zero-width chars & BOM
              .replace(/\u00A0/g, ' ') // NBSP -> normal space
              .replace(/\s+/g, ' ')
              .trim()
              .toLowerCase();
          const q = name ? norm(name) : '';
          const nameScore = (n: string | null | undefined): number => {
            if (!name) return 0;
            const cand = norm(n);
            if (cand === q) return 2; // exact (normalized)
            if (cand.includes(q) || q.includes(cand)) return 1; // substring either way
            return 0;
          };
          /** Pick the best name match from a list, or undefined if none score > 0. */
          const bestByName = <T extends { name: string | null }>(list: T[]): T | undefined => {
            let best: T | undefined;
            let bestScore = 0;
            for (const item of list) {
              const s = nameScore(item.name);
              if (s > bestScore) { best = item; bestScore = s; }
            }
            return best;
          };

          if (source === 'mail') {
            const mail = await client.getMail(messageId);
            const atts = mail.attachments.filter((a) => a.id);
            if (atts.length === 0) return { error: 'That mail has no attachments.' };
            const chosen =
              (attachmentId && atts.find((a) => a.id === attachmentId)) ||
              (typeof index === 'number' ? atts[index] : undefined) ||
              (name && bestByName(atts)) ||
              (!attachmentId && !name && index == null ? atts[0] : undefined);
            if (!chosen) {
              return {
                error: `No matching attachment. Available: ${atts.map((a, i) => `[${i}] ${a.name} (${a.contentType ?? '?'})`).join('; ')}`,
              };
            }
            const cacheKey = `mail:${messageId}:${chosen.id}`;
            const got = await hostedContentCache.getOneVia(api, cacheKey, () =>
              client.getMailAttachmentRaw(messageId, chosen.id),
            );
            if (!got) return { error: `Attachment "${chosen.name ?? 'attachment'}" has no downloadable content (may be an item/reference attachment).` };
            return {
              success: true, source, messageId, name: chosen.name ?? 'attachment', mediaType: got.mediaType,
              bytes: Math.floor((got.base64.length * 3) / 4),
              _modelContent: [partForMedia(got.base64, got.mediaType, chosen.name ?? undefined)],
            };
          }

          // Teams
          if (!chatId) return { error: 'chatId is required when source="teams".' };
          const raw = await client.getMessage(chatId, messageId);
          const msg = normalizeMessage(raw, tokenCache.getObjectId());
          // Keep url-less entries in the list so name matching still works and we
          // can give a targeted error; only require a url at download time.
          const candidates = [
            ...msg.attachments.map((a) => ({ name: a.name, contentType: a.contentType, url: a.url })),
            ...msg.files.map((a) => ({ name: a.name, contentType: a.contentType, url: a.url })),
          ];
          if (candidates.length === 0) {
            return { error: 'That message has no file attachments. (Inline pasted images use get-teams-image.)' };
          }
          const chosen =
            (typeof index === 'number' ? candidates[index] : undefined) ||
            (name && bestByName(candidates)) ||
            (!name && index == null ? candidates[0] : undefined);
          if (!chosen) {
            return {
              error: `No matching attachment. Available: ${candidates.map((a, i) => `[${i}] ${a.name}`).join('; ')}`,
            };
          }
          if (!chosen.url) {
            return {
              error: `Attachment "${chosen.name ?? 'file'}" has no direct download URL — it is likely an inline pasted image. Use get-teams-image for this message instead.`,
            };
          }
          let base64: string;
          let mediaType: string;
          if (chosen.url.startsWith('https://graph.microsoft.com/')) {
            // Cached (memory→disk) to avoid re-fetching on repeat questions.
            const got = await hostedContentCache.getOne(api, client, chosen.url);
            if (!got) {
              return { error: `Attachment "${chosen.name ?? 'file'}" could not be fetched (may have expired or been removed).` };
            }
            ({ base64, mediaType } = got);
          } else {
            // SharePoint/OneDrive reference — cached by its contentUrl.
            const got = await hostedContentCache.getOneVia(api, `ref:${chosen.url}`, () =>
              client.downloadReferenceAttachment(chosen.url!),
            );
            if (!got) {
              return { error: `Attachment "${chosen.name ?? 'file'}" could not be downloaded from SharePoint/OneDrive.` };
            }
            ({ base64, mediaType } = got);
          }
          return {
            success: true, source, chatId, messageId, name: chosen.name ?? 'attachment', mediaType,
            bytes: Math.floor((base64.length * 3) / 4),
            _modelContent: [partForMedia(base64, mediaType, chosen.name ?? undefined)],
          };
        } catch (err) {
          return errResult(err);
        }
      },
    },

    {
      name: 'create-group-chat',
      description:
        'Create a new Teams group chat with the given members (plus the signed-in user) and optionally send an initial message. Members may be emails, UPNs, AAD ids, or unambiguous names. Requires at least 2 other members.',
      inputSchema: {
        type: 'object',
        properties: {
          members: {
            type: 'array',
            items: { type: 'string' },
            minItems: 2,
            description: 'People to add (excluding yourself).',
          },
          topic: { type: 'string', description: 'Optional chat title.' },
          text: { type: 'string', description: 'Optional first message to post after creation.' },
        },
        required: ['members'],
        additionalProperties: false,
      },
      execute: async (input) => {
        try {
          const client = await ensureAuthenticated();
          const { members, topic, text } = input as {
            members: string[];
            topic?: string;
            text?: string;
          };
          const resolved = [];
          for (const ref of members) {
            resolved.push(await resolveUser(client, ref));
          }
          const chat = await client.createGroupChat(topic ?? null, resolved.map((u) => u.id));
          let messageId: string | null = null;
          if (text) {
            const m = await client.sendMessage(chat.id, text, 'text');
            messageId = m.id;
          }
          return {
            success: true,
            chatId: chat.id,
            topic: chat.topic ?? topic ?? null,
            members: resolved.map((u) => ({
              id: u.id,
              displayName: u.displayName ?? null,
              email: u.mail ?? u.userPrincipalName ?? null,
            })),
            messageId,
          };
        } catch (err) {
          return errResult(err);
        }
      },
    },
  ];
}

export const ALL_TOOL_NAMES = [
  'auth-status',
  'find-user',
  'list-chats',
  'get-chat-messages',
  'search-messages',
  'send-message',
  'send-dm',
  'react-to-message',
  'edit-message',
  'delete-message',
  'list-teams',
  'list-channels',
  'list-channel-messages',
  'get-channel-message',
  'search-channel-messages',
  'edit-channel-message',
  'delete-channel-message',
  'forward-message',
  'mark-chat-read',
  'get-presence',
  'set-presence',
  'set-status-message',
  'invoke-card-action',
  'get-teams-image',
  'get-attachment',
  'list-folders',
  'list-mail',
  'get-mail',
  'search-mail',
  'send-mail',
  'reply-to-mail',
  'mark-mail',
  'archive-mail',
  'delete-mail',
  'create-group-chat',
];
