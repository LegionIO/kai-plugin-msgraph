import type { PluginAPI, GraphUser, GraphChatType } from '../shared/types.js';
import { GraphClient, normalizeChat, normalizeMessage } from './graph-client.js';
import * as tokenCache from './token-cache.js';
import { buildMessageBody, withMessageRef, type PendingImage } from '../shared/markdown.js';
import { getLogger } from './logger-singleton.js';

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

async function buildOutgoing(
  client: GraphClient,
  chatId: string,
  input: { text: string; contentType?: 'text' | 'html'; replyToMessageId?: string; images?: PendingImage[] },
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
  const { ensureAuthenticated } = deps;

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
          return { success: true, count: chats.length, chats };
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
        },
        required: ['chatId'],
        additionalProperties: false,
      },
      execute: async (input) => {
        try {
          const client = await ensureAuthenticated();
          const { chatId, top } = input as { chatId: string; top?: number };
          const msgs = await client.getChatMessages(chatId, clampTop(top, 25, 50));
          const myId = tokenCache.getObjectId();
          const messages = msgs
            .filter((m) => m.messageType === 'message' || m.messageType == null)
            .map((m) => normalizeMessage(m, myId));
          return { success: true, chatId, count: messages.length, messages };
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
            description: 'Inline images to embed as hostedContents.',
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
        },
        required: ['chatId', 'text'],
        additionalProperties: false,
      },
      execute: async (input) => {
        try {
          const client = await ensureAuthenticated();
          const { chatId, text, contentType, replyToMessageId, images } = input as {
            chatId: string;
            text: string;
            contentType?: 'text' | 'html';
            replyToMessageId?: string;
            images?: Array<{ contentType: string; contentBytes: string; name?: string }>;
          };
          const imgs: PendingImage[] | undefined = images?.map((i, idx) => ({
            id: `img${idx}`,
            contentType: i.contentType,
            contentBytes: i.contentBytes,
            name: i.name,
          }));
          const body = await buildOutgoing(client, chatId, { text, contentType, replyToMessageId, images: imgs });
          const m = await client.sendMessageRaw(chatId, body);
          return {
            success: true,
            chatId,
            messageId: m.id,
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
        },
        required: ['to', 'text'],
        additionalProperties: false,
      },
      execute: async (input) => {
        try {
          const client = await ensureAuthenticated();
          const { to, text, contentType, replyToMessageId } = input as {
            to: string;
            text: string;
            contentType?: 'text' | 'html';
            replyToMessageId?: string;
          };
          const user = await resolveUser(client, to);
          const chat = await client.getOrCreateOneOnOne(user.id);
          const body = await buildOutgoing(client, chat.id, { text, contentType, replyToMessageId });
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
        "Edit one of the signed-in user's own messages. Markdown/mention syntax is supported (same as send-message).",
      inputSchema: {
        type: 'object',
        properties: {
          chatId: { type: 'string' },
          messageId: { type: 'string' },
          text: { type: 'string' },
        },
        required: ['chatId', 'messageId', 'text'],
        additionalProperties: false,
      },
      execute: async (input) => {
        try {
          const client = await ensureAuthenticated();
          const { chatId, messageId, text } = input as { chatId: string; messageId: string; text: string };
          const p = buildMessageBody(text);
          await client.editMessage(chatId, messageId, { body: p.body });
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
  'forward-message',
  'mark-chat-read',
  'get-presence',
  'create-group-chat',
];
