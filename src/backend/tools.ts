import type { PluginAPI, GraphUser, GraphChatType } from '../shared/types.js';
import { GraphClient, normalizeChat, normalizeMessage } from './graph-client.js';
import * as tokenCache from './token-cache.js';
import { buildMessageBody } from '../shared/markdown.js';
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
              'Message body. Markdown (**bold**, *italic*, `code`, ```fenced blocks```) is auto-converted to Teams HTML unless contentType is set explicitly.',
          },
          contentType: {
            type: 'string',
            enum: ['text', 'html'],
            description: 'Force a specific body content type. Omit to auto-detect markdown.',
          },
        },
        required: ['chatId', 'text'],
        additionalProperties: false,
      },
      execute: async (input) => {
        try {
          const client = await ensureAuthenticated();
          const { chatId, text, contentType } = input as {
            chatId: string;
            text: string;
            contentType?: 'text' | 'html';
          };
          const m = contentType
            ? await client.sendMessage(chatId, text, contentType)
            : await client.sendMessageRaw(chatId, buildMessageBody(text));
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
          text: { type: 'string' },
          contentType: { type: 'string', enum: ['text', 'html'] },
        },
        required: ['to', 'text'],
        additionalProperties: false,
      },
      execute: async (input) => {
        try {
          const client = await ensureAuthenticated();
          const { to, text, contentType } = input as {
            to: string;
            text: string;
            contentType?: 'text' | 'html';
          };
          const user = await resolveUser(client, to);
          const chat = await client.getOrCreateOneOnOne(user.id);
          const m = contentType
            ? await client.sendMessage(chat.id, text, contentType)
            : await client.sendMessageRaw(chat.id, buildMessageBody(text));
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
  'create-group-chat',
];
