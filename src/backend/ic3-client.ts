/**
 * Teams-internal chat service (IC3) client.
 *
 * Microsoft Graph exposes no way to invoke Adaptive Card actions (Action.Submit,
 * Action.Execute, task/fetch) — those go through the Teams client's own IC3
 * messaging service to the Bot Framework. The FOCI refresh token we already hold
 * can mint tokens for both audiences that endpoint needs, so we call it directly:
 *
 *   POST {chatServiceAfd}/v1/agents/28:{botId}/invoke
 *   Authorization: Bearer <ic3.teams.office.com token>
 *
 * Region + skypetoken come from POST teams.microsoft.com/api/authsvc/v1.0/authz
 * with a spaces-audience token; the same spaces token is embedded in `messageback`
 * bodies so IC3 can post the card-response into the conversation on our behalf.
 *
 * This is Teams' undocumented client protocol. It can change without notice; every
 * caller must be prepared to fall back to the Teams deep-link redirect.
 */

import { randomUUID } from 'crypto';
import {
  CLIENT_ID_TEAMS,
  IC3_SCOPE,
  SPACES_SCOPE,
  PRESENCE_SCOPE,
  TEAMS_AUTHSVC_URL,
  TEAMS_CHATSVC_FALLBACK,
  TEAMS_UPS_FALLBACK,
  TEAMS_REGISTRAR_FALLBACK,
  TEAMS_MT_FALLBACK,
} from '../shared/constants.js';
import type { PluginAPI } from '../shared/types.js';
import { acquireFociAccessToken } from './auth.js';
import * as tokenCache from './token-cache.js';
import { getLogger } from './logger-singleton.js';

export interface IC3Region {
  chatServiceAfd: string;
  presenceUPS: string;
  registrarUrl: string;
  middleTier: string;
  skypeToken: string | null;
  expiresAt: number;
  session: number;
}

let region: IC3Region | null = null;

export function clearIC3State(): void {
  region = null;
}

const TRUSTED_SUFFIXES = [
  '.teams.microsoft.com',
  '.office.com',
  '.skype.com',
  '.cloud.microsoft',
];

function trustedUrl(candidate: string | undefined, fallback: string): string {
  if (!candidate) return fallback;
  try {
    const u = new URL(candidate);
    if (u.protocol !== 'https:') return fallback;
    const h = u.hostname.toLowerCase();
    if (h === 'teams.microsoft.com' || TRUSTED_SUFFIXES.some((s) => h.endsWith(s))) {
      return candidate;
    }
  } catch { /* fall through */ }
  getLogger().warn(`IC3: rejecting untrusted region URL "${candidate}", using ${fallback}`);
  return fallback;
}

export async function ensureRegion(api: PluginAPI): Promise<IC3Region> {
  const session = tokenCache.currentSession();
  if (region && region.session === session && region.expiresAt > Date.now() + 60_000) {
    return region;
  }
  const spaces = await acquireFociAccessToken(api, CLIENT_ID_TEAMS, SPACES_SCOPE);
  let chatServiceAfd = TEAMS_CHATSVC_FALLBACK;
  let presenceUPS = TEAMS_UPS_FALLBACK;
  let registrarUrl = TEAMS_REGISTRAR_FALLBACK;
  let middleTier = TEAMS_MT_FALLBACK;
  let skypeToken: string | null = null;
  let ttlMs = 30 * 60_000;
  try {
    const resp = await api.fetch(TEAMS_AUTHSVC_URL, {
      method: 'POST',
      headers: { Authorization: `Bearer ${spaces}` },
    });
    if (resp.ok) {
      const j = (await resp.json()) as {
        tokens?: { skypeToken?: string; expiresIn?: number };
        regionGtms?: {
          chatServiceAfd?: string;
          chatService?: string;
          presenceUPS?: string;
          calling_registrarUrl?: string;
          middleTier?: string;
          mtImageService?: string;
        };
      };
      const g = j.regionGtms ?? {};
      chatServiceAfd = trustedUrl(g.chatServiceAfd || g.chatService, TEAMS_CHATSVC_FALLBACK);
      presenceUPS = trustedUrl(g.presenceUPS, TEAMS_UPS_FALLBACK);
      registrarUrl = trustedUrl(g.calling_registrarUrl, TEAMS_REGISTRAR_FALLBACK);
      middleTier = trustedUrl(g.middleTier || g.mtImageService, TEAMS_MT_FALLBACK);
      skypeToken = j.tokens?.skypeToken ?? null;
      if (typeof j.tokens?.expiresIn === 'number') ttlMs = Math.max(5 * 60_000, j.tokens.expiresIn * 1000 - 60_000);
    } else {
      getLogger().warn(`IC3 authz ${resp.status}; using fallback region`);
    }
  } catch (err) {
    getLogger().warn(`IC3 authz failed (${err}); using fallback region`);
  }
  region = { chatServiceAfd, presenceUPS, registrarUrl, middleTier, skypeToken, expiresAt: Date.now() + ttlMs, session };
  return region;
}

function spacesToken(api: PluginAPI): Promise<string> {
  return acquireFociAccessToken(api, CLIENT_ID_TEAMS, SPACES_SCOPE);
}

export interface BotProfile { displayName: string; description: string | null }

/**
 * Resolve friendly display names for bots via Teams middle-tier fetchShortProfile.
 * Graph's `from.application.displayName` is the Azure Bot resource name (e.g.
 * "ucap-smartbot-prod-rg-smartbot"); this returns the Teams-app display name
 * (e.g. "SmartBot"), which is what the Teams client shows.
 */
export async function getBotProfiles(api: PluginAPI, botIds: string[]): Promise<Map<string, BotProfile>> {
  const out = new Map<string, BotProfile>();
  if (botIds.length === 0) return out;
  const [{ middleTier }, tok] = await Promise.all([ensureRegion(api), spacesToken(api)]);
  const resp = await api.fetch(
    `${middleTier}/beta/users/fetchShortProfile?isMailAddress=false&enableGuest=true&skypeTeamsInfo=true&includeBots=true`,
    {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${tok}`,
        'content-type': 'application/json;charset=UTF-8',
        'x-ms-client-type': 'cdlworker',
        'x-ms-client-version': '1415/26052906121',
      },
      body: JSON.stringify(botIds.map((id) => `28:${id}`)),
    },
  );
  if (!resp.ok) throw new Error(`fetchShortProfile → ${resp.status}`);
  const j = (await resp.json()) as { value?: Array<{ mri?: string; displayName?: string; givenName?: string; description?: string }> };
  for (const p of j.value ?? []) {
    const id = p.mri?.replace(/^28:/, '');
    const name = p.displayName || p.givenName;
    if (id && name) out.set(id, { displayName: name, description: p.description ?? null });
  }
  return out;
}

export async function getBotIcon(api: PluginAPI, botId: string): Promise<string | null> {
  const [{ middleTier, skypeToken }, tok] = await Promise.all([ensureRegion(api), spacesToken(api)]);
  const asDataUrl = (r: Response, buf: Buffer) =>
    `data:${r.headers.get('content-type') || 'image/png'};base64,${buf.toString('base64')}`;
  // MT profilepicturev2/28:{id} returns Teams' generic hexagon placeholder for
  // bots; the real app icon is fetchShortProfile.imageUri on asyncgw (AMS),
  // which authenticates with the skype token.
  try {
    const p = await api.fetch(
      `${middleTier}/beta/users/fetchShortProfile?isMailAddress=false&includeBots=true`,
      {
        method: 'POST',
        headers: {
          Authorization: `Bearer ${tok}`,
          'content-type': 'application/json;charset=UTF-8',
          'x-ms-client-type': 'cdlworker',
        },
        body: JSON.stringify([`28:${botId}`]),
      },
    );
    if (p.ok) {
      const j = (await p.json()) as { value?: Array<{ imageUri?: string }> };
      const uri = j.value?.[0]?.imageUri;
      if (uri && skypeToken && trustedUrl(uri, '') === uri) {
        const ir = await api.fetch(uri, { headers: { Authorization: `skype_token ${skypeToken}` } });
        if (ir.ok) {
          const buf = Buffer.from(await ir.arrayBuffer());
          if (buf.length > 0) return asDataUrl(ir, buf);
        }
      }
    }
  } catch { /* fall through */ }
  const myId = tokenCache.getObjectId();
  const resp = await api.fetch(
    `${middleTier}/beta/users/${myId}/profilepicturev2/${encodeURIComponent(`28:${botId}`)}?size=HR64x64`,
    { headers: { Authorization: `Bearer ${tok}` } },
  );
  if (!resp.ok) return null;
  const buf = Buffer.from(await resp.arrayBuffer());
  return buf.length ? asDataUrl(resp, buf) : null;
}

export function ic3Token(api: PluginAPI): Promise<string> {
  return acquireFociAccessToken(api, CLIENT_ID_TEAMS, IC3_SCOPE);
}

function presenceToken(api: PluginAPI): Promise<string> {
  return acquireFociAccessToken(api, CLIENT_ID_TEAMS, PRESENCE_SCOPE);
}

async function upsRequest<T = void>(
  api: PluginAPI,
  method: string,
  path: string,
  body?: unknown,
): Promise<T> {
  const [rgn, tok] = await Promise.all([ensureRegion(api), presenceToken(api)]);
  const resp = await api.fetch(`${rgn.presenceUPS}${path}`, {
    method,
    headers: {
      Authorization: `Bearer ${tok}`,
      'Content-Type': 'application/json',
      'x-ms-client-user-agent': 'Teams-V2-Desktop',
      'x-ms-correlation-id': randomUUID(),
    },
    body: body === undefined ? undefined : JSON.stringify(body),
  });
  if (!resp.ok) {
    const t = await resp.text().catch(() => '');
    throw new IC3Error(`UPS ${method} ${path} → ${resp.status} ${t.slice(0, 200)}`, resp.status);
  }
  const ct = resp.headers.get('content-type') ?? '';
  return (ct.includes('json') ? await resp.json() : undefined) as T;
}

export type UpsAvailability =
  | 'Available'
  | 'Busy'
  | 'DoNotDisturb'
  | 'BeRightBack'
  | 'Away'
  | 'Offline';

export interface UpsMyPresence {
  availability: string;
  activity: string;
  note?: { message?: string; expiry?: string } | null;
  forcedAvailability?: { availability?: string; activity?: string; expiry?: string } | null;
  calendarData?: { isOutOfOffice?: boolean } | null;
}

export function getMyPresence(api: PluginAPI): Promise<UpsMyPresence> {
  return upsRequest<UpsMyPresence>(api, 'GET', '/v1/me/presence');
}

/** Set the user-forced presence. Passing null resets to automatic (device-derived). */
export async function setForcedAvailability(
  api: PluginAPI,
  availability: UpsAvailability | null,
): Promise<void> {
  if (availability === null) {
    // DELETE returns 401/40102 in some deployments; PUT Available is the observed reset path.
    await upsRequest(api, 'PUT', '/v1/me/forceavailability/', { availability: 'Available', activity: 'Available' });
    return;
  }
  const activity = availability === 'Offline' ? 'OffWork' : availability;
  await upsRequest(api, 'PUT', '/v1/me/forceavailability/', { availability, activity });
}

/** Set or clear the status message. `pinned` appends the show-when-messaging marker Teams uses. */
export async function setStatusNote(
  api: PluginAPI,
  message: string,
  opts: { pinned?: boolean; expiry?: string } = {},
): Promise<void> {
  const body =
    message.trim().length === 0
      ? { message: '', expiry: '9999-12-31T23:59:59Z' }
      : {
          message: opts.pinned ? `${message}<pinnednote></pinnednote>` : message,
          expiry: opts.expiry ?? '9999-12-31T23:59:59Z',
        };
  await upsRequest(api, 'PUT', '/v1/me/publishnote', body);
}

let lastTypingSent: Record<string, number> = {};

async function postControl(api: PluginAPI, chatId: string, messagetype: string): Promise<void> {
  const [rgn, ic3] = await Promise.all([ensureRegion(api), ic3Token(api)]);
  const resp = await api.fetch(
    `${rgn.chatServiceAfd}/v1/users/ME/conversations/${encodeURIComponent(chatId)}/messages`,
    {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${ic3}`,
        'Content-Type': 'application/json',
        'x-ms-migration': 'True',
        clientinfo: CLIENT_INFO,
      },
      body: JSON.stringify({ messagetype, contenttype: 'Application/Message', content: '' }),
    },
  );
  if (!resp.ok) getLogger().warn(`${messagetype} ${resp.status}`);
}

/** Emit a Control/Typing indicator into a chat, throttled to once per 4s per chat. */
export async function sendTyping(api: PluginAPI, chatId: string): Promise<void> {
  const now = Date.now();
  if (now - (lastTypingSent[chatId] ?? 0) < 4000) return;
  lastTypingSent[chatId] = now;
  await postControl(api, chatId, 'Control/Typing');
}

/** Clear the typing indicator immediately (called on send). */
export async function sendClearTyping(api: PluginAPI, chatId: string): Promise<void> {
  if (lastTypingSent[chatId] === undefined) return;
  delete lastTypingSent[chatId];
  await postControl(api, chatId, 'Control/ClearTyping');
}

export function clearTypingThrottle(): void {
  lastTypingSent = {};
}

/** Per-member read positions for a chat: userId → last-read message id (arrival timestamp). */
export async function getConsumptionHorizons(
  api: PluginAPI,
  chatId: string,
): Promise<Record<string, string>> {
  const [rgn, ic3] = await Promise.all([ensureRegion(api), ic3Token(api)]);
  const resp = await api.fetch(
    `${rgn.chatServiceAfd}/v1/threads/${encodeURIComponent(chatId)}/consumptionhorizons`,
    { headers: { Authorization: `Bearer ${ic3}`, 'x-ms-migration': 'True' } },
  );
  if (!resp.ok) throw new IC3Error(`consumptionhorizons ${resp.status}`, resp.status);
  const j = (await resp.json()) as {
    consumptionhorizons?: Array<{ id?: string; consumptionhorizon?: string }>;
  };
  const out: Record<string, string> = {};
  for (const h of j.consumptionhorizons ?? []) {
    const uid = (h.id ?? '').replace(/^8:orgid:/, '');
    const first = (h.consumptionhorizon ?? '').split(';')[0];
    if (uid && first) out[uid] = first;
  }
  return out;
}

export class IC3Error extends Error {
  name = 'IC3Error' as const;
  constructor(message: string, public readonly status: number, public readonly body?: unknown) {
    super(message);
  }
}

const CLIENT_INFO =
  'os=unknown; osVer=0; proc=x86; lcid=en-us; deviceType=1; country=us; clientName=skypeteams; clientVer=1415/1.0.0.0; utcOffset=+00:00; timezone=UTC';

async function post(api: PluginAPI, botId: string, body: unknown): Promise<unknown> {
  const [rgn, ic3] = await Promise.all([
    ensureRegion(api),
    acquireFociAccessToken(api, CLIENT_ID_TEAMS, IC3_SCOPE),
  ]);
  const url = `${rgn.chatServiceAfd}/v1/agents/28:${botId}/invoke`;
  const resp = await api.fetch(url, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${ic3}`,
      'Content-Type': 'application/json',
      behavioroverride: 'redirectAs404',
      clientinfo: CLIENT_INFO,
      'x-ms-migration': 'True',
      ...(rgn.skypeToken ? { 'x-skypetoken': rgn.skypeToken } : {}),
    },
    body: JSON.stringify(body),
  });
  const text = await resp.text();
  let json: unknown;
  try { json = text ? JSON.parse(text) : undefined; } catch { json = text; }
  if (!resp.ok) {
    throw new IC3Error(
      `IC3 invoke ${resp.status}: ${typeof json === 'object' ? JSON.stringify(json).slice(0, 300) : text.slice(0, 300)}`,
      resp.status,
      json,
    );
  }
  return json;
}

function randomClientMessageId(): string {
  // Teams uses a positive int64 string.
  const hi = Math.floor(Math.random() * 0x7fffffff);
  const lo = Math.floor(Math.random() * 0xffffffff);
  return (BigInt(hi) * 4294967296n + BigInt(lo)).toString();
}

export interface InvokeContext {
  botId: string;
  chatId: string;
  messageId: string;
}

export interface TaskModuleResponse {
  type: 'continue' | 'message' | string;
  title?: string;
  card?: unknown;
  width?: number | string;
  height?: number | string;
  url?: string;
  message?: string;
}

function parseTaskResponse(raw: unknown): TaskModuleResponse | null {
  const task = (raw as { task?: { type?: string; value?: Record<string, unknown> } } | undefined)?.task;
  if (!task) return null;
  const v = task.value ?? {};
  const cardWrap = v.card as { content?: unknown } | undefined;
  return {
    type: task.type ?? 'continue',
    title: typeof v.title === 'string' ? v.title : undefined,
    card: cardWrap?.content ?? undefined,
    width: v.width as number | string | undefined,
    height: v.height as number | string | undefined,
    url: typeof v.url === 'string' ? v.url : undefined,
    message: typeof (raw as { value?: string }).value === 'string' ? (raw as { value: string }).value : undefined,
  };
}

/** Plain Action.Submit — bot replies asynchronously into the chat. */
export async function invokeMessageback(
  api: PluginAPI,
  ctx: InvokeContext,
  value: Record<string, unknown>,
): Promise<void> {
  const spaces = await acquireFociAccessToken(api, CLIENT_ID_TEAMS, SPACES_SCOPE);
  await post(api, ctx.botId, {
    name: 'messageback',
    appId: ctx.botId,
    messageType: 'RichText/Media_Card',
    value,
    imdisplayname: tokenCache.getDisplayName() ?? '',
    userAadToken: spaces,
    serverMessageId: ctx.messageId,
    clientMessageId: randomClientMessageId(),
    conversation: { id: ctx.chatId },
  });
}

/** Dynamic ChoiceSet typeahead (Input.ChoiceSet with choices.data / Data.Query). */
export async function invokeSearch(
  api: PluginAPI,
  ctx: InvokeContext,
  dataset: string,
  queryText: string,
): Promise<Array<{ title: string; value: string }>> {
  const raw = (await post(api, ctx.botId, {
    type: 'invoke',
    name: 'application/search',
    conversation: { id: ctx.chatId },
    imdisplayname: tokenCache.getDisplayName() ?? '',
    replyToId: ctx.messageId,
    value: {
      queryText,
      queryOptions: { skip: 0, top: 15 },
      dataset,
      context: { theme: 'default' },
    },
  })) as { value?: { results?: Array<{ title?: string; value?: string }> } } | undefined;
  return (raw?.value?.results ?? [])
    .filter((r): r is { title: string; value: string } => !!r?.value)
    .map((r) => ({ title: r.title ?? r.value, value: r.value }));
}

/** task/fetch or task/submit — synchronous, may return the next dialog card. */
export async function invokeTask(
  api: PluginAPI,
  ctx: InvokeContext,
  name: 'task/fetch' | 'task/submit',
  data: Record<string, unknown>,
): Promise<TaskModuleResponse | null> {
  const raw = await post(api, ctx.botId, {
    type: 'invoke',
    name,
    conversation: { id: ctx.chatId },
    imdisplayname: tokenCache.getDisplayName() ?? '',
    replyToId: ctx.messageId,
    value: { data, context: { theme: 'default' } },
  });
  return parseTaskResponse(raw);
}

/** Action.Execute (Universal Actions) — synchronous card refresh. */
export async function invokeExecute(
  api: PluginAPI,
  ctx: InvokeContext,
  verb: string | null,
  data: Record<string, unknown>,
): Promise<unknown> {
  return post(api, ctx.botId, {
    type: 'invoke',
    name: 'adaptiveCard/action',
    conversation: { id: ctx.chatId },
    imdisplayname: tokenCache.getDisplayName() ?? '',
    replyToId: ctx.messageId,
    value: { action: { type: 'Action.Execute', verb: verb ?? undefined, data }, trigger: 'manual' },
  });
}
