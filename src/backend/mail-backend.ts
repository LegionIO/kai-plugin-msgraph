import { randomUUID } from 'node:crypto';
import { GraphClient } from './graph-client.js';
import * as tokenCache from './token-cache.js';
import * as photoCache from './photo-cache.js';
import { acquireFociAccessToken } from './auth.js';
import { DiskCache } from './disk-cache.js';
import * as mediaServer from './media-server.js';
import { getLogger } from './logger-singleton.js';
import {
  MAIL_NAV_ID, MAIL_PANEL_ID, MAIL_DELTA_POLL_SECONDS, WELL_KNOWN_FOLDERS,
  CLIENT_ID_OFFICE, OUTLOOK_SCOPE, OUTLOOK_CLOUD_SETTINGS_URL, OUTLOOK_ROAMING_SIG_URL, OWA_SERVICE_URL,
} from '../shared/constants.js';
import type {
  PluginAPI,
  MsgraphPluginState,
  NormalizedMailSummary,
  NormalizedMail,
  MailComposeState,
  MailSignature,
  OutgoingMail,
  MailAddress,
  UserPreferences,
} from '../shared/types.js';

type EnsureFn = (allowInteractive?: boolean) => Promise<GraphClient>;

let ensureAuthenticated: EnsureFn = () => Promise.reject(new Error('mail-backend not initialised'));
let deltaTimer: ReturnType<typeof setInterval> | null = null;
// Set true by disposeMail(), reset false by initMail(). Async continuations
// (loadMailFolders, runDelta) re-check this after their awaits so they don't
// call api.state.set() on a torn-down generation ("Plugin is no longer active").
let mailDeactivated = false;
let mailLoadSeq = 0;
let mailListSeq = 0;
let mailSearchSeq = 0;
let deltaLink: string | null = null;
let deltaCache: DiskCache<string> | null = null;
let senderIdCache: DiskCache<string | null> | null = null;
const inlineCache = new Map<string, string | null>();
const senderInFlight = new Set<string>();

mediaServer.register('mailinline', (k) => inlineCache.get(k) ?? undefined);

const st = (api: PluginAPI) => api.state.get() as Partial<MsgraphPluginState>;

export function mailInitialState(): Pick<
  MsgraphPluginState,
  | 'mailFolders' | 'mailFoldersExpanded' | 'mailSenderIds' | 'activeMailFolder' | 'mailList' | 'mailListNextLink' | 'mailSearch'
  | 'activeMailId' | 'activeMail' | 'mailInlineAttachments' | 'composingMail' | 'mailSignature'
  | 'loadingMailFolders' | 'loadingMailList' | 'loadingMail' | 'sendingMail' | 'mailError'
> {
  return {
    mailFolders: [],
    mailFoldersExpanded: [],
    mailSenderIds: {},
    activeMailFolder: 'inbox',
    mailList: [],
    mailListNextLink: null,
    mailSearch: null,
    activeMailId: null,
    activeMail: null,
    mailInlineAttachments: {},
    composingMail: null,
    mailSignature: null,
    loadingMailFolders: false,
    loadingMailList: false,
    loadingMail: false,
    sendingMail: false,
    mailError: null,
  };
}

export function initMail(api: PluginAPI, ensure: EnsureFn): void {
  mailDeactivated = false;
  ensureAuthenticated = ensure;
  deltaCache = new DiskCache<string>(api.pluginName, 'mail-delta', { hardTtlMs: 7 * 24 * 60 * 60_000, maxEntries: 20, sync: true });
  deltaLink = deltaCache.get('inbox')?.v ?? null;
  senderIdCache = new DiskCache<string | null>(
    api.pluginName,
    'mail-senders',
    { hardTtlMs: 14 * 24 * 60 * 60_000, maxEntries: 2000 },
    () => {
      const seed: Record<string, string | null> = {};
      for (const [k, e] of senderIdCache!.entries()) seed[k] = e.v;
      if (Object.keys(seed).length) {
        api.state.set('mailSenderIds', { ...(st(api).mailSenderIds ?? {}), ...seed });
      }
    },
  );
}

/** Resolve sender email addresses → AAD user ids (for avatar photos). Cached, throttled. */
function ensureSenderPhotos(api: PluginAPI, client: GraphClient, messages: NormalizedMailSummary[]): void {
  const known = st(api).mailSenderIds ?? {};
  const need = new Set<string>();
  const haveIds = new Set<string>();
  for (const m of messages) {
    const addr = m.from?.address?.toLowerCase();
    if (!addr) continue;
    if (addr in known) {
      const id = known[addr];
      if (id) haveIds.add(id);
    } else if (!senderInFlight.has(addr)) {
      need.add(addr);
    }
  }
  if (haveIds.size) photoCache.ensure(api, client, haveIds);
  if (need.size === 0) return;
  const session = tokenCache.currentSession();
  const batch = [...need].slice(0, 20);
  for (const a of batch) senderInFlight.add(a);
  void (async () => {
    const found: Record<string, string | null> = {};
    const ids: string[] = [];
    for (const addr of batch) {
      if (session !== tokenCache.currentSession()) break;
      try {
        const u = await client.getUserByEmail(addr);
        found[addr] = u?.id ?? null;
        if (u?.id) ids.push(u.id);
      } catch {
        found[addr] = null;
      } finally {
        senderInFlight.delete(addr);
      }
      // Only persist successful resolutions so a lookup miss (or a bug in the
      // lookup path) doesn't stick for the whole TTL.
      if (found[addr]) senderIdCache?.set(addr, found[addr]);
    }
    if (session !== tokenCache.currentSession()) return;
    api.state.set('mailSenderIds', { ...(st(api).mailSenderIds ?? {}), ...found });
    if (ids.length) photoCache.ensure(api, client, ids);
  })();
}

export function updateMailNavBadge(api: PluginAPI): void {
  const inbox = (st(api).mailFolders ?? []).find((f) => f.wellKnownName === 'inbox' || f.id === 'inbox');
  api.ui.registerNavigationItem({
    id: MAIL_NAV_ID,
    label: 'Outlook',
    icon: { lucide: 'mail' },
    visible: true,
    priority: 1,
    badge: inbox && inbox.unreadItemCount > 0 ? inbox.unreadItemCount : undefined,
    target: { type: 'action', targetId: `panel:${MAIL_PANEL_ID}`, action: 'set-view', data: { view: 'mail' }, panelId: MAIL_PANEL_ID },
  });
}

export async function loadMailFolders(api: PluginAPI, allowInteractive = false): Promise<void> {
  if (mailDeactivated) return;
  if (!tokenCache.hasRefreshToken() && !tokenCache.isTokenValid()) return;
  const session = tokenCache.currentSession();
  api.state.set('loadingMailFolders', true);
  try {
    const client = await ensureAuthenticated(allowInteractive);
    const { folders: all } = await client.listMailFolders();
    if (mailDeactivated || session !== tokenCache.currentSession()) return;
    const order = new Map(WELL_KNOWN_FOLDERS.map((n, i) => [n, i]));
    const sorted = [...all].sort((a, b) => {
      const ai = order.get((a.wellKnownName ?? '') as typeof WELL_KNOWN_FOLDERS[number]) ?? 99;
      const bi = order.get((b.wellKnownName ?? '') as typeof WELL_KNOWN_FOLDERS[number]) ?? 99;
      return ai !== bi ? ai - bi : a.displayName.localeCompare(b.displayName);
    });
    // Preserve any already-expanded subtrees across refreshes.
    const prev = st(api).mailFolders ?? [];
    const expanded = new Set(st(api).mailFoldersExpanded ?? []);
    const withChildren: typeof sorted = [];
    for (const f of sorted) {
      withChildren.push(f);
      if (expanded.has(f.id)) {
        for (const c of prev.filter((p) => p.parentId === f.id)) withChildren.push(c);
      }
    }
    api.state.set('mailFolders', withChildren);
    updateMailNavBadge(api);
  } catch (err) {
    if (!mailDeactivated && session === tokenCache.currentSession()) {
      api.state.set('mailError', err instanceof Error ? err.message : String(err));
    }
  } finally {
    if (!mailDeactivated && session === tokenCache.currentSession()) api.state.set('loadingMailFolders', false);
  }
}

export async function loadMailList(api: PluginAPI, folderId: string): Promise<void> {
  const seq = ++mailListSeq;
  const prevFolder = st(api).activeMailFolder;
  api.state.set('activeMailFolder', folderId);
  api.state.set('mailError', null);
  if (prevFolder !== folderId) {
    api.state.set('mailList', []);
    api.state.set('mailListNextLink', null);
    api.state.set('activeMailId', null);
    api.state.set('activeMail', null);
  }
  api.state.set('loadingMailList', true);
  try {
    const client = await ensureAuthenticated(false);
    const { messages, nextLink } = await client.listMail(folderId);
    if (seq !== mailListSeq) return;
    api.state.set('mailList', messages);
    api.state.set('mailListNextLink', nextLink);
    ensureSenderPhotos(api, client, messages);
  } catch (err) {
    if (seq === mailListSeq) api.state.set('mailError', err instanceof Error ? err.message : String(err));
  } finally {
    if (seq === mailListSeq) api.state.set('loadingMailList', false);
  }
}

async function loadMoreMail(api: PluginAPI): Promise<void> {
  const link = st(api).mailListNextLink;
  if (!link) return;
  const seq = mailListSeq;
  api.state.set('loadingMailList', true);
  try {
    const client = await ensureAuthenticated(false);
    const { messages, nextLink } = await client.listMailPage(link);
    if (seq !== mailListSeq) return;
    const cur = st(api).mailList ?? [];
    const seen = new Set(cur.map((m) => m.id));
    api.state.set('mailList', [...cur, ...messages.filter((m) => !seen.has(m.id))]);
    api.state.set('mailListNextLink', nextLink);
    ensureSenderPhotos(api, client, messages);
  } finally {
    if (seq === mailListSeq) api.state.set('loadingMailList', false);
  }
}

async function loadMailBody(api: PluginAPI, messageId: string): Promise<void> {
  const seq = ++mailLoadSeq;
  api.state.set('activeMailId', messageId);
  api.state.set('activeMail', null);
  api.state.set('mailInlineAttachments', {});
  api.state.set('loadingMail', true);
  api.state.set('mailError', null);
  try {
    const client = await ensureAuthenticated(false);
    const mail = await client.getMail(messageId);
    if (seq !== mailLoadSeq) return;
    api.state.set('activeMail', mail);
    // Optimistic mark-read + reflect in list.
    if (!mail.isRead) {
      void client.patchMail(messageId, { isRead: true }).catch((e) => getLogger().warn(`mark-read: ${e}`));
      patchList(api, messageId, { isRead: true });
      bumpFolderUnread(api, st(api).activeMailFolder ?? 'inbox', -1);
    }
    // Inline images referenced by cid: — fetch and expose via media server.
    const inline = mail.attachments.filter((a) => a.isInline);
    if (inline.length) void fetchInline(api, client, messageId, inline, seq);
  } catch (err) {
    if (seq === mailLoadSeq) api.state.set('mailError', err instanceof Error ? err.message : String(err));
  } finally {
    if (seq === mailLoadSeq) api.state.set('loadingMail', false);
  }
}

async function fetchInline(
  api: PluginAPI,
  client: GraphClient,
  messageId: string,
  atts: NormalizedMail['attachments'],
  seq: number,
): Promise<void> {
  const out: Record<string, string | null> = {};
  for (const a of atts) {
    if (seq !== mailLoadSeq) return;
    try {
      const { contentId, dataUrl } = await client.getMailAttachment(messageId, a.id);
      const key = (contentId ?? a.id).replace(/^<|>$/g, '');
      inlineCache.set(key, dataUrl);
      out[key] = mediaServer.urlFor('mailinline', key) ?? dataUrl;
    } catch (err) {
      getLogger().warn(`inline attachment ${a.id}: ${err}`);
      out[a.id] = null;
    }
    if (seq === mailLoadSeq) api.state.set('mailInlineAttachments', { ...(st(api).mailInlineAttachments ?? {}), ...out });
  }
}

function patchList(api: PluginAPI, id: string, patch: Partial<NormalizedMailSummary>): void {
  const cur = st(api).mailList ?? [];
  const idx = cur.findIndex((m) => m.id === id);
  if (idx < 0) return;
  const next = [...cur];
  next[idx] = { ...next[idx], ...patch };
  api.state.set('mailList', next);
}

function removeFromList(api: PluginAPI, id: string): void {
  const cur = st(api).mailList ?? [];
  api.state.set('mailList', cur.filter((m) => m.id !== id));
  if (st(api).activeMailId === id) {
    api.state.set('activeMailId', null);
    api.state.set('activeMail', null);
  }
}

function bumpFolderUnread(api: PluginAPI, folderId: string, delta: number): void {
  const cur = st(api).mailFolders ?? [];
  const next = cur.map((f) =>
    f.id === folderId || f.wellKnownName === folderId
      ? { ...f, unreadItemCount: Math.max(0, f.unreadItemCount + delta) }
      : f,
  );
  api.state.set('mailFolders', next);
  updateMailNavBadge(api);
}

// ── Delta poll ──

export function startMailPoll(api: PluginAPI): void {
  if (deltaTimer) return;
  void runDelta(api).catch((e) => getLogger().warn(`initial mail delta: ${e}`));
  deltaTimer = setInterval(() => {
    if (!tokenCache.hasRefreshToken()) return;
    void runDelta(api).catch((e) => getLogger().warn(`mail delta: ${e}`));
  }, MAIL_DELTA_POLL_SECONDS * 1000);
}

export function stopMailPoll(): void {
  if (deltaTimer) { clearInterval(deltaTimer); deltaTimer = null; }
  deltaLink = null;
  mailLoadSeq++; mailListSeq++; mailSearchSeq++;
  inlineCache.clear();
}

export function disposeMail(): void {
  mailDeactivated = true;
  stopMailPoll();
  deltaCache?.dispose();
  senderIdCache?.dispose();
}

async function runDelta(api: PluginAPI): Promise<void> {
  const session = tokenCache.currentSession();
  const client = await ensureAuthenticated(false);
  const initial = deltaLink === null;
  const { changes, deltaLink: next } = await client.mailDelta('inbox', deltaLink);
  if (mailDeactivated || session !== tokenCache.currentSession()) return;
  if (next) { deltaLink = next; deltaCache?.set('inbox', next); }
  if (initial) return; // first call just establishes the cursor
  const folder = st(api).activeMailFolder ?? 'inbox';
  const isInboxView = folder === 'inbox' || (st(api).mailFolders ?? []).some((f) => f.id === folder && f.wellKnownName === 'inbox');
  const cur = st(api).mailList ?? [];
  const byId = new Map(cur.map((m) => [m.id, m]));
  let unreadDelta = 0;
  const newArrivals: NormalizedMailSummary[] = [];
  for (const c of changes) {
    if ('removedId' in c) {
      if (byId.has(c.removedId) && !byId.get(c.removedId)!.isRead) unreadDelta--;
      byId.delete(c.removedId);
      continue;
    }
    const prev = byId.get(c.id);
    if (!prev) {
      byId.set(c.id, c);
      if (!c.isRead) unreadDelta++;
      newArrivals.push(c);
    } else {
      if (prev.isRead !== c.isRead) unreadDelta += c.isRead ? -1 : 1;
      byId.set(c.id, { ...prev, ...c });
    }
  }
  if (isInboxView && changes.length) {
    const merged = [...byId.values()].sort(
      (a, b) => (b.receivedDateTime ?? '').localeCompare(a.receivedDateTime ?? ''),
    );
    api.state.set('mailList', merged);
  }
  if (unreadDelta !== 0) bumpFolderUnread(api, 'inbox', unreadDelta);
  else if (changes.length) void loadMailFolders(api).catch(() => {});
  for (const m of newArrivals) {
    try {
      api.events?.emit('mail-received', {
        messageId: m.id,
        subject: m.subject,
        from: m.from,
        preview: m.bodyPreview,
        receivedDateTime: m.receivedDateTime,
        hasAttachments: m.hasAttachments,
      });
    } catch { /* older host */ }
  }
}

// ── Signature ──

function prefs(api: PluginAPI): Partial<UserPreferences> {
  return ((api.config.getPluginData().preferences ?? {}) as Partial<UserPreferences>);
}

function decodeJwt(t: string): Record<string, unknown> {
  try { return JSON.parse(Buffer.from(t.split('.')[1], 'base64url').toString('utf-8')); }
  catch { return {}; }
}

async function outlookHeaders(api: PluginAPI): Promise<Record<string, string>> {
  const tok = await acquireFociAccessToken(api, CLIENT_ID_OFFICE, OUTLOOK_SCOPE);
  const c = decodeJwt(tok);
  const anchor = c.puid && c.tid ? `PUID:${c.puid}@${c.tid}` : `SMTP:${tokenCache.getEmail() ?? ''}`;
  return {
    Authorization: `Bearer ${tok}`,
    'x-anchormailbox': anchor,
    'x-routingparameter-sessionkey': anchor,
    'x-outlook-client': 'owa',
  };
}

type OwaResponseItem = { ResponseCode?: string; ResponseClass?: string; MessageText?: string };

async function owaAction(api: PluginAPI, action: string, body: Record<string, unknown>): Promise<void> {
  const H = await outlookHeaders(api);
  const payload = {
    __type: `${action}JsonRequest:#Exchange`,
    Header: { __type: 'JsonRequestHeaders:#Exchange', RequestServerVersion: 'V2018_01_08' },
    Body: { __type: `${action}Request:#Exchange`, ...body },
  };
  const resp = await fetch(`${OWA_SERVICE_URL}?action=${action}&app=Mail`, {
    method: 'POST',
    headers: {
      ...H,
      action,
      'content-type': 'application/json; charset=utf-8',
      'x-req-source': 'Mail',
      'x-owa-correlationid': randomUUID(),
      'x-owa-urlpostdata': encodeURIComponent(JSON.stringify(payload)),
      // Graph mailFolder ids are immutable-format; this makes OWA accept them as FolderId:#Exchange.
      prefer: 'IdType="ImmutableId"',
    },
  });
  if (!resp.ok) throw new Error(`OWA ${action} → ${resp.status} ${(await resp.text()).slice(0, 200)}`);
  const j = (await resp.json()) as { Body?: { ResponseMessages?: { Items?: OwaResponseItem[] } } };
  const item = j.Body?.ResponseMessages?.Items?.[0];
  if (item?.ResponseClass && item.ResponseClass !== 'Success') {
    throw new Error(`OWA ${action} → ${item.ResponseCode}: ${item.MessageText ?? item.ResponseClass}`);
  }
}

const owaFolderId = (id: string) => ({ __type: 'FolderId:#Exchange', Id: id });

function setSignatureIfChanged(api: PluginAPI, next: MailSignature | null): void {
  const cur = st(api).mailSignature ?? null;
  if (cur === next) return;
  if (cur && next
    && cur.html === next.html
    && cur.autoAddOnNew === next.autoAddOnNew
    && cur.autoAddOnReply === next.autoAddOnReply
    && cur.source === next.source) return;
  api.state.set('mailSignature', next);
}

export async function loadSignature(api: PluginAPI): Promise<void> {
  const p = prefs(api);
  if (p.mailSignatureHtml?.trim()) {
    setSignatureIfChanged(api, {
      html: p.mailSignatureHtml,
      autoAddOnNew: p.mailSignatureAutoNew ?? true,
      autoAddOnReply: p.mailSignatureAutoReply ?? true,
      source: 'config',
    });
    return;
  }
  try {
    const H = await outlookHeaders(api);
    const resp = await api.fetch(
      `${OUTLOOK_CLOUD_SETTINGS_URL}?settingname=signaturehtml,roaming_signature_list,roaming_new_signature,roaming_reply_signature`,
      { headers: H },
    );
    if (!resp.ok) throw new Error(`OutlookCloudSettings ${resp.status}`);
    const raw = (await resp.json()) as unknown;
    const list: Array<{ name?: string; value?: unknown }> =
      Array.isArray(raw) ? raw : ((raw as { settings?: unknown[] }).settings as never) ?? [];
    const get = (n: string) => list.find((s) => s.name?.toLowerCase() === n)?.value as string | undefined;

    let html = get('signaturehtml');
    const newName = get('roaming_new_signature');
    const replyName = get('roaming_reply_signature');
    const usesRoaming = !!get('roaming_signature_list');

    if (!html && newName) {
      const r2 = await api.fetch(
        `${OUTLOOK_ROAMING_SIG_URL}/GetSignature?signatureName=${encodeURIComponent(newName)}&api-version=1.0`,
        { headers: H },
      );
      if (r2.ok) {
        const j = (await r2.json()) as { html?: string; Html?: string; signatureHtml?: string };
        html = j.html ?? j.Html ?? j.signatureHtml;
      }
    }
    if (html?.trim()) {
      setSignatureIfChanged(api, {
        html,
        autoAddOnNew: true,
        autoAddOnReply: usesRoaming ? !!replyName : true,
        source: 'owa',
      });
      getLogger().info(`mail-signature: loaded from Outlook (${usesRoaming ? 'roaming' : 'legacy'})`);
    } else {
      setSignatureIfChanged(api, null);
      getLogger().info('mail-signature: no signature configured in Outlook; use plugin Settings to set one.');
    }
  } catch (err) {
    getLogger().warn(`mail-signature fetch failed: ${err}`);
  }
}

// ── Compose helpers ──

function quoteHeader(m: NormalizedMail): string {
  const when = m.receivedDateTime ? new Date(m.receivedDateTime).toLocaleString() : '';
  const from = m.from ? `${m.from.name ?? m.from.address} &lt;${m.from.address}&gt;` : '';
  return (
    `<div style="border-left:2px solid #ccc;margin:12px 0 0;padding:4px 0 4px 12px;color:#555;font-size:12px">` +
    `<b>From:</b> ${from}<br><b>Sent:</b> ${when}<br><b>Subject:</b> ${escapeHtml(m.subject)}</div>`
  );
}

function escapeHtml(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function makeCompose(
  mode: MailComposeState['mode'],
  src: NormalizedMail | null,
  me: MailAddress,
  sig: MailSignature | null,
): MailComposeState {
  const isReplyish = mode === 'reply' || mode === 'replyAll' || mode === 'forward';
  const wantSig = sig && (isReplyish ? sig.autoAddOnReply : sig.autoAddOnNew);
  const signatureHtml = wantSig ? `<br><br>${sig.html}` : null;
  if (!src || mode === 'new') {
    return { mode: 'new', sourceId: null, to: [], cc: [], subject: '', quotedHtml: null, signatureHtml };
  }
  const isMe = (a: MailAddress) => a.address.toLowerCase() === me.address.toLowerCase();
  const subj = src.subject.replace(/^(re|fw|fwd):\s*/i, '');
  if (mode === 'forward') {
    return {
      mode, sourceId: src.id, to: [], cc: [],
      subject: `FW: ${subj}`,
      quotedHtml: quoteHeader(src) + src.bodyHtml,
      signatureHtml,
    };
  }
  const to = mode === 'replyAll'
    ? [src.from, ...src.toRecipients].filter((a): a is MailAddress => !!a && !isMe(a))
    : src.from ? [src.from] : [];
  const cc = mode === 'replyAll' ? src.ccRecipients.filter((a) => !isMe(a)) : [];
  return {
    mode, sourceId: src.id,
    to: dedupeAddrs(to), cc: dedupeAddrs(cc),
    subject: `RE: ${subj}`,
    quotedHtml: quoteHeader(src) + src.bodyHtml,
    signatureHtml,
  };
}

function dedupeAddrs(list: MailAddress[]): MailAddress[] {
  const seen = new Set<string>();
  return list.filter((a) => {
    const k = a.address.toLowerCase();
    if (seen.has(k)) return false;
    seen.add(k);
    return true;
  });
}

// ── Panel actions ──

export async function handleMailAction(api: PluginAPI, action: string, data?: unknown): Promise<void> {
  const log = getLogger();
  try {
    switch (action) {
      case 'refresh-mail': {
        await Promise.all([loadMailFolders(api), loadMailList(api, st(api).activeMailFolder ?? 'inbox')]);
        break;
      }
      case 'select-folder': {
        const { folderId } = data as { folderId: string };
        mailSearchSeq++;
        api.state.set('mailSearch', null);
        await loadMailList(api, folderId);
        break;
      }
      case 'mark-folder-read': {
        const { folderId } = data as { folderId: string };
        const realId = (st(api).mailFolders ?? []).find((f) => f.wellKnownName === folderId)?.id ?? folderId;
        try {
          await owaAction(api, 'MarkAllItemsAsRead', {
            ReadFlag: true,
            SuppressReadReceipts: true,
            FolderIds: [owaFolderId(realId)],
            ItemIdsToExclude: [],
          });
          log.info(`mark-folder-read(${folderId}): OWA bulk`);
        } catch (e) {
          log.warn(`OWA MarkAllItemsAsRead failed, falling back to Graph batch: ${(e as Error).message}`);
          const client = await ensureAuthenticated(false);
          const n = await client.markFolderRead(folderId);
          log.info(`mark-folder-read(${folderId}): ${n} messages via Graph`);
        }
        const cur = st(api).mailFolders ?? [];
        api.state.set('mailFolders', cur.map((f) =>
          f.id === folderId || f.wellKnownName === folderId ? { ...f, unreadItemCount: 0 } : f,
        ));
        updateMailNavBadge(api);
        if (st(api).activeMailFolder === folderId || (cur.find((f) => f.id === folderId)?.wellKnownName === st(api).activeMailFolder)) {
          api.state.set('mailList', (st(api).mailList ?? []).map((m) => ({ ...m, isRead: true })));
        }
        break;
      }
      case 'empty-folder': {
        const { folderId } = data as { folderId: string };
        const folder = (st(api).mailFolders ?? []).find((f) => f.id === folderId || f.wellKnownName === folderId);
        const hardDelete = folder?.wellKnownName === 'deleteditems' || folder?.wellKnownName === 'junkemail';
        const realId = folder?.id ?? folderId;
        try {
          await owaAction(api, 'EmptyFolder', {
            FolderIds: [owaFolderId(realId)],
            DeleteType: hardDelete ? 'HardDelete' : 'MoveToDeletedItems',
            DeleteSubFolders: false,
            AllowSearchFolder: true,
            SuppressReadReceipt: true,
            ItemIdsToExclude: [],
            BulkActionRequest: true,
            BulkActionBatchSize: 200,
          });
          log.info(`empty-folder(${folderId}): OWA bulk (${hardDelete ? 'HardDelete' : 'MoveToDeletedItems'})`);
        } catch (e) {
          log.warn(`OWA EmptyFolder failed, falling back to Graph batch: ${(e as Error).message}`);
          const client = await ensureAuthenticated(false);
          const n = await client.emptyFolder(folderId, hardDelete);
          log.info(`empty-folder(${folderId}): ${n} messages ${hardDelete ? 'deleted' : 'moved to Deleted Items'} via Graph`);
        }
        await loadMailFolders(api);
        if (st(api).activeMailFolder === folderId || st(api).activeMailFolder === folder?.wellKnownName) {
          await loadMailList(api, st(api).activeMailFolder ?? folderId);
        }
        break;
      }
      case 'create-folder': {
        const { parentId, name } = data as { parentId: string; name: string };
        if (!name.trim()) break;
        const client = await ensureAuthenticated(false);
        await client.createChildFolder(parentId, name.trim());
        // Ensure parent is expanded so the new child shows.
        const expanded = new Set(st(api).mailFoldersExpanded ?? []);
        if (!expanded.has(parentId)) {
          await handleMailAction(api, 'toggle-folder', { folderId: parentId });
        } else {
          // Refresh the parent's children.
          expanded.delete(parentId);
          api.state.set('mailFoldersExpanded', [...expanded]);
          api.state.set('mailFolders', (st(api).mailFolders ?? []).filter((f) => f.parentId !== parentId));
          await handleMailAction(api, 'toggle-folder', { folderId: parentId });
        }
        break;
      }
      case 'rename-folder': {
        const { folderId, name } = data as { folderId: string; name: string };
        if (!name.trim()) break;
        const client = await ensureAuthenticated(false);
        await client.renameFolder(folderId, name.trim());
        api.state.set('mailFolders', (st(api).mailFolders ?? []).map((f) =>
          f.id === folderId ? { ...f, displayName: name.trim() } : f,
        ));
        break;
      }
      case 'delete-folder': {
        const { folderId } = data as { folderId: string };
        const client = await ensureAuthenticated(false);
        await client.deleteFolder(folderId);
        api.state.set('mailFolders', (st(api).mailFolders ?? []).filter((f) => f.id !== folderId && f.parentId !== folderId));
        if (st(api).activeMailFolder === folderId) await loadMailList(api, 'inbox');
        void loadMailFolders(api);
        break;
      }
      case 'toggle-folder': {
        const { folderId } = data as { folderId: string };
        const expanded = new Set(st(api).mailFoldersExpanded ?? []);
        const cur = st(api).mailFolders ?? [];
        if (expanded.has(folderId)) {
          // Collapse: remove this id (and any descendants) from expanded, drop descendant rows.
          const drop = new Set<string>([folderId]);
          let changed = true;
          while (changed) {
            changed = false;
            for (const f of cur) {
              if (f.parentId && drop.has(f.parentId) && !drop.has(f.id)) { drop.add(f.id); changed = true; }
            }
          }
          for (const id of drop) expanded.delete(id);
          expanded.delete(folderId);
          api.state.set('mailFoldersExpanded', [...expanded]);
          api.state.set('mailFolders', cur.filter((f) => !f.parentId || !drop.has(f.parentId)));
        } else {
          const parent = cur.find((f) => f.id === folderId);
          if (!parent) break;
          const client = await ensureAuthenticated(false);
          const children = await client.listChildFolders(folderId, parent.depth + 1);
          const idx = cur.findIndex((f) => f.id === folderId);
          const next = [...cur.slice(0, idx + 1), ...children, ...cur.slice(idx + 1).filter((f) => f.parentId !== folderId)];
          expanded.add(folderId);
          api.state.set('mailFoldersExpanded', [...expanded]);
          api.state.set('mailFolders', next);
        }
        break;
      }
      case 'load-more-mail': {
        await loadMoreMail(api);
        break;
      }
      case 'select-mail': {
        const { messageId } = data as { messageId: string };
        await loadMailBody(api, messageId);
        break;
      }
      case 'mark-mail': {
        const { messageId, isRead, flag } = data as { messageId: string; isRead?: boolean; flag?: boolean };
        const client = await ensureAuthenticated(false);
        const patch: { isRead?: boolean; flag?: 'flagged' | 'notFlagged' } = {};
        if (isRead !== undefined) patch.isRead = isRead;
        if (flag !== undefined) patch.flag = flag ? 'flagged' : 'notFlagged';
        await client.patchMail(messageId, patch);
        patchList(api, messageId, {
          ...(isRead !== undefined ? { isRead } : {}),
          ...(flag !== undefined ? { flagged: flag } : {}),
        });
        if (isRead !== undefined) bumpFolderUnread(api, st(api).activeMailFolder ?? 'inbox', isRead ? -1 : 1);
        if (st(api).activeMailId === messageId && st(api).activeMail) {
          api.state.set('activeMail', { ...st(api).activeMail!, ...(isRead !== undefined ? { isRead } : {}), ...(flag !== undefined ? { flagged: flag } : {}) });
        }
        break;
      }
      case 'archive-mail': {
        const { messageId } = data as { messageId: string };
        const client = await ensureAuthenticated(false);
        await client.moveMail(messageId, 'archive');
        removeFromList(api, messageId);
        void loadMailFolders(api);
        break;
      }
      case 'delete-mail': {
        const { messageId } = data as { messageId: string };
        const client = await ensureAuthenticated(false);
        await client.moveMail(messageId, 'deleteditems');
        removeFromList(api, messageId);
        void loadMailFolders(api);
        break;
      }
      case 'move-mail': {
        const { messageId, folderId } = data as { messageId: string; folderId: string };
        const client = await ensureAuthenticated(false);
        await client.moveMail(messageId, folderId);
        removeFromList(api, messageId);
        void loadMailFolders(api);
        break;
      }
      case 'search-mail': {
        const { query } = data as { query: string };
        const q = query.trim();
        const seq = ++mailSearchSeq;
        if (!q) { api.state.set('mailSearch', null); break; }
        api.state.set('mailSearch', { query: q, loading: true, results: [] });
        try {
          const client = await ensureAuthenticated(false);
          const results = await client.searchMail(q, 40);
          if (seq === mailSearchSeq) {
            api.state.set('mailSearch', { query: q, loading: false, results });
            ensureSenderPhotos(api, client, results);
          }
        } catch (err) {
          if (seq === mailSearchSeq) {
            api.state.set('mailSearch', {
              query: q, loading: false, results: [],
              error: err instanceof Error ? err.message : String(err),
            });
          }
        }
        break;
      }
      case 'clear-mail-search': {
        mailSearchSeq++;
        api.state.set('mailSearch', null);
        break;
      }
      case 'compose-mail': {
        const { mode } = data as { mode: MailComposeState['mode'] };
        const me: MailAddress = { name: tokenCache.getDisplayName(), address: tokenCache.getEmail() ?? '' };
        api.state.set('composingMail', makeCompose(mode, st(api).activeMail ?? null, me, st(api).mailSignature ?? null));
        break;
      }
      case 'close-compose': {
        api.state.set('composingMail', null);
        break;
      }
      case 'send-mail': {
        const { mail } = data as { mail: OutgoingMail };
        const compose = st(api).composingMail;
        api.state.set('sendingMail', true);
        try {
          const client = await ensureAuthenticated(true);
          if (compose && compose.mode !== 'new' && compose.sourceId) {
            await client.replyMail(compose.sourceId, compose.mode, mail);
          } else {
            if (mail.to.length === 0) throw new Error('At least one recipient is required');
            await client.sendMail(mail);
          }
          api.state.set('composingMail', null);
          void loadMailFolders(api);
        } finally {
          api.state.set('sendingMail', false);
        }
        break;
      }
      case 'download-attachment': {
        const { messageId, attachmentId } = data as { messageId: string; attachmentId: string; name: string };
        const client = await ensureAuthenticated(false);
        const { dataUrl } = await client.getMailAttachment(messageId, attachmentId);
        const key = `dl:${attachmentId}`;
        inlineCache.set(key, dataUrl);
        const url = mediaServer.urlFor('mailinline', key) ?? dataUrl;
        api.state.set('mailInlineAttachments', { ...(st(api).mailInlineAttachments ?? {}), [key]: url });
        break;
      }
      case 'open-in-outlook': {
        const { url } = data as { url: string };
        try {
          const u = new URL(url);
          if (u.protocol === 'https:' && (u.hostname === 'outlook.office.com' || u.hostname === 'outlook.office365.com')) {
            await api.shell.openExternal(url);
            break;
          }
        } catch { /* fall through */ }
        throw new Error('Refusing to open non-Outlook URL');
      }
      default:
        log.warn(`Unknown mail action: ${action}`);
    }
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    log.error(`Mail action '${action}' failed: ${msg}`);
    api.state.set('mailError', msg);
  }
}
