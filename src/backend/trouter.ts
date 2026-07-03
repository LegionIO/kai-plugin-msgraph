/**
 * Trouter — Teams' real-time push channel.
 *
 * A raw WebSocket to go.trouter.teams.microsoft.com/v4/c speaking a socket.io-0.9
 * wire format. After connect the client emits `user.authenticate` with the
 * skypetoken; the server replies with `trouter.connected` carrying a per-session
 * `surl`. Registering that surl with the Teams registrar causes IC3 to push chat
 * events (new messages, edits, reactions, typing, read receipts) to the socket as
 * type-3 frames shaped like HTTP requests. Each is acked with a type-3 response.
 *
 * Undocumented Teams-client protocol; if any leg fails the plugin still has its
 * poll fallback.
 */

import { randomUUID } from 'crypto';
import WS from 'ws';
import { TROUTER_CONNECT_URL, TROUTER_CLIENT_VERSION, CLIENT_ID_TEAMS, PRESENCE_SCOPE } from '../shared/constants.js';
import type { PluginAPI } from '../shared/types.js';
import { ensureRegion, ic3Token } from './ic3-client.js';
import { acquireFociAccessToken } from './auth.js';
import { getLogger } from './logger-singleton.js';

// ── Event shapes emitted to the handler ──

export type TrouterEvent =
  | { kind: 'connected' }
  | { kind: 'disconnected'; willRetry: boolean }
  | { kind: 'error'; message: string }
  | { kind: 'message'; chatId: string; messageId: string; fromUserId: string | null; fromName: string | null; own: boolean; preview: string | null }
  | {
      kind: 'messageUpdate';
      chatId: string;
      messageId: string;
      reaction: { type: string; userId: string | null } | null;
    }
  | { kind: 'typing'; chatId: string; fromUserId: string; fromName: string | null }
  | { kind: 'clearTyping'; chatId: string; fromUserId: string }
  | { kind: 'readReceipt'; chatId: string; userId: string; lastReadMessageId: string }
  | { kind: 'conversationUpdate'; chatId: string }
  | { kind: 'presence'; userId: string; availability: string; activity: string };

export type TrouterHandler = (ev: TrouterEvent) => void;

// ── socket.io 0.9 framing ──

const FRAME = /^(\d):(\d*)\+?:[^:]*(?::([\s\S]*))?$/;

function parseFrame(s: string): { type: string; id: string; data: string } | null {
  const m = FRAME.exec(s);
  return m ? { type: m[1], id: m[2] ?? '', data: m[3] ?? '' } : null;
}

function mriToUserId(mri: string | null | undefined): string | null {
  if (!mri) return null;
  const m = /8:orgid:([0-9a-fA-F-]{36})/.exec(mri);
  return m ? m[1] : null;
}

function chatIdFromLink(link: string | null | undefined): string | null {
  if (!link) return null;
  const i = link.lastIndexOf('/');
  return decodeURIComponent(i >= 0 ? link.slice(i + 1) : link);
}

// ── Listener ──

const HEARTBEAT_MS = 25_000;
const MAX_BACKOFF_MS = 60_000;
const REGISTRAR_REFRESH_MS = 50 * 60_000; // TTL is 3600s; refresh before expiry

export class TrouterListener {
  private ws: WS | null = null;
  private readonly epid = randomUUID();
  private readonly corId = randomUUID();
  private conNum = 0;
  private gen = 0;
  private hbTimer: ReturnType<typeof setInterval> | null = null;
  private regTimer: ReturnType<typeof setInterval> | null = null;
  private retryTimer: ReturnType<typeof setTimeout> | null = null;
  private backoffMs = 1000;
  private stopped = false;
  private tlsInsecureFallback = false;
  private connectparams: Record<string, unknown> = {};
  private surl: string | null = null;
  private presenceSubs = new Set<string>();

  constructor(
    private readonly api: PluginAPI,
    private readonly myUserId: string | null,
    private readonly onEvent: TrouterHandler,
  ) {}

  /** Subscribe to live presence updates for these userIds (idempotent, additive). */
  subscribePresence(userIds: Iterable<string>): void {
    const add: string[] = [];
    for (const id of userIds) {
      if (!id || this.presenceSubs.has(id)) continue;
      this.presenceSubs.add(id);
      add.push(id);
    }
    if (add.length && this.surl) void this.sendPresenceSub(add, false);
  }

  start(): void {
    this.stopped = false;
    void this.connect();
  }

  stop(): void {
    this.stopped = true;
    this.gen++;
    if (this.retryTimer) { clearTimeout(this.retryTimer); this.retryTimer = null; }
    if (this.hbTimer) { clearInterval(this.hbTimer); this.hbTimer = null; }
    if (this.regTimer) { clearInterval(this.regTimer); this.regTimer = null; }
    this.surl = null;
    try { this.ws?.close(); } catch { /* ignore */ }
    this.ws = null;
  }

  private live(g: number): boolean {
    return !this.stopped && g === this.gen;
  }

  private async connect(): Promise<void> {
    const g = ++this.gen;
    if (this.stopped) return;
    let rgn;
    try {
      rgn = await ensureRegion(this.api);
    } catch (err) {
      if (!this.live(g)) return;
      getLogger().warn(`trouter: region acquire failed (${err}); retrying`);
      return this.scheduleReconnect();
    }
    if (!this.live(g)) return;
    if (!rgn.skypeToken) {
      getLogger().warn('trouter: no skypetoken; real-time disabled');
      this.onEvent({ kind: 'disconnected', willRetry: false });
      return;
    }
    const skypeToken = rgn.skypeToken;

    const tc = encodeURIComponent(
      JSON.stringify({ cv: TROUTER_CLIENT_VERSION, ua: 'TeamsCDL', hr: '', v: '1415/1.0.0.0' }),
    );
    const url =
      `${TROUTER_CONNECT_URL}?tc=${tc}&timeout=40&epid=${this.epid}&ccid=` +
      `&dom=teams.microsoft.com&cor_id=${this.corId}&con_num=${Date.now()}_${this.conNum++}`;

    getLogger().info(`trouter: connecting (epid=${this.epid}, tlsVerify=${!this.tlsInsecureFallback})`);
    const ws = new WS(url, {
      origin: 'https://teams.microsoft.com',
      // Verify by default. Corporate TLS interception may break Node's CA set;
      // on a cert error we retry once with verification disabled, surfaced via
      // state.realtimeError so the user knows.
      rejectUnauthorized: !this.tlsInsecureFallback,
      handshakeTimeout: 15_000,
    });
    this.ws = ws;

    const send = (frame: string) => {
      try { if (ws.readyState === WS.OPEN) ws.send(frame); } catch { /* closing */ }
    };
    const emit = (name: string, ...args: unknown[]) => send(`5:::${JSON.stringify({ name, args })}`);
    const ack = (id: string) => { if (id) send(`6:::${id}`); };
    const respond = (reqId: number | string, status: number) =>
      send(`3:::${JSON.stringify({ id: reqId, status, headers: {}, body: '' })}`);

    ws.on('open', () => {
      if (!this.live(g)) { try { ws.close(); } catch { /* ignore */ } return; }
      if (this.hbTimer) clearInterval(this.hbTimer);
      this.hbTimer = setInterval(() => send('2::'), HEARTBEAT_MS);
    });

    ws.on('message', (data) => {
      if (!this.live(g)) return;
      const raw = typeof data === 'string' ? data : data.toString('utf-8');
      const f = parseFrame(raw);
      if (!f) return;
      switch (f.type) {
        case '1': // connect
          emit('user.authenticate', {
            headers: { 'X-Skypetoken': skypeToken },
            connectparams: this.connectparams,
          });
          return;
        case '2': // heartbeat
          send('2::');
          return;
        case '5': { // event
          let j: { name?: string; args?: unknown[] };
          try { j = JSON.parse(f.data); } catch { ack(f.id); return; }
          if (j.name === 'trouter.connected') {
            const info = (j.args?.[0] ?? {}) as {
              surl?: string;
              connectparams?: Record<string, unknown>;
            };
            if (info.connectparams) this.connectparams = info.connectparams;
            ack(f.id);
            emit('user.activity', { state: 'active', cv: `${this.corId}.${this.conNum}` });
            this.backoffMs = 1000;
            if (info.surl) void this.register(g, info.surl);
            this.onEvent({ kind: 'connected' });
          } else if (j.name === 'trouter.reconnect') {
            const info = (j.args?.[0] ?? {}) as { connectparams?: Record<string, unknown> };
            if (info.connectparams) this.connectparams = info.connectparams;
            ack(f.id);
            try { ws.close(); } catch { /* ignore */ }
          } else {
            ack(f.id);
          }
          return;
        }
        case '3': { // pushed request
          let req: { id?: number; url?: string; body?: string };
          try { req = JSON.parse(f.data); } catch { return; }
          this.handlePush(req);
          if (req.id !== undefined) respond(req.id, 200);
          return;
        }
        case '0': // disconnect
        case '7': // error
          getLogger().warn(`trouter: server ${f.type === '0' ? 'disconnect' : 'error'}: ${f.data}`);
          return;
      }
    });

    ws.on('close', (code, reason) => {
      if (this.hbTimer) { clearInterval(this.hbTimer); this.hbTimer = null; }
      if (this.regTimer) { clearInterval(this.regTimer); this.regTimer = null; }
      if (this.ws === ws) this.ws = null;
      getLogger().info(`trouter: closed (${code} ${reason?.toString?.() ?? ''})`);
      if (!this.live(g)) return;
      this.onEvent({ kind: 'disconnected', willRetry: true });
      this.scheduleReconnect();
    });
    ws.on('error', (e) => {
      const msg = e instanceof Error ? e.message : String(e);
      const code = (e as NodeJS.ErrnoException).code ?? '';
      const isCert = /CERT|SELF_SIGNED|UNABLE_TO_VERIFY|DEPTH_ZERO/i.test(code) ||
        /certificate|self[- ]signed/i.test(msg);
      if (isCert && !this.tlsInsecureFallback) {
        this.tlsInsecureFallback = true;
        getLogger().warn(
          `trouter: TLS verification failed (${code || msg}); retrying without verification. ` +
          `Set NODE_EXTRA_CA_CERTS to your corporate CA to avoid this.`,
        );
        if (this.live(g)) this.onEvent({ kind: 'error', message: `TLS unverified fallback: ${code || msg}` });
        return;
      }
      getLogger().warn(`trouter: ws error ${msg}`);
      if (this.live(g)) this.onEvent({ kind: 'error', message: msg });
    });
  }

  private scheduleReconnect(): void {
    if (this.stopped || this.retryTimer) return;
    const delay = this.backoffMs;
    this.backoffMs = Math.min(MAX_BACKOFF_MS, this.backoffMs * 2);
    this.retryTimer = setTimeout(() => {
      this.retryTimer = null;
      void this.connect();
    }, delay);
  }

  private async register(g: number, surl: string): Promise<void> {
    try {
      const [rgn, ic3] = await Promise.all([ensureRegion(this.api), ic3Token(this.api)]);
      if (!this.live(g)) return;
      const resp = await this.api.fetch(rgn.registrarUrl, {
        method: 'POST',
        headers: {
          Authorization: `Bearer ${ic3}`,
          'Content-Type': 'application/json',
          'X-MS-Migration': 'True',
        },
        body: JSON.stringify({
          clientDescription: {
            appId: 'TeamsCDLWebWorker',
            aesKey: '',
            languageId: 'en-US',
            platform: 'electron',
            templateKey: 'TeamsCDLWebWorker_2.6',
            platformUIVersion: '1415/1.0.0.0',
          },
          registrationId: this.epid,
          nodeId: '',
          transports: { TROUTER: [{ context: '', path: surl, ttl: 3600 }] },
        }),
      });
      if (!resp.ok) getLogger().warn(`trouter: registrar ${resp.status}`);
      else getLogger().info(`trouter: registered surl (registrar ${resp.status})`);
    } catch (err) {
      getLogger().warn(`trouter: registrar failed (${err})`);
    }
    if (!this.live(g)) return;
    this.surl = surl;
    if (this.presenceSubs.size) void this.sendPresenceSub([...this.presenceSubs], true);
    if (this.regTimer) clearInterval(this.regTimer);
    this.regTimer = setInterval(() => {
      if (this.surl && this.live(g)) void this.register(g, this.surl);
    }, REGISTRAR_REFRESH_MS);
  }

  private async sendPresenceSub(userIds: string[], purge: boolean): Promise<void> {
    if (!this.surl) return;
    try {
      const [rgn, tok] = await Promise.all([
        ensureRegion(this.api),
        acquireFociAccessToken(this.api, CLIENT_ID_TEAMS, PRESENCE_SCOPE),
      ]);
      const resp = await this.api.fetch(`${rgn.presenceUPS}/v1/pubsub/subscriptions/${this.epid}`, {
        method: 'POST',
        headers: {
          Authorization: `Bearer ${tok}`,
          'Content-Type': 'application/json',
          'x-ms-client-user-agent': 'Teams-V2-Desktop',
          'x-ms-correlation-id': randomUUID(),
          'x-ms-endpoint-id': this.epid,
        },
        body: JSON.stringify({
          trouterUri: `${this.surl}unifiedPresenceService`,
          shouldPurgePreviousSubscriptions: purge,
          subscriptionsToAdd: userIds.map((id) => ({ mri: `8:orgid:${id}`, source: 'ups' })),
          subscriptionsToRemove: [],
        }),
      });
      if (!resp.ok) getLogger().warn(`trouter: ups pubsub ${resp.status}`);
    } catch (err) {
      getLogger().warn(`trouter: ups pubsub failed (${err})`);
    }
  }

  private handlePush(req: { url?: string; body?: string }): void {
    if (!req.body) return;
    if (req.url?.endsWith('/unifiedPresenceService')) {
      try {
        type PresItem = {
          mri?: string;
          availability?: string;
          activity?: string;
          presence?: { availability?: string; activity?: string };
        };
        const raw = JSON.parse(req.body) as PresItem | { presence?: PresItem[] };
        const items: PresItem[] = Array.isArray((raw as { presence?: PresItem[] }).presence)
          ? (raw as { presence: PresItem[] }).presence
          : [raw as PresItem];
        for (const it of items) {
          const uid = mriToUserId(it.mri);
          const avail = it.presence?.availability ?? it.availability;
          const act = it.presence?.activity ?? it.activity ?? avail;
          if (uid && avail) {
            this.onEvent({ kind: 'presence', userId: uid, availability: avail, activity: act ?? avail });
          }
        }
      } catch { /* ignore */ }
      return;
    }
    let body: {
      resourceType?: string;
      resource?: {
        id?: string | number;
        to?: string | null;
        conversationLink?: string | null;
        messagetype?: string;
        content?: string;
        from?: string;
        imdisplayname?: string;
        properties?: { emotions?: Array<{ key?: string; users?: Array<{ mri?: string; time?: number }> }> };
      };
    };
    try { body = JSON.parse(req.body); } catch { return; }
    const res = body.resource ?? {};
    const chatId = res.to || chatIdFromLink(res.conversationLink) ||
      (body as { resource?: { id?: string } }).resource?.id || '';
    if (!chatId || chatId.startsWith('48:')) return; // skip activity-feed pseudo-thread
    const mt = res.messagetype ?? '';
    const rt = body.resourceType ?? '';
    const fromUserId = mriToUserId(res.from);
    const own = !!this.myUserId && fromUserId === this.myUserId;

    if (mt === 'Control/Typing') {
      if (!own && fromUserId) {
        this.onEvent({ kind: 'typing', chatId, fromUserId, fromName: res.imdisplayname ?? null });
      }
      return;
    }
    if (mt === 'Control/ClearTyping') {
      if (!own && fromUserId) this.onEvent({ kind: 'clearTyping', chatId, fromUserId });
      return;
    }
    if (mt === 'ThreadActivity/MemberConsumptionHorizonUpdate') {
      try {
        const c = JSON.parse(res.content ?? '{}') as { user?: string; consumptionhorizon?: string };
        const uid = mriToUserId(c.user);
        const last = (c.consumptionhorizon ?? '').split(';')[0];
        if (uid && last) this.onEvent({ kind: 'readReceipt', chatId, userId: uid, lastReadMessageId: last });
      } catch { /* ignore */ }
      return;
    }
    if (rt === 'NewMessage') {
      const preview = (res.content ?? '').replace(/<[^>]+>/g, '').replace(/&[a-z#0-9]+;/gi, ' ').trim().slice(0, 200) || null;
      this.onEvent({
        kind: 'message',
        chatId,
        messageId: String(res.id ?? ''),
        fromUserId,
        fromName: res.imdisplayname ?? null,
        own,
        preview,
      });
      return;
    }
    if (rt === 'MessageUpdate') {
      let reaction: { type: string; userId: string | null } | null = null;
      const props = (res as { properties?: { emotions?: Array<{ key?: string; users?: Array<{ mri?: string; time?: number }> }> } }).properties;
      const emo = props?.emotions;
      if (Array.isArray(emo) && emo.length) {
        let best: { key: string; mri: string | null; time: number } | null = null;
        for (const e of emo) {
          for (const u of e.users ?? []) {
            const t = Number(u.time ?? 0);
            if (!best || t > best.time) best = { key: e.key ?? '', mri: u.mri ?? null, time: t };
          }
        }
        if (best?.key) reaction = { type: best.key, userId: mriToUserId(best.mri) };
      }
      this.onEvent({ kind: 'messageUpdate', chatId, messageId: String(res.id ?? ''), reaction });
      return;
    }
    if (rt === 'ConversationUpdate') {
      const cid = (body.resource as { id?: string } | undefined)?.id;
      if (cid && !cid.startsWith('48:')) this.onEvent({ kind: 'conversationUpdate', chatId: cid });
      return;
    }
  }
}
