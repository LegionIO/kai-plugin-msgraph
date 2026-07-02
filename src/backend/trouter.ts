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
import { TROUTER_CONNECT_URL, TROUTER_CLIENT_VERSION } from '../shared/constants.js';
import type { PluginAPI } from '../shared/types.js';
import { ensureRegion, ic3Token } from './ic3-client.js';
import { getLogger } from './logger-singleton.js';

// ── Event shapes emitted to the handler ──

export type TrouterEvent =
  | { kind: 'connected' }
  | { kind: 'disconnected'; willRetry: boolean }
  | { kind: 'error'; message: string }
  | { kind: 'message'; chatId: string; messageId: string; fromUserId: string | null; fromName: string | null; own: boolean }
  | { kind: 'messageUpdate'; chatId: string; messageId: string }
  | { kind: 'typing'; chatId: string; fromUserId: string; fromName: string | null }
  | { kind: 'readReceipt'; chatId: string; userId: string; lastReadMessageId: string }
  | { kind: 'conversationUpdate'; chatId: string };

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

export class TrouterListener {
  private ws: WS | null = null;
  private readonly epid = randomUUID();
  private readonly corId = randomUUID();
  private conNum = 0;
  private hbTimer: ReturnType<typeof setInterval> | null = null;
  private retryTimer: ReturnType<typeof setTimeout> | null = null;
  private backoffMs = 1000;
  private stopped = false;
  private connectparams: Record<string, unknown> = {};

  constructor(
    private readonly api: PluginAPI,
    private readonly myUserId: string | null,
    private readonly onEvent: TrouterHandler,
  ) {}

  start(): void {
    this.stopped = false;
    void this.connect();
  }

  stop(): void {
    this.stopped = true;
    if (this.retryTimer) { clearTimeout(this.retryTimer); this.retryTimer = null; }
    if (this.hbTimer) { clearInterval(this.hbTimer); this.hbTimer = null; }
    try { this.ws?.close(); } catch { /* ignore */ }
    this.ws = null;
  }

  private async connect(): Promise<void> {
    if (this.stopped) return;
    let rgn;
    try {
      rgn = await ensureRegion(this.api);
    } catch (err) {
      getLogger().warn(`trouter: region acquire failed (${err}); retrying`);
      return this.scheduleReconnect();
    }
    if (!rgn.skypeToken) {
      getLogger().warn('trouter: no skypetoken; real-time disabled');
      return;
    }
    const skypeToken = rgn.skypeToken;

    const tc = encodeURIComponent(
      JSON.stringify({ cv: TROUTER_CLIENT_VERSION, ua: 'TeamsCDL', hr: '', v: '1415/1.0.0.0' }),
    );
    const url =
      `${TROUTER_CONNECT_URL}?tc=${tc}&timeout=40&epid=${this.epid}&ccid=` +
      `&dom=teams.microsoft.com&cor_id=${this.corId}&con_num=${Date.now()}_${this.conNum++}`;

    getLogger().info(`trouter: connecting (epid=${this.epid})`);
    // Corporate TLS interception (Zscaler et al.) breaks Node's bundled CA set on
    // *.trouter.teams.microsoft.com. Honor NODE_EXTRA_CA_CERTS if present; otherwise
    // fall back to no-verify since api.fetch (Electron net) already validated
    // teams.microsoft.com when fetching the skypetoken this connection authenticates with.
    const ws = new WS(url, {
      origin: 'https://teams.microsoft.com',
      rejectUnauthorized: !!process.env.NODE_EXTRA_CA_CERTS,
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
      if (this.hbTimer) clearInterval(this.hbTimer);
      this.hbTimer = setInterval(() => send('2::'), HEARTBEAT_MS);
    });

    ws.on('message', (data) => {
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
            if (info.surl) void this.register(info.surl);
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
      if (this.ws === ws) this.ws = null;
      getLogger().info(`trouter: closed (${code} ${reason?.toString?.() ?? ''})`);
      this.onEvent({ kind: 'disconnected', willRetry: !this.stopped });
      if (!this.stopped) this.scheduleReconnect();
    });
    ws.on('error', (e) => {
      const msg = e instanceof Error ? e.message : String(e);
      getLogger().warn(`trouter: ws error ${msg}`);
      this.onEvent({ kind: 'error', message: msg });
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

  private async register(surl: string): Promise<void> {
    try {
      const [rgn, ic3] = await Promise.all([ensureRegion(this.api), ic3Token(this.api)]);
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
  }

  private handlePush(req: { url?: string; body?: string }): void {
    if (!req.body) return;
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

    if (mt === 'Control/Typing' || mt === 'Control/ClearTyping') {
      if (!own && fromUserId) {
        this.onEvent({ kind: 'typing', chatId, fromUserId, fromName: res.imdisplayname ?? null });
      }
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
      this.onEvent({
        kind: 'message',
        chatId,
        messageId: String(res.id ?? ''),
        fromUserId,
        fromName: res.imdisplayname ?? null,
        own,
      });
      return;
    }
    if (rt === 'MessageUpdate') {
      this.onEvent({ kind: 'messageUpdate', chatId, messageId: String(res.id ?? '') });
      return;
    }
    if (rt === 'ConversationUpdate') {
      const cid = (body.resource as { id?: string } | undefined)?.id;
      if (cid && !cid.startsWith('48:')) this.onEvent({ kind: 'conversationUpdate', chatId: cid });
      return;
    }
  }
}
