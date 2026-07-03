/**
 * Localhost HTTP server that serves cached profile photos and hosted-content
 * images so the renderer receives short URLs instead of multi-KB data-URLs
 * over IPC. Bound to 127.0.0.1 with a random path prefix; falls back cleanly
 * (baseUrl() returns null) if the port can't be opened, in which case callers
 * publish data-URLs as before.
 */

import { createServer, type Server } from 'http';
import { randomBytes } from 'crypto';
import { getLogger } from './logger-singleton.js';

type Lookup = (key: string) => string | null | undefined;

let server: Server | null = null;
let base: string | null = null;
let prefix = '';
const routes = new Map<string, Lookup>();

export function register(route: string, lookup: Lookup): void {
  routes.set(route, lookup);
}

export function baseUrl(): string | null {
  return base;
}

export function urlFor(route: string, key: string): string | null {
  return base ? `${base}/${route}/${encodeURIComponent(key)}` : null;
}

const DATA_URL = /^data:([^;,]+)?(;base64)?,(.*)$/s;

function decode(dataUrl: string): { mime: string; body: Buffer } | null {
  const m = DATA_URL.exec(dataUrl);
  if (!m) return null;
  const mime = m[1] || 'application/octet-stream';
  const body = m[2] ? Buffer.from(m[3], 'base64') : Buffer.from(decodeURIComponent(m[3]), 'utf-8');
  return { mime, body };
}

export function start(): void {
  if (server) return;
  prefix = randomBytes(12).toString('hex');
  const srv = createServer((req, res) => {
    try {
      const url = new URL(req.url ?? '/', 'http://127.0.0.1');
      const parts = url.pathname.split('/').filter(Boolean);
      if (parts[0] !== prefix || parts.length < 3) {
        res.writeHead(404).end();
        return;
      }
      const route = parts[1];
      const key = decodeURIComponent(parts.slice(2).join('/'));
      const lookup = routes.get(route);
      const dataUrl = lookup?.(key);
      if (!dataUrl) {
        res.writeHead(404).end();
        return;
      }
      const dec = decode(dataUrl);
      if (!dec) {
        res.writeHead(500).end();
        return;
      }
      res.writeHead(200, {
        'Content-Type': dec.mime,
        'Content-Length': String(dec.body.length),
        'Cache-Control': 'private, max-age=3600',
        'Access-Control-Allow-Origin': '*',
      });
      res.end(dec.body);
    } catch {
      res.writeHead(500).end();
    }
  });
  srv.on('error', (err) => {
    getLogger().warn(`media-server: ${err}; falling back to data-URLs`);
    base = null;
    server = null;
  });
  srv.listen(0, '127.0.0.1', () => {
    const addr = srv.address();
    if (addr && typeof addr === 'object') {
      base = `http://127.0.0.1:${addr.port}/${prefix}`;
      getLogger().info(`media-server: listening on ${base}`);
    }
  });
  server = srv;
}

export function stop(): void {
  try { server?.close(); } catch { /* ignore */ }
  server = null;
  base = null;
}
