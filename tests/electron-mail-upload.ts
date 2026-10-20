// Opt-in real Electron transport check; all requests stay on a synthetic loopback server.
import assert from 'node:assert/strict';
import { createServer } from 'node:http';
import { writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { GraphClient } from '../src/backend/graph-client.ts';
import { buildMsgraphTools } from '../src/backend/tools.ts';
import { disposeScheduler } from '../src/backend/request-scheduler.ts';
import { setLogger } from '../src/backend/logger-singleton.ts';
import * as tokens from '../src/backend/token-cache.ts';
import { CLIENT_ID_OFFICE } from '../src/shared/constants.ts';
import type { PluginAPI } from '../src/shared/types.ts';

const { app, net } = require('electron');
app.setPath('userData', join(__dirname, 'profile'));
const uploadUrl = 'https://outlook.office.com/test-upload?authtoken=synthetic';
const csv = Buffer.alloc(5_598_817, 'x');
csv.write('name,value\n');

app.whenReady().then(async () => {
  app.dock?.hide();
  const file = join(__dirname, 'report.csv');
  await writeFile(file, csv);
  setLogger({ info() {}, warn() {}, error() {} });
  const scopes = 'Mail.ReadWrite Mail.Send';
  const jwt = `e30.${Buffer.from(JSON.stringify({ oid: 'test-user', tid: 'test-tenant', appid: CLIENT_ID_OFFICE, aud: 'https://graph.microsoft.com', scp: scopes })).toString('base64url')}.test`;
  tokens.set({ accessToken: jwt, refreshToken: 'synthetic', expiresAt: Date.now() + 3600_000,
    clientId: CLIENT_ID_OFFICE, objectId: 'test-user', tenantId: 'test-tenant', scopes, email: '', displayName: null });
  let chunks: Buffer[] = [];
  let paths: string[] = [];
  let serverError: unknown;
  const server = createServer(async (req, res) => {
    try {
      const received: Buffer[] = [];
      for await (const chunk of req) received.push(Buffer.from(chunk));
      const bytes = Buffer.concat(received);
      paths.push(req.url!);
      res.setHeader('Content-Type', 'application/json');
      if (req.url === '/test-upload') {
        assert.equal(req.headers.authorization, undefined);
        assert.equal(req.headers['content-length'], String(bytes.length));
        assert.equal(req.headers['transfer-encoding'], undefined);
        const offset = chunks.reduce((n, b) => n + b.length, 0);
        assert.equal(req.headers['content-range'], `bytes ${offset}-${offset + bytes.length - 1}/${csv.length}`);
        chunks.push(bytes);
        const end = offset + bytes.length;
        if (end === csv.length) { res.writeHead(201); res.end(); }
        else res.end(JSON.stringify({ nextExpectedRanges: [String(end)] }));
      } else if (req.url?.endsWith('/createUploadSession')) {
        assert.equal(JSON.parse(bytes.toString()).AttachmentItem.size, csv.length);
        res.writeHead(201); res.end(JSON.stringify({ uploadUrl }));
      } else if (req.url?.endsWith('/send')) {
        assert.deepEqual(Buffer.concat(chunks), csv);
        res.writeHead(202); res.end();
      } else if (req.url === '/v1.0/me/messages' || req.url === '/v1.0/me/messages/original/createReply') {
        res.writeHead(201); res.end(JSON.stringify({ id: 'draft' }));
      } else throw new Error(`Unexpected local request ${req.url}`);
    } catch (error) {
      serverError = error;
      res.writeHead(400); res.end('{}');
    }
  });
  await new Promise<void>(r => server.listen(0, '127.0.0.1', r));
  const address = server.address() as { port: number };
  try {
    // Direct utility-process fetch and the streamed-body shape used by Kai's Node broker.
    for (const streamed of [false, true]) {
      const api = {
        fetch: async (input: string | URL, init?: RequestInit) => {
          const url = new URL(String(input));
          assert.ok(url.origin === 'https://graph.microsoft.com' || url.href === uploadUrl);
          const local = `http://127.0.0.1:${address.port}${url.pathname}`;
          if (!streamed || !init?.body) return net.fetch(local, init);
          const request = new Request(local, { ...init, duplex: 'half' } as RequestInit);
          return net.fetch(local, { ...init, headers: request.headers, body: request.body, duplex: 'half' });
        },
        config: { getPluginData: () => ({ preferences: { adaptiveFociRouting: false, maxThrottleWaitSeconds: 0 } }) },
        state: { set() {}, get: () => ({}) },
      } as unknown as PluginAPI;
      const tools = buildMsgraphTools({ api, ensureAuthenticated: async () => new GraphClient(api, false) });
      try {
        for (const name of ['send-mail', 'reply-to-mail']) {
          chunks = []; paths = []; serverError = undefined;
          const result = await tools.find(t => t.name === name)!.execute({
            to: ['synthetic@example.test'], subject: 'Synthetic test', body: 'Synthetic test',
            messageId: 'original', filePaths: [file],
          }) as { success?: boolean; error?: string };
          assert.equal(serverError, undefined);
          assert.equal(result.success, true, result.error);
          assert.equal(chunks.length, 2);
          assert.deepEqual(Buffer.concat(chunks), csv);
          assert.equal(paths.filter(p => p.endsWith('/send')).length, 1);
          assert.equal(paths.at(-1), '/v1.0/me/messages/draft/send');
          console.log(JSON.stringify({ tool: name, streamed, bytes: csv.length, chunks: chunks.length, electron: process.versions.electron, passed: true }));
        }
      } finally { disposeScheduler(api); }
    }
  } finally {
    server.closeAllConnections();
    await new Promise<void>(r => server.close(() => r()));
    tokens.invalidateSession();
  }
  app.quit();
}).catch((error: unknown) => { console.error(error); app.exit(1); });
