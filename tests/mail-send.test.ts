import assert from 'node:assert/strict';
import { afterEach, beforeEach, test } from 'node:test';
import { mkdtemp, rm, truncate, writeFile } from 'node:fs/promises';
import { homedir, tmpdir } from 'node:os';
import { join, relative } from 'node:path';
import { buildMsgraphTools } from '../src/backend/tools.ts';
import { GraphClient, MAX_MAIL_ATTACHMENT_BYTES } from '../src/backend/graph-client.ts';
import { graphPolicy } from '../src/backend/request-policy.ts';
import { disposeScheduler } from '../src/backend/request-scheduler.ts';
import { clearFociTokens } from '../src/backend/auth.ts';
import * as tokens from '../src/backend/token-cache.ts';
import { setLogger } from '../src/backend/logger-singleton.ts';
import { CLIENT_ID_OFFICE } from '../src/shared/constants.ts';
import type { PluginAPI } from '../src/shared/types.ts';

const draftPath = '/v1.0/me/messages/draft%2F1';
const uploadUrl = 'https://outlook.office.com/upload/session?authtoken=test-upload-credential';
const recipient = { emailAddress: { address: 'recipient@example.test' } };
const newMail = { to: ['recipient@example.test'], subject: 'Report', body: 'Attached report.' };
const reply = { messageId: 'original/1', body: 'Attached report.' };
const smallAttachment = { name: 'note.txt', contentType: 'text/plain', contentBytes: Buffer.from('hello').toString('base64') };
let dir: string;
const apis: PluginAPI[] = [];

beforeEach(async () => {
  dir = await mkdtemp(join(tmpdir(), 'msgraph-send-test-'));
  tokens.invalidateSession();
  clearFociTokens();
  const scopes = 'Mail.ReadWrite Mail.Send';
  const jwt = `e30.${Buffer.from(JSON.stringify({ oid: 'user', tid: 'tenant', appid: CLIENT_ID_OFFICE, aud: 'https://graph.microsoft.com', scp: scopes })).toString('base64url')}.test`;
  tokens.set({ accessToken: jwt, refreshToken: 'test-refresh', clientId: CLIENT_ID_OFFICE,
    expiresAt: Date.now() + 3600_000, objectId: 'user', tenantId: 'tenant', scopes, email: '', displayName: null });
  setLogger({ info() {}, warn() {}, error() {} });
});
afterEach(async () => {
  for (const api of apis.splice(0)) disposeScheduler(api);
  tokens.invalidateSession();
  clearFociTokens();
  await rm(dir, { recursive: true, force: true });
});

interface Request {
  url: string;
  path: string;
  method: string;
  headers: Headers;
  body?: Record<string, any>;
  bytes?: Buffer;
}

function harness(intercept?: (req: Request) => Response | undefined) {
  const requests: Request[] = [];
  const api = {
    fetch: async (input: string | URL, init?: RequestInit) => {
      const url = String(input);
      const req: Request = { url, path: new URL(url).pathname, method: init?.method ?? 'GET', headers: new Headers(init?.headers) };
      if (typeof init?.body === 'string') req.body = JSON.parse(init.body);
      else if (init?.body instanceof Uint8Array) req.bytes = Buffer.from(init.body);
      requests.push(req);
      const override = intercept?.(req);
      if (override) return override;
      if (url === uploadUrl) {
        const range = /^bytes (\d+)-(\d+)\/(\d+)$/.exec(req.headers.get('content-range') ?? '')!;
        const end = Number(range[2]) + 1;
        return end === Number(range[3]) ? new Response(null, { status: 201 }) : Response.json({ nextExpectedRanges: [String(end)] });
      }
      if (req.path.endsWith('/attachments/createUploadSession')) return Response.json({ uploadUrl }, { status: 201 });
      if (req.path.endsWith('/attachments')) return Response.json({ id: 'attachment-1' }, { status: 201 });
      if (req.path === '/v1.0/me/messages' || /\/create(Reply|ReplyAll|Forward)$/.test(req.path)) {
        return Response.json({ id: 'draft/1' }, { status: 201 });
      }
      if (/\/(sendMail|send|reply|replyAll|forward)$/.test(req.path)) return new Response(null, { status: 202 });
      throw new Error(`Unexpected mocked request: ${req.method} ${req.path}`);
    },
    config: { getPluginData: () => ({ preferences: { adaptiveFociRouting: false, maxThrottleWaitSeconds: 0 } }) },
    state: { set() {}, get: () => ({}) },
  } as unknown as PluginAPI;
  apis.push(api);
  const client = new GraphClient(api, false);
  const tools = buildMsgraphTools({ api, ensureAuthenticated: async () => client });
  const run = async (name: string, args: unknown) => await tools.find(t => t.name === name)!.execute(args) as Record<string, unknown>;
  return { requests, run, client };
}

test('send-mail uploads a 5.6 MB local CSV in ordered chunks before sending once', async () => {
  const { run, requests } = harness();
  const csv = Buffer.alloc(5_600_000, 'x');
  csv.write('name,value\n');
  const path = join(dir, 'report.csv');
  await writeFile(path, csv);
  const result = await run('send-mail', { ...newMail, filePaths: [path] });
  assert.deepEqual(result, { success: true, to: newMail.to, subject: newMail.subject });
  assert.deepEqual(requests.map(r => r.path), ['/v1.0/me/messages', `${draftPath}/attachments/createUploadSession`, '/upload/session', '/upload/session', `${draftPath}/send`]);
  assert.equal(requests[0].body?.attachments, undefined);
  assert.deepEqual(requests[0].body?.toRecipients, [recipient]);
  assert.deepEqual(requests[1].body?.AttachmentItem, { attachmentType: 'file', name: 'report.csv', size: csv.length, contentType: 'text/csv' });
  const chunks = requests.filter(r => r.url === uploadUrl);
  assert.deepEqual(chunks.map(r => r.headers.get('content-range')), ['bytes 0-3276799/5600000', 'bytes 3276800-5599999/5600000']);
  assert.deepEqual(Buffer.concat(chunks.map(r => r.bytes!)), csv);
  for (const chunk of chunks) {
    assert.equal(chunk.headers.has('authorization'), false);
    assert.equal(chunk.headers.get('content-type'), 'application/octet-stream');
    assert.equal(chunk.headers.get('content-length'), String(chunk.bytes!.length));
  }
});

for (const mode of ['reply', 'replyAll', 'forward'] as const) {
  test(`${mode} attaches local and embedded files to the correct draft without clearing recipients`, async () => {
    const { run, requests } = harness();
    const path = join(dir, 'report.csv');
    await writeFile(path, 'a,b\n1,2\n');
    const result = await run('reply-to-mail', { ...reply, mode, ...(mode === 'forward' ? { to: newMail.to } : {}),
      attachments: [smallAttachment], filePaths: [`~/${relative(homedir(), path)}`] });
    assert.equal(result.success, true);
    const action = mode === 'reply' ? 'createReply' : mode === 'replyAll' ? 'createReplyAll' : 'createForward';
    assert.equal(requests[0].path, `/v1.0/me/messages/original%2F1/${action}`);
    assert.equal(requests[0].body?.message.toRecipients, undefined);
    assert.equal(requests[0].body?.message.subject, undefined);
    assert.equal(requests[0].body?.message.attachments, undefined);
    assert.match(requests[0].body?.message.body.content, /Attached report/);
    assert.deepEqual(requests[0].body?.toRecipients, mode === 'forward' ? [recipient] : undefined);
    assert.deepEqual(requests.slice(1).map(r => r.path), [`${draftPath}/attachments`, `${draftPath}/attachments`, `${draftPath}/send`]);
    assert.deepEqual(requests[1].body, { '@odata.type': '#microsoft.graph.fileAttachment', ...smallAttachment });
    assert.equal(requests[2].body?.name, 'report.csv');
    assert.equal(Buffer.from(requests[2].body?.contentBytes, 'base64').toString(), 'a,b\n1,2\n');
  });
}

test('large reply attachments use upload sessions and send the reply draft', async () => {
  const { run, requests } = harness();
  const path = join(dir, 'large.csv');
  await writeFile(path, Buffer.alloc(5_600_000, 65));
  assert.equal((await run('reply-to-mail', { ...reply, filePaths: [path] })).success, true);
  assert.equal(requests[0].path, '/v1.0/me/messages/original%2F1/createReply');
  assert.equal(requests.filter(r => r.url === uploadUrl).length, 2);
  assert.equal(requests.at(-1)?.path, `${draftPath}/send`);
});

test('upload progress also accepts the offset-range form', async () => {
  const { run, requests } = harness(req => req.url === uploadUrl && req.headers.get('content-range')?.startsWith('bytes 0-')
    ? Response.json({ nextExpectedRanges: ['3276800-'] }) : undefined);
  const path = join(dir, 'large.csv');
  await writeFile(path, Buffer.alloc(5_600_000, 65));
  assert.equal((await run('send-mail', { ...newMail, filePaths: [path] })).success, true);
  assert.equal(requests.at(-1)?.path, `${draftPath}/send`);
});

test('sending without attachments retains the direct send and reply paths', async () => {
  const { run, requests } = harness();
  assert.equal((await run('send-mail', newMail)).success, true);
  assert.equal((await run('reply-to-mail', reply)).success, true);
  assert.equal((await run('reply-to-mail', { ...reply, mode: 'forward', to: newMail.to })).success, true);
  assert.deepEqual(requests.map(r => r.path), ['/v1.0/me/sendMail', '/v1.0/me/messages/original%2F1/reply', '/v1.0/me/messages/original%2F1/forward']);
  assert.equal(requests[0].body?.saveToSentItems, true);
  assert.equal(requests[1].body?.message.toRecipients, undefined);
  assert.equal(requests[2].body?.message.toRecipients, undefined);
  assert.deepEqual(requests[2].body?.toRecipients, [recipient]);
});

test('inline attachment content IDs survive the draft attachment path', async () => {
  const { client, requests } = harness();
  await client.sendMail({ to: [{ name: null, address: 'recipient@example.test' }], subject: 'Inline', bodyHtml: '<img src="cid:image-1">',
    attachments: [{ ...smallAttachment, contentId: 'image-1' }] });
  assert.equal(requests[1].body?.contentId, 'image-1');
  assert.equal(requests[1].body?.isInline, true);
});

test('all local files are checked before creating a draft or sending', async () => {
  const { run, requests } = harness();
  const good = join(dir, 'good.csv');
  await writeFile(good, 'valid');
  for (const paths of [[good, join(dir, 'missing.csv')], ['relative.csv'], [dir], ['/dev/null']]) {
    assert.ok((await run('send-mail', { ...newMail, filePaths: paths })).error);
  }
  const huge = join(dir, 'oversized.csv');
  await writeFile(huge, '');
  await truncate(huge, MAX_MAIL_ATTACHMENT_BYTES + 1);
  assert.match(String((await run('send-mail', { ...newMail, filePaths: [huge] })).error), /attachment limit/);
  assert.match(String((await run('reply-to-mail', { ...reply, mode: 'forward', filePaths: [good] })).error), /recipient/);
  assert.deepEqual(requests, []);
});

test('a small attachment failure retains the draft and never sends', async () => {
  const { run, requests } = harness(req => req.path.endsWith('/attachments') ? Response.json({ error: { message: 'Rejected' } }, { status: 403 }) : undefined);
  const result = await run('send-mail', { ...newMail, attachments: [smallAttachment] });
  assert.match(String(result.error), /draft draft\/1 was retained and the email was not sent/);
  assert.equal(requests.length, 2);
  assert.equal(result.success, undefined);
});

for (const fault of ['network', 'server', 'progress', 'early-completion', 'no-completion'] as const) {
  test(`large upload ${fault} failure cannot send or leak the upload credential`, async () => {
    const { run, requests } = harness(req => {
      if (req.url !== uploadUrl) return;
      if (fault === 'network') throw new Error(`connection failed at ${uploadUrl}`);
      if (fault === 'server') return new Response(null, { status: 503 });
      if (fault === 'progress') return Response.json({ nextExpectedRanges: ['0-'] });
      if (fault === 'early-completion') return new Response(null, { status: 201 });
      const end = Number(/-(\d+)\//.exec(req.headers.get('content-range')!)![1]) + 1;
      return Response.json({ nextExpectedRanges: [`${end}-`] });
    });
    const path = join(dir, 'large.csv');
    await writeFile(path, Buffer.alloc(5_600_000, 65));
    const result = await run('send-mail', { ...newMail, filePaths: [path] });
    assert.match(String(result.error), /was retained and the email was not sent/);
    assert.equal(String(result.error).includes('test-upload-credential'), false);
    assert.equal(requests.some(r => r.path.endsWith('/send')), false);
    assert.equal(requests.filter(r => r.url === uploadUrl).length, fault === 'no-completion' ? 2 : 1);
  });
}

test('an uncertain final send is not repeated and tells the caller to check Sent Items', async () => {
  const { run, requests } = harness(req => req.path.endsWith('/send') ? new Response(null, { status: 503 }) : undefined);
  const result = await run('send-mail', { ...newMail, attachments: [smallAttachment] });
  assert.match(String(result.error), /Check Sent Items and the draft before retrying/);
  assert.equal(result.success, undefined);
  assert.equal(requests.filter(r => r.path.endsWith('/send')).length, 1);
});

test('draft creation failure does not upload or send', async () => {
  const { run, requests } = harness(req => req.path === '/v1.0/me/messages' ? Response.json({ error: { message: 'Missing Mail.ReadWrite' } }, { status: 403 }) : undefined);
  assert.ok((await run('send-mail', { ...newMail, attachments: [smallAttachment] })).error);
  assert.equal(requests.length, 1);
});

test('draft preparation requires Mail.ReadWrite while final sending requires Mail.Send', () => {
  for (const path of ['/me/messages', '/me/messages/id/createReply', '/me/messages/id/attachments/createUploadSession']) {
    assert.deepEqual(graphPolicy('POST', path).scopes, [['Mail.ReadWrite']]);
  }
  const send = graphPolicy('POST', '/me/messages/id/send');
  assert.deepEqual(send.scopes, [['Mail.Send']]);
  assert.equal(send.readOnly, false);
  assert.equal(send.operation, 'send');
});
