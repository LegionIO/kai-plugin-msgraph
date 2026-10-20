import assert from 'node:assert/strict';
import { afterEach, beforeEach, test } from 'node:test';
import { mkdtemp, readFile, rm, stat, symlink, writeFile } from 'node:fs/promises';
import { homedir, tmpdir } from 'node:os';
import { join, relative } from 'node:path';
import { buildMsgraphTools } from '../src/backend/tools.ts';
import { GraphClient } from '../src/backend/graph-client.ts';
import type { PluginAPI } from '../src/shared/types.ts';
import * as hostedContentCache from '../src/backend/hosted-content-cache.ts';
import { setLogger } from '../src/backend/logger-singleton.ts';

// Includes binary bytes and base64 padding to catch encoding and size regressions.
const payload = Buffer.from([0, 255, 128, 10, 13, 65, 66]);
const mediaType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const base64 = payload.toString('base64');
let dir: string;

beforeEach(async () => {
  dir = await mkdtemp(join(tmpdir(), 'msgraph-attachment-'));
  hostedContentCache.clear();
  setLogger({ info() {}, warn() {}, error() {} });
});
afterEach(async () => {
  hostedContentCache.clear();
  await rm(dir, { recursive: true, force: true });
});

function harness(options: { mime?: string; content?: string; teamsUrl?: string } = {}) {
  const mime = options.mime ?? mediaType;
  const calls: string[] = [];
  const download = async (kind: string) => {
    calls.push(kind);
    return { base64: options.content ?? base64, mediaType: mime, name: 'report.xlsx', contentId: null };
  };
  const client = {
    getMail: async () => ({ attachments: [{ id: 'attachment-1', name: 'report.xlsx', contentType: mime }] }),
    getMailAttachmentRaw: async () => download('mail'),
    getMessage: async () => ({
      id: 'message-1',
      body: { contentType: 'text', content: '' },
      attachments: [{ id: 'attachment-1', name: 'report.xlsx', contentType: 'reference', contentUrl: options.teamsUrl }],
    }),
    downloadReferenceAttachment: async () => download('reference'),
    getHostedContentRaw: async () => download('graph'),
  } as unknown as GraphClient;
  const api = { state: { set() {} } } as unknown as PluginAPI;
  const tool = buildMsgraphTools({ api, ensureAuthenticated: async () => {
    calls.push('auth');
    return client;
  } }).find((t) => t.name === 'get-attachment')!;
  const execute = async (args: Record<string, unknown> = {}) => await tool.execute({
    source: 'mail', messageId: 'message-1', ...args,
  }) as Record<string, unknown>;
  return { execute, calls };
}

test('mail defaults to direct model file content and reports the exact byte count', async () => {
  const { execute } = harness();
  const result = await execute();
  assert.equal(result.success, true);
  assert.equal(result.bytes, payload.length);
  assert.deepEqual(result._modelContent, [{ type: 'file', data: base64, mediaType, filename: 'report.xlsx' }]);
  assert.equal(Object.hasOwn(result, 'path'), false);
});

test('mail images still arrive as native image content when savePath is omitted', async () => {
  const { execute } = harness({ mime: 'image/png' });
  const result = await execute();
  assert.deepEqual(result._modelContent, [{ type: 'image', data: base64, mediaType: 'image/png' }]);
});

test('saving mail creates parent directories and returns only metadata plus the absolute path', async () => {
  const { execute } = harness();
  const path = join(dir, 'attachments', 'report.xlsx');
  const result = await execute({ attachmentId: 'attachment-1', savePath: path });
  assert.deepEqual(result, { success: true, source: 'mail', messageId: 'message-1', name: 'report.xlsx', mediaType, bytes: payload.length, path });
  assert.deepEqual(await readFile(path), payload);
  assert.equal((await stat(path)).mode & 0o777, 0o600);
  assert.equal((await stat(join(dir, 'attachments'))).mode & 0o777, 0o700);
});

test('saving an image also omits model content and expands a home-relative path', async () => {
  const { execute } = harness({ mime: 'image/png' });
  const path = join(dir, 'image.png');
  const result = await execute({ savePath: `~/${relative(homedir(), path)}` });
  assert.equal(result.path, path);
  assert.equal(Object.hasOwn(result, '_modelContent'), false);
  assert.deepEqual(await readFile(path), payload);
});

test('a direct request can be retried in save mode using the cached bytes', async () => {
  const { execute, calls } = harness();
  await execute();
  const path = join(dir, 'retry.xlsx');
  const saved = await execute({ savePath: path });
  assert.equal(saved.path, path);
  assert.equal(Object.hasOwn(saved, '_modelContent'), false);
  assert.equal(calls.filter((c) => c === 'mail').length, 1);
  assert.deepEqual(await readFile(path), payload);
});

for (const [kind, url] of [
  ['reference', 'https://example.sharepoint.com/report.xlsx'],
  ['graph', 'https://graph.microsoft.com/v1.0/chats/chat-1/messages/message-1/hostedContents/file/$value'],
] as const) {
  test(`Teams ${kind} attachments support both direct content and saving`, async () => {
    const { execute, calls } = harness({ teamsUrl: url });
    const args = { source: 'teams', chatId: 'chat-1', name: 'report.xlsx' };
    const direct = await execute(args);
    assert.equal(direct.success, true);
    assert.deepEqual(direct._modelContent, [{ type: 'file', data: base64, mediaType, filename: 'report.xlsx' }]);
    const path = join(dir, 'teams.xlsx');
    const saved = await execute({ ...args, savePath: path });
    assert.deepEqual(saved, { success: true, source: 'teams', chatId: 'chat-1', messageId: 'message-1', name: 'report.xlsx', mediaType, bytes: payload.length, path });
    assert.deepEqual(await readFile(path), payload);
    assert.equal(calls.filter((c) => c === kind).length, 1);
  });
}

test('save mode refuses existing files and symlinks without replacing their contents', async () => {
  const { execute } = harness();
  const existing = join(dir, 'existing.xlsx');
  const link = join(dir, 'link.xlsx');
  await writeFile(existing, 'keep this file');
  await symlink(existing, link);
  for (const savePath of [existing, link]) {
    const result = await execute({ savePath });
    assert.match(String(result.error), /File already exists/);
    assert.equal(Object.hasOwn(result, '_modelContent'), false);
    assert.equal(result.success, undefined);
  }
  assert.equal(await readFile(existing, 'utf8'), 'keep this file');
});

test('concurrent saves to the same path allow only one writer', async () => {
  const { execute } = harness();
  const path = join(dir, 'same.xlsx');
  const results = await Promise.all([execute({ savePath: path }), execute({ savePath: path })]);
  assert.equal(results.filter((r) => r.success === true).length, 1);
  assert.equal(results.filter((r) => /File already exists/.test(String(r.error))).length, 1);
  assert.deepEqual(await readFile(path), payload);
});

test('invalid or ambiguous save paths fail before authentication or downloading', async () => {
  const { execute, calls } = harness();
  for (const savePath of ['', '  ', 'report.xlsx', '../report.xlsx', 1, '/tmp/bad\0path']) {
    const result = await execute({ savePath });
    assert.match(String(result.error), /savePath must be/);
  }
  assert.deepEqual(calls, []);
});

test('undownloadable attachments do not create a destination', async () => {
  const { execute } = harness({ content: '' });
  const path = join(dir, 'missing', 'report.xlsx');
  const result = await execute({ savePath: path });
  assert.match(String(result.error), /no downloadable content/);
  await assert.rejects(stat(join(dir, 'missing')), { code: 'ENOENT' });
});

test('filesystem errors do not report success or fall back to piping contents to the model', async () => {
  const { execute } = harness();
  const parent = join(dir, 'not-a-directory');
  await writeFile(parent, 'preserve');
  const result = await execute({ savePath: join(parent, 'report.xlsx') });
  assert.match(String(result.error), /ENOTDIR|EEXIST/);
  assert.equal(result.success, undefined);
  assert.equal(Object.hasOwn(result, '_modelContent'), false);
  assert.equal(await readFile(parent, 'utf8'), 'preserve');
});
