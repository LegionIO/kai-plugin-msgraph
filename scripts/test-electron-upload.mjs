import { build } from 'esbuild';
import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { spawnSync } from 'node:child_process';

const electron = process.argv[2] ?? process.env.MSGRAPH_ELECTRON_PATH;
if (!electron) throw new Error('Pass the Electron executable path as an argument or set MSGRAPH_ELECTRON_PATH.');
const dir = await mkdtemp(join(tmpdir(), 'msgraph-electron-upload-'));
try {
  const outfile = join(dir, 'test.cjs');
  await build({
    entryPoints: [resolve('tests/electron-mail-upload.ts')], outfile,
    bundle: true, platform: 'node', format: 'cjs', target: 'node22',
    external: ['electron', 'bufferutil', 'utf-8-validate'],
  });
  const env = { ...process.env };
  delete env.ELECTRON_RUN_AS_NODE;
  const result = spawnSync(electron, [outfile], { env, stdio: 'inherit', timeout: 30_000 });
  if (result.error) throw result.error;
  process.exitCode = result.status ?? 1;
} finally {
  await rm(dir, { recursive: true, force: true });
}
