import { build } from 'esbuild';
import { mkdtemp, readdir, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { spawnSync } from 'node:child_process';

const outdir = await mkdtemp(join(tmpdir(), 'msgraph-tests-'));
try {
  const files = (await readdir('tests')).filter((f) => /\.test\.tsx?$/.test(f));
  await build({
    entryPoints: files.map((f) => resolve('tests', f)),
    outdir, bundle: true, platform: 'node', format: 'esm', target: 'node22',
    external: ['bufferutil', 'utf-8-validate'],
    banner: { js: "import { createRequire } from 'node:module'; const require = createRequire(import.meta.url);" },
  });
  const result = spawnSync(process.execPath, ['--test', ...files.map((f) => join(outdir, f.replace(/\.tsx?$/, '.js')))], { stdio: 'inherit' });
  process.exitCode = result.status ?? 1;
} finally {
  await rm(outdir, { recursive: true, force: true });
}
