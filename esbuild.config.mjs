import esbuild from 'esbuild';
import { fileURLToPath } from 'url';
import { dirname, resolve } from 'path';
import { createRequire } from 'module';
import { copyFileSync, mkdirSync, readFileSync } from 'fs';
import { homedir } from 'os';

const __dirname = dirname(fileURLToPath(import.meta.url));
const require = createRequire(import.meta.url);
const isWatch = process.argv.includes('--watch');
const isDev = process.argv.includes('--dev');

const manifest = JSON.parse(readFileSync(resolve(__dirname, 'plugin.json'), 'utf-8'));
const pluginName = manifest.name;

const outputDir = isDev
  ? resolve(homedir(), '.kai', 'plugins', pluginName)
  : resolve(__dirname, 'dist');

const builtins = new Set([
  'fs', 'path', 'child_process', 'crypto', 'events', 'stream', 'util',
  'http', 'https', 'net', 'os', 'url', 'zlib', 'buffer', 'process',
  'assert', 'constants', 'dns', 'domain', 'dgram', 'querystring',
  'readline', 'repl', 'string_decoder', 'sys', 'timers', 'tls', 'tty', 'vm',
]);

const reactGlobalPlugin = {
  name: 'react-global',
  setup(build) {
    build.onResolve({ filter: /^react(-dom)?(\/.*)?$/ }, args => ({
      path: args.path,
      namespace: 'react-global',
    }));
    build.onLoad({ filter: /.*/, namespace: 'react-global' }, () => ({
      contents: `
        const R = () => globalThis.React;
        export default new Proxy({}, { get: (_, k) => R()[k] });
        export const useState = (...a) => R().useState(...a);
        export const useEffect = (...a) => R().useEffect(...a);
        export const useLayoutEffect = (...a) => R().useLayoutEffect(...a);
        export const useRef = (...a) => R().useRef(...a);
        export const useCallback = (...a) => R().useCallback(...a);
        export const useMemo = (...a) => R().useMemo(...a);
        export const useContext = (...a) => R().useContext(...a);
        export const useReducer = (...a) => R().useReducer(...a);
        export const createElement = (...a) => R().createElement(...a);
        export const createContext = (...a) => R().createContext(...a);
        export const forwardRef = (...a) => R().forwardRef(...a);
        export const memo = (...a) => R().memo(...a);
        export const Fragment = Symbol.for('react.fragment');
      `,
      loader: 'js',
    }));
  },
};

const localNodeModulesPlugin = {
  name: 'local-node-modules',
  setup(build) {
    build.onResolve({ filter: /^[^./]/ }, args => {
      if (args.path.startsWith('node:')) return null;
      const packageName = args.path.startsWith('@')
        ? args.path.split('/').slice(0, 2).join('/')
        : args.path.split('/')[0];
      if (builtins.has(packageName)) return null;
      try {
        const resolved = require.resolve(args.path, {
          paths: [resolve(__dirname, 'node_modules', '..')]
        });
        return { path: resolved };
      } catch {
        return null;
      }
    });
  },
};

const backendOptions = {
  entryPoints: ['./src/backend/index.ts'],
  bundle: true,
  platform: 'node',
  format: 'esm',
  outfile: resolve(outputDir, 'backend.js'),
  external: [],
  sourcemap: true,
  target: 'node18',
  plugins: [localNodeModulesPlugin],
};

const frontendOptions = {
  entryPoints: ['./src/frontend/index.ts'],
  bundle: true,
  platform: 'browser',
  format: 'esm',
  outfile: resolve(outputDir, 'frontend.js'),
  sourcemap: true,
  target: 'es2020',
  jsx: 'transform',
  jsxFactory: 'React.createElement',
  jsxFragment: 'React.Fragment',
  plugins: [reactGlobalPlugin, localNodeModulesPlugin],
};

mkdirSync(outputDir, { recursive: true });
copyFileSync(resolve(__dirname, 'plugin.json'), resolve(outputDir, 'plugin.json'));

if (isWatch) {
  const backendCtx = await esbuild.context(backendOptions);
  const frontendCtx = await esbuild.context(frontendOptions);
  await Promise.all([backendCtx.watch(), frontendCtx.watch()]);
  console.log(`Watching for changes... (output: ${outputDir})`);
} else {
  await Promise.all([
    esbuild.build(backendOptions),
    esbuild.build(frontendOptions)
  ]).catch(() => process.exit(1));

  if (isDev) {
    console.log(`Built to ~/.kai/plugins/${pluginName}/`);
  } else {
    console.log('Built backend.js and frontend.js to dist/');
  }
}
