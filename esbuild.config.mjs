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
    build.onLoad({ filter: /.*/, namespace: 'react-global' }, (args) => ({
      contents: args.path.includes('jsx-runtime') || args.path.includes('jsx-dev-runtime')
        ? `
        const R = () => globalThis.React;
        export const Fragment = Symbol.for('react.fragment');
        export function jsx(type, props, key) {
          return R().createElement(type, key !== undefined ? { ...props, key } : props);
        }
        export const jsxs = jsx;
        export const jsxDEV = jsx;
        `
        : args.path.startsWith('react-dom')
        ? `
        const RD = () => globalThis.ReactDOM ?? {};
        // Fallback portal: render in place if host doesn't expose ReactDOM.
        export const createPortal = (children, _container) =>
          (RD().createPortal ? RD().createPortal(children, _container) : children);
        export const flushSync = (fn) => (RD().flushSync ? RD().flushSync(fn) : fn());
        export default new Proxy({}, { get: (_, k) => RD()[k] });
        `
        : `
        const R = () => globalThis.React;
        // Self-contained base classes so 'class X extends React.Component' can evaluate
        // before Kai calls register({React}). React's reconciler only checks
        // prototype.isReactComponent / isPureReactComponent — the runtime updater is
        // injected onto the instance by React itself.
        const noopUpdater = { isMounted: () => false, enqueueSetState() {}, enqueueReplaceState() {}, enqueueForceUpdate() {} };
        export function Component(props, context, updater) {
          this.props = props; this.context = context; this.refs = {};
          this.updater = updater || noopUpdater;
        }
        Component.prototype.isReactComponent = {};
        Component.prototype.setState = function (s, cb) { this.updater.enqueueSetState(this, s, cb, 'setState'); };
        Component.prototype.forceUpdate = function (cb) { this.updater.enqueueForceUpdate(this, cb, 'forceUpdate'); };
        export function PureComponent(props, context, updater) { Component.call(this, props, context, updater); }
        PureComponent.prototype = Object.create(Component.prototype);
        PureComponent.prototype.constructor = PureComponent;
        PureComponent.prototype.isPureReactComponent = true;
        export default new Proxy({}, { get: (_, k) => (k === 'Component' ? Component : k === 'PureComponent' ? PureComponent : R()?.[k]) });
        export const useState = (...a) => R().useState(...a);
        export const useEffect = (...a) => R().useEffect(...a);
        export const useLayoutEffect = (...a) => R().useLayoutEffect(...a);
        export const useRef = (...a) => R().useRef(...a);
        export const useCallback = (...a) => R().useCallback(...a);
        export const useMemo = (...a) => R().useMemo(...a);
        export const useContext = (...a) => R().useContext(...a);
        export const useReducer = (...a) => R().useReducer(...a);
        export const useId = (...a) => R().useId(...a);
        export const useSyncExternalStore = (...a) => R().useSyncExternalStore(...a);
        export const useDeferredValue = (...a) => R().useDeferredValue(...a);
        export const useTransition = (...a) => R().useTransition(...a);
        export const useImperativeHandle = (...a) => R().useImperativeHandle(...a);
        export const useInsertionEffect = (...a) => (R().useInsertionEffect ?? R().useLayoutEffect)(...a);
        export const useDebugValue = () => {};
        export const startTransition = (fn) => (R().startTransition ? R().startTransition(fn) : fn());
        export const createElement = (...a) => R().createElement(...a);
        export const cloneElement = (...a) => R().cloneElement(...a);
        export const createContext = (...a) => R().createContext(...a);
        export const forwardRef = (...a) => R().forwardRef(...a);
        export const memo = (...a) => R().memo(...a);
        export const isValidElement = (...a) => R().isValidElement(...a);
        export const Children = new Proxy({}, { get: (_, k) => R().Children[k] });
        export const Fragment = Symbol.for('react.fragment');
        export const Suspense = Symbol.for('react.suspense');
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
  // ws optionally requires these native addons; let its try/catch see the real ENOENT.
  external: ['bufferutil', 'utf-8-validate'],
  sourcemap: true,
  target: 'node18',
  // Bundled CJS deps (ws) call require('events') etc. esbuild's ESM __require shim
  // throws unless a real require is in scope — provide one via createRequire.
  banner: {
    js: "import { createRequire as __msgraph_createRequire } from 'module'; var require = __msgraph_createRequire(import.meta.url);",
  },
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
