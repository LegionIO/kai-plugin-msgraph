import { test } from 'node:test';
import assert from 'node:assert/strict';
import { createMfaActivityGuard, normalizeMfaActivityPolicy } from '../src/backend/mfa-activity';

function fixture(idle: () => Promise<number | null> = async () => 3_600, data: Record<string, any> = {}) {
  const state: Record<string, any> = {};
  const actions = new Map<string, any>();
  let tools: any[] = [];
  let finish!: (result: any) => void;
  let helpers: any;
  let opens = 0, shows = 0, closes = 0, cookieMfa = false;
  const original: any = {
    config: { getPluginData: () => data, setPluginData: (key: string, value: any) => { data[key] = value; } },
    state: { set: (key: string, value: any) => { state[key] = value; } },
    onAction: (id: string, action: any) => actions.set(id, action),
    tools: { register: (registered: any[]) => { tools = registered; } },
    auth: { openAuthWindow: async (options: any) => {
      opens++;
      assert.equal(options.showAfterMs, undefined);
      const result = new Promise(resolve => { finish = resolve; });
      options.onReady({ getURL: async () => 'https://login.microsoftonline.com/common', executeJavaScript: async () => cookieMfa,
        show() { shows++; }, close() { closes++; finish({ success: false }); } });
      return result;
    } },
  };
  const guard = createMfaActivityGuard(original, idle);
  guard.api.onAction('panel', () => {});
  return { guard, data, state, actions, get tools() { return tools; }, get opens() { return opens; },
    get shows() { return shows; }, get closes() { return closes; }, get helpers() { return helpers; },
    finish: (result: any) => finish(result),
    cookieMfa: () => { cookieMfa = true; },
    open: () => guard.api.auth.openAuthWindow({ url: 'https://login.microsoftonline.com', partition: 'test', showOnCreate: false,
      showAfterMs: 30_000, onReady: (value: any) => { helpers = value; } }),
  };
}
async function drain() { for (let i = 0; i < 30; i++) await Promise.resolve(); }

test('default computer activity, configurable threshold, and unavailable idle detection', async () => {
  let idle: number | null = 299;
  const f = fixture(async () => idle);
  assert.equal(await f.guard.isActive(), true);
  f.actions.get('panel')('save-mfa-activity-policy', { idleMinutes: 1 });
  assert.equal(await f.guard.isActive(), false);
  idle = null;
  f.actions.get('panel')('save-mfa-activity-policy', { idleMinutes: 5 });
  assert.equal(await f.guard.isActive(), false);
  f.actions.get('panel')('mfa-user-activity', { source: 'test-ui', active: true });
  assert.equal(await f.guard.isActive(), true);
  f.actions.get('panel')('mfa-user-activity', { source: 'test-ui', active: false });
  assert.equal(await f.guard.isActive(), false);
  assert.deepEqual(normalizeMfaActivityPolicy({ idleMinutes: NaN }), { source: 'computer', idleMinutes: 5, allowAgent: true });
  f.guard.dispose();
});

test('an open panel and polling do not count; plugin input expires', async t => {
  t.mock.timers.enable({ apis: ['Date'], now: 100_000 });
  const f = fixture(async () => 0, { mfaActivityPolicy: { source: 'plugin', idleMinutes: 1 } });
  try {
    assert.equal(await f.guard.isActive(), false);
    f.actions.get('panel')('refresh');
    assert.equal(await f.guard.isActive(), false);
    f.actions.get('panel')('mfa-user-activity', { source: 'ui', active: true });
    assert.equal(await f.guard.isActive(), true);
    t.mock.timers.tick(60_000);
    assert.equal(await f.guard.isActive(), false);
  } finally { f.guard.dispose(); t.mock.timers.reset(); }
});

test('only a live interactive tool counts, never scheduled or completed tools', async () => {
  const f = fixture();
  let release!: () => void;
  f.guard.api.tools.register([{ name: 'test', execute: () => new Promise<void>(resolve => { release = resolve; }) }]);
  let task = f.tools[0].execute({}, { toolCallId: 'scheduled', isHeadless: true });
  assert.equal(await f.guard.isActive(), false);
  release(); await task;
  const abort = new AbortController();
  task = f.tools[0].execute({}, { toolCallId: 'interactive', abortSignal: abort.signal });
  assert.equal(await f.guard.isActive(), true);
  f.actions.get('panel')('save-mfa-activity-policy', { allowAgent: false });
  assert.equal(await f.guard.isActive(), false);
  f.actions.get('panel')('save-mfa-activity-policy', { allowAgent: true });
  abort.abort();
  assert.equal(await f.guard.isActive(), false);
  release(); await task;
  assert.equal(await f.guard.isActive(), false);
  f.guard.dispose();
});

test('MFA waits without revealing a window and resumes when computer activity returns', async t => {
  t.mock.timers.enable({ apis: ['setTimeout', 'Date'], now: 100_000 });
  let idle = 3_600;
  const f = fixture(async () => idle);
  try {
    const login = f.open(); await drain();
    let resumed = false;
    const challenge = f.helpers.waitForMfaActivity().then((value: boolean) => { resumed = value; });
    await drain();
    assert.equal(f.state.mfaActivityPaused, true);
    assert.equal(resumed, false);
    t.mock.timers.tick(30_000); await drain();
    assert.equal(f.shows, 0);
    idle = 0;
    t.mock.timers.tick(1_000); await drain(); await challenge;
    assert.equal(resumed, true);
    f.finish({ success: true }); await login;
    assert.deepEqual(f.data.mfaPendingPartitions, []);
    assert.equal(f.state.mfaActivityPaused, false);
  } finally { f.guard.dispose(); t.mock.timers.reset(); }
});

test('expired windows cancel the wait and background retries cannot replay credentials', async () => {
  const f = fixture();
  try {
    const login = f.open(); await drain();
    const challenge = f.helpers.waitForMfaActivity(); await drain();
    f.finish({ success: false }); await login;
    assert.equal(await challenge, false);
    assert.equal((await f.open()).success, false);
    assert.equal(f.opens, 1);
    // Persisted latch still applies after plugin restart.
    const restarted = fixture(async () => 3_600, f.data);
    assert.equal((await restarted.open()).success, false);
    assert.equal(restarted.opens, 0);
    restarted.guard.dispose();
    f.actions.get('panel')('mfa-user-activity', { source: 'ui', active: true });
    const retry = f.open(); await drain();
    assert.equal(f.opens, 2);
    f.finish({ success: true }); await retry;
  } finally { f.guard.dispose(); }
});

test('deactivation closes a paused window and cancels its wait', async () => {
  const f = fixture();
  const login = f.open(); await drain();
  const challenge = f.helpers.waitForMfaActivity(); await drain();
  f.guard.dispose(); await drain();
  assert.equal(await challenge, false);
  assert.equal(f.closes, 1);
  await login;
  assert.equal((await f.open()).success, false);
});

test('cookie-only refreshes that encounter MFA also suppress subsequent unattended attempts', async () => {
  const f = fixture();
  try {
    f.cookieMfa();
    const login = f.open(); await drain();
    assert.equal(f.state.mfaActivityPaused, true);
    assert.deepEqual(f.data.mfaPendingPartitions, ['test']);
    f.finish({ success: false }); await login;
    assert.equal((await f.open()).success, false);
    assert.equal(f.opens, 1);
    assert.equal(f.shows, 0);
  } finally { f.guard.dispose(); }
});

test('concurrent MFA checks share a slow operating system probe', async t => {
  t.mock.timers.enable({ apis: ['Date'], now: 100_000 });
  let reads = 0;
  let finish!: (seconds: number) => void;
  const f = fixture(() => { reads++; return new Promise(resolve => { finish = resolve; }); });
  try {
    const first = f.guard.isActive(); await drain();
    t.mock.timers.tick(3_000);
    const second = f.guard.isActive(); await drain();
    assert.equal(reads, 1);
    finish(0);
    assert.equal(await first, true); assert.equal(await second, true);
  } finally { f.guard.dispose(); t.mock.timers.reset(); }
});
