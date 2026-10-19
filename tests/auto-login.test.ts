import assert from 'node:assert/strict';
import { test } from 'node:test';
import { runInNewContext } from 'node:vm';
import { performAutoLogin } from '../src/backend/auto-login.ts';

type Element = { textContent?: string; offsetParent?: object | null; click?: () => void; dispatchEvent?: () => void };
function fixture(initial: string) {
  let page = initial;
  let number = '';
  let description = 'Check your Microsoft Authenticator app';
  let hiddenApproval = false;
  let submitCount = 0;
  const elements = () => {
    const visible = (textContent = '', click = () => {}) => ({ offsetParent: {}, textContent, click, dispatchEvent() {} });
    const found: Record<string, Element> = {};
    if (page === 'username') {
      found['input[name="loginfmt"]'] = visible();
      found['input[type="submit"]#idSIButton9'] = visible('', () => { submitCount++; page = 'push'; });
    }
    if (page === 'push') {
      found['#idDiv_SAOTCAS_Description'] = visible(description);
      if (number) found['#idRichContext_DisplaySign'] = visible(number);
    }
    if (page === 'totp' || page === 'sms') {
      found['input#idTxtBx_SAOTCC_OTC'] = visible();
      found['input#idSubmit_SAOTCC_Continue'] = visible('', () => { submitCount++; page = 'complete'; });
      if (hiddenApproval) found['#idDiv_SAOTCAS_Description'] = { offsetParent: null, textContent: 'Approve sign-in' };
    }
    if (page === 'kmsi') {
      found['#KmsiBanner'] = visible();
      found['input#idSIButton9'] = visible('', () => { submitCount++; page = 'complete'; });
    }
    return found;
  };
  const scripts: string[] = [];
  const approvals: Array<string | undefined> = [];
  const codes: string[] = [];
  const fallbacks: string[] = [];
  let shown = 0;
  let onApproval = (_number?: string) => { page = 'complete'; };
  const helpers = {
    getURL: () => 'https://login.microsoftonline.com/organizations/login',
    show: () => { shown++; },
    executeJavaScript: async (script: string) => {
      scripts.push(script);
      const els = elements();
      class HTMLInputElement {}
      Object.defineProperty(HTMLInputElement.prototype, 'value', { set() {} });
      return runInNewContext(script, {
        window: { HTMLInputElement, location: { protocol: 'https:', port: '', hostname: page === 'complete' ? 'example.test' : 'login.microsoftonline.com' } },
        document: { querySelector: (selector: string) => els[selector] ?? null, body: { innerText: page === 'totp' ? 'Enter a code from your authenticator app' : 'Enter code' } },
        Event: class {},
      });
    },
  };
  const callbacks = {
    onMfaCodeNeeded: async (type: 'sms' | 'totp') => { codes.push(type); return '123456'; },
    onMfaApprovalNeeded: (n?: string) => { approvals.push(n); onApproval(n); },
    onMfaApprovalComplete() {},
    onFallback: (reason: string) => { fallbacks.push(reason); },
  };
  return { helpers, callbacks, approvals, codes, scripts, fallbacks,
    page: (p: string) => { page = p; },
    number: (n: string) => { number = n; },
    description: (d: string) => { description = d; },
    hiddenApproval: () => { hiddenApproval = true; },
    onApproval: (f: typeof onApproval) => { onApproval = f; },
    submits: () => submitCount, shown: () => shown,
  };
}
const credentials = { username: 'user@example.test', password: 'test-password' };
async function flush() { for (let i = 0; i < 20; i++) await Promise.resolve(); }

test('cached SSO can open directly on a push prompt without the word approve', async () => {
  const f = fixture('push'); f.number('42');
  assert.equal((await performAutoLogin(f.helpers, credentials, f.callbacks)).success, true);
  assert.deepEqual(f.approvals, ['42']); assert.equal(f.submits(), 0); assert.equal(f.shown(), 0);
});

test('a visible number-matching challenge is detected even without description text', async () => {
  const f = fixture('push'); f.number('27'); f.description('');
  await performAutoLogin(f.helpers, credentials, f.callbacks); assert.deepEqual(f.approvals, ['27']);
});

test('late Authenticator number updates the displayed challenge', async (t) => {
  t.mock.timers.enable({ apis: ['setTimeout', 'Date'] });
  const f = fixture('push'); f.onApproval((n) => { if (n === '64') f.page('complete'); });
  const result = performAutoLogin(f.helpers, credentials, f.callbacks); await flush();
  assert.deepEqual(f.approvals, [undefined]);
  f.number('64'); t.mock.timers.tick(500); await flush();
  t.mock.timers.tick(500); await flush();
  assert.equal((await result).success, true); assert.deepEqual(f.approvals, [undefined, '64']);
});

for (const type of ['totp', 'sms'] as const) test(`direct ${type} entry is requested and submitted`, async () => {
  const f = fixture(type); f.hiddenApproval();
  assert.equal((await performAutoLogin(f.helpers, credentials, f.callbacks)).success, true);
  assert.deepEqual(f.codes, [type]); assert.equal(f.submits(), 1); assert.deepEqual(f.approvals, []);
});

test('passwordless username submission can transition straight to MFA', async () => {
  const f = fixture('username'); f.number('42');
  assert.equal((await performAutoLogin(f.helpers, credentials, f.callbacks)).success, true);
  assert.equal(f.submits(), 1); assert.deepEqual(f.approvals, ['42']);
  assert.ok(!f.scripts.some((script) => script.includes('test-password')));
});

test('a push challenge can switch to code entry in the same login', async () => {
  const f = fixture('push'); f.number('42'); f.onApproval(() => f.page('totp'));
  assert.equal((await performAutoLogin(f.helpers, credentials, f.callbacks)).success, true);
  assert.deepEqual(f.approvals, ['42']); assert.deepEqual(f.codes, ['totp']);
});

test('an existing stay-signed-in page completes without filling a password', async () => {
  const f = fixture('kmsi');
  assert.equal((await performAutoLogin(f.helpers, credentials, f.callbacks)).success, true);
  assert.equal(f.submits(), 1); assert.deepEqual(f.codes, []); assert.deepEqual(f.approvals, []);
});

test('cancelling push polling stops automation without revealing a stale window', async () => {
  const f = fixture('push'); f.onApproval(() => {}); const controller = new AbortController();
  const result = performAutoLogin(f.helpers, credentials, f.callbacks, controller.signal); await flush();
  controller.abort();
  assert.equal((await result).fallbackToManual, false); assert.deepEqual(f.fallbacks, []); assert.equal(f.shown(), 0);
});

test('push timeout reveals the same window for manual completion', async (t) => {
  t.mock.timers.enable({ apis: ['setTimeout', 'Date'] });
  const f = fixture('push'); f.onApproval(() => {});
  const result = performAutoLogin(f.helpers, credentials, f.callbacks); await flush();
  t.mock.timers.tick(120_000); await flush();
  assert.equal((await result).fallbackToManual, true); assert.equal(f.shown(), 1);
});
