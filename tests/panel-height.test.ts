import assert from 'node:assert/strict';
import { test } from 'node:test';
import { measurePanelHeight } from '../src/frontend/hooks.ts';

test('panel height stays inside the host after scrolling, with all wrapper insets', t => {
  let scrollTop = 0;
  const host = {
    getBoundingClientRect: () => ({ top: 64 }),
    clientTop: 0, clientHeight: 986,
    get scrollTop() { return scrollTop; },
    style: { paddingBottom: '0px' },
  };
  const wrapper = { parentElement: host, style: { paddingBottom: '20px' } };
  const card = { parentElement: wrapper, style: { borderBottomWidth: '1px' } };
  const panel = {
    getBoundingClientRect: () => ({ top: 77 - scrollTop }),
    parentElement: card,
  };
  Object.defineProperty(globalThis, 'window', { configurable: true, value: {
    getComputedStyle: (el: { style: object }) => el.style,
    innerHeight: 1050,
  } });
  t.after(() => { Reflect.deleteProperty(globalThis, 'window'); });
  const measure = () => measurePanelHeight(panel as unknown as HTMLElement, host as unknown as HTMLElement);
  assert.equal(measure(), 952);
  scrollTop = 363;
  assert.equal(measure(), 952);
  host.clientHeight = 250;
  assert.equal(measure(), 216, 'short windows must not force a 320px panel into the host');
  host.clientHeight = 20;
  assert.equal(measure(), 0);
});

test('a panel without a scroll host uses the viewport', t => {
  Object.defineProperty(globalThis, 'window', { configurable: true, value: { innerHeight: 800 } });
  t.after(() => { Reflect.deleteProperty(globalThis, 'window'); });
  const panel = { getBoundingClientRect: () => ({ top: 100 }) };
  assert.equal(measurePanelHeight(panel as HTMLElement, null), 700);
});
