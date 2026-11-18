import assert from 'node:assert/strict';
import { test } from 'node:test';
import React from 'react';
import { renderToStaticMarkup } from 'react-dom/server';
import { TokenExpiry } from '../src/frontend/components/TokenExpiry.tsx';

test('token expiry shows the deadline, remaining time and refresh explanation', t => {
  t.mock.timers.enable({ apis: ['Date'], now: 1_800_000_000_000 });
  const expiresAt = Date.now() + 65 * 60_000;
  const render = () => renderToStaticMarkup(<TokenExpiry expiresAt={expiresAt} />);
  const html = render();
  assert.match(html, /1h 5m left/);
  assert.ok(html.includes(new Date(expiresAt).toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' })));
  assert.match(html, /Silent refresh can extend this time/);
  t.mock.timers.tick(64 * 60_000 + 30_000);
  assert.match(render(), /1m left/);
  t.mock.timers.tick(30_000);
  assert.match(render(), /Token expired/);
  assert.doesNotMatch(render(), /left/);
});

test('missing expiry is not guessed', () => {
  assert.equal(renderToStaticMarkup(<TokenExpiry expiresAt={null} />), '');
});
