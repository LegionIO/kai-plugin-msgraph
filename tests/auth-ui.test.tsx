import assert from 'node:assert/strict';
import { test } from 'node:test';
import React from 'react';
import { renderToStaticMarkup } from 'react-dom/server';
import { SettingsView } from '../src/frontend/components/SettingsView.tsx';
import { PanelView } from '../src/frontend/components/PanelView.tsx';
import { MailPanelView } from '../src/frontend/components/MailPanelView.tsx';
import type { MsgraphPluginState } from '../src/shared/types.ts';

function render(Component: any, overrides = {}) {
  const pluginState = {
    auth: { isAuthenticated: false, isSigningIn: true, autoLoginStatus: 'Waiting for Microsoft Authenticator', error: null },
    mfa: { needed: false, type: null, approvalNumber: null },
    ...overrides,
  } as MsgraphPluginState;
  return renderToStaticMarkup(<Component pluginName="msgraph" pluginState={pluginState} onAction={() => {}} />);
}
test('Settings exposes bounded throttle wait and adaptive routing defaults', () => {
  const html=render(SettingsView);
  assert.match(html,/Maximum throttle wait \(seconds\)/);
  assert.match(html,/min="0" max="300" step="1" value="30"/);
  assert.match(html,/Adaptive FOCI routing/);
});
for (const [name, Component] of [['Settings', SettingsView], ['Teams', PanelView], ['Outlook', MailPanelView]] as const) {
  test(`${name} shows sign-in feedback and disables repeated clicks`, () => {
    const html = render(Component);
    assert.match(html, /disabled="" aria-busy="true"/);
    assert.match(html, /Signing in…/);
    assert.match(html, /role="status"/);
    assert.match(html, /Waiting for Microsoft Authenticator/);
    assert.match(html, /Show sign-in window/);
  });
  test(`${name} displays the Authenticator number during login`, () => {
    const html = render(Component, { mfa: { needed: true, type: 'push', approvalNumber: '42' } });
    assert.match(html, /role="dialog"/); assert.match(html, />42<\/span>/);
    assert.match(html, /Enter this number in the Authenticator app/);
  });
  for (const type of ['sms', 'totp']) test(`${name} displays the ${type} code entry component`, () => {
    const html = render(Component, { mfa: { needed: true, type, approvalNumber: null } });
    assert.match(html, /role="dialog"/); assert.match(html, /type="password"/);
    assert.match(html, /autoComplete="one-time-code"/); assert.match(html, /Verify/);
  });
  test(`${name} shows an actionable failure with an enabled sign-in button`, () => {
    const html = render(Component, { auth: { isAuthenticated: false, isSigningIn: false, autoLoginStatus: null, error: 'Sign-in timed out. Please try again.' } });
    assert.match(html, /role="alert"/); assert.match(html, /Sign-in timed out/);
    assert.match(html, /aria-busy="false"/); assert.doesNotMatch(html, /Signing in…/);
  });
}
