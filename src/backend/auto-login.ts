/**
 * Azure AD auto-login form automation.
 * Adapted from kai-plugin-pim/src/backend/auto-login.ts (same login pages).
 */

import {
  AUTO_LOGIN_POLL_INTERVAL_MS,
  AUTO_LOGIN_STEP_TIMEOUT_MS,
  AUTO_LOGIN_MFA_TIMEOUT_MS,
} from '../shared/constants.js';

const SELECTORS = {
  USERNAME_INPUT: 'input[name="loginfmt"]',
  NEXT_BUTTON: 'input[type="submit"]#idSIButton9',
  PASSWORD_INPUT: 'input[name="passwd"]',
  SIGN_IN_BUTTON: 'input[type="submit"]#idSIButton9',
  ACCOUNT_PICKER_TILES: '#tilesHolder',
  ACCOUNT_PICKER_HEADER: '#loginHeader',
  ACCOUNT_PICKER_OTHER_TILE: '#otherTile',
  MFA_PROOF_SECTION: '#idDiv_SAOTCS_Proofs',
  MFA_PROOF_TITLE: '#idDiv_SAOTCS_Title',
  MFA_CODE_INPUT: 'input#idTxtBx_SAOTCC_OTC',
  MFA_CODE_VERIFY_BUTTON: 'input#idSubmit_SAOTCC_Continue',
  MFA_APPROVAL_INDICATOR: '#idDiv_SAOTCAS_Description',
  STAY_SIGNED_IN_MARKER: '#KmsiBanner',
  STAY_SIGNED_IN_YES: 'input#idSIButton9',
  STAY_SIGNED_IN_NO: 'input#idBtn_Back',
  ERROR_BANNER: '#usernameError, #passwordError, #errorText, .alert-error, #service_exception_message',
} as const;

export type LoginPageState =
  | 'username_needed'
  | 'password_needed'
  | 'account_picker'
  | 'mfa_method_picker'
  | 'mfa_sms_code'
  | 'mfa_authenticator_push'
  | 'mfa_totp_code'
  | 'stay_signed_in'
  | 'error'
  | 'loading'
  | 'not_login_page'
  | 'unknown';

interface PageStateResult {
  state: LoginPageState;
  errorText?: string;
  approvalNumber?: string;
}

export interface AutoLoginHelpers {
  executeJavaScript: (code: string) => Promise<unknown>;
  getURL: () => string | Promise<string>;
  show: () => void;
}

export interface AutoLoginCallbacks {
  onMfaCodeNeeded: (type: 'sms' | 'totp') => Promise<string>;
  onMfaApprovalNeeded: (approvalNumber?: string) => void;
  onMfaApprovalComplete: () => void;
  onFallback: (reason: string) => void;
}

export interface AutoLoginResult {
  success: boolean;
  fallbackToManual: boolean;
  reason?: string;
}

const MICROSOFT_LOGIN_HOSTS = new Set(['login.microsoftonline.com', 'login.live.com']);

export function isMicrosoftLoginHost(url: string): boolean {
  try {
    const u = new URL(url);
    return (
      u.protocol === 'https:' &&
      (u.port === '' || u.port === '443') &&
      MICROSOFT_LOGIN_HOSTS.has(u.hostname)
    );
  } catch {
    return false;
  }
}

function sleep(ms: number, signal?: AbortSignal): Promise<void> {
  signal?.throwIfAborted();
  return new Promise((resolve, reject) => {
    const onAbort = () => { clearTimeout(timer); reject(signal?.reason); };
    const timer = setTimeout(() => { signal?.removeEventListener('abort', onAbort); resolve(); }, ms);
    signal?.addEventListener('abort', onAbort, { once: true });
  });
}

async function detectPageState(helpers: AutoLoginHelpers): Promise<PageStateResult> {
  try {
    return await helpers.executeJavaScript(`
      (function() {
        var loc = window.location;
        if (loc.protocol !== 'https:' ||
            (loc.port !== '' && loc.port !== '443') ||
            (loc.hostname !== 'login.microsoftonline.com' && loc.hostname !== 'login.live.com')) {
          return { state: 'not_login_page' };
        }
        var errorEl = document.querySelector('${SELECTORS.ERROR_BANNER}');
        if (errorEl && errorEl.offsetParent !== null && errorEl.textContent.trim()) {
          return { state: 'error', errorText: errorEl.textContent.trim() };
        }
        var kmsiBanner = document.querySelector('${SELECTORS.STAY_SIGNED_IN_MARKER}');
        if (kmsiBanner && kmsiBanner.offsetParent !== null) {
          return { state: 'stay_signed_in' };
        }
        var tilesHolder = document.querySelector('${SELECTORS.ACCOUNT_PICKER_TILES}');
        var loginHeader = document.querySelector('${SELECTORS.ACCOUNT_PICKER_HEADER}');
        if (tilesHolder && tilesHolder.offsetParent !== null && loginHeader) {
          var headerText = loginHeader.textContent || '';
          if (headerText.includes('Pick an account') || headerText.includes('pick an account')) {
            return { state: 'account_picker' };
          }
        }
        var mfaProofSection = document.querySelector('${SELECTORS.MFA_PROOF_SECTION}');
        var mfaTitle = document.querySelector('${SELECTORS.MFA_PROOF_TITLE}');
        if (mfaProofSection && mfaProofSection.offsetParent !== null && mfaTitle) {
          var titleText = mfaTitle.textContent || '';
          if (titleText.includes('Verify your identity') || titleText.includes('verify your identity')) {
            return { state: 'mfa_method_picker' };
          }
        }
        var approvalDesc = document.querySelector('${SELECTORS.MFA_APPROVAL_INDICATOR}');
        var signEl = document.querySelector('#idRichContext_DisplaySign');
        if ((signEl && signEl.offsetParent !== null && signEl.textContent.trim()) ||
            (approvalDesc && approvalDesc.offsetParent !== null && approvalDesc.textContent.trim())) {
          var approvalNumber = signEl ? (signEl.textContent || '').trim() : '';
          return { state: 'mfa_authenticator_push', approvalNumber: approvalNumber || undefined };
        }
        var mfaCodeInput = document.querySelector('${SELECTORS.MFA_CODE_INPUT}');
        if (mfaCodeInput && mfaCodeInput.offsetParent !== null) {
          var pageText = document.body.innerText || '';
          if (pageText.includes('authenticator app') || pageText.includes('verification code from')) {
            return { state: 'mfa_totp_code' };
          }
          return { state: 'mfa_sms_code' };
        }
        var passwordInput = document.querySelector('${SELECTORS.PASSWORD_INPUT}');
        if (passwordInput && passwordInput.offsetParent !== null) {
          return { state: 'password_needed' };
        }
        var usernameInput = document.querySelector('${SELECTORS.USERNAME_INPUT}');
        if (usernameInput && usernameInput.offsetParent !== null) {
          return { state: 'username_needed' };
        }
        return { state: 'loading' };
      })();
    `) as PageStateResult;
  } catch {
    return { state: 'loading' };
  }
}

async function waitForState(
  helpers: AutoLoginHelpers,
  expectedStates: LoginPageState[],
  timeoutMs: number,
  signal?: AbortSignal,
  onState?: (state: PageStateResult) => void,
): Promise<PageStateResult> {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    signal?.throwIfAborted();
    const result = await detectPageState(helpers);
    signal?.throwIfAborted();
    onState?.(result);
    if (expectedStates.includes(result.state)) return result;
    if (result.state === 'error') return result;
    await sleep(AUTO_LOGIN_POLL_INTERVAL_MS, signal);
  }
  throw new Error(`Timed out waiting for states: [${expectedStates.join(', ')}]`);
}

async function clickAccountTile(helpers: AutoLoginHelpers, username: string): Promise<boolean> {
  return await helpers.executeJavaScript(`
    (function() {
      var uname = ${JSON.stringify(username)};
      var matchingTile = document.querySelector('[data-test-id="' + CSS.escape(uname) + '"]');
      if (matchingTile) { matchingTile.click(); return true; }
      var tiles = document.querySelectorAll('${SELECTORS.ACCOUNT_PICKER_TILES} .tile .table');
      for (var i = 0; i < tiles.length; i++) {
        var text = tiles[i].textContent || '';
        if (text.includes(uname)) { tiles[i].click(); return true; }
      }
      var otherTile = document.querySelector('${SELECTORS.ACCOUNT_PICKER_OTHER_TILE}');
      if (otherTile) { otherTile.click(); return false; }
      return false;
    })();
  `) as boolean;
}

async function clickMfaMethod(helpers: AutoLoginHelpers): Promise<string> {
  return await helpers.executeJavaScript(`
    (function() {
      var proofs = document.querySelector('${SELECTORS.MFA_PROOF_SECTION}');
      if (!proofs) return '';
      var tiles = proofs.querySelectorAll('.tile .table[data-value]');
      if (tiles.length === 0) tiles = proofs.querySelectorAll('.tile .table[role="button"]');
      if (tiles.length === 0) return '';
      for (var i = 0; i < tiles.length; i++) {
        var value = tiles[i].getAttribute('data-value') || '';
        if (value === 'OneWaySMS' || value === 'TwoWaySMS' || value.includes('SMS')) {
          tiles[i].click(); return value || 'sms';
        }
      }
      var firstValue = tiles[0].getAttribute('data-value') || 'unknown';
      tiles[0].click();
      return firstValue;
    })();
  `) as string;
}

async function fillAndSubmitUsername(helpers: AutoLoginHelpers, username: string): Promise<void> {
  await helpers.executeJavaScript(`
    (function() {
      var input = document.querySelector('${SELECTORS.USERNAME_INPUT}');
      if (!input) throw new Error('Username input not found');
      var nativeInputValueSetter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;
      nativeInputValueSetter.call(input, ${JSON.stringify(username)});
      input.dispatchEvent(new Event('input', { bubbles: true }));
      input.dispatchEvent(new Event('change', { bubbles: true }));
      var btn = document.querySelector('${SELECTORS.NEXT_BUTTON}');
      if (btn) btn.click();
    })();
  `);
}

async function fillAndSubmitPassword(helpers: AutoLoginHelpers, password: string): Promise<void> {
  await helpers.executeJavaScript(`
    (function() {
      var input = document.querySelector('${SELECTORS.PASSWORD_INPUT}');
      if (!input) throw new Error('Password input not found');
      var nativeInputValueSetter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;
      nativeInputValueSetter.call(input, ${JSON.stringify(password)});
      input.dispatchEvent(new Event('input', { bubbles: true }));
      input.dispatchEvent(new Event('change', { bubbles: true }));
      var btn = document.querySelector('${SELECTORS.SIGN_IN_BUTTON}');
      if (btn) btn.click();
    })();
  `);
}

async function fillAndSubmitMfaCode(helpers: AutoLoginHelpers, code: string): Promise<void> {
  await helpers.executeJavaScript(`
    (function() {
      var input = document.querySelector('${SELECTORS.MFA_CODE_INPUT}');
      if (!input) throw new Error('MFA code input not found');
      var nativeInputValueSetter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;
      nativeInputValueSetter.call(input, ${JSON.stringify(code)});
      input.dispatchEvent(new Event('input', { bubbles: true }));
      input.dispatchEvent(new Event('change', { bubbles: true }));
      var btn = document.querySelector('${SELECTORS.MFA_CODE_VERIFY_BUTTON}');
      if (btn) btn.click();
    })();
  `);
}

async function clickStaySignedIn(helpers: AutoLoginHelpers, yes: boolean): Promise<void> {
  const selector = yes ? SELECTORS.STAY_SIGNED_IN_YES : SELECTORS.STAY_SIGNED_IN_NO;
  await helpers.executeJavaScript(`
    (function() {
      var btn = document.querySelector('${selector}');
      if (btn) btn.click();
    })();
  `);
}

function fallback(
  helpers: AutoLoginHelpers,
  callbacks: AutoLoginCallbacks,
  reason: string,
): AutoLoginResult {
  callbacks.onFallback(reason);
  helpers.show();
  return { success: false, fallbackToManual: true, reason };
}

async function handleMfa(
  helpers: AutoLoginHelpers,
  startState: PageStateResult,
  callbacks: AutoLoginCallbacks,
  signal?: AbortSignal,
): Promise<{ result: AutoLoginResult | null; currentState: PageStateResult }> {
  let currentState = startState;

  if (currentState.state === 'mfa_method_picker') {
    const clickedMethod = await clickMfaMethod(helpers);
    if (!clickedMethod) {
      return { result: fallback(helpers, callbacks, 'No MFA method available to click'), currentState };
    }
    let afterMethodPick: PageStateResult;
    try {
      afterMethodPick = await waitForState(helpers,
        ['mfa_sms_code', 'mfa_totp_code', 'mfa_authenticator_push', 'stay_signed_in', 'not_login_page', 'error'],
        AUTO_LOGIN_STEP_TIMEOUT_MS, signal,
      );
    } catch {
      signal?.throwIfAborted();
      return { result: fallback(helpers, callbacks, 'Timed out waiting for MFA method to load'), currentState };
    }
    if (afterMethodPick.state === 'error') {
      return { result: fallback(helpers, callbacks, `MFA method error: ${afterMethodPick.errorText}`), currentState };
    }
    currentState = afterMethodPick;
  }

  if (currentState.state === 'mfa_sms_code' || currentState.state === 'mfa_totp_code') {
    const mfaType = currentState.state === 'mfa_sms_code' ? 'sms' : 'totp';
    let code: string;
    try {
      code = await callbacks.onMfaCodeNeeded(mfaType);
    } catch {
      signal?.throwIfAborted();
      return { result: fallback(helpers, callbacks, 'MFA code entry cancelled'), currentState };
    }
    signal?.throwIfAborted();
    await fillAndSubmitMfaCode(helpers, code);
    let afterMfa: PageStateResult;
    try {
      afterMfa = await waitForState(helpers,
        ['stay_signed_in', 'not_login_page', 'error'],
        AUTO_LOGIN_STEP_TIMEOUT_MS * 2, signal,
      );
    } catch {
      signal?.throwIfAborted();
      return { result: fallback(helpers, callbacks, 'Timed out waiting for MFA verification'), currentState };
    }
    callbacks.onMfaApprovalComplete();
    if (afterMfa.state === 'error') {
      return { result: fallback(helpers, callbacks, `MFA error: ${afterMfa.errorText}`), currentState };
    }
    currentState = afterMfa;
  }

  if (startState.state === 'mfa_authenticator_push' || currentState.state === 'mfa_authenticator_push') {
    let approvalNumber = currentState.approvalNumber || startState.approvalNumber;
    callbacks.onMfaApprovalNeeded(approvalNumber);
    let afterApproval: PageStateResult;
    try {
      afterApproval = await waitForState(helpers,
        ['stay_signed_in', 'not_login_page', 'mfa_sms_code', 'mfa_totp_code', 'mfa_method_picker', 'error'],
        AUTO_LOGIN_MFA_TIMEOUT_MS, signal,
        (page) => {
          if (page.state === 'mfa_authenticator_push' && page.approvalNumber && page.approvalNumber !== approvalNumber) {
            approvalNumber = page.approvalNumber;
            callbacks.onMfaApprovalNeeded(approvalNumber);
          }
        },
      );
    } catch {
      signal?.throwIfAborted();
      callbacks.onMfaApprovalComplete();
      return { result: fallback(helpers, callbacks, 'MFA push approval timed out'), currentState };
    }
    callbacks.onMfaApprovalComplete();
    if (afterApproval.state === 'error') {
      return { result: fallback(helpers, callbacks, `MFA approval error: ${afterApproval.errorText}`), currentState };
    }
    currentState = afterApproval;
    if (['mfa_sms_code', 'mfa_totp_code', 'mfa_method_picker'].includes(currentState.state)) {
      return handleMfa(helpers, currentState, callbacks, signal);
    }
  }

  return { result: null, currentState };
}

export async function performAutoLogin(
  helpers: AutoLoginHelpers,
  credentials: { username: string; password: string },
  callbacks: AutoLoginCallbacks,
  signal?: AbortSignal,
): Promise<AutoLoginResult> {
  const states: LoginPageState[] = [
    'username_needed', 'password_needed', 'account_picker', 'mfa_method_picker',
    'mfa_sms_code', 'mfa_authenticator_push', 'mfa_totp_code', 'stay_signed_in', 'not_login_page',
  ];
  try {
    // Cached SSO and passwordless accounts can start at MFA or skip password
    // after the username. Handle each observed step instead of assuming a form order.
    let current = await waitForState(helpers, states, AUTO_LOGIN_STEP_TIMEOUT_MS, signal);
    for (let step = 0; step < 12; step++) {
      signal?.throwIfAborted();
      if (current.state === 'error') return fallback(helpers, callbacks, `Login page error: ${current.errorText}`);
      if (current.state === 'not_login_page') return { success: true, fallbackToManual: false };
      if (current.state === 'stay_signed_in') {
        await clickStaySignedIn(helpers, true);
        return { success: true, fallbackToManual: false };
      }
      if (current.state.startsWith('mfa_')) {
        const mfa = await handleMfa(helpers, current, callbacks, signal);
        if (mfa.result) return mfa.result;
        current = mfa.currentState;
        continue;
      }
      const previous = current.state;
      if (previous === 'account_picker') await clickAccountTile(helpers, credentials.username);
      else if (previous === 'username_needed') await fillAndSubmitUsername(helpers, credentials.username);
      else if (previous === 'password_needed') await fillAndSubmitPassword(helpers, credentials.password);
      current = await waitForState(helpers, states.filter((s) => s !== previous), AUTO_LOGIN_STEP_TIMEOUT_MS, signal);
    }
    return fallback(helpers, callbacks, 'Sign-in needs attention');
  } catch (err) {
    if (signal?.aborted) return { success: false, fallbackToManual: false, reason: 'Sign-in finished' };
    return fallback(helpers, callbacks, `Unable to finish automatic sign-in: ${err}`);
  }
}
