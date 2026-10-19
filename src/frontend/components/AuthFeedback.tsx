import React, { useEffect, useState } from 'react';
import type { MsgraphPluginState } from '../../shared/types.ts';
import { MfaDialog } from './MfaDialog.tsx';
import { MfaApprovalDialog } from './MfaApprovalDialog.tsx';

type Props = {
  state: Partial<MsgraphPluginState>;
  onAction: (action: string, data?: unknown) => void;
};

export function SignInButton({ state, onAction }: Props) {
  const [requested, setRequested] = useState(false);
  const busy = !!state.auth?.isSigningIn;
  useEffect(() => {
    if (busy || state.auth?.error || state.auth?.isAuthenticated) setRequested(false);
  }, [busy, state.auth?.error, state.auth?.isAuthenticated]);
  return (
    <button
      type="button"
      disabled={busy || requested}
      aria-busy={busy || requested}
      onClick={() => { setRequested(true); onAction('login'); }}
      className="px-4 py-2 text-sm font-medium bg-primary text-primary-foreground rounded-lg hover:bg-primary/90 disabled:opacity-60 disabled:cursor-wait"
    >
      {busy || requested ? 'Signing in…' : 'Sign in'}
    </button>
  );
}

/** Shared by Settings, Teams, and Outlook; backend state survives view changes. */
export function AuthFeedback({ state, onAction }: Props) {
  const mfa = state.mfa;
  return (
    <>
      {state.auth?.isSigningIn && (
        <div className="text-xs text-muted-foreground space-y-2" role="status" aria-live="polite">
          <p>{state.auth.autoLoginStatus || 'Signing in…'}</p>
          <div className="flex justify-center gap-3">
            <button type="button" className="text-primary underline" onClick={() => onAction('show-sign-in')}>Show sign-in window</button>
            <button type="button" className="underline" onClick={() => onAction('cancel-mfa')}>Cancel sign-in</button>
          </div>
        </div>
      )}
      {state.auth?.error && <p role="alert" className="text-xs text-destructive">{state.auth.error}</p>}
      {mfa?.needed && mfa.type === 'push' && (
        <MfaApprovalDialog
          approvalNumber={mfa.approvalNumber}
          onShow={() => onAction('show-sign-in')}
          onCancel={() => onAction('cancel-mfa')}
        />
      )}
      {mfa?.needed && (mfa.type === 'sms' || mfa.type === 'totp') && (
        <MfaDialog
          type={mfa.type}
          onSubmit={(code) => onAction('submit-mfa-code', { code })}
          onCancel={() => onAction('cancel-mfa')}
        />
      )}
    </>
  );
}
