import React, { useEffect, useRef } from 'react';

/** Only actual input counts; mounting a panel, polling, and focus alone do not. */
export function withMfaActivity(Component: React.ComponentType<any>, settings = false) {
  return function MfaActivityView(props: any) {
    const source = useRef(`mfa-${Math.random().toString(36).slice(2)}`);
    const last = useRef(-Infinity);
    const action = useRef(props.onAction);
    action.current = props.onAction;
    useEffect(() => {
      const inactive = () => { last.current = -Infinity; action.current?.('mfa-user-activity', { source: source.current, active: false }); };
      const visibility = () => { if (document.visibilityState !== 'visible') inactive(); };
      window.addEventListener('blur', inactive);
      document.addEventListener('visibilitychange', visibility);
      return () => {
        window.removeEventListener('blur', inactive);
        document.removeEventListener('visibilitychange', visibility);
        inactive();
      };
    }, []);
    const activity = (event: React.SyntheticEvent) => {
      if (!event.isTrusted || document.visibilityState !== 'visible' || !document.hasFocus()) return;
      if (Date.now() - last.current < 1_000) return;
      last.current = Date.now();
      props.onAction('mfa-user-activity', { source: source.current, active: true });
    };
    return <div style={{ display: 'contents' }} onPointerDownCapture={activity} onPointerMoveCapture={activity}
      onKeyDownCapture={activity} onWheelCapture={activity}>
      {props.pluginState?.mfaActivityPaused && <p role="status" className="rounded-lg border border-border bg-muted p-3 text-xs text-muted-foreground">
        MFA is paused while you are away. It resumes when you are active; if the sign-in expires, choose Sign in again.
      </p>}
      <Component {...props} />
      {settings && <MfaActivitySettings state={props.pluginState} onAction={props.onAction} />}
    </div>;
  };
}

export function MfaActivitySettings({ state, onAction }: { state?: any; onAction: (action: string, data?: unknown) => void }) {
  const policy = state?.mfaActivityPolicy ?? { source: 'computer', idleMinutes: 5, allowAgent: true };
  const save = (patch: any) => onAction('save-mfa-activity-policy', { ...policy, ...patch });
  return <fieldset className="my-4 space-y-3 rounded-lg border border-border p-3">
    <legend className="px-1 text-xs font-medium text-muted-foreground">MFA activity</legend>
    <p className="text-xs text-muted-foreground">Saved credentials can sign in while you are away. MFA waits for activity. Microsoft may send an initial push immediately after sign-in.</p>
    <label className="block text-xs">Activity source
      <select aria-label="MFA activity source" value={policy.source} onChange={e => save({ source: e.target.value })}
        className="ml-2 rounded border border-border bg-card p-1 text-foreground">
        <option value="computer">Keyboard or mouse anywhere on the computer</option>
        <option value="plugin">Interaction with this plugin</option>
      </select>
    </label>
    <label className="block text-xs">Consider inactive after
      <select aria-label="MFA idle timeout" value={policy.idleMinutes} onChange={e => save({ idleMinutes: Number(e.target.value) })}
        className="mx-2 rounded border border-border bg-card p-1 text-foreground">
        {[1, 2, 5, 10, 15, 30, 60].map(n => <option key={n} value={n}>{n} {n === 1 ? 'minute' : 'minutes'}</option>)}
      </select>
    </label>
    <label className="flex items-center gap-2 text-xs"><input type="checkbox" checked={policy.allowAgent}
      onChange={e => save({ allowAgent: e.target.checked })} />Count an interactive AI agent using this plugin</label>
    <p className="text-xs text-muted-foreground">Scheduled agents do not count. Computer activity uses the operating system’s idle time. Linux requires a supported desktop idle service or xprintidle on X11. If unavailable, interaction with this plugin still counts.</p>
  </fieldset>;
}
