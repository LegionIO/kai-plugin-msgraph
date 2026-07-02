import React, { useEffect, useMemo, useRef, useState } from 'react';
import * as AC from 'adaptivecards';
import type { CardActionKind } from '../../shared/types.ts';

export interface CardInvokeRequest {
  kind: CardActionKind;
  data: Record<string, unknown>;
  verb?: string | null;
  title?: string | null;
}

interface MsTeamsBlock {
  type?: string;
  text?: string;
  displayText?: string;
  value?: unknown;
}

/** Separate the msteams routing block from the rest of action.data. */
function splitMsTeams(raw: unknown): { data: Record<string, unknown>; msteams: MsTeamsBlock | null } {
  if (!raw || typeof raw !== 'object') return { data: {}, msteams: null };
  const { msteams, ...rest } = raw as Record<string, unknown> & { msteams?: MsTeamsBlock };
  return { data: rest, msteams: (msteams && typeof msteams === 'object') ? msteams : null };
}

let cssInjected = false;
function ensureCss() {
  if (cssInjected || typeof document === 'undefined') return;
  const s = document.createElement('style');
  s.setAttribute('data-msgraph-ac', '');
  s.textContent = `
.msg-ac,.msg-ac .ac-adaptiveCard{background:transparent !important;color:inherit;font-family:inherit;font-size:12px;line-height:1.45;box-shadow:none !important;padding:0 !important}
.msg-ac .ac-container,.msg-ac .ac-columnSet,.msg-ac .ac-column{background:transparent !important;min-width:0}
.msg-ac .ac-textBlock{color:inherit !important;margin:0 0 4px}
.msg-ac .ac-textBlock[aria-level]{font-weight:600}
.msg-ac .ac-horizontal-separator{border:0;border-top:1px solid rgba(127,127,127,.25);margin:8px 0}
.msg-ac .ac-pushButton,.msg-ac button{font:inherit;font-size:11px;padding:5px 10px;border-radius:6px;border:1px solid rgba(127,127,127,.35);background:transparent;color:inherit;cursor:pointer}
.msg-ac .ac-pushButton:hover,.msg-ac button:hover{background:rgba(127,127,127,.12)}
.msg-ac .ac-actionSet{gap:6px}
.msg-ac a,.msg-ac .ac-anchor{color:#2563eb;text-decoration:underline;cursor:pointer}
.dark .msg-ac a,.dark .msg-ac .ac-anchor{color:#60a5fa}
.msg-ac img{max-width:100%}
.msg-ac .ac-factset .ac-fact-title{opacity:.85;padding-right:14px}
.msg-ac [style*="font-family: Courier"],.msg-ac [style*="font-family:Courier"],.msg-ac [style*="monospace"]{font-family:ui-monospace,SFMono-Regular,Menlo,monospace !important;background:rgba(127,127,127,.15);padding:1px 4px;border-radius:3px}
.msg-ac .ac-input,.msg-ac input,.msg-ac select,.msg-ac textarea{background:transparent;color:inherit;border:1px solid rgba(127,127,127,.35);border-radius:6px;padding:4px 6px;font:inherit;font-size:11px}
`;
  document.head.appendChild(s);
  cssInjected = true;
}

function fgColors(primary: string, subtle: string, accent: string) {
  const c = (d: string, s: string) => ({ default: d, subtle: s });
  return {
    default: c(primary, subtle),
    dark: c(primary, subtle),
    light: c(primary, subtle),
    accent: c(accent, accent),
    good: c('#10b981', '#34d399'),
    warning: c('#f59e0b', '#fbbf24'),
    attention: c('#ef4444', '#f87171'),
  };
}

function makeHostConfig(dark: boolean) {
  const fg = dark
    ? fgColors('#e5e7eb', '#9ca3af', '#60a5fa')
    : fgColors('#1f2937', '#6b7280', '#2563eb');
  return new AC.HostConfig({
    fontFamily: 'inherit',
    containerStyles: {
      default: { backgroundColor: 'transparent', foregroundColors: fg },
      emphasis: {
        backgroundColor: dark ? 'rgba(255,255,255,0.06)' : 'rgba(0,0,0,0.04)',
        foregroundColors: fg,
      },
      accent: { backgroundColor: 'transparent', foregroundColors: fg },
      good: { backgroundColor: 'transparent', foregroundColors: fg },
      attention: { backgroundColor: 'transparent', foregroundColors: fg },
      warning: { backgroundColor: 'transparent', foregroundColors: fg },
    },
    spacing: { small: 4, default: 8, medium: 12, large: 16, extraLarge: 20, padding: 0 },
    separator: { lineThickness: 1, lineColor: dark ? 'rgba(255,255,255,0.12)' : 'rgba(0,0,0,0.12)' },
  });
}

export function AdaptiveCard({
  contentJson,
  onOpenUrl,
  teamsDeepLink,
  onOpenInTeams,
  onInvoke,
  mode = 'message',
  pending = false,
  frameless = false,
}: {
  contentJson: string;
  onOpenUrl: (url: string) => void;
  teamsDeepLink: string | null;
  onOpenInTeams: (url: string) => void;
  /** When provided, Submit/Execute actions are routed here instead of the Teams redirect. */
  onInvoke?: (req: CardInvokeRequest) => void;
  /** In a task-module dialog, plain Action.Submit means task/submit, not messageback. */
  mode?: 'message' | 'taskModule';
  pending?: boolean;
  frameless?: boolean;
}) {
  const hostRef = useRef<HTMLDivElement>(null);
  const [error, setError] = useState<string | null>(null);
  const [redirect, setRedirect] = useState<{ label: string; secs: number } | null>(null);
  const redirectTimer = useRef<ReturnType<typeof setInterval> | null>(null);

  const onOpenUrlRef = useRef(onOpenUrl);
  const onOpenInTeamsRef = useRef(onOpenInTeams);
  const onInvokeRef = useRef(onInvoke);
  const deepLinkRef = useRef(teamsDeepLink);
  const modeRef = useRef(mode);
  useEffect(() => {
    onOpenUrlRef.current = onOpenUrl;
    onOpenInTeamsRef.current = onOpenInTeams;
    onInvokeRef.current = onInvoke;
    deepLinkRef.current = teamsDeepLink;
    modeRef.current = mode;
  }, [onOpenUrl, onOpenInTeams, onInvoke, teamsDeepLink, mode]);

  const startRedirect = (label: string) => {
    const link = deepLinkRef.current;
    if (!link) return;
    if (redirectTimer.current) clearInterval(redirectTimer.current);
    setRedirect({ label, secs: 5 });
    redirectTimer.current = setInterval(() => {
      setRedirect((r) => {
        if (!r) return null;
        if (r.secs <= 1) {
          if (redirectTimer.current) clearInterval(redirectTimer.current);
          onOpenInTeamsRef.current(link);
          return null;
        }
        return { ...r, secs: r.secs - 1 };
      });
    }, 1000);
  };
  useEffect(() => () => { if (redirectTimer.current) clearInterval(redirectTimer.current); }, []);

  const payload = useMemo(() => {
    try {
      return JSON.parse(contentJson);
    } catch {
      return null;
    }
  }, [contentJson]);

  useEffect(() => {
    ensureCss();
    const host = hostRef.current;
    if (!host) return;
    host.innerHTML = '';
    setError(null);
    if (!payload) {
      setError('Invalid card payload');
      return;
    }
    try {
      const dark =
        typeof document !== 'undefined' && document.documentElement.classList.contains('dark');
      const card = new AC.AdaptiveCard();
      card.hostConfig = makeHostConfig(dark);
      card.onExecuteAction = (action) => {
        if (action instanceof AC.OpenUrlAction && action.url) {
          onOpenUrlRef.current(action.url);
          return;
        }
        if (action instanceof AC.ToggleVisibilityAction || action instanceof AC.ShowCardAction) {
          return; // handled internally by the renderer
        }
        const title = action.title || action.getJsonTypeName?.() || 'This action';
        const invoke = onInvokeRef.current;
        if (!invoke) {
          startRedirect(title);
          return;
        }
        if (action instanceof AC.ExecuteAction) {
          const { data } = splitMsTeams(action.data);
          invoke({ kind: 'execute', data, verb: action.verb ?? null, title });
          return;
        }
        if (action instanceof AC.SubmitAction) {
          const { data, msteams } = splitMsTeams(action.data);
          const t = msteams?.type?.toLowerCase();
          if (t === 'signin') {
            const u = typeof msteams?.value === 'string' ? msteams.value : null;
            if (u) onOpenUrlRef.current(u); else startRedirect(title);
            return;
          }
          if (
            t === 'task/fetch' ||
            (t === 'invoke' && (msteams?.value as { type?: string } | undefined)?.type === 'task/fetch')
          ) {
            invoke({ kind: 'task/fetch', data, title });
            return;
          }
          if (modeRef.current === 'taskModule') {
            invoke({ kind: 'task/submit', data, title });
            return;
          }
          if (t === 'messageback' || t === 'imback') {
            const val =
              msteams?.value && typeof msteams.value === 'object'
                ? { ...data, ...(msteams.value as Record<string, unknown>) }
                : data;
            invoke({ kind: 'messageback', data: val, title });
            return;
          }
          invoke({ kind: 'messageback', data, title });
          return;
        }
        startRedirect(title);
      };
      card.parse(payload);
      const el = card.render();
      if (el) {
        el.classList.add('msg-ac');
        host.appendChild(el);
      } else {
        setError('Card rendered empty');
      }
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    }
    // Only rebuild when the card content itself changes.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [payload]);

  return (
    <div
      className={frameless ? 'text-foreground' : 'my-1.5 rounded-lg border border-border bg-card text-foreground'}
      style={{ whiteSpace: 'normal', position: 'relative', overflow: 'hidden' }}
    >
      <div ref={hostRef} className={frameless ? '' : 'px-3 py-2.5'} />
      {error && <div className="px-3 pb-2 text-[11px] text-destructive">Card error: {error}</div>}
      {pending && !redirect && (
        <div
          style={{
            position: 'absolute', inset: 0, display: 'flex', alignItems: 'center', justifyContent: 'center',
            background: 'rgba(20,20,24,0.35)', backdropFilter: 'blur(1px)', WebkitBackdropFilter: 'blur(1px)',
          }}
        >
          <div className="w-5 h-5 border-2 border-primary border-t-transparent rounded-full animate-spin" />
        </div>
      )}
      {redirect && (
        <div
          style={{
            position: 'absolute',
            inset: 0,
            display: 'flex',
            flexDirection: 'column',
            alignItems: 'center',
            justifyContent: 'center',
            gap: 10,
            padding: 16,
            textAlign: 'center',
            background: 'rgba(20,20,24,0.55)',
            backdropFilter: 'blur(3px)',
            WebkitBackdropFilter: 'blur(3px)',
          }}
          className="text-foreground"
        >
          <div className="text-xs">
            <b>{redirect.label}</b> isn't supported in Kai.
          </div>
          <div className="text-[11px] text-muted-foreground">
            Opening this message in Teams in <b className="text-foreground">{redirect.secs}</b>…
          </div>
          <div className="flex gap-2">
            <button
              type="button"
              onClick={() => {
                if (redirectTimer.current) clearInterval(redirectTimer.current);
                setRedirect(null);
                const link = deepLinkRef.current;
                if (link) onOpenInTeamsRef.current(link);
              }}
              className="px-3 py-1.5 text-xs font-medium bg-primary text-primary-foreground rounded-lg hover:bg-primary/90 transition-colors"
            >
              Open now
            </button>
            <button
              type="button"
              onClick={() => {
                if (redirectTimer.current) clearInterval(redirectTimer.current);
                setRedirect(null);
              }}
              className="px-3 py-1.5 text-xs font-medium text-muted-foreground bg-muted border border-border rounded-lg hover:bg-muted/80 transition-colors"
            >
              Cancel
            </button>
          </div>
        </div>
      )}
    </div>
  );
}
