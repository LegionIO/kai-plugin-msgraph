import React, { useEffect, useRef, useState } from 'react';
import type { AuthStatus, Presence } from '../../shared/types.ts';
import { Avatar } from './Avatar.tsx';

const AVAIL_OPTIONS: Array<{ value: string; label: string; color: string }> = [
  { value: 'Available', label: 'Available', color: '#6bb700' },
  { value: 'Busy', label: 'Busy', color: '#c4314b' },
  { value: 'DoNotDisturb', label: 'Do not disturb', color: '#c4314b' },
  { value: 'BeRightBack', label: 'Be right back', color: '#ffaa44' },
  { value: 'Away', label: 'Appear away', color: '#ffaa44' },
  { value: 'Offline', label: 'Appear offline', color: '#8a8886' },
];

export function SelfMenu({
  auth,
  photo,
  presence,
  onClose,
  onLogout,
  onAction,
}: {
  auth: AuthStatus;
  photo: string | null | undefined;
  presence: Presence | undefined;
  onClose: () => void;
  onLogout: () => void;
  onAction: (action: string, data?: unknown) => void;
}) {
  const ref = useRef<HTMLDivElement>(null);
  useEffect(() => {
    const h = (e: MouseEvent) => {
      if (ref.current && !ref.current.contains(e.target as Node)) onClose();
    };
    const k = (e: KeyboardEvent) => { if (e.key === 'Escape') onClose(); };
    document.addEventListener('mousedown', h);
    document.addEventListener('keydown', k);
    return () => {
      document.removeEventListener('mousedown', h);
      document.removeEventListener('keydown', k);
    };
  }, [onClose]);

  const initialNote = presence?.statusMessage ?? '';
  const [note, setNote] = useState(initialNote);
  const [pinned, setPinned] = useState(true);
  const [savingNote, setSavingNote] = useState(false);
  const [savingPresence, setSavingPresence] = useState<string | null>(null);
  const noteDirty = note !== initialNote;

  const currentAvail = presence?.availability ?? '';

  return (
    <div
      ref={ref}
      style={{ width: 300, maxWidth: 'calc(100% - 1rem)' }}
      className="absolute right-2 top-11 z-40 overflow-hidden rounded-xl border border-border bg-card shadow-2xl p-4"
    >
      <div className="flex items-center gap-3">
        {auth.objectId && (
          <Avatar
            id={auth.objectId}
            name={auth.displayName ?? auth.email ?? 'Me'}
            photo={photo}
            presence={presence}
            size={10}
          />
        )}
        <div className="min-w-0">
          <div className="text-sm font-semibold text-foreground truncate">
            {auth.displayName ?? auth.email}
          </div>
          <div className="text-[11px] text-muted-foreground truncate">{auth.email}</div>
          {auth.jobTitle && (
            <div className="text-[11px] text-muted-foreground truncate">{auth.jobTitle}</div>
          )}
        </div>
      </div>

      <div className="mt-3 rounded-lg border border-border/60 bg-muted/40 px-3 py-2.5">
        <div className="text-[10px] uppercase tracking-wide text-muted-foreground mb-1.5">Set presence</div>
        <div className="flex flex-col gap-0.5">
          {AVAIL_OPTIONS.map((o) => {
            const active = currentAvail === o.value;
            const busy = savingPresence === o.value;
            return (
              <button
                key={o.value}
                type="button"
                disabled={savingPresence !== null}
                onClick={() => {
                  setSavingPresence(o.value);
                  Promise.resolve(onAction('set-presence', { availability: o.value })).finally(
                    () => setSavingPresence(null),
                  );
                }}
                className={`flex items-center gap-2 px-2 py-1 rounded-md text-xs text-left transition-colors ${
                  active ? 'bg-primary/15 text-foreground' : 'hover:bg-muted text-foreground/90'
                } ${savingPresence !== null ? 'opacity-60' : ''}`}
              >
                <span
                  style={{
                    width: 10, height: 10, borderRadius: '50%', flexShrink: 0,
                    background: o.value === 'Offline' ? 'transparent' : o.color,
                    border: o.value === 'Offline' ? `2px solid ${o.color}` : 'none',
                  }}
                />
                <span className="flex-1">{o.label}</span>
                {busy && <span className="w-3 h-3 border-2 border-primary border-t-transparent rounded-full animate-spin" />}
                {active && !busy && (
                  <svg className="w-3 h-3" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={3} strokeLinecap="round">
                    <path d="M20 6 9 17l-5-5" />
                  </svg>
                )}
              </button>
            );
          })}
          <button
            type="button"
            disabled={savingPresence !== null}
            onClick={() => {
              setSavingPresence('__reset');
              Promise.resolve(onAction('set-presence', { availability: null })).finally(
                () => setSavingPresence(null),
              );
            }}
            className="flex items-center gap-2 px-2 py-1 rounded-md text-xs text-left text-muted-foreground hover:bg-muted"
          >
            <span style={{ width: 10, flexShrink: 0 }} />
            <span className="flex-1">Reset (automatic)</span>
            {savingPresence === '__reset' && <span className="w-3 h-3 border-2 border-primary border-t-transparent rounded-full animate-spin" />}
          </button>
        </div>

        <div className="text-[10px] uppercase tracking-wide text-muted-foreground mt-3 mb-1">Status message</div>
        <textarea
          value={note}
          onChange={(e) => setNote(e.target.value)}
          rows={3}
          placeholder="What's your status?"
          style={{ resize: 'vertical', maxHeight: 140 }}
          className="w-full rounded-md border border-border bg-background px-2 py-1.5 text-xs text-foreground focus:outline-none focus:ring-1 focus:ring-primary"
        />
        <label className="flex items-center gap-1.5 mt-1.5 text-[11px] text-muted-foreground cursor-pointer">
          <input
            type="checkbox"
            checked={pinned}
            onChange={(e) => setPinned(e.target.checked)}
            className="w-3 h-3"
          />
          Show when people message me
        </label>
        <div className="flex gap-1.5 mt-2">
          <button
            type="button"
            disabled={!noteDirty || savingNote}
            onClick={() => {
              setSavingNote(true);
              Promise.resolve(onAction('set-status-message', { message: note, pinned })).finally(
                () => setSavingNote(false),
              );
            }}
            className="flex-1 px-2.5 py-1.5 text-[11px] font-medium bg-primary text-primary-foreground rounded-md hover:bg-primary/90 disabled:opacity-40 transition-colors"
          >
            {savingNote ? 'Saving…' : 'Save'}
          </button>
          {initialNote && (
            <button
              type="button"
              disabled={savingNote}
              onClick={() => {
                setNote('');
                setSavingNote(true);
                Promise.resolve(onAction('set-status-message', { message: '', pinned: false })).finally(
                  () => setSavingNote(false),
                );
              }}
              className="px-2.5 py-1.5 text-[11px] font-medium text-muted-foreground bg-muted border border-border rounded-md hover:bg-muted/80 disabled:opacity-40 transition-colors"
            >
              Clear
            </button>
          )}
        </div>
      </div>

      <button
        type="button"
        onClick={onLogout}
        className="mt-3 w-full px-3 py-1.5 text-xs font-medium text-muted-foreground bg-muted border border-border rounded-lg hover:bg-muted/80 transition-colors"
      >
        Sign out
      </button>
    </div>
  );
}
