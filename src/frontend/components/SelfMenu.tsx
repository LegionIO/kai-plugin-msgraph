import React, { useEffect, useRef } from 'react';
import type { AuthStatus, Presence } from '../../shared/types.ts';
import { Avatar } from './Avatar.tsx';

const AVAIL_LABEL: Record<string, string> = {
  Available: 'Available',
  AvailableIdle: 'Available (idle)',
  Away: 'Away',
  BeRightBack: 'Be right back',
  Busy: 'Busy',
  BusyIdle: 'Busy (idle)',
  DoNotDisturb: 'Do not disturb',
  Offline: 'Offline',
};

export function SelfMenu({
  auth,
  photo,
  presence,
  onClose,
  onLogout,
}: {
  auth: AuthStatus;
  photo: string | null | undefined;
  presence: Presence | undefined;
  onClose: () => void;
  onLogout: () => void;
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

  const avail = presence?.availability ?? '';
  const activity = presence?.activity ?? '';
  const availLabel = AVAIL_LABEL[avail] ?? avail.replace(/([A-Z])/g, ' $1').trim();
  const activityLabel =
    activity && activity !== avail ? activity.replace(/([A-Z])/g, ' $1').trim() : null;

  return (
    <div
      ref={ref}
      style={{ width: 288, maxWidth: 'calc(100% - 1rem)' }}
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

      <div className="mt-3 rounded-lg border border-border/60 bg-muted/40 px-3 py-2">
        <div className="text-[10px] uppercase tracking-wide text-muted-foreground mb-0.5">Presence</div>
        <div className="text-xs text-foreground">
          {availLabel || '—'}
          {activityLabel && <span className="text-muted-foreground"> · {activityLabel}</span>}
        </div>
        {presence?.statusMessage && (
          <>
            <div className="text-[10px] uppercase tracking-wide text-muted-foreground mt-2 mb-0.5">
              Status message
            </div>
            <div
              style={{ maxHeight: 128 }}
              className="text-xs text-foreground/90 whitespace-pre-wrap break-words overflow-y-auto"
            >
              {presence.statusMessage}
            </div>
          </>
        )}
        <div className="text-[10px] text-muted-foreground mt-2">
          Editing presence/status requires <code>Presence.ReadWrite</code>, which isn't available via
          the current sign-in path — read-only for now.
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
