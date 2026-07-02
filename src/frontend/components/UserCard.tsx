import React, { useEffect } from 'react';
import type { MsgraphPluginState } from '../../shared/types.ts';
import { Avatar } from './Avatar.tsx';

export function UserCard({
  state,
  photos,
  presence,
  onAction,
}: {
  state: MsgraphPluginState;
  photos: Record<string, string | null>;
  presence: MsgraphPluginState['presence'];
  onAction: (action: string, data?: unknown) => void;
}) {
  const card = state.userCard;
  useEffect(() => {
    const k = (e: KeyboardEvent) => { if (e.key === 'Escape') onAction('close-user-card'); };
    document.addEventListener('keydown', k);
    return () => document.removeEventListener('keydown', k);
  }, [onAction]);

  if (!card) return null;
  const p = presence[card.userId];
  const availLabel = p?.availability?.replace(/([A-Z])/g, ' $1').trim();
  const activityLabel = p?.activity && p.activity !== p.availability
    ? p.activity.replace(/([A-Z])/g, ' $1').trim()
    : null;

  return (
    <div className="fixed inset-0 bg-black/60 flex items-center justify-center z-50" onMouseDown={() => onAction('close-user-card')}>
      <div
        style={{ width: 320, maxWidth: 'calc(100% - 2rem)' }}
        className="bg-card rounded-xl shadow-2xl border border-border p-4"
        onMouseDown={(e) => e.stopPropagation()}
      >
        <div className="flex items-center gap-3">
          <Avatar id={card.userId} name={card.displayName ?? '?'} photo={photos[card.userId]} presence={p} size={10} />
          <div className="min-w-0 flex-1">
            <div className="text-sm font-semibold text-foreground truncate">
              {card.loading ? 'Loading…' : card.displayName ?? '(unknown)'}
            </div>
            {card.email && <div className="text-[11px] text-muted-foreground truncate">{card.email}</div>}
            {card.jobTitle && <div className="text-[11px] text-muted-foreground truncate">{card.jobTitle}</div>}
          </div>
          <button type="button" onClick={() => onAction('close-user-card')} className="text-muted-foreground hover:text-foreground">×</button>
        </div>

        {p && (
          <div className="mt-3 rounded-lg border border-border/60 bg-muted/40 px-3 py-2">
            {p && (
              <>
                <div className="text-[10px] uppercase tracking-wide text-muted-foreground mb-0.5">Presence</div>
                <div className="text-xs text-foreground">
                  {availLabel || '—'}
                  {activityLabel && <span className="text-muted-foreground"> · {activityLabel}</span>}
                </div>
              </>
            )}
            {p?.statusMessage && (
              <>
                <div className="text-[10px] uppercase tracking-wide text-muted-foreground mt-2 mb-0.5">Status message</div>
                <div style={{ maxHeight: 96 }} className="text-xs text-foreground/90 whitespace-pre-wrap break-words overflow-y-auto">
                  {p.statusMessage}
                </div>
              </>
            )}
          </div>
        )}

        {card.error && <div className="mt-3 text-[11px] text-destructive">{card.error}</div>}

        <div className="mt-4 flex gap-2">
          <button
            type="button"
            onClick={() => onAction('open-chat-with', { userId: card.userId })}
            className="flex-1 px-3 py-1.5 text-xs font-medium bg-primary text-primary-foreground rounded-lg hover:bg-primary/90 transition-colors"
          >
            Open chat
          </button>
          <button
            type="button"
            onClick={() => onAction('close-user-card')}
            className="px-3 py-1.5 text-xs font-medium text-muted-foreground bg-muted border border-border rounded-lg hover:bg-muted/80 transition-colors"
          >
            Close
          </button>
        </div>
      </div>
    </div>
  );
}
