import React, { useMemo } from 'react';
import { AdaptiveCard } from './AdaptiveCard.tsx';
import type { TaskModuleState } from '../../shared/types.ts';

export function TaskModuleDialog({
  tm,
  onAction,
}: {
  tm: TaskModuleState;
  onAction: (action: string, data?: unknown) => void;
}) {
  const contentJson = useMemo(() => JSON.stringify(tm.card ?? {}), [tm.card]);
  const width =
    typeof tm.width === 'number' ? Math.min(720, Math.max(320, tm.width)) : 520;

  return (
    <div
      style={{
        position: 'fixed', inset: 0, zIndex: 60,
        display: 'flex', alignItems: 'center', justifyContent: 'center',
        background: 'rgba(10,10,14,0.55)',
        backdropFilter: 'blur(2px)', WebkitBackdropFilter: 'blur(2px)',
      }}
      onMouseDown={(e) => {
        if (e.target === e.currentTarget) onAction('close-task-module');
      }}
    >
      <div
        className="bg-card text-foreground border border-border rounded-xl shadow-xl flex flex-col"
        style={{ width, maxWidth: '92vw', maxHeight: '85vh' }}
        onMouseDown={(e) => e.stopPropagation()}
      >
        <div className="flex items-center justify-between gap-3 px-4 py-2.5 border-b border-border">
          <div className="text-sm font-semibold truncate">{tm.title ?? 'App dialog'}</div>
          <button
            type="button"
            onClick={() => onAction('close-task-module')}
            className="p-1 rounded-md text-muted-foreground hover:text-foreground hover:bg-muted"
            aria-label="Close"
          >
            <svg className="w-4 h-4" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={2} strokeLinecap="round">
              <path d="M18 6 6 18" /><path d="M6 6l12 12" />
            </svg>
          </button>
        </div>

        <div className="px-4 py-3 overflow-y-auto" style={{ minHeight: 120 }}>
          {tm.card ? (
            <AdaptiveCard
              contentJson={contentJson}
              mode="taskModule"
              frameless
              pending={tm.submitting}
              teamsDeepLink={`https://teams.microsoft.com/l/message/${tm.chatId}/${tm.messageId}?context=${encodeURIComponent('{"contextType":"chat"}')}`}
              onOpenUrl={(url) => window.open?.(url, '_blank')}
              onOpenInTeams={(url) => onAction('open-in-teams', { url })}
              onInvoke={(req) => {
                if (req.kind === 'task/fetch') {
                  onAction('invoke-card-action', {
                    kind: 'task/fetch',
                    chatId: tm.chatId,
                    messageId: tm.messageId,
                    botId: tm.botId,
                    data: req.data,
                    title: req.title,
                  });
                } else {
                  onAction('submit-task-module', { data: req.data });
                }
              }}
            />
          ) : (
            <div className="flex items-center justify-center py-8">
              <div className="w-5 h-5 border-2 border-primary border-t-transparent rounded-full animate-spin" />
            </div>
          )}
          {tm.error && (
            <div className="mt-3 text-[11px] text-destructive bg-destructive/10 px-3 py-1.5 rounded-md">
              {tm.error}
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
