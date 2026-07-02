import React, { useEffect, useRef, useState } from 'react';
import type { MsgraphPluginState } from '../../shared/types.ts';
import { Avatar } from './Avatar.tsx';

type Person = { id: string; displayName: string; email: string | null };

export function ForwardDialog({
  target,
  state,
  photos,
  onAction,
}: {
  target: NonNullable<MsgraphPluginState['forwardTarget']>;
  state: MsgraphPluginState;
  photos: Record<string, string | null>;
  onAction: (action: string, data?: unknown) => void;
}) {
  const [query, setQuery] = useState('');
  const [recipients, setRecipients] = useState<Person[]>([]);
  const debounce = useRef<ReturnType<typeof setTimeout> | null>(null);

  const search = state.peopleSearch;
  const suggestions = (search?.query === query.trim() ? search.results : []).filter(
    (p) => !recipients.some((r) => r.id === p.id),
  );

  useEffect(() => {
    if (debounce.current) clearTimeout(debounce.current);
    debounce.current = setTimeout(() => onAction('search-people', { query }), 200);
    return () => { if (debounce.current) clearTimeout(debounce.current); };
  }, [query]);

  const close = () => onAction('set-forward-target', null);
  const submit = () => {
    if (!recipients.length) return;
    onAction('forward-message', { source: target, recipients });
  };

  return (
    <div className="fixed inset-0 bg-black/60 flex items-center justify-center z-50" onMouseDown={close}>
      <div
        style={{ width: 420, maxWidth: 'calc(100% - 2rem)' }}
        className="bg-card rounded-xl shadow-2xl border border-border p-4"
        onMouseDown={(e) => e.stopPropagation()}
      >
        <div className="flex items-center justify-between mb-3">
          <h3 className="text-sm font-semibold text-foreground">Forward message</h3>
          <button type="button" onClick={close} className="text-muted-foreground hover:text-foreground">×</button>
        </div>

        <div className="mb-3 rounded-lg border-l-2 border-border bg-muted/40 px-2.5 py-1.5 text-[11px]">
          {target.senderName && <div className="font-medium opacity-80">{target.senderName}</div>}
          <div className="opacity-75 line-clamp-3 whitespace-pre-wrap break-words">{target.text ?? '(message)'}</div>
        </div>

        <div className="flex flex-wrap items-center gap-1.5 px-2 py-1.5 bg-muted border border-border rounded-lg">
          {recipients.map((r) => (
            <span key={r.id} className="inline-flex items-center gap-1.5 pl-1 pr-1.5 py-0.5 rounded-full bg-primary/15 text-xs">
              <Avatar id={r.id} name={r.displayName} photo={photos[r.id]} size={6} />
              <span className="text-foreground">{r.displayName}</span>
              <button type="button" onClick={() => setRecipients((rs) => rs.filter((x) => x.id !== r.id))} className="text-muted-foreground hover:text-foreground">×</button>
            </span>
          ))}
          <input
            className="flex-1 min-w-[120px] bg-transparent text-xs text-foreground placeholder:text-muted-foreground outline-none py-1"
            placeholder={recipients.length ? 'Add more…' : 'Forward to…'}
            value={query}
            onChange={(e) => setQuery(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === 'Enter' && suggestions[0]) {
                e.preventDefault();
                setRecipients((r) => [...r, suggestions[0]]);
                setQuery('');
              }
            }}
            autoFocus
          />
        </div>
        {query.trim().length >= 1 && (
          <div style={{ maxHeight: 168 }} className="mt-1 overflow-y-auto rounded-lg border border-border bg-card">
            {search?.loading && <div className="px-3 py-2 text-[11px] text-muted-foreground animate-pulse">Searching…</div>}
            {suggestions.map((p) => (
              <button
                key={p.id}
                type="button"
                onClick={() => { setRecipients((r) => [...r, p]); setQuery(''); }}
                className="w-full flex items-center gap-2 px-2 py-1.5 hover:bg-muted/60 text-left"
              >
                <Avatar id={p.id} name={p.displayName} photo={photos[p.id]} size={7} />
                <div className="min-w-0">
                  <div className="text-xs text-foreground truncate">{p.displayName}</div>
                  <div className="text-[10px] text-muted-foreground truncate">{p.email}</div>
                </div>
              </button>
            ))}
          </div>
        )}

        <div className="mt-4 flex justify-end gap-2">
          <button type="button" onClick={close} className="px-3 py-1.5 text-xs font-medium text-muted-foreground bg-muted border border-border rounded-lg hover:bg-muted/80">
            Cancel
          </button>
          <button type="button" onClick={submit} disabled={!recipients.length} className="px-3 py-1.5 text-xs font-medium bg-primary text-primary-foreground rounded-lg hover:bg-primary/90 disabled:opacity-50">
            Forward
          </button>
        </div>
      </div>
    </div>
  );
}
