import React, { useEffect, useRef, useState } from 'react';
import type { MsgraphPluginState, Presence } from '../../shared/types.ts';
import { Avatar } from './Avatar.tsx';
import { RichComposer } from '../editor/RichComposer.tsx';
import type { SerializedPayload } from '../editor/serialize.ts';

type Person = { id: string; displayName: string; email: string | null };

const noop = () => {};

export function NewChatDialog({
  state,
  photos,
  presence,
  onAction,
  onClose,
}: {
  state: MsgraphPluginState;
  photos: Record<string, string | null>;
  presence: Record<string, Presence>;
  onAction: (action: string, data?: unknown) => void;
  onClose: () => void;
}) {
  const [query, setQuery] = useState('');
  const [recipients, setRecipients] = useState<Person[]>([]);
  const recipientsRef = useRef(recipients);
  useEffect(() => { recipientsRef.current = recipients; }, [recipients]);
  const [topic, setTopic] = useState('');
  const debounce = useRef<ReturnType<typeof setTimeout> | null>(null);

  const search = state.peopleSearch;
  const suggestions = (search?.query === query.trim() ? search.results : []).filter(
    (p) => !recipients.some((r) => r.id === p.id),
  );

  useEffect(() => {
    if (debounce.current) clearTimeout(debounce.current);
    debounce.current = setTimeout(() => onAction('search-people', { query }), 250);
    return () => {
      if (debounce.current) clearTimeout(debounce.current);
    };
  }, [query]);

  const add = (p: Person) => {
    setRecipients((r) => (r.some((x) => x.id === p.id) ? r : [...r, p]));
    setQuery('');
  };
  const remove = (id: string) => setRecipients((r) => r.filter((x) => x.id !== id));

  const submit = (payload?: SerializedPayload) => {
    const rs = recipientsRef.current;
    if (rs.length === 0) return;
    onAction('compose-new-chat', {
      recipients: rs,
      topic: rs.length > 1 ? topic.trim() || undefined : undefined,
      payload,
    });
    onClose();
  };

  return (
    <div className="fixed inset-0 bg-black/60 flex items-center justify-center z-50" onMouseDown={onClose}>
      <div
        style={{ width: 520, maxWidth: 'calc(100% - 2rem)' }}
        className="bg-card rounded-xl shadow-2xl border border-border p-4"
        onMouseDown={(e) => e.stopPropagation()}
      >
        <div className="flex items-center justify-between mb-3">
          <h3 className="text-sm font-semibold text-foreground">New chat</h3>
          <button type="button" onClick={onClose} className="text-muted-foreground hover:text-foreground">×</button>
        </div>

        {/* Recipients */}
        <div className="flex flex-wrap items-center gap-1.5 px-2 py-1.5 bg-muted border border-border rounded-lg">
          {recipients.map((r) => (
            <span key={r.id} className="inline-flex items-center gap-1.5 pl-1 pr-1.5 py-0.5 rounded-full bg-primary/15 text-xs">
              <Avatar id={r.id} name={r.displayName} photo={photos[r.id]} size={6} />
              <span className="text-foreground">{r.displayName}</span>
              <button type="button" onClick={() => remove(r.id)} className="text-muted-foreground hover:text-foreground">×</button>
            </span>
          ))}
          <input
            className="flex-1 min-w-[120px] bg-transparent text-xs text-foreground placeholder:text-muted-foreground outline-none py-1"
            placeholder={recipients.length ? 'Add more…' : 'Type a name or email…'}
            value={query}
            onChange={(e) => setQuery(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === 'Backspace' && !query && recipients.length) remove(recipients[recipients.length - 1].id);
              if (e.key === 'Enter' && suggestions[0]) { e.preventDefault(); add(suggestions[0]); }
            }}
            autoFocus
          />
        </div>
        {query.trim().length >= 2 && (
          <div style={{ maxHeight: 168 }} className="mt-1 overflow-y-auto rounded-lg border border-border bg-card">
            {search?.loading && <div className="px-3 py-2 text-[11px] text-muted-foreground animate-pulse">Searching…</div>}
            {suggestions.map((p) => (
              <button
                key={p.id}
                type="button"
                onClick={() => add(p)}
                className="w-full flex items-center gap-2 px-2 py-1.5 hover:bg-muted/60 text-left"
              >
                <Avatar id={p.id} name={p.displayName} photo={photos[p.id]} size={7} />
                <div className="min-w-0">
                  <div className="text-xs text-foreground truncate">{p.displayName}</div>
                  <div className="text-[10px] text-muted-foreground truncate">{p.email}</div>
                </div>
              </button>
            ))}
            {!search?.loading && suggestions.length === 0 && (
              <div className="px-3 py-2 text-[11px] text-muted-foreground">No matches</div>
            )}
          </div>
        )}

        {recipients.length > 1 && (
          <input
            className="mt-3 w-full px-2.5 py-1.5 text-xs bg-muted border border-border rounded-lg text-foreground placeholder:text-muted-foreground focus:border-primary transition-colors"
            placeholder="Group name (optional)"
            value={topic}
            onChange={(e) => setTopic(e.target.value)}
          />
        )}

        <div className="mt-3">
          <RichComposer
            chatId="__new-chat__"
            bare
            placeholder="First message (optional)"
            sending={false}
            disabled={recipients.length === 0}
            replyTo={null}
            editing={null}
            hostedContents={state.hostedContents ?? {}}
            onClearReply={noop}
            onCancelEdit={noop}
            onSaveEdit={noop}
            onSend={submit}
            onSearchPeople={(q) => onAction('search-people', { query: q })}
            peopleSearch={state.peopleSearch ?? null}
            photos={photos}
            presence={presence}
          />
        </div>

        <div className="mt-3 flex items-center justify-end gap-2">
          <button
            type="button"
            onClick={onClose}
            className="px-3 py-1.5 text-xs font-medium text-muted-foreground bg-muted border border-border rounded-lg hover:bg-muted/80 transition-colors"
          >
            Cancel
          </button>
          <button
            type="button"
            onClick={() => submit()}
            disabled={recipients.length === 0}
            className="px-3 py-1.5 text-xs font-medium text-foreground border border-border rounded-lg hover:bg-muted disabled:opacity-50 transition-colors"
          >
            {recipients.length > 1 ? 'Create without message' : 'Open chat'}
          </button>
        </div>
      </div>
    </div>
  );
}

export async function fileToBase64(f: File): Promise<string> {
  const buf = await f.arrayBuffer();
  const bytes = new Uint8Array(buf);
  let bin = '';
  for (let i = 0; i < bytes.length; i++) bin += String.fromCharCode(bytes[i]);
  return btoa(bin);
}
