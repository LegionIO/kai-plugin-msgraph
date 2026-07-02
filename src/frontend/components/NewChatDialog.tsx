import React, { useEffect, useRef, useState } from 'react';
import type { MsgraphPluginState } from '../../shared/types.ts';
import type { PendingImage } from '../../shared/markdown.ts';
import { Avatar } from './Avatar.tsx';

type Person = { id: string; displayName: string; email: string | null };

export function NewChatDialog({
  state,
  photos,
  onAction,
  onClose,
}: {
  state: MsgraphPluginState;
  photos: Record<string, string | null>;
  onAction: (action: string, data?: unknown) => void;
  onClose: () => void;
}) {
  const [query, setQuery] = useState('');
  const [recipients, setRecipients] = useState<Person[]>([]);
  const [topic, setTopic] = useState('');
  const [text, setText] = useState('');
  const [images, setImages] = useState<PendingImage[]>([]);
  const debounce = useRef<ReturnType<typeof setTimeout> | null>(null);
  const fileRef = useRef<HTMLInputElement>(null);

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

  const attach = async (files: FileList | null) => {
    if (!files) return;
    const next: PendingImage[] = [];
    for (const f of Array.from(files)) {
      if (!f.type.startsWith('image/')) continue;
      const b64 = await fileToBase64(f);
      next.push({ id: crypto.randomUUID?.() ?? String(Date.now() + Math.random()), contentType: f.type, contentBytes: b64, name: f.name });
    }
    setImages((i) => [...i, ...next]);
  };

  const canSend = recipients.length > 0;
  const submit = () => {
    if (!canSend) return;
    onAction('compose-new-chat', {
      recipients,
      topic: recipients.length > 1 ? topic.trim() || undefined : undefined,
      text: text.trim() || undefined,
      images: images.length ? images : undefined,
    });
    onClose();
  };

  return (
    <div className="fixed inset-0 bg-black/60 flex items-center justify-center z-50" onMouseDown={onClose}>
      <div
        style={{ width: 460, maxWidth: 'calc(100% - 2rem)' }}
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

        <textarea
          className="mt-3 w-full resize-none px-3 py-2 text-sm bg-muted border border-border rounded-lg text-foreground placeholder:text-muted-foreground focus:border-primary transition-colors"
          rows={3}
          placeholder="First message (optional) — supports **bold**, *italic*, `code`, ```blocks```"
          value={text}
          onChange={(e) => setText(e.target.value)}
        />

        {images.length > 0 && (
          <div className="mt-2 flex flex-wrap gap-1.5">
            {images.map((img) => (
              <div key={img.id} className="relative">
                <img
                  src={`data:${img.contentType};base64,${img.contentBytes}`}
                  alt={img.name}
                  style={{ width: 56, height: 56 }}
                  className="rounded-md object-cover border border-border"
                />
                <button
                  type="button"
                  onClick={() => setImages((i) => i.filter((x) => x.id !== img.id))}
                  style={{ width: 16, height: 16 }}
                  className="absolute -top-1 -right-1 rounded-full bg-card border border-border text-[10px] leading-none flex items-center justify-center text-muted-foreground hover:text-foreground"
                >
                  ×
                </button>
              </div>
            ))}
          </div>
        )}

        <div className="mt-3 flex items-center justify-between">
          <div>
            <input ref={fileRef} type="file" accept="image/*" multiple hidden onChange={(e) => attach(e.target.files)} />
            <button
              type="button"
              onClick={() => fileRef.current?.click()}
              className="p-1.5 rounded-md text-muted-foreground hover:text-foreground hover:bg-muted transition-colors"
              title="Attach image"
            >
              <svg className="w-4 h-4" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={2} strokeLinecap="round" strokeLinejoin="round">
                <rect x="3" y="3" width="18" height="18" rx="2" /><circle cx="9" cy="9" r="2" /><path d="m21 15-5-5L5 21" />
              </svg>
            </button>
          </div>
          <div className="flex gap-2">
            <button
              type="button"
              onClick={onClose}
              className="px-3 py-1.5 text-xs font-medium text-muted-foreground bg-muted border border-border rounded-lg hover:bg-muted/80 transition-colors"
            >
              Cancel
            </button>
            <button
              type="button"
              onClick={submit}
              disabled={!canSend}
              className="px-3 py-1.5 text-xs font-medium bg-primary text-primary-foreground rounded-lg hover:bg-primary/90 disabled:opacity-50 transition-colors"
            >
              {recipients.length > 1 ? 'Create' : text.trim() || images.length ? 'Send' : 'Open'}
            </button>
          </div>
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
