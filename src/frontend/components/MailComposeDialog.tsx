import React, { useRef, useState } from 'react';
import type { MailAddress, MailComposeState, MsgraphPluginState, OutgoingMail } from '../../shared/types.ts';
import { fileToBase64 } from './NewChatDialog.tsx';
import { MailEditor, type MailEditorHandle } from '../editor/MailEditor.tsx';

export function MailComposeDialog({
  compose,
  sending,
  peopleSearch,
  onAction,
}: {
  compose: MailComposeState;
  sending: boolean;
  peopleSearch: MsgraphPluginState['peopleSearch'];
  onAction: (action: string, data?: unknown) => void;
}) {
  const [to, setTo] = useState<MailAddress[]>(compose.to);
  const [cc, setCc] = useState<MailAddress[]>(compose.cc);
  const [showCc, setShowCc] = useState(compose.cc.length > 0);
  const [subject, setSubject] = useState(compose.subject);
  const [files, setFiles] = useState<File[]>([]);
  const [includeSig, setIncludeSig] = useState(!!compose.signatureHtml);
  const fileRef = useRef<HTMLInputElement>(null);
  const editorRef = useRef<MailEditorHandle>(null);

  const canSend = !sending && subject.trim().length > 0 && (to.length > 0 || compose.mode === 'reply' || compose.mode === 'replyAll');

  const send = async () => {
    const [serialized, fileAtts] = await Promise.all([
      editorRef.current?.serialize() ?? Promise.resolve({ html: '', inlineImages: [], isEmpty: true }),
      Promise.all(
        files.map(async (f) => ({
          name: f.name,
          contentType: f.type || 'application/octet-stream',
          contentBytes: await fileToBase64(f),
        })),
      ),
    ]);
    let bodyHtml = serialized.html;
    if (includeSig && compose.signatureHtml) bodyHtml += compose.signatureHtml;
    if (compose.quotedHtml) bodyHtml += `<br>${compose.quotedHtml}`;
    const attachments = [
      ...serialized.inlineImages.map((img) => ({
        name: img.name,
        contentType: img.contentType,
        contentBytes: img.contentBytes,
        contentId: img.cid,
        isInline: true,
      })),
      ...fileAtts,
    ];
    const mail: OutgoingMail = { to, cc, subject, bodyHtml, attachments: attachments.length ? attachments : undefined };
    onAction('send-mail', { mail });
  };

  const modeLabel =
    compose.mode === 'reply' ? 'Reply'
    : compose.mode === 'replyAll' ? 'Reply all'
    : compose.mode === 'forward' ? 'Forward'
    : 'New message';

  return (
    <div
      style={{
        position: 'fixed', inset: 0, zIndex: 60,
        display: 'flex', alignItems: 'center', justifyContent: 'center',
        background: 'rgba(10,10,14,0.55)', backdropFilter: 'blur(2px)', WebkitBackdropFilter: 'blur(2px)',
      }}
      onMouseDown={(e) => { if (e.target === e.currentTarget) onAction('close-compose'); }}
    >
      <div
        className="bg-card text-foreground border border-border rounded-xl shadow-2xl flex flex-col"
        style={{ width: 680, maxWidth: '94vw', maxHeight: '88vh' }}
        onMouseDown={(e) => e.stopPropagation()}
      >
        <div className="flex items-center justify-between gap-3 px-4 py-2.5 border-b border-border">
          <div className="text-sm font-semibold">{modeLabel}</div>
          <button
            type="button"
            onClick={() => onAction('close-compose')}
            className="p-1 rounded-md text-muted-foreground hover:text-foreground hover:bg-muted"
            aria-label="Close"
          >
            <svg className="w-4 h-4" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={2} strokeLinecap="round">
              <path d="M18 6 6 18" /><path d="M6 6l12 12" />
            </svg>
          </button>
        </div>

        <div className="px-4 py-3 flex flex-col gap-2 overflow-y-auto">
          <RecipientField
            label="To"
            value={to}
            onChange={setTo}
            peopleSearch={peopleSearch}
            onSearch={(q) => onAction('search-people', { query: q })}
            trailing={
              !showCc && (
                <button type="button" className="text-[11px] text-muted-foreground hover:text-foreground" onClick={() => setShowCc(true)}>
                  Cc
                </button>
              )
            }
          />
          {showCc && (
            <RecipientField
              label="Cc"
              value={cc}
              onChange={setCc}
              peopleSearch={peopleSearch}
              onSearch={(q) => onAction('search-people', { query: q })}
            />
          )}
          <div className="flex items-center gap-2">
            <span className="text-[11px] text-muted-foreground w-10 shrink-0">Subject</span>
            <input
              value={subject}
              onChange={(e) => setSubject(e.target.value)}
              placeholder="Subject"
              className="flex-1 rounded-md border border-border bg-background px-2 py-1.5 text-xs focus:outline-none focus:ring-1 focus:ring-primary"
            />
          </div>
          <MailEditor ref={editorRef} minHeight={200} />
          {files.length > 0 && (
            <div className="flex flex-wrap gap-1.5">
              {files.map((f, i) => (
                <span
                  key={i}
                  className="inline-flex items-center gap-1.5 rounded-md border border-border bg-muted/50 px-2 py-1 text-[11px]"
                >
                  <svg className="w-3 h-3 opacity-70" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={2}>
                    <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" /><polyline points="14 2 14 8 20 8" />
                  </svg>
                  <span className="max-w-[160px] truncate">{f.name}</span>
                  <span className="opacity-60">{(f.size / 1024).toFixed(0)} KB</span>
                  <button type="button" onClick={() => setFiles(files.filter((_, j) => j !== i))} className="opacity-60 hover:opacity-100">×</button>
                </span>
              ))}
            </div>
          )}
          {compose.signatureHtml && (
            <label className="flex items-center gap-1.5 text-[11px] text-muted-foreground cursor-pointer">
              <input type="checkbox" checked={includeSig} onChange={(e) => setIncludeSig(e.target.checked)} className="w-3 h-3" />
              Append signature
            </label>
          )}
          {compose.quotedHtml && (
            <details className="text-[11px] text-muted-foreground">
              <summary className="cursor-pointer">Quoted message will be appended</summary>
            </details>
          )}
        </div>

        <div className="flex items-center justify-between gap-2 px-4 py-2.5 border-t border-border">
          <div className="flex items-center gap-1.5">
            <input
              ref={fileRef}
              type="file"
              multiple
              className="hidden"
              onChange={(e) => {
                const list = Array.from(e.target.files ?? []);
                if (list.length) setFiles((f) => [...f, ...list]);
                if (fileRef.current) fileRef.current.value = '';
              }}
            />
            <button
              type="button"
              onClick={() => fileRef.current?.click()}
              className="p-1.5 rounded-md text-muted-foreground hover:text-foreground hover:bg-muted"
              title="Attach files"
            >
              <svg className="w-4 h-4" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={2} strokeLinecap="round" strokeLinejoin="round">
                <path d="m21.44 11.05-9.19 9.19a6 6 0 0 1-8.49-8.49l8.57-8.57A4 4 0 1 1 18 8.84l-8.59 8.57a2 2 0 0 1-2.83-2.83l8.49-8.48" />
              </svg>
            </button>
          </div>
          <div className="flex gap-2">
            <button
              type="button"
              onClick={() => onAction('close-compose')}
              className="px-3 py-1.5 text-xs font-medium text-muted-foreground bg-muted border border-border rounded-lg hover:bg-muted/80"
            >
              Discard
            </button>
            <button
              type="button"
              disabled={!canSend}
              onClick={() => void send()}
              className="px-4 py-1.5 text-xs font-medium bg-primary text-primary-foreground rounded-lg hover:bg-primary/90 disabled:opacity-40"
            >
              {sending ? 'Sending…' : 'Send'}
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}

function RecipientField({
  label,
  value,
  onChange,
  peopleSearch,
  onSearch,
  trailing,
}: {
  label: string;
  value: MailAddress[];
  onChange: (next: MailAddress[]) => void;
  peopleSearch: MsgraphPluginState['peopleSearch'];
  onSearch: (q: string) => void;
  trailing?: React.ReactNode;
}) {
  const [input, setInput] = useState('');
  const [focus, setFocus] = useState(false);
  const inputRef = useRef<HTMLInputElement>(null);

  const commit = (a: MailAddress) => {
    if (!a.address || value.some((v) => v.address.toLowerCase() === a.address.toLowerCase())) return;
    onChange([...value, a]);
    setInput('');
    onSearch('');
  };
  const commitRaw = () => {
    const addr = input.trim().replace(/[,;]+$/, '');
    if (/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(addr)) commit({ name: null, address: addr });
  };

  const results = focus && input.trim() && peopleSearch?.query === input.trim() ? peopleSearch.results : [];

  return (
    <div className="relative flex items-start gap-2">
      <span className="text-[11px] text-muted-foreground w-10 shrink-0 pt-1.5">{label}</span>
      <div
        className="flex-1 flex flex-wrap items-center gap-1 rounded-md border border-border bg-background px-2 py-1 min-h-[30px]"
        onClick={() => inputRef.current?.focus()}
      >
        {value.map((a, i) => (
          <span
            key={a.address + i}
            className="inline-flex items-center gap-1 rounded-full bg-primary/15 border border-primary/30 px-2 py-0.5 text-[11px]"
          >
            <span>{a.name ?? a.address}</span>
            <button
              type="button"
              onClick={(e) => { e.stopPropagation(); onChange(value.filter((_, j) => j !== i)); }}
              className="opacity-60 hover:opacity-100"
            >×</button>
          </span>
        ))}
        <input
          ref={inputRef}
          value={input}
          onChange={(e) => { setInput(e.target.value); onSearch(e.target.value.trim()); }}
          onFocus={() => setFocus(true)}
          onBlur={() => { setTimeout(() => setFocus(false), 150); commitRaw(); }}
          onKeyDown={(e) => {
            if (e.key === 'Enter' || e.key === ',' || e.key === ';' || e.key === 'Tab') {
              if (results[0]?.email) { e.preventDefault(); commit({ name: results[0].displayName, address: results[0].email }); }
              else if (input.includes('@')) { e.preventDefault(); commitRaw(); }
            } else if (e.key === 'Backspace' && !input && value.length) {
              onChange(value.slice(0, -1));
            }
          }}
          placeholder={value.length === 0 ? 'name or email' : ''}
          className="flex-1 min-w-[120px] bg-transparent border-none outline-none text-xs py-0.5"
        />
      </div>
      {trailing && <div className="pt-1.5 shrink-0">{trailing}</div>}
      {focus && results.length > 0 && (
        <div
          style={{ position: 'absolute', left: 48, right: 0, top: '100%', zIndex: 50, marginTop: 4 }}
          className="rounded-lg border border-border bg-card shadow-xl overflow-hidden"
        >
          {results.slice(0, 6).map((r) => (
            <button
              key={r.id}
              type="button"
              onMouseDown={(e) => { e.preventDefault(); if (r.email) commit({ name: r.displayName, address: r.email }); }}
              className="w-full text-left px-3 py-1.5 text-xs hover:bg-muted flex flex-col"
            >
              <span className="font-medium">{r.displayName}</span>
              <span className="text-[10px] text-muted-foreground">{r.email}</span>
            </button>
          ))}
        </div>
      )}
    </div>
  );
}
