import React, { useEffect, useMemo, useRef, useState } from 'react';
import type { PluginComponentProps } from '../hooks.ts';
import { usePanelHeight } from '../hooks.ts';
import type { MsgraphPluginState, NormalizedMailSummary, MailFolder } from '../../shared/types.ts';
import { MailBody } from './MailBody.tsx';
import { MailComposeDialog } from './MailComposeDialog.tsx';
import { FolderContextMenu } from './FolderContextMenu.tsx';
import { Avatar } from './Avatar.tsx';

type Props = PluginComponentProps<MsgraphPluginState>;

const FOLDER_ICON: Record<string, React.ReactNode> = {
  inbox: (<><rect x="2" y="4" width="20" height="16" rx="2" /><path d="m2 7 10 6 10-6" /></>),
  drafts: (<><path d="M12 20h9" /><path d="M16.5 3.5a2.121 2.121 0 0 1 3 3L7 19l-4 1 1-4Z" /></>),
  sentitems: (<><path d="m22 2-7 20-4-9-9-4Z" /><path d="M22 2 11 13" /></>),
  archive: (<><rect x="2" y="3" width="20" height="5" rx="1" /><path d="M4 8v11a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8M10 12h4" /></>),
  deleteditems: (<><path d="M3 6h18M8 6V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2m3 0v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6" /></>),
  junkemail: (<><circle cx="12" cy="12" r="10" /><path d="m4.9 4.9 14.2 14.2" /></>),
};
const GENERIC_FOLDER = (<path d="M4 20h16a2 2 0 0 0 2-2V8a2 2 0 0 0-2-2h-7.9a2 2 0 0 1-1.7-.9l-.8-1.2A2 2 0 0 0 8 3H4a2 2 0 0 0-2 2v13a2 2 0 0 0 2 2Z" />);

function fmtWhen(iso: string | null): string {
  if (!iso) return '';
  const d = new Date(iso);
  const now = new Date();
  if (d.toDateString() === now.toDateString()) return d.toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' });
  if (d.getFullYear() === now.getFullYear()) return d.toLocaleDateString([], { month: 'short', day: 'numeric' });
  return d.toLocaleDateString([], { month: 'short', day: 'numeric', year: '2-digit' });
}

function fmtFull(iso: string | null): string {
  if (!iso) return '';
  return new Date(iso).toLocaleString([], { dateStyle: 'medium', timeStyle: 'short' });
}

export function MailPanelView({ pluginState, onAction }: Props) {
  const s = pluginState ?? ({} as MsgraphPluginState);
  const [panelRef, panelHeight] = usePanelHeight();
  const [search, setSearch] = useState('');
  const [ctxMenu, setCtxMenu] = useState<{ folder: MailFolder; x: number; y: number } | null>(null);
  const searchDebounce = useRef<ReturnType<typeof setTimeout> | null>(null);

  const folders = s.mailFolders ?? [];
  const activeFolder = s.activeMailFolder ?? 'inbox';
  const mail = s.activeMail;
  const list = s.mailSearch ? s.mailSearch.results : (s.mailList ?? []);
  const searching = !!s.mailSearch;

  useEffect(() => {
    return () => { if (searchDebounce.current) clearTimeout(searchDebounce.current); };
  }, []);

  const onSearchInput = (v: string) => {
    setSearch(v);
    if (searchDebounce.current) clearTimeout(searchDebounce.current);
    searchDebounce.current = setTimeout(() => {
      if (v.trim()) onAction('search-mail', { query: v.trim() });
      else onAction('clear-mail-search');
    }, 300);
  };

  const meId = s.auth?.objectId ?? null;

  if (!s.auth?.isAuthenticated) {
    return (
      <div ref={panelRef} className="flex flex-col items-center justify-center gap-4 p-10 text-center" style={panelHeight ? { height: panelHeight } : undefined}>
        <div className="text-sm font-medium text-foreground">Outlook</div>
        <p className="text-xs text-muted-foreground max-w-xs">
          Uses the same Microsoft sign-in as the Teams panel.
        </p>
        <button
          type="button"
          onClick={() => onAction('login')}
          className="px-4 py-1.5 text-xs font-medium bg-primary text-primary-foreground rounded-lg hover:bg-primary/90"
        >
          Sign in
        </button>
        {s.auth?.autoLoginStatus && (
          <p className="text-xs text-muted-foreground animate-pulse">{s.auth.autoLoginStatus}</p>
        )}
      </div>
    );
  }

  return (
    <div ref={panelRef} className="flex flex-col text-foreground" style={panelHeight ? { height: panelHeight } : undefined}>
      {/* Top bar */}
      <div className="flex items-center justify-between gap-3 border-b border-border/50 px-3 py-2 shrink-0">
        <div className="flex items-center gap-2">
          <h2 className="text-sm font-semibold text-foreground">Outlook</h2>
          <button
            type="button"
            title="Refresh"
            onClick={() => onAction('refresh-mail')}
            className="p-1 rounded-md text-muted-foreground hover:text-foreground hover:bg-muted"
          >
            <svg className="w-4 h-4" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={2} strokeLinecap="round">
              <path d="M21 12a9 9 0 1 1-3-6.7L21 8" /><path d="M21 3v5h-5" />
            </svg>
          </button>
        </div>
        <div className="flex items-center gap-2.5 min-w-0">
          <div className="text-right min-w-0 hidden sm:block">
            <div className="text-xs font-medium text-foreground truncate">{s.auth?.displayName ?? s.auth?.email}</div>
            {s.auth?.displayName && (
              <div className="text-[10px] text-muted-foreground truncate">{s.auth?.email}</div>
            )}
          </div>
          {meId && (
            <Avatar id={meId} name={s.auth?.displayName ?? s.auth?.email ?? 'Me'} photo={(s.photos ?? {})[meId]} presence={(s.presence ?? {})[meId]} size={8} />
          )}
        </div>
      </div>

      <div className="flex flex-1 min-h-0">
        {/* Left: search+compose header, then folders | message list */}
        <div
          className="flex flex-col border-r border-border/50 min-h-0 shrink-0 overflow-hidden"
          style={{ width: 510, minWidth: 400, maxWidth: 510 }}
        >
          <div className="flex items-center gap-1.5 px-3 pt-3 pb-2 shrink-0">
            <div className="relative flex-1 min-w-0">
              <input
                value={search}
                onChange={(e) => onSearchInput(e.target.value)}
                placeholder="Search mail…"
                style={{ paddingLeft: 26, paddingRight: search ? 24 : 12, paddingTop: 6, paddingBottom: 6 }}
                className="w-full text-xs bg-muted border border-border rounded-lg text-foreground placeholder:text-muted-foreground focus:border-primary transition-colors"
              />
              <svg
                className="w-3.5 h-3.5 absolute left-2 text-muted-foreground"
                style={{ top: '50%', transform: 'translateY(-50%)' }}
                viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={2}
              >
                <circle cx="11" cy="11" r="8" /><path d="m21 21-4.3-4.3" />
              </svg>
              {search && (
                <button
                  type="button"
                  onClick={() => onSearchInput('')}
                  className="absolute right-1.5 text-muted-foreground hover:text-foreground text-xs"
                  style={{ top: '50%', transform: 'translateY(-50%)' }}
                >×</button>
              )}
            </div>
            <button
              type="button"
              title="New mail"
              onClick={() => onAction('compose-mail', { mode: 'new' })}
              className="p-1.5 rounded-lg border border-border bg-muted text-muted-foreground hover:text-foreground hover:border-primary transition-colors shrink-0"
            >
              <svg className="w-4 h-4" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={2} strokeLinecap="round" strokeLinejoin="round">
                <path d="M12 20h9" /><path d="M16.5 3.5a2.121 2.121 0 0 1 3 3L7 19l-4 1 1-4Z" />
              </svg>
            </button>
          </div>
          <div className="flex flex-1 min-h-0">
            {/* Folder tree */}
            <div className="flex flex-col border-r border-border/50 min-h-0 shrink-0" style={{ width: 190 }}>
              <div className="flex-1 overflow-y-auto py-1">
                {folders.map((f) => (
                  <FolderRow
                    key={f.id}
                    f={f}
                    active={!searching && (activeFolder === f.id || activeFolder === f.wellKnownName)}
                    expanded={(s.mailFoldersExpanded ?? []).includes(f.id)}
                    onClick={() => { setSearch(''); onAction('select-folder', { folderId: f.wellKnownName ?? f.id }); }}
                    onToggle={() => onAction('toggle-folder', { folderId: f.id })}
                    onContextMenu={(e) => {
                      e.preventDefault();
                      setCtxMenu({ folder: f, x: e.clientX, y: e.clientY });
                    }}
                  />
                ))}
                {s.loadingMailFolders && folders.length === 0 && (
                  <div className="px-3 py-2 text-[11px] text-muted-foreground">Loading…</div>
                )}
              </div>
            </div>

            {/* Message list */}
            <div className="flex flex-col min-h-0 flex-1 min-w-0 overflow-hidden">
              <div className="flex-1 overflow-y-auto">
            {searching && s.mailSearch?.loading && (
              <div className="px-3 py-3 text-[11px] text-muted-foreground">Searching…</div>
            )}
            {searching && s.mailSearch?.error && (
              <div className="px-3 py-3 text-[11px] text-destructive">{s.mailSearch.error}</div>
            )}
            {list.map((m) => {
              const addr = m.from?.address?.toLowerCase();
              const uid = addr ? (s.mailSenderIds ?? {})[addr] : null;
              return (
                <MailRow
                  key={m.id}
                  m={m}
                  active={s.activeMailId === m.id}
                  senderId={uid ?? null}
                  photo={uid ? (s.photos ?? {})[uid] : undefined}
                  onAction={onAction}
                />
              );
            })}
            {!searching && s.mailListNextLink && (
              <button
                type="button"
                onClick={() => onAction('load-more-mail')}
                disabled={s.loadingMailList}
                className="w-full px-3 py-2 text-[11px] text-muted-foreground hover:text-foreground hover:bg-muted/50 disabled:opacity-50"
              >
                {s.loadingMailList ? 'Loading…' : 'Load more'}
              </button>
            )}
            {list.length === 0 && !s.loadingMailList && !s.mailSearch?.loading && (
              <div className="px-3 py-6 text-center text-[11px] text-muted-foreground">
                {searching ? 'No results' : 'No messages'}
              </div>
            )}
              </div>
            </div>
          </div>
        </div>

        {/* Reading pane */}
        <div className="flex-1 min-w-0 flex flex-col min-h-0">
          {mail ? (
            <ReadingPane
              mail={mail}
              inline={s.mailInlineAttachments ?? {}}
              senderId={mail.from?.address ? (s.mailSenderIds ?? {})[mail.from.address.toLowerCase()] ?? null : null}
              photos={s.photos ?? {}}
              onAction={onAction}
            />
          ) : s.loadingMail ? (
            <div className="flex-1 flex items-center justify-center">
              <div className="w-5 h-5 border-2 border-primary border-t-transparent rounded-full animate-spin" />
            </div>
          ) : (
            <div className="flex-1 flex items-center justify-center text-[11px] text-muted-foreground">
              Select a message
            </div>
          )}
          {s.mailError && (
            <div className="mx-4 mb-2 text-xs text-destructive bg-destructive/10 px-3 py-1.5 rounded-md">
              {s.mailError}
            </div>
          )}
        </div>
      </div>

      {ctxMenu && (
        <FolderContextMenu
          folder={ctxMenu.folder}
          x={ctxMenu.x}
          y={ctxMenu.y}
          onAction={onAction}
          onClose={() => setCtxMenu(null)}
        />
      )}
      {s.composingMail && (
        <MailComposeDialog
          compose={s.composingMail}
          sending={!!s.sendingMail}
          peopleSearch={s.peopleSearch ?? null}
          onAction={onAction}
        />
      )}
    </div>
  );
}

function FolderRow({
  f,
  active,
  expanded,
  onClick,
  onToggle,
  onContextMenu,
}: {
  f: MailFolder;
  active: boolean;
  expanded: boolean;
  onClick: () => void;
  onToggle: () => void;
  onContextMenu: (e: React.MouseEvent) => void;
}) {
  const icon = f.wellKnownName ? FOLDER_ICON[f.wellKnownName] : undefined;
  const hasChildren = f.childFolderCount > 0;
  return (
    <div
      className={`w-full flex items-center gap-1 pr-2 py-1.5 text-left text-xs transition-colors cursor-pointer ${
        active ? 'bg-primary/15 text-foreground font-medium' : 'text-foreground/80 hover:bg-muted'
      }`}
      style={{ paddingLeft: 6 + f.depth * 14 }}
      onClick={onClick}
      onContextMenu={onContextMenu}
    >
      <button
        type="button"
        onClick={(e) => { e.stopPropagation(); if (hasChildren) onToggle(); }}
        className={`w-4 h-4 flex items-center justify-center shrink-0 rounded ${
          hasChildren ? 'text-muted-foreground hover:text-foreground hover:bg-muted' : 'invisible'
        }`}
        aria-label={expanded ? 'Collapse' : 'Expand'}
      >
        <svg
          className="w-3 h-3"
          viewBox="0 0 24 24"
          fill="none"
          stroke="currentColor"
          strokeWidth={2.5}
          strokeLinecap="round"
          strokeLinejoin="round"
          style={{ transform: expanded ? 'rotate(90deg)' : undefined, transition: 'transform 120ms' }}
        >
          <path d="m9 18 6-6-6-6" />
        </svg>
      </button>
      <svg
        className="w-3.5 h-3.5 shrink-0 opacity-70"
        viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={2} strokeLinecap="round" strokeLinejoin="round"
      >
        {icon ?? GENERIC_FOLDER}
      </svg>
      <span className="flex-1 truncate">{f.displayName}</span>
      {f.unreadItemCount > 0 && (
        <span
          style={{ minWidth: 18, height: 16, fontSize: 9, lineHeight: '16px' }}
          className="rounded-full bg-primary text-primary-foreground text-center px-1 font-semibold"
        >
          {f.unreadItemCount > 999 ? '999+' : f.unreadItemCount}
        </span>
      )}
    </div>
  );
}

function MailRow({
  m,
  active,
  senderId,
  photo,
  onAction,
}: {
  m: NormalizedMailSummary;
  active: boolean;
  senderId: string | null;
  photo: string | null | undefined;
  onAction: Props['onAction'];
}) {
  const [hover, setHover] = useState(false);
  const senderName = m.from?.name ?? m.from?.address ?? '?';
  return (
    <div
      onClick={() => onAction('select-mail', { messageId: m.id })}
      onMouseEnter={() => setHover(true)}
      onMouseLeave={() => setHover(false)}
      className={`relative px-3 py-2 border-b border-border/30 cursor-pointer transition-colors overflow-hidden ${
        active ? 'bg-primary/10' : hover ? 'bg-muted/60' : ''
      }`}
    >
      <div className="flex items-start gap-2.5 min-w-0">
        <div className="relative shrink-0" style={{ marginTop: 2 }}>
          <Avatar id={senderId ?? m.from?.address ?? '?'} name={senderName} photo={photo} size={8} />
          {!m.isRead && (
            <span
              className="bg-primary rounded-full absolute"
              style={{ width: 8, height: 8, top: -1, right: -1, boxShadow: '0 0 0 2px var(--card, #111)' }}
            />
          )}
        </div>
        <div className={`flex-1 min-w-0 ${m.isRead ? '' : 'font-semibold'}`}>
          <div className="flex items-baseline justify-between gap-2">
            <div className="text-xs truncate text-foreground">{m.from?.name ?? m.from?.address ?? '(unknown)'}</div>
            <div className="text-[10px] text-muted-foreground shrink-0">{fmtWhen(m.receivedDateTime)}</div>
          </div>
          <div className="text-[11px] truncate text-foreground/90 flex items-center gap-1">
            {m.importance === 'high' && <span style={{ color: '#ef4444' }} title="High importance">!</span>}
            {m.hasAttachments && (
              <svg className="w-3 h-3 shrink-0 opacity-60" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={2}>
                <path d="m21.44 11.05-9.19 9.19a6 6 0 0 1-8.49-8.49l8.57-8.57A4 4 0 1 1 18 8.84l-8.59 8.57a2 2 0 0 1-2.83-2.83l8.49-8.48" />
              </svg>
            )}
            <span className="truncate">{m.subject}</span>
          </div>
          <div className="text-[10px] text-muted-foreground truncate">{m.bodyPreview}</div>
        </div>
        {m.flagged && <span style={{ color: '#ef4444' }} className="shrink-0 text-xs" title="Flagged">⚑</span>}
      </div>
      {hover && (
        <MailRowActions m={m} onAction={onAction} />
      )}
    </div>
  );
}

function MailRowActions({ m, onAction }: { m: NormalizedMailSummary; onAction: Props['onAction'] }) {
  const [label, setLabel] = useState<string | null>(null);
  const btn = (title: string, onClick: () => void, svg: React.ReactNode) => (
    <button
      type="button"
      aria-label={title}
      onClick={onClick}
      onMouseEnter={() => setLabel(title)}
      onMouseLeave={() => setLabel((l) => (l === title ? null : l))}
      className="p-1 rounded text-muted-foreground hover:text-foreground hover:bg-muted"
    >
      <svg className="w-3.5 h-3.5" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={2} strokeLinecap="round" strokeLinejoin="round">
        {svg}
      </svg>
    </button>
  );
  return (
    <div
      style={{ position: 'absolute', right: 4, top: 4 }}
      className="flex flex-col items-end gap-0.5"
      onClick={(e) => e.stopPropagation()}
    >
      <div className="flex gap-0.5 rounded-md border border-border bg-card shadow-sm p-0.5">
        {btn(
          m.isRead ? 'Mark unread' : 'Mark read',
          () => onAction('mark-mail', { messageId: m.id, isRead: !m.isRead }),
          <><rect x="2" y="4" width="20" height="16" rx="2" /><path d="m2 7 10 6 10-6" /></>,
        )}
        {btn(
          m.flagged ? 'Unflag' : 'Flag',
          () => onAction('mark-mail', { messageId: m.id, flag: !m.flagged }),
          <path d="M4 22V4a2 2 0 0 1 2-2h10l4 4v9H6" />,
        )}
        {btn(
          'Archive',
          () => onAction('archive-mail', { messageId: m.id }),
          <><rect x="2" y="3" width="20" height="5" rx="1" /><path d="M4 8v11a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8M10 12h4" /></>,
        )}
        {btn(
          'Delete',
          () => onAction('delete-mail', { messageId: m.id }),
          <path d="M3 6h18M8 6V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2m3 0v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6" />,
        )}
      </div>
      {label && (
        <span
          style={{
            fontSize: 10, lineHeight: '14px', padding: '2px 6px', borderRadius: 4, whiteSpace: 'nowrap',
            background: '#18181b', color: '#fafafa', border: '1px solid rgba(127,127,127,.3)',
            boxShadow: '0 4px 12px rgba(0,0,0,.35)',
          }}
        >
          {label}
        </span>
      )}
    </div>
  );
}

function ReadingPane({
  mail,
  inline,
  senderId,
  photos,
  onAction,
}: {
  mail: NonNullable<MsgraphPluginState['activeMail']>;
  inline: Record<string, string | null>;
  senderId: string | null;
  photos: Record<string, string | null>;
  onAction: Props['onAction'];
}) {
  const nonInline = useMemo(() => mail.attachments.filter((a) => !a.isInline), [mail.attachments]);
  return (
    <div className="flex-1 min-h-0 flex flex-col">
      <div className="px-5 pt-4 pb-3 border-b border-border/40 shrink-0">
        <div className="flex items-start justify-between gap-3">
          <h3 className="text-base font-semibold text-foreground break-words min-w-0">
            {mail.importance === 'high' && <span style={{ color: '#ef4444' }} className="mr-1">!</span>}
            {mail.subject}
          </h3>
          <div className="flex gap-1 shrink-0">
            <ActionBtn title="Reply" onClick={() => onAction('compose-mail', { mode: 'reply' })}>
              <path d="m9 17-5-5 5-5" /><path d="M20 18v-2a4 4 0 0 0-4-4H4" />
            </ActionBtn>
            <ActionBtn title="Reply all" onClick={() => onAction('compose-mail', { mode: 'replyAll' })}>
              <path d="m7 17-5-5 5-5" /><path d="m12 17-5-5 5-5" /><path d="M22 18v-2a4 4 0 0 0-4-4H7" />
            </ActionBtn>
            <ActionBtn title="Forward" onClick={() => onAction('compose-mail', { mode: 'forward' })}>
              <path d="m15 17 5-5-5-5" /><path d="M4 18v-2a4 4 0 0 1 4-4h12" />
            </ActionBtn>
            <ActionBtn title="Archive" onClick={() => onAction('archive-mail', { messageId: mail.id })}>
              <rect x="2" y="3" width="20" height="5" rx="1" /><path d="M4 8v11a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8M10 12h4" />
            </ActionBtn>
            <ActionBtn title="Delete" onClick={() => onAction('delete-mail', { messageId: mail.id })}>
              <path d="M3 6h18M8 6V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2m3 0v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6" />
            </ActionBtn>
            {mail.webLink && (
              <ActionBtn title="Open in Outlook" onClick={() => onAction('open-in-outlook', { url: mail.webLink })}>
                <path d="M7 17 17 7" /><path d="M7 7h10v10" />
              </ActionBtn>
            )}
          </div>
        </div>
        <div className="flex items-center gap-2.5 mt-2">
          <Avatar
            id={senderId ?? mail.from?.address ?? '?'}
            name={mail.from?.name ?? mail.from?.address ?? '?'}
            photo={senderId ? photos[senderId] : undefined}
            size={8}
          />
          <div className="min-w-0 flex-1">
            <div className="text-xs font-medium text-foreground truncate">
              {mail.from?.name ?? mail.from?.address}
              {mail.from?.name && <span className="text-muted-foreground font-normal"> &lt;{mail.from.address}&gt;</span>}
            </div>
            <div className="text-[10px] text-muted-foreground truncate">
              To: {mail.toRecipients.map((r) => r.name ?? r.address).join(', ')}
              {mail.ccRecipients.length > 0 && ` · Cc: ${mail.ccRecipients.map((r) => r.name ?? r.address).join(', ')}`}
            </div>
          </div>
          <div className="text-[10px] text-muted-foreground shrink-0">{fmtFull(mail.receivedDateTime)}</div>
        </div>
        {nonInline.length > 0 && (
          <div className="flex flex-wrap gap-1.5 mt-2.5">
            {nonInline.map((a) => {
              const dl = inline[`dl:${a.id}`];
              return (
                <a
                  key={a.id}
                  href={dl ?? undefined}
                  download={a.name}
                  onClick={(e) => {
                    if (!dl) { e.preventDefault(); onAction('download-attachment', { messageId: mail.id, attachmentId: a.id, name: a.name }); }
                  }}
                  className="inline-flex items-center gap-1.5 rounded-md border border-border bg-muted/40 px-2 py-1 text-[11px] hover:bg-muted"
                  style={{ textDecoration: 'none', color: 'inherit' }}
                >
                  <svg className="w-3.5 h-3.5 opacity-70" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={2}>
                    <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" /><polyline points="14 2 14 8 20 8" />
                  </svg>
                  <span className="max-w-[200px] truncate">{a.name}</span>
                  <span className="opacity-60">{(a.size / 1024).toFixed(0)} KB</span>
                  {!dl && <span className="opacity-60">↓</span>}
                </a>
              );
            })}
          </div>
        )}
      </div>
      <div className="flex-1 overflow-y-auto px-5 py-4">
        <MailBody
          html={mail.bodyHtml}
          inlineAttachments={inline}
          onOpenLink={(url) => onAction('open-external', { url })}
        />
      </div>
    </div>
  );
}

function ActionBtn({ title, onClick, children }: { title: string; onClick: () => void; children: React.ReactNode }) {
  const [hover, setHover] = useState(false);
  return (
    <button
      type="button"
      aria-label={title}
      onClick={onClick}
      onMouseEnter={() => setHover(true)}
      onMouseLeave={() => setHover(false)}
      className="relative p-1.5 rounded-md text-muted-foreground hover:text-foreground hover:bg-muted"
    >
      <svg className="w-4 h-4" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={2} strokeLinecap="round" strokeLinejoin="round">
        {children}
      </svg>
      {hover && (
        <span
          style={{
            position: 'absolute', top: '100%', right: 0, marginTop: 6, whiteSpace: 'nowrap',
            fontSize: 10, lineHeight: '14px', padding: '3px 7px', borderRadius: 5, zIndex: 20,
            background: '#18181b', color: '#fafafa', border: '1px solid rgba(127,127,127,.3)',
            boxShadow: '0 4px 12px rgba(0,0,0,.35)', pointerEvents: 'none',
          }}
        >
          {title}
        </span>
      )}
    </button>
  );
}
