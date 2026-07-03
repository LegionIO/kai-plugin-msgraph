import React, { useEffect, useRef, useState } from 'react';
import type { MailFolder } from '../../shared/types.ts';

type Mode = 'menu' | 'new' | 'rename' | 'confirm-empty' | 'confirm-delete';

export function FolderContextMenu({
  folder,
  x,
  y,
  onAction,
  onClose,
}: {
  folder: MailFolder;
  x: number;
  y: number;
  onAction: (action: string, data?: unknown) => void;
  onClose: () => void;
}) {
  const ref = useRef<HTMLDivElement>(null);
  const [mode, setMode] = useState<Mode>('menu');
  const [name, setName] = useState(folder.displayName);
  const [busy, setBusy] = useState(false);

  useEffect(() => {
    const onDown = (e: MouseEvent) => {
      if (ref.current && !ref.current.contains(e.target as Node)) onClose();
    };
    const onKey = (e: KeyboardEvent) => { if (e.key === 'Escape') onClose(); };
    document.addEventListener('mousedown', onDown);
    document.addEventListener('keydown', onKey);
    return () => {
      document.removeEventListener('mousedown', onDown);
      document.removeEventListener('keydown', onKey);
    };
  }, [onClose]);

  // Keep the menu on-screen.
  const pos = (() => {
    const w = 200, h = 260;
    const vw = typeof window !== 'undefined' ? window.innerWidth : 9999;
    const vh = typeof window !== 'undefined' ? window.innerHeight : 9999;
    return { left: Math.min(x, vw - w - 8), top: Math.min(y, vh - h - 8) };
  })();

  const isWellKnown = !!folder.wellKnownName;
  const isTrashLike = folder.wellKnownName === 'deleteditems' || folder.wellKnownName === 'junkemail';

  const run = (action: string, payload: Record<string, unknown>) => {
    setBusy(true);
    Promise.resolve(onAction(action, payload)).finally(() => { setBusy(false); onClose(); });
  };

  return (
    <div
      ref={ref}
      style={{ position: 'fixed', ...pos, width: 200, zIndex: 70 }}
      className="rounded-lg border border-border bg-card shadow-2xl py-1 text-xs"
      onContextMenu={(e) => e.preventDefault()}
    >
      {mode === 'menu' && (
        <>
          <div className="px-3 py-1.5 text-[10px] text-muted-foreground border-b border-border/50 mb-1">
            {folder.totalItemCount.toLocaleString()} item{folder.totalItemCount === 1 ? '' : 's'}
            {folder.unreadItemCount > 0 && ` · ${folder.unreadItemCount} unread`}
          </div>
          <Item
            disabled={folder.unreadItemCount === 0 || busy}
            onClick={() => run('mark-folder-read', { folderId: folder.id })}
          >
            Mark all as read
          </Item>
          <Item disabled={folder.totalItemCount === 0 || busy} onClick={() => setMode('confirm-empty')}>
            Empty folder…
          </Item>
          <Separator />
          <Item onClick={() => { setName(''); setMode('new'); }}>New subfolder…</Item>
          <Item disabled={isWellKnown} onClick={() => { setName(folder.displayName); setMode('rename'); }}>
            Rename…
          </Item>
          <Item disabled={isWellKnown} danger onClick={() => setMode('confirm-delete')}>
            Delete folder…
          </Item>
          <Separator />
          <Item onClick={() => { onAction('select-folder', { folderId: folder.wellKnownName ?? folder.id }); onClose(); }}>
            Open
          </Item>
        </>
      )}

      {(mode === 'new' || mode === 'rename') && (
        <div className="px-2.5 py-2">
          <div className="text-[10px] text-muted-foreground mb-1.5">
            {mode === 'new' ? `New folder under "${folder.displayName}"` : 'Rename folder'}
          </div>
          <input
            autoFocus
            value={name}
            onChange={(e) => setName(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === 'Enter' && name.trim()) {
                run(mode === 'new' ? 'create-folder' : 'rename-folder',
                    mode === 'new' ? { parentId: folder.id, name } : { folderId: folder.id, name });
              }
            }}
            placeholder="Folder name"
            className="w-full rounded-md border border-border bg-background px-2 py-1 text-xs focus:outline-none focus:ring-1 focus:ring-primary"
          />
          <div className="flex gap-1.5 mt-2">
            <button
              type="button"
              disabled={!name.trim() || busy}
              onClick={() =>
                run(mode === 'new' ? 'create-folder' : 'rename-folder',
                    mode === 'new' ? { parentId: folder.id, name } : { folderId: folder.id, name })
              }
              className="flex-1 px-2 py-1 text-[11px] font-medium bg-primary text-primary-foreground rounded-md hover:bg-primary/90 disabled:opacity-40"
            >
              {busy ? '…' : mode === 'new' ? 'Create' : 'Rename'}
            </button>
            <button type="button" onClick={() => setMode('menu')} className="px-2 py-1 text-[11px] text-muted-foreground bg-muted border border-border rounded-md">
              Cancel
            </button>
          </div>
        </div>
      )}

      {(mode === 'confirm-empty' || mode === 'confirm-delete') && (
        <div className="px-2.5 py-2">
          <div className="text-[11px] text-foreground mb-2">
            {mode === 'confirm-empty'
              ? isTrashLike
                ? `Permanently delete all ${folder.totalItemCount.toLocaleString()} items in "${folder.displayName}"?`
                : `Move all ${folder.totalItemCount.toLocaleString()} items in "${folder.displayName}" to Deleted Items?`
              : `Delete folder "${folder.displayName}"? It will be moved to Deleted Items.`}
          </div>
          <div className="flex gap-1.5">
            <button
              type="button"
              disabled={busy}
              onClick={() =>
                run(mode === 'confirm-empty' ? 'empty-folder' : 'delete-folder', { folderId: folder.id })
              }
              className="flex-1 px-2 py-1 text-[11px] font-medium rounded-md disabled:opacity-40"
              style={{ background: '#ef4444', color: '#fff' }}
            >
              {busy ? '…' : mode === 'confirm-empty' ? 'Empty' : 'Delete'}
            </button>
            <button type="button" onClick={() => setMode('menu')} className="px-2 py-1 text-[11px] text-muted-foreground bg-muted border border-border rounded-md">
              Cancel
            </button>
          </div>
        </div>
      )}
    </div>
  );
}

function Item({
  children,
  onClick,
  disabled,
  danger,
}: {
  children: React.ReactNode;
  onClick: () => void;
  disabled?: boolean;
  danger?: boolean;
}) {
  return (
    <button
      type="button"
      disabled={disabled}
      onClick={onClick}
      className={`w-full text-left px-3 py-1.5 transition-colors ${
        disabled
          ? 'text-muted-foreground/50 cursor-not-allowed'
          : danger
            ? 'text-foreground hover:bg-destructive/15 hover:text-destructive'
            : 'text-foreground hover:bg-muted'
      }`}
    >
      {children}
    </button>
  );
}

function Separator() {
  return <div className="my-1 border-t border-border/50" />;
}
