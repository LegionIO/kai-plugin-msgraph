import React, { useState } from 'react';
import type { PluginComponentProps } from '../hooks.ts';
import type { MsgraphPluginState, ToolPermissions, UserPreferences } from '../../shared/types.ts';
import { DEFAULT_TOOL_PERMISSIONS } from '../../shared/types.ts';

type Config = {
  preferences?: Partial<UserPreferences>;
  toolPermissions?: Partial<ToolPermissions>;
};

type Props = PluginComponentProps<MsgraphPluginState, Config>;

const TOOL_LABELS: Record<keyof ToolPermissions, string> = {
  authStatus: 'auth-status — check sign-in state',
  findUser: 'find-user — search directory',
  listChats: 'list-chats — list DMs & group chats',
  getChatMessages: 'get-chat-messages — read a chat',
  searchMessages: 'search-messages — full-text search',
  sendMessage: 'send-message — send to a chat by id',
  sendDm: 'send-dm — resolve person + send DM',
  reactToMessage: 'react-to-message — add/remove a reaction',
  editMessage: 'edit-message — edit one of your own messages',
  deleteMessage: 'delete-message — soft-delete one of your own messages',
  forwardMessage: 'forward-message — forward a message to a person or chat',
  markChatRead: 'mark-chat-read — clear a chat\'s unread indicator',
  getPresence: 'get-presence — availability/activity/status for users',
  setPresence: 'set-presence — change your own availability (visible to others)',
  setStatusMessage: 'set-status-message — set/clear your Teams status note',
  invokeCardAction: 'invoke-card-action — click a bot card button (posts on your behalf)',
  listFolders: 'list-folders — list mailbox folders & counts',
  listMail: 'list-mail — list messages in a folder',
  getMail: 'get-mail — read a single message body',
  searchMail: 'search-mail — full-text mailbox search',
  sendMail: 'send-mail — send a new email as you',
  replyToMail: 'reply-to-mail — reply/reply-all/forward a message',
  markMail: 'mark-mail — mark read/unread/flag',
  archiveMail: 'archive-mail — move to Archive',
  deleteMail: 'delete-mail — move to Deleted Items',
  createGroupChat: 'create-group-chat — new group + optional first message',
};

const SIG_EDITOR_CSS = `
.msg-sig-editor img[data-sig-unresolved] {
  min-width: 24px; min-height: 24px;
  outline: 2px dashed #f97316; outline-offset: 1px;
  background: #fff7ed
    url("data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='%23f97316' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><rect x='3' y='3' width='18' height='18' rx='2'/><circle cx='9' cy='9' r='2'/><path d='m21 15-5-5L5 21'/></svg>")
    center/18px no-repeat;
  cursor: pointer;
}
.msg-sig-editor img[data-sig-unresolved]:hover { outline-color: #ea580c; background-color: #ffedd5; }
`;

function SignatureEditor({ value, onChange }: { value: string; onChange: (html: string) => void }) {
  const ref = React.useRef<HTMLDivElement>(null);
  const fileRef = React.useRef<HTMLInputElement>(null);
  const targetImgRef = React.useRef<HTMLImageElement | null>(null);
  const [showRaw, setShowRaw] = React.useState(false);
  const [raw, setRaw] = React.useState(value);
  const [missing, setMissing] = React.useState(0);
  React.useEffect(() => { setRaw(value); if (ref.current && !showRaw) { ref.current.innerHTML = value; markUnresolved(ref.current); } }, [value]);
  React.useEffect(() => { if (ref.current && !showRaw) { ref.current.innerHTML = raw; markUnresolved(ref.current); } }, [showRaw]);

  const blobToDataUrl = (b: Blob): Promise<string> =>
    new Promise((res, rej) => {
      const fr = new FileReader();
      fr.onload = () => res(String(fr.result));
      fr.onerror = () => rej(fr.error);
      fr.readAsDataURL(b);
    });

  const markUnresolved = (root: HTMLElement) => {
    let n = 0;
    for (const img of Array.from(root.querySelectorAll<HTMLImageElement>('img'))) {
      const src = img.getAttribute('src') ?? '';
      if (src.startsWith('data:')) img.removeAttribute('data-sig-unresolved');
      else { img.setAttribute('data-sig-unresolved', ''); n++; }
    }
    setMissing(n);
  };

  const serialize = (root: HTMLElement) => {
    const clone = root.cloneNode(true) as HTMLElement;
    clone.querySelectorAll('[data-sig-unresolved]').forEach((el) => el.removeAttribute('data-sig-unresolved'));
    return clone.innerHTML.trim();
  };

  const resolveImages = async (root: HTMLElement, clipboardImages: string[]) => {
    const imgs = Array.from(root.querySelectorAll<HTMLImageElement>('img'));
    let ci = 0;
    await Promise.all(
      imgs.map(async (img) => {
        const src = img.getAttribute('src') ?? '';
        if (src.startsWith('data:')) return;
        // Office-origin HTML references images via file:///…/clip_imageNNN, cid:, or
        // auth-gated attachment URLs — the real bitmaps ride alongside in
        // clipboardData.items. Consume those in document order.
        if (!/^https?:/i.test(src)) {
          const d = clipboardImages[ci++];
          if (d) img.setAttribute('src', d);
          return;
        }
        try {
          const r = await fetch(src);
          if (!r.ok) throw new Error(String(r.status));
          img.setAttribute('src', await blobToDataUrl(await r.blob()));
        } catch {
          const d = clipboardImages[ci++];
          if (d) img.setAttribute('src', d);
        }
      }),
    );
  };

  const handlePaste = async (e: React.ClipboardEvent<HTMLDivElement>) => {
    const cd = e.clipboardData;
    const html = cd.getData('text/html');
    const text = cd.getData('text/plain');
    const imageFiles: File[] = [];
    for (const item of Array.from(cd.items)) {
      if (item.kind === 'file' && item.type.startsWith('image/')) {
        const f = item.getAsFile();
        if (f) imageFiles.push(f);
      }
    }
    if (!html && !text && !imageFiles.length) return;
    e.preventDefault();
    const clipboardImages = await Promise.all(imageFiles.map(blobToDataUrl));
    const el = ref.current;
    if (!el) return;
    const content = html
      ? html.replace(/<!--\[if[\s\S]*?endif\]-->/gi, '').replace(/<\/?o:p[^>]*>/gi, '')
      : text
        ? text.replace(/\n/g, '<br>')
        : `<img src="${clipboardImages[0]}">`;
    const isEmpty = !el.textContent?.trim() && !el.querySelector('img,table');
    if (isEmpty) el.innerHTML = content;
    else document.execCommand('insertHTML', false, content);
    trimEmptyEdges(el);
    await resolveImages(el, clipboardImages);
    markUnresolved(el);
    const out = serialize(el);
    setRaw(out);
    onChange(out);
  };

  const trimEmptyEdges = (root: HTMLElement) => {
    const noText = (n: Node) => !(n.textContent ?? '').replace(/[\s ​]/g, '');
    const isBlank = (n: ChildNode) =>
      (n.nodeType === Node.TEXT_NODE && noText(n)) ||
      (n.nodeType === Node.ELEMENT_NODE && /^(P|DIV|BR|SPAN|HEAD|META)$/.test((n as Element).tagName)
        && noText(n) && !(n as Element).querySelector('img,table,hr'));
    while (root.firstChild && isBlank(root.firstChild)) root.removeChild(root.firstChild);
    while (root.lastChild && isBlank(root.lastChild)) root.removeChild(root.lastChild);
  };

  const applyFileToImg = async (img: HTMLImageElement, file: File) => {
    img.setAttribute('src', await blobToDataUrl(file));
    img.removeAttribute('data-sig-unresolved');
    if (!img.getAttribute('alt')) img.setAttribute('alt', file.name);
    const el = ref.current; if (!el) return;
    markUnresolved(el);
    const out = serialize(el);
    setRaw(out); onChange(out);
  };

  const fillUnresolved = async (files: File[]) => {
    const el = ref.current; if (!el || !files.length) return;
    const targets = Array.from(el.querySelectorAll<HTMLImageElement>('img[data-sig-unresolved]'));
    for (let i = 0; i < files.length; i++) {
      const img = targets[i];
      const url = await blobToDataUrl(files[i]);
      if (img) { img.setAttribute('src', url); img.removeAttribute('data-sig-unresolved'); }
      else el.insertAdjacentHTML('beforeend', `<img src="${url}" alt="${files[i].name}">`);
    }
    markUnresolved(el);
    const out = serialize(el);
    setRaw(out); onChange(out);
  };

  const commit = async () => {
    const el = ref.current;
    if (!el) return;
    await resolveImages(el, []);
    markUnresolved(el);
    const html = serialize(el);
    setRaw(html);
    onChange(html);
  };

  return (
    <div className="rounded-lg border border-border overflow-hidden">
      <style>{SIG_EDITOR_CSS}</style>
      <input
        ref={fileRef}
        type="file"
        accept="image/*"
        multiple
        hidden
        onChange={(e) => {
          const files = Array.from(e.target.files ?? []).filter((f) => f.type.startsWith('image/'));
          if (targetImgRef.current && files[0]) void applyFileToImg(targetImgRef.current, files[0]);
          else void fillUnresolved(files);
          targetImgRef.current = null;
          e.target.value = '';
        }}
      />
      <div className="flex items-center justify-between px-2 py-1 border-b border-border bg-muted/30">
        <span className="text-[10px] text-muted-foreground">
          {showRaw
            ? 'Raw HTML'
            : missing > 0
              ? `${missing} image${missing === 1 ? '' : 's'} missing — click a dashed box (or drop a file on it) to set the image`
              : 'Paste your signature here — formatting is kept'}
        </span>
        <div className="flex gap-1">
          <button
            type="button"
            onClick={() => setShowRaw((v) => !v)}
            className="px-1.5 py-0.5 text-[10px] rounded border border-border hover:bg-muted"
          >
            {showRaw ? 'Preview' : 'HTML'}
          </button>
          <button
            type="button"
            onClick={() => { setRaw(''); onChange(''); if (ref.current) ref.current.innerHTML = ''; }}
            className="px-1.5 py-0.5 text-[10px] rounded border border-border hover:bg-muted"
          >
            Clear
          </button>
        </div>
      </div>
      {showRaw ? (
        <textarea
          value={raw}
          onChange={(e) => setRaw(e.target.value)}
          onBlur={() => onChange(raw)}
          rows={6}
          style={{ fontFamily: 'ui-monospace,monospace', resize: 'vertical' }}
          className="w-full bg-background px-2 py-1.5 text-[11px] border-0 focus:outline-none"
        />
      ) : (
        <div
          ref={ref}
          contentEditable
          suppressContentEditableWarning
          onBlur={() => void commit()}
          onPaste={(e) => void handlePaste(e)}
          onClick={(e) => {
            const t = e.target as HTMLElement;
            if (t.tagName === 'IMG' && t.hasAttribute('data-sig-unresolved')) {
              e.preventDefault();
              targetImgRef.current = t as HTMLImageElement;
              fileRef.current?.click();
            }
          }}
          onDragOver={(e) => { if (Array.from(e.dataTransfer.types).includes('Files')) e.preventDefault(); }}
          onDrop={(e) => {
            const files = Array.from(e.dataTransfer.files).filter((f) => f.type.startsWith('image/'));
            if (!files.length) return;
            e.preventDefault();
            const t = e.target as HTMLElement;
            const img = t.tagName === 'IMG' ? (t as HTMLImageElement) : t.closest('img');
            if (img) void applyFileToImg(img as HTMLImageElement, files[0]);
            else void fillUnresolved(files);
          }}
          style={{
            minHeight: 90, maxHeight: 280, overflowY: 'auto',
            background: '#fff', color: '#1f2937', colorScheme: 'light',
            fontFamily: 'Calibri, Arial, sans-serif', fontSize: 14.7, lineHeight: 'normal',
          }}
          className="msg-sig-editor px-3 py-2 focus:outline-none"
        />
      )}
    </div>
  );
}

export function SettingsView({ pluginState, pluginConfig, onAction }: Props) {
  const s = pluginState ?? ({} as MsgraphPluginState);
  const cfg = pluginConfig ?? {};
  const prefs = cfg.preferences ?? {};
  const perms = { ...DEFAULT_TOOL_PERMISSIONS, ...(cfg.toolPermissions ?? {}) };

  const [username, setUsername] = useState(s.credentials?.username ?? '');
  const [password, setPassword] = useState('');

  return (
    <div className="flex flex-col gap-6 p-1 text-sm text-foreground">
      <Section title="Account">
        {s.auth?.isAuthenticated ? (
          <div className="flex items-center justify-between gap-3">
            <div>
              <div className="font-medium">{s.auth.displayName ?? s.auth.email}</div>
              <div className="text-xs text-muted-foreground">
                {s.auth.email} · access token{' '}
                {s.auth.minutesRemaining != null ? `${s.auth.minutesRemaining}m remaining` : 'expired'} · auto-refresh
                enabled
              </div>
            </div>
            <button
              type="button"
              onClick={() => onAction('logout')}
              className="px-3 py-1.5 text-xs font-medium text-muted-foreground bg-muted border border-border rounded-lg hover:bg-muted/80 transition-colors"
            >
              Sign out
            </button>
          </div>
        ) : (
          <div className="flex items-center justify-between gap-3">
            <span className="text-xs text-muted-foreground">Not signed in</span>
            <button
              type="button"
              onClick={() => onAction('login')}
              className="px-3 py-1.5 text-xs font-medium bg-primary text-primary-foreground rounded-lg hover:bg-primary/90 transition-colors"
            >
              Sign in
            </button>
          </div>
        )}
      </Section>

      <Section title="Auto-login credentials">
        <p className="text-xs text-muted-foreground">
          Optional. When saved, the sign-in window is filled automatically (including MFA prompts). Stored via{' '}
          {s.credentials?.encryptionMethod === 'os-keychain' ? 'OS keychain' : 'AES-256-GCM (local key)'}.
        </p>
        <div className="flex gap-2 mt-3">
          <input
            className="flex-1 px-2.5 py-1.5 text-xs bg-muted border border-border rounded-lg text-foreground placeholder:text-muted-foreground focus:border-primary transition-colors"
            placeholder="user@company.com"
            value={username}
            onChange={(e) => setUsername(e.target.value)}
          />
          <input
            className="flex-1 px-2.5 py-1.5 text-xs bg-muted border border-border rounded-lg text-foreground placeholder:text-muted-foreground focus:border-primary transition-colors"
            type="password"
            placeholder="Password"
            value={password}
            onChange={(e) => setPassword(e.target.value)}
          />
          <button
            type="button"
            disabled={!username.trim() || !password}
            onClick={() => {
              onAction('save-credentials', { username: username.trim(), password });
              setPassword('');
            }}
            className="px-3 py-1.5 text-xs font-medium bg-primary text-primary-foreground rounded-lg hover:bg-primary/90 disabled:opacity-50 transition-colors"
          >
            Save
          </button>
          {s.credentials?.hasCredentials && (
            <button
              type="button"
              onClick={() => onAction('clear-credentials')}
              className="px-3 py-1.5 text-xs font-medium text-muted-foreground bg-muted border border-border rounded-lg hover:bg-muted/80 transition-colors"
            >
              Clear
            </button>
          )}
        </div>
        {s.credentials?.hasCredentials && (
          <p className="text-xs text-muted-foreground mt-2">Saved for {s.credentials.username}</p>
        )}
      </Section>

      <Section title="AI tools">
        <p className="text-xs text-muted-foreground">Enable or disable individual tools exposed to the AI.</p>
        <div className="mt-3 flex flex-col gap-2">
          {(Object.keys(TOOL_LABELS) as Array<keyof ToolPermissions>).map((key) => (
            <label key={key} className="flex items-center gap-2 text-xs">
              <input
                type="checkbox"
                checked={perms[key]}
                onChange={(e) => onAction('set-tool-permission', { key, value: e.target.checked })}
                className="accent-primary"
              />
              <span>{TOOL_LABELS[key]}</span>
            </label>
          ))}
        </div>
      </Section>

      <Section title="Cache">
        <p className="text-xs text-muted-foreground">
          Profile photos, bot icons, and people-search results are cached on disk at
          <code className="mx-1">~/.kai/plugin-caches/msgraph/</code> and refreshed in the
          background when older than 24h.
        </p>
        <button
          type="button"
          onClick={() => onAction('clear-cache')}
          className="mt-3 px-3 py-1.5 text-xs font-medium text-muted-foreground bg-muted border border-border rounded-lg hover:bg-muted/80 transition-colors"
        >
          Clear cache
        </button>
      </Section>

      <Section title="Preferences">
        <label className="flex items-center gap-2 text-xs">
          <input
            type="checkbox"
            checked={prefs.notifications ?? true}
            onChange={(e) => onAction('set-preference', { key: 'notifications', value: e.target.checked })}
            className="accent-primary"
          />
          <span>Desktop notifications</span>
        </label>
        <label className="flex items-center gap-2 text-xs mt-3">
          <span>Chat refresh interval (s)</span>
          <input
            className="w-20 px-2 py-1 text-xs bg-muted border border-border rounded-lg text-foreground focus:border-primary transition-colors"
            type="number"
            min={15}
            value={prefs.pollIntervalSeconds ?? 30}
            onChange={(e) =>
              onAction('set-preference', { key: 'pollIntervalSeconds', value: Number(e.target.value) || 30 })
            }
          />
        </label>
        <label className="flex items-center gap-2 text-xs mt-3">
          <input
            type="checkbox"
            checked={prefs.debugLogging ?? false}
            onChange={(e) => onAction('set-preference', { key: 'debugLogging', value: e.target.checked })}
            className="accent-primary"
          />
          <span>Debug logging</span>
        </label>
      </Section>

      <Section title="Mail signature">
        <p className="text-[11px] text-muted-foreground mb-2">
          {s.mailSignature?.source === 'owa'
            ? 'Loaded from your Outlook settings. You can override it below.'
            : s.mailSignature?.source === 'config'
              ? 'Using the signature below.'
              : 'No signature found in your Outlook settings. Paste yours into the box below (rich formatting from Outlook/Word is preserved).'}
        </p>
        <SignatureEditor
          value={prefs.mailSignatureHtml ?? s.mailSignature?.html ?? ''}
          onChange={(html) => onAction('set-preference', { key: 'mailSignatureHtml', value: html })}
        />
        <div className="flex gap-4 mt-2">
          <label className="flex items-center gap-1.5 text-[11px]">
            <input
              type="checkbox"
              checked={prefs.mailSignatureAutoNew ?? true}
              onChange={(e) => onAction('set-preference', { key: 'mailSignatureAutoNew', value: e.target.checked })}
              className="accent-primary"
            />
            Add to new messages
          </label>
          <label className="flex items-center gap-1.5 text-[11px]">
            <input
              type="checkbox"
              checked={prefs.mailSignatureAutoReply ?? true}
              onChange={(e) => onAction('set-preference', { key: 'mailSignatureAutoReply', value: e.target.checked })}
              className="accent-primary"
            />
            Add to replies/forwards
          </label>
        </div>
      </Section>
    </div>
  );
}

function Section({ title, children }: { title: string; children: React.ReactNode }) {
  return (
    <section className="rounded-xl border border-border p-4">
      <h3 className="text-sm font-semibold text-foreground mb-2">{title}</h3>
      {children}
    </section>
  );
}
