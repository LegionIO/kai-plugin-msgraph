import React, { useMemo, useState, useEffect, useRef, useLayoutEffect } from 'react';
import type { PluginComponentProps } from '../hooks.ts';
import { usePanelHeight } from '../hooks.ts';
import type { MsgraphPluginState, NormalizedChat, NormalizedMessage } from '../../shared/types.ts';
import { MfaDialog } from './MfaDialog.tsx';
import { MfaApprovalDialog } from './MfaApprovalDialog.tsx';
import { Avatar, AvatarStack } from './Avatar.tsx';
import { ReactionPicker } from './ReactionPicker.tsx';

type Props = PluginComponentProps<MsgraphPluginState>;

function fmtTime(iso: string | null): string {
  if (!iso) return '';
  const d = new Date(iso);
  const now = new Date();
  const sameDay = d.toDateString() === now.toDateString();
  return sameDay
    ? d.toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' })
    : d.toLocaleDateString([], { month: 'short', day: 'numeric' });
}

function chatTitle(c: NormalizedChat): string {
  if (c.topic) return c.topic;
  if (c.members.length > 0) return c.members.map((m) => m.displayName).join(', ');
  return c.type === 'group' ? 'Group chat' : 'Chat';
}

export function PanelView({ pluginState, onAction }: Props) {
  const s = pluginState ?? ({} as MsgraphPluginState);
  const photos = s.photos ?? {};
  const [filter, setFilter] = useState('');
  const [draft, setDraft] = useState('');
  const [reactionTarget, setReactionTarget] = useState<string | null>(null);
  const scrollRef = useRef<HTMLDivElement>(null);
  const stickToBottomRef = useRef(true);
  const [panelRef, panelHeight] = usePanelHeight();

  const chats = useMemo(() => {
    const list = s.chats ?? [];
    if (!filter.trim()) return list;
    const f = filter.toLowerCase();
    return list.filter(
      (ch) =>
        chatTitle(ch).toLowerCase().includes(f) ||
        ch.members.some(
          (m) => m.displayName.toLowerCase().includes(f) || (m.email ?? '').toLowerCase().includes(f),
        ),
    );
  }, [s.chats, filter]);

  const activeChat = useMemo(
    () => (s.chats ?? []).find((ch) => ch.id === s.activeChatId) ?? null,
    [s.chats, s.activeChatId],
  );

  const messages = s.activeChatMessages ?? [];

  // Reset stick-to-bottom whenever the active chat changes.
  useEffect(() => {
    stickToBottomRef.current = true;
    setReactionTarget(null);
  }, [s.activeChatId]);

  // Auto-scroll: pin to bottom after layout when new messages arrive or chat/height changes,
  // unless the user has scrolled up.
  useLayoutEffect(() => {
    const el = scrollRef.current;
    if (!el || !stickToBottomRef.current) return;
    el.scrollTop = el.scrollHeight;
    requestAnimationFrame(() => {
      if (scrollRef.current) scrollRef.current.scrollTop = scrollRef.current.scrollHeight;
    });
  }, [messages.length, s.activeChatId, panelHeight]);

  const onThreadScroll = () => {
    const el = scrollRef.current;
    if (!el) return;
    stickToBottomRef.current = el.scrollHeight - el.scrollTop - el.clientHeight < 40;
  };

  const send = () => {
    if (!draft.trim() || !s.activeChatId) return;
    stickToBottomRef.current = true;
    onAction('send-message', { chatId: s.activeChatId, text: draft.trim() });
    setDraft('');
  };

  const react = (messageId: string, reactionType: string, remove = false) => {
    if (!s.activeChatId) return;
    onAction('react-to-message', { chatId: s.activeChatId, messageId, reactionType, remove });
    setReactionTarget(null);
  };

  const mfa = s.mfa ?? { needed: false, type: null, approvalNumber: null };

  if (!s.auth?.isAuthenticated) {
    return (
      <div
        ref={panelRef}
        style={panelHeight ? { height: panelHeight } : undefined}
        className="flex flex-col items-center justify-center gap-3 p-6 text-center overflow-hidden"
      >
        <div className="w-12 h-12 rounded-xl bg-primary/15 flex items-center justify-center">
          <svg className="w-6 h-6 text-primary" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={2}>
            <path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z" />
          </svg>
        </div>
        <div>
          <h3 className="text-base font-semibold text-foreground">Microsoft Teams</h3>
          <p className="text-xs text-muted-foreground max-w-xs mt-1">
            Sign in with your work account to list, read, and send Teams chats from Kai.
          </p>
        </div>
        {s.auth?.autoLoginStatus && (
          <p className="text-xs text-muted-foreground animate-pulse">{s.auth.autoLoginStatus}</p>
        )}
        {s.error && (
          <p className="text-xs text-destructive bg-destructive/10 px-3 py-1.5 rounded-md">{s.error}</p>
        )}
        <button
          type="button"
          onClick={() => onAction('login')}
          className="px-4 py-2 text-sm font-medium bg-primary text-primary-foreground rounded-lg hover:bg-primary/90 transition-colors"
        >
          Sign in
        </button>

        {mfa.needed && mfa.type === 'push' && <MfaApprovalDialog approvalNumber={mfa.approvalNumber} />}
        {mfa.needed && (mfa.type === 'sms' || mfa.type === 'totp') && (
          <MfaDialog
            type={mfa.type}
            onSubmit={(code) => onAction('submit-mfa-code', { code })}
            onCancel={() => onAction('cancel-mfa')}
          />
        )}
      </div>
    );
  }

  return (
    <div
      ref={panelRef}
      style={panelHeight ? { height: panelHeight } : undefined}
      className="flex min-h-0 overflow-hidden"
    >
      {/* Sidebar */}
      <div className="flex w-[280px] min-w-[220px] flex-col border-r border-border/50 min-h-0">
        <div className="flex items-center justify-between gap-2 px-3 pt-3 pb-2 shrink-0">
          <div className="flex items-baseline gap-2 min-w-0">
            <h2 className="text-sm font-semibold text-foreground">Teams</h2>
            <span className="text-[10px] text-muted-foreground truncate">{s.auth?.email}</span>
          </div>
          <div className="flex items-center gap-1">
            <button
              type="button"
              title="Refresh"
              onClick={() => onAction('refresh-chats')}
              className="p-1 rounded-md text-muted-foreground hover:text-foreground hover:bg-muted transition-colors"
            >
              <svg className="w-4 h-4" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={2}>
                <path d="M21 12a9 9 0 1 1-3-6.7L21 8" /><path d="M21 3v5h-5" />
              </svg>
            </button>
            <button
              type="button"
              title="Sign out"
              onClick={() => onAction('logout')}
              className="p-1 rounded-md text-muted-foreground hover:text-foreground hover:bg-muted transition-colors"
            >
              <svg className="w-4 h-4" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={2}>
                <path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4" /><path d="m16 17 5-5-5-5" /><path d="M21 12H9" />
              </svg>
            </button>
          </div>
        </div>
        <div className="px-3 pb-2 shrink-0">
          <input
            className="w-full px-2.5 py-1.5 text-xs bg-muted border border-border rounded-lg text-foreground placeholder:text-muted-foreground focus:border-primary transition-colors"
            placeholder="Filter chats…"
            value={filter}
            onChange={(e) => setFilter(e.target.value)}
          />
        </div>
        <div className="flex-1 min-h-0 overflow-y-auto overscroll-contain px-1.5 pb-2">
          {s.loadingChats && chats.length === 0 && (
            <div className="p-3 text-xs text-muted-foreground">Loading…</div>
          )}
          {chats.map((ch) => (
            <ChatRow
              key={ch.id}
              chat={ch}
              photos={photos}
              active={ch.id === s.activeChatId}
              onClick={() => onAction('select-chat', { chatId: ch.id })}
            />
          ))}
          {!s.loadingChats && chats.length === 0 && (
            <div className="p-3 text-xs text-muted-foreground">No chats</div>
          )}
        </div>
      </div>

      {/* Thread */}
      <div className="flex flex-1 min-w-0 min-h-0 flex-col">
        {!activeChat ? (
          <div className="m-auto text-xs text-muted-foreground">Select a chat</div>
        ) : (
          <>
            <div className="flex items-center justify-between gap-3 border-b border-border/50 px-4 py-2.5 shrink-0">
              <div className="flex items-center gap-3 min-w-0">
                {activeChat.type === 'oneOnOne' && activeChat.members[0] ? (
                  <Avatar
                    id={activeChat.members[0].id}
                    name={activeChat.members[0].displayName}
                    photo={photos[activeChat.members[0].id]}
                    size={9}
                  />
                ) : (
                  <AvatarStack members={activeChat.members} photos={photos} />
                )}
                <div className="min-w-0">
                  <div className="text-sm font-semibold text-foreground truncate">{chatTitle(activeChat)}</div>
                  <div className="text-[10px] text-muted-foreground truncate">
                    {activeChat.type} · {activeChat.members.length} member{activeChat.members.length === 1 ? '' : 's'}
                  </div>
                </div>
              </div>
              {activeChat.webUrl && (
                <button
                  type="button"
                  title="Open in Teams"
                  onClick={() => onAction('open-in-teams', { url: activeChat.webUrl })}
                  className="p-1.5 rounded-md text-muted-foreground hover:text-foreground hover:bg-muted transition-colors"
                >
                  <svg className="w-4 h-4" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={2}>
                    <path d="M7 17 17 7" /><path d="M7 7h10v10" />
                  </svg>
                </button>
              )}
            </div>

            <div
              ref={scrollRef}
              onScroll={onThreadScroll}
              className="flex-1 min-h-0 overflow-y-auto overscroll-contain p-4 flex flex-col gap-2"
            >
              {s.loadingMessages && messages.length === 0 && (
                <div className="text-xs text-muted-foreground">Loading messages…</div>
              )}
              {messages.map((m, i) => (
                <MessageBubble
                  key={m.id}
                  m={m}
                  photos={photos}
                  showHeader={i === 0 || messages[i - 1].fromId !== m.fromId}
                  pickerOpen={reactionTarget === m.id}
                  onOpenPicker={() => setReactionTarget(m.id)}
                  onClosePicker={() => setReactionTarget(null)}
                  onReact={(type, remove) => react(m.id, type, remove)}
                />
              ))}
            </div>

            {s.error && (
              <div className="mx-4 mb-2 text-xs text-destructive bg-destructive/10 px-3 py-1.5 rounded-md">
                {s.error}
              </div>
            )}

            <div className="flex items-end gap-2 border-t border-border/50 p-3 shrink-0">
              <textarea
                className="flex-1 resize-none px-3 py-2 text-sm bg-muted border border-border rounded-lg text-foreground placeholder:text-muted-foreground focus:border-primary transition-colors max-h-32"
                rows={1}
                placeholder="Message…"
                value={draft}
                onChange={(e) => setDraft(e.target.value)}
                onKeyDown={(e) => {
                  if (e.key === 'Enter' && !e.shiftKey) {
                    e.preventDefault();
                    send();
                  }
                }}
              />
              <button
                type="button"
                disabled={!draft.trim() || s.sendingMessage}
                onClick={send}
                className="px-3.5 py-2 text-sm font-medium bg-primary text-primary-foreground rounded-lg hover:bg-primary/90 disabled:opacity-50 transition-colors"
              >
                {s.sendingMessage ? '…' : 'Send'}
              </button>
            </div>
          </>
        )}
      </div>

      {mfa.needed && mfa.type === 'push' && <MfaApprovalDialog approvalNumber={mfa.approvalNumber} />}
      {mfa.needed && (mfa.type === 'sms' || mfa.type === 'totp') && (
        <MfaDialog
          type={mfa.type}
          onSubmit={(code) => onAction('submit-mfa-code', { code })}
          onCancel={() => onAction('cancel-mfa')}
        />
      )}
    </div>
  );
}

function ChatRow({
  chat,
  photos,
  active,
  onClick,
}: {
  chat: NormalizedChat;
  photos: Record<string, string | null>;
  active: boolean;
  onClick: () => void;
}) {
  return (
    <div
      role="button"
      tabIndex={0}
      onClick={onClick}
      className={`flex items-center gap-2.5 px-2 py-2 rounded-lg cursor-pointer transition-colors ${
        active ? 'bg-primary/10' : 'hover:bg-muted/60'
      }`}
    >
      {chat.type === 'oneOnOne' && chat.members[0] ? (
        <Avatar id={chat.members[0].id} name={chat.members[0].displayName} photo={photos[chat.members[0].id]} />
      ) : (
        <AvatarStack members={chat.members} photos={photos} max={2} />
      )}
      <div className="min-w-0 flex-1">
        <div className="flex items-baseline justify-between gap-2">
          <div className="text-xs font-medium text-foreground truncate">{chatTitle(chat)}</div>
          <div className="text-[10px] text-muted-foreground shrink-0">{fmtTime(chat.lastUpdated)}</div>
        </div>
        <div className="text-[11px] text-muted-foreground truncate">
          {chat.lastMessageFrom ? `${chat.lastMessageFrom}: ` : ''}
          {chat.lastMessagePreview ?? ''}
        </div>
      </div>
    </div>
  );
}

function MessageBubble({
  m,
  photos,
  showHeader,
  pickerOpen,
  onOpenPicker,
  onClosePicker,
  onReact,
}: {
  m: NormalizedMessage;
  photos: Record<string, string | null>;
  showHeader: boolean;
  pickerOpen: boolean;
  onOpenPicker: () => void;
  onClosePicker: () => void;
  onReact: (type: string, remove?: boolean) => void;
}) {
  const [hover, setHover] = useState(false);

  return (
    <div
      className={`group flex gap-2 ${m.fromMe ? 'flex-row-reverse' : ''}`}
      onMouseEnter={() => setHover(true)}
      onMouseLeave={() => setHover(false)}
    >
      <div className="w-7 shrink-0">
        {showHeader && !m.fromMe && m.fromId && (
          <Avatar id={m.fromId} name={m.fromName ?? '?'} photo={photos[m.fromId]} size={7} />
        )}
      </div>
      <div className={`relative flex flex-col max-w-[75%] ${m.fromMe ? 'items-end' : 'items-start'}`}>
        {showHeader && !m.fromMe && m.fromName && (
          <div className="text-[10px] font-medium text-muted-foreground px-1 mb-0.5">{m.fromName}</div>
        )}
        <div
          onContextMenu={(e) => {
            e.preventDefault();
            onOpenPicker();
          }}
          className={`relative px-3 py-1.5 rounded-2xl text-sm whitespace-pre-wrap break-words ${
            m.fromMe
              ? 'bg-primary text-primary-foreground rounded-br-md'
              : 'bg-muted text-foreground rounded-bl-md'
          }`}
        >
          {m.text || <span className="opacity-50 italic">(no text)</span>}
          {m.attachments.length > 0 && (
            <div className="mt-1 space-y-0.5">
              {m.attachments.map((a, i) => (
                <div key={i} className="text-[10px] opacity-80">📎 {a.name ?? a.contentType ?? 'attachment'}</div>
              ))}
            </div>
          )}
          {m.reactions.length > 0 && (
            <div
              className={`absolute -bottom-2.5 ${m.fromMe ? 'left-1' : 'right-1'} flex gap-0.5`}
            >
              {m.reactions.map((r) => (
                <span
                  key={r.type}
                  title={r.users.join(', ')}
                  className="inline-flex items-center gap-0.5 px-1.5 py-0.5 rounded-full bg-card border border-border shadow-sm text-[10px] leading-none"
                >
                  <span>{r.emoji}</span>
                  {r.count > 1 && <span className="text-muted-foreground">{r.count}</span>}
                </span>
              ))}
            </div>
          )}
          {pickerOpen && (
            <ReactionPicker
              align={m.fromMe ? 'right' : 'left'}
              onSelect={(t) => onReact(t)}
              onClose={onClosePicker}
            />
          )}
        </div>
        <div className={`text-[9px] text-muted-foreground px-1 ${m.reactions.length > 0 ? 'mt-3' : 'mt-0.5'}`}>
          {fmtTime(m.createdDateTime)}
        </div>
      </div>
      {/* Hover affordance to open picker */}
      <div className={`self-center shrink-0 transition-opacity ${hover || pickerOpen ? 'opacity-100' : 'opacity-0'}`}>
        <button
          type="button"
          title="React"
          onClick={() => (pickerOpen ? onClosePicker() : onOpenPicker())}
          className="w-6 h-6 flex items-center justify-center rounded-full text-muted-foreground hover:text-foreground hover:bg-muted transition-colors"
        >
          <svg className="w-3.5 h-3.5" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={2}>
            <circle cx="12" cy="12" r="10" />
            <path d="M8 14s1.5 2 4 2 4-2 4-2" />
            <line x1="9" y1="9" x2="9.01" y2="9" />
            <line x1="15" y1="9" x2="15.01" y2="9" />
          </svg>
        </button>
      </div>
    </div>
  );
}
