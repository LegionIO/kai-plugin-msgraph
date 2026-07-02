import React, { useMemo, useState, useEffect, useRef, useLayoutEffect } from 'react';
import type { PluginComponentProps } from '../hooks.ts';
import { usePanelHeight } from '../hooks.ts';
import type { MsgraphPluginState, NormalizedChat, NormalizedMessage, Presence, BodySegment } from '../../shared/types.ts';
import { MfaDialog } from './MfaDialog.tsx';
import { MfaApprovalDialog } from './MfaApprovalDialog.tsx';
import { Avatar, AvatarStack } from './Avatar.tsx';
import { ReactionPicker } from './ReactionPicker.tsx';
import { SelfMenu } from './SelfMenu.tsx';
import { NewChatDialog } from './NewChatDialog.tsx';
import { UserCard } from './UserCard.tsx';
import { RichComposer } from '../editor/RichComposer.tsx';
import { highlight } from '../highlight.ts';
import { ensureSyntaxThemeInjected } from '../editor/syntax-theme.ts';
import { AdaptiveCard } from './AdaptiveCard.tsx';

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
  useEffect(ensureSyntaxThemeInjected, []);
  const s = pluginState ?? ({} as MsgraphPluginState);
  const photos = s.photos ?? {};
  const presence = s.presence ?? {};
  const hostedContents = s.hostedContents ?? {};
  const [filter, setFilter] = useState('');
  const [searchMode, setSearchMode] = useState<'people' | 'content'>('people');
  const [unreadOnly, setUnreadOnly] = useState(false);
  const [reactionTarget, setReactionTarget] = useState<string | null>(null);
  const [selfMenuOpen, setSelfMenuOpen] = useState(false);
  const [newChatOpen, setNewChatOpen] = useState(false);
  const searchDebounce = useRef<ReturnType<typeof setTimeout> | null>(null);
  const scrollRef = useRef<HTMLDivElement>(null);
  const stickToBottomRef = useRef(true);
  const [panelRef, panelHeight] = usePanelHeight();

  const localMatches = useMemo(() => {
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

  const remote = s.remoteSearch;
  const chats = useMemo(() => {
    let list = localMatches;
    if (filter.trim()) {
      const seen = new Set(localMatches.map((c) => c.id));
      const extra = (remote?.query === filter.trim() ? remote.results : []).filter((c) => !seen.has(c.id));
      list = [...localMatches, ...extra];
    }
    if (unreadOnly) list = list.filter((c) => c.unread);
    return list;
  }, [localMatches, remote, filter, unreadOnly]);

  // Debounced remote search when the local filter is sparse.
  useEffect(() => {
    if (searchDebounce.current) clearTimeout(searchDebounce.current);
    const q = filter.trim();
    if (q.length < 2) {
      onAction('clear-search');
      return;
    }
    searchDebounce.current = setTimeout(() => {
      onAction('search-chats', { query: q, mode: searchMode });
    }, 350);
    return () => {
      if (searchDebounce.current) clearTimeout(searchDebounce.current);
    };
  }, [filter, searchMode]);

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

  const scrollToBottomIfSticky = () => {
    const el = scrollRef.current;
    if (!el || !stickToBottomRef.current) return;
    el.scrollTop = el.scrollHeight;
    requestAnimationFrame(() => {
      if (scrollRef.current && stickToBottomRef.current) {
        scrollRef.current.scrollTop = scrollRef.current.scrollHeight;
      }
    });
  };

  // Auto-scroll: pin to bottom after layout when new messages arrive, chat/height changes,
  // or hosted images finish loading — unless the user has scrolled up.
  useLayoutEffect(
    scrollToBottomIfSticky,
    [messages.length, s.activeChatId, panelHeight, Object.keys(hostedContents).length],
  );

  const onThreadScroll = () => {
    const el = scrollRef.current;
    if (!el) return;
    stickToBottomRef.current = el.scrollHeight - el.scrollTop - el.clientHeight < 40;
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

  const meId = s.auth?.objectId ?? null;

  return (
    <div
      ref={panelRef}
      style={panelHeight ? { height: panelHeight } : undefined}
      className="relative flex flex-col min-h-0 overflow-hidden"
    >
      {/* Top bar */}
      <div className="flex items-center justify-between gap-3 border-b border-border/50 px-3 py-2 shrink-0">
        <div className="flex items-center gap-2">
          <h2 className="text-sm font-semibold text-foreground">Teams</h2>
          <button
            type="button"
            title="New chat"
            onClick={() => setNewChatOpen(true)}
            className="p-1 rounded-md text-muted-foreground hover:text-foreground hover:bg-muted transition-colors"
          >
            <svg className="w-4 h-4" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={2} strokeLinecap="round" strokeLinejoin="round">
              <path d="M12 20h9" /><path d="M16.5 3.5a2.121 2.121 0 0 1 3 3L7 19l-4 1 1-4Z" />
            </svg>
          </button>
          <button
            type="button"
            title="Refresh"
            onClick={() => onAction('refresh-chats')}
            className="p-1 rounded-md text-muted-foreground hover:text-foreground hover:bg-muted transition-colors"
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
            <button
              type="button"
              onClick={() => setSelfMenuOpen((v) => !v)}
              className="rounded-full hover:opacity-90 transition-opacity"
              title={s.auth?.displayName ?? s.auth?.email ?? undefined}
            >
              <Avatar
                id={meId}
                name={s.auth?.displayName ?? s.auth?.email ?? 'Me'}
                photo={photos[meId]}
                presence={presence[meId]}
                size={8}
              />
            </button>
          )}
          <button
            type="button"
            title="Sign out"
            onClick={() => onAction('logout')}
            className="p-1 rounded-md text-muted-foreground hover:text-foreground hover:bg-muted transition-colors"
          >
            <svg className="w-4 h-4" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={2} strokeLinecap="round">
              <path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4" /><path d="m16 17 5-5-5-5" /><path d="M21 12H9" />
            </svg>
          </button>
        </div>
      </div>
      {selfMenuOpen && s.auth && (
        <SelfMenu
          auth={s.auth}
          photo={meId ? photos[meId] : null}
          presence={meId ? presence[meId] : undefined}
          onClose={() => setSelfMenuOpen(false)}
          onLogout={() => {
            setSelfMenuOpen(false);
            onAction('logout');
          }}
        />
      )}

      <div className="flex flex-1 min-h-0">
      {/* Sidebar */}
      <div className="flex w-[280px] min-w-[220px] flex-col border-r border-border/50 min-h-0">
        <div className="px-3 pt-3 pb-2 shrink-0">
          <div className="flex items-center gap-1.5 mb-2">
            <button
              type="button"
              onClick={() => setUnreadOnly(false)}
              className={`px-2 py-0.5 rounded-full text-[11px] transition-colors ${
                !unreadOnly ? 'bg-primary/15 text-primary font-medium' : 'text-muted-foreground hover:bg-muted'
              }`}
            >
              All
            </button>
            <button
              type="button"
              onClick={() => setUnreadOnly(true)}
              className={`px-2 py-0.5 rounded-full text-[11px] transition-colors ${
                unreadOnly ? 'bg-primary/15 text-primary font-medium' : 'text-muted-foreground hover:bg-muted'
              }`}
            >
              Unread
            </button>
          </div>
          <div className="relative">
            <input
              style={{ paddingLeft: 12, paddingRight: 32, paddingTop: 6, paddingBottom: 6 }}
              className="w-full text-xs bg-muted border border-border rounded-lg text-foreground placeholder:text-muted-foreground focus:border-primary transition-colors"
              placeholder={searchMode === 'people' ? 'Search people & chats…' : 'Search in messages…'}
              value={filter}
              onChange={(e) => setFilter(e.target.value)}
            />
            <button
              type="button"
              title={searchMode === 'people' ? 'Searching by person — click for content search' : 'Searching in message content — click for people search'}
              onClick={() => setSearchMode((m) => (m === 'people' ? 'content' : 'people'))}
              style={{ position: 'absolute', right: 4, top: '50%', transform: 'translateY(-50%)' }}
              className="p-1 rounded-md text-muted-foreground hover:text-foreground hover:bg-background/60 transition-colors"
            >
              {searchMode === 'people' ? (
                <svg className="w-3.5 h-3.5" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={2} strokeLinecap="round" strokeLinejoin="round">
                  <path d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2" /><circle cx="9" cy="7" r="4" /><path d="M22 21v-2a4 4 0 0 0-3-3.87" /><path d="M16 3.13a4 4 0 0 1 0 7.75" />
                </svg>
              ) : (
                <svg className="w-3.5 h-3.5" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={2} strokeLinecap="round" strokeLinejoin="round">
                  <path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z" /><path d="M8 9h8M8 13h5" />
                </svg>
              )}
            </button>
          </div>
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
              presence={presence}
              active={ch.id === s.activeChatId}
              onClick={() => onAction('select-chat', { chatId: ch.id })}
            />
          ))}
          {filter.trim() && remote?.loading && (
            <div className="px-3 py-2 text-[11px] text-muted-foreground animate-pulse">Searching directory…</div>
          )}
          {filter.trim() && remote?.error && (
            <div className="px-3 py-2 text-[11px] text-destructive">Search failed: {remote.error}</div>
          )}
          {!s.loadingChats && chats.length === 0 && !(filter.trim() && remote?.loading) && (
            <div className="p-3 text-xs text-muted-foreground">
              {filter.trim() ? 'No matching chats' : 'No chats'}
            </div>
          )}
          {!filter.trim() && s.chatsNextLink && (
            <button
              type="button"
              onClick={() => onAction('load-more-chats')}
              disabled={s.loadingMoreChats}
              className="mx-1.5 mt-1 mb-2 w-[calc(100%-12px)] px-2 py-1.5 text-[11px] font-medium text-muted-foreground bg-muted border border-border rounded-lg hover:bg-muted/80 disabled:opacity-50 transition-colors"
            >
              {s.loadingMoreChats ? 'Loading…' : 'Load more'}
            </button>
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
                    presence={presence[activeChat.members[0].id]}
                    size={9}
                  />
                ) : (
                  <AvatarStack members={activeChat.members} photos={photos} presence={presence} />
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
                  hostedContents={hostedContents}
                  showHeader={i === 0 || messages[i - 1].fromId !== m.fromId}
                  pickerOpen={reactionTarget === m.id}
                  onOpenPicker={() => setReactionTarget(m.id)}
                  onClosePicker={() => setReactionTarget(null)}
                  onReact={(type, remove) => react(m.id, type, remove)}
                  onContentResize={scrollToBottomIfSticky}
                  onMentionClick={(userId) => onAction('load-user-card', { userId })}
                  onOpenInTeams={(url) => onAction('open-in-teams', { url })}
                />
              ))}
            </div>

            {s.error && (
              <div className="mx-4 mb-2 text-xs text-destructive bg-destructive/10 px-3 py-1.5 rounded-md">
                {s.error}
              </div>
            )}

            <RichComposer
              chatId={s.activeChatId ?? ''}
              sending={!!s.sendingMessage}
              onSend={(payload) => {
                if (!s.activeChatId) return;
                stickToBottomRef.current = true;
                onAction('send-message', { chatId: s.activeChatId, payload });
              }}
              onSearchPeople={(q) => onAction('search-people', { query: q })}
              peopleSearch={s.peopleSearch ?? null}
              photos={photos}
              presence={presence}
            />
          </>
        )}
      </div>

      </div>

      {s.userCard && (
        <UserCard state={s} photos={photos} presence={presence} onAction={onAction} />
      )}
      {newChatOpen && (
        <NewChatDialog
          state={s}
          photos={photos}
          onAction={onAction}
          onClose={() => setNewChatOpen(false)}
        />
      )}
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

function Segments({
  segments,
  fromMe,
  onMentionClick,
}: {
  segments: BodySegment[];
  fromMe: boolean;
  onMentionClick: (userId: string) => void;
}) {
  return (
    <>
      {segments.map((seg, i) => {
        if (seg.type === 'text') return <span key={i}>{seg.text}</span>;
        if (seg.type === 'br') return <br key={i} />;
        if (seg.type === 'mention') {
          return (
            <button
              key={i}
              type="button"
              onClick={() => seg.userId && onMentionClick(seg.userId)}
              disabled={!seg.userId}
              style={{
                color: fromMe ? undefined : '#2563eb',
                textDecoration: 'underline',
                textDecorationColor: fromMe ? 'currentColor' : 'rgba(37,99,235,0.4)',
                textUnderlineOffset: 2,
                cursor: seg.userId ? 'pointer' : 'default',
              }}
              className={`inline font-medium hover:opacity-80 transition-opacity ${
                fromMe ? 'text-primary-foreground' : ''
              }`}
            >
              {seg.displayName}
            </button>
          );
        }
        if (seg.type === 'code') {
          return (
            <code
              key={i}
              style={{ fontFamily: 'ui-monospace, SFMono-Regular, Menlo, monospace', fontSize: '0.85em' }}
              className={`px-1 py-0.5 rounded ${fromMe ? 'bg-primary-foreground/20' : 'bg-background/60 border border-border/60'}`}
            >
              {seg.code}
            </code>
          );
        }
        if (seg.type === 'codeblock') {
          return <CodeBlock key={i} code={seg.code} lang={seg.lang} />;
        }
        if (seg.type === 'link') {
          return (
            <a
              key={i}
              href={seg.href}
              onClick={(e) => e.preventDefault()}
              title={seg.href}
              style={{ color: fromMe ? undefined : '#2563eb', textDecoration: 'underline', cursor: 'pointer' }}
              className={fromMe ? 'text-primary-foreground' : ''}
            >
              {seg.text || seg.href}
            </a>
          );
        }
        if (seg.type === 'hr') {
          return <div key={i} className="my-2 border-t border-border/60" />;
        }
        if (seg.type === 'heading') {
          const size = seg.level === 1 ? '1.15em' : seg.level === 2 ? '1.05em' : '1em';
          return (
            <div key={i} style={{ fontSize: size, fontWeight: 600, margin: '6px 0 2px' }}>
              <Segments segments={seg.segments} fromMe={fromMe} onMentionClick={onMentionClick} />
            </div>
          );
        }
        if (seg.type === 'blockquote') {
          return (
            <div
              key={i}
              className={`my-1 pl-2.5 border-l-2 ${fromMe ? 'border-primary-foreground/40' : 'border-border'} opacity-90`}
            >
              <Segments segments={seg.segments} fromMe={fromMe} onMentionClick={onMentionClick} />
            </div>
          );
        }
        if (seg.type === 'table') {
          return (
            <div
              key={i}
              className="my-1.5 rounded-lg border border-border bg-card text-foreground overflow-auto"
              style={{ whiteSpace: 'normal', maxHeight: 360 }}
            >
              <table style={{ borderCollapse: 'collapse', width: '100%', fontSize: '0.85em' }}>
                {seg.header && (
                  <thead>
                    <tr>
                      {seg.header.map((h, hi) => (
                        <th
                          key={hi}
                          className="text-left font-semibold border-b border-border px-2.5 py-1.5 bg-muted/40"
                        >
                          {h}
                        </th>
                      ))}
                    </tr>
                  </thead>
                )}
                <tbody>
                  {seg.rows.map((r, ri) => (
                    <tr key={ri} className="border-b border-border/50 last:border-0">
                      {r.map((c, ci) => (
                        <td key={ci} className="px-2.5 py-1 align-top">{c}</td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          );
        }
        return null;
      })}
    </>
  );
}

function CodeBlock({ code, lang }: { code: string; lang: string | null }) {
  const { html, displayName } = useMemo(() => highlight(code, lang), [code, lang]);
  const lineCount = useMemo(() => code.split('\n').length, [code]);
  const gutterWidth = 12 + String(lineCount).length * 7;
  return (
    <div
      className="my-1.5 rounded-lg border border-border bg-card text-foreground overflow-hidden"
      style={{ whiteSpace: 'normal' }}
    >
      <div className="flex items-center justify-between px-2.5 py-1 border-b border-border/60 bg-muted/40">
        <span className="text-[10px] font-medium uppercase tracking-wide text-muted-foreground">
          {displayName ?? lang ?? 'code'}
        </span>
        <button
          type="button"
          title="Copy"
          onClick={() => navigator.clipboard?.writeText(code)}
          className="p-0.5 rounded text-muted-foreground hover:text-foreground hover:bg-muted transition-colors"
        >
          <svg className="w-3.5 h-3.5" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={2} strokeLinecap="round" strokeLinejoin="round">
            <rect x="9" y="9" width="13" height="13" rx="2" /><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1" />
          </svg>
        </button>
      </div>
      <div
        style={{
          fontFamily: 'ui-monospace, SFMono-Regular, Menlo, monospace',
          fontSize: '0.8em',
          lineHeight: 1.5,
          maxHeight: 320,
        }}
        className="overflow-auto flex"
      >
        <div
          style={{ minWidth: gutterWidth, userSelect: 'none' }}
          className="text-right pr-2 pl-1.5 py-1.5 text-muted-foreground/50 shrink-0 border-r border-border/40"
        >
          {Array.from({ length: lineCount }, (_, i) => (
            <div key={i}>{i + 1}</div>
          ))}
        </div>
        <code
          style={{ whiteSpace: 'pre', display: 'block' }}
          className="px-3 py-1.5 flex-1"
          dangerouslySetInnerHTML={{ __html: html }}
        />
      </div>
    </div>
  );
}

function ChatRow({
  chat,
  photos,
  presence,
  active,
  onClick,
}: {
  chat: NormalizedChat;
  photos: Record<string, string | null>;
  presence: Record<string, Presence>;
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
        <Avatar
          id={chat.members[0].id}
          name={chat.members[0].displayName}
          photo={photos[chat.members[0].id]}
          presence={presence[chat.members[0].id]}
        />
      ) : (
        <AvatarStack members={chat.members} photos={photos} presence={presence} max={2} />
      )}
      <div className="min-w-0 flex-1">
        <div className="flex items-baseline justify-between gap-2">
          <div
            className={`text-xs truncate ${chat.unread ? 'font-semibold text-foreground' : 'font-medium text-foreground'}`}
          >
            {chatTitle(chat)}
          </div>
          <div className="text-[10px] text-muted-foreground shrink-0">{fmtTime(chat.lastUpdated)}</div>
        </div>
        <div
          className={`text-[11px] truncate ${chat.unread ? 'text-foreground/80' : 'text-muted-foreground'}`}
        >
          {chat.lastMessageFrom ? `${chat.lastMessageFrom}: ` : ''}
          {chat.lastMessagePreview ?? ''}
        </div>
      </div>
      {chat.unread && <span className="w-2 h-2 rounded-full bg-primary shrink-0" />}
    </div>
  );
}

function MessageBubble({
  m,
  photos,
  hostedContents,
  showHeader,
  pickerOpen,
  onOpenPicker,
  onClosePicker,
  onReact,
  onContentResize,
  onMentionClick,
  onOpenInTeams,
}: {
  m: NormalizedMessage;
  photos: Record<string, string | null>;
  hostedContents: Record<string, string | null>;
  showHeader: boolean;
  pickerOpen: boolean;
  onOpenPicker: () => void;
  onClosePicker: () => void;
  onReact: (type: string, remove?: boolean) => void;
  onContentResize: () => void;
  onMentionClick: (userId: string) => void;
  onOpenInTeams: (url: string) => void;
}) {
  const [hover, setHover] = useState(false);

  if (m.systemEvent) {
    return (
      <div className="flex items-center gap-2 my-0.5 text-[10px] text-muted-foreground">
        <div className="flex-1 border-t border-border/40" />
        <span className="px-2 whitespace-nowrap">{m.systemEvent}{m.createdDateTime ? ` · ${fmtTime(m.createdDateTime)}` : ''}</span>
        <div className="flex-1 border-t border-border/40" />
      </div>
    );
  }

  const hasBody = m.segments.length > 0;
  const hasContent =
    hasBody || m.hostedImages.length > 0 || m.attachments.length > 0 || m.files.length > 0 ||
    m.cards.length > 0 || !!m.replyTo || !!m.forwarded;
  const imageOnly = !hasBody && !m.replyTo && !m.forwarded && m.cards.length === 0 && m.files.length === 0 && m.hostedImages.length > 0;

  return (
    <div
      className="group flex gap-2"
      style={{ flexDirection: m.fromMe ? 'row-reverse' : 'row' }}
      onMouseEnter={() => setHover(true)}
      onMouseLeave={() => setHover(false)}
    >
      <div className="w-7 shrink-0">
        {showHeader && !m.fromMe && m.fromId && (
          <Avatar id={m.fromId} name={m.fromName ?? '?'} photo={photos[m.fromId]} size={7} />
        )}
      </div>
      <div
        className="relative flex flex-col"
        style={{ maxWidth: '75%', alignItems: m.fromMe ? 'flex-end' : 'flex-start' }}
      >
        {showHeader && !m.fromMe && m.fromName && (
          <div className="text-[10px] font-medium text-muted-foreground px-1 mb-0.5">
            {m.fromName}
            {m.fromApp && (
              <span className="ml-1 px-1 rounded bg-muted text-[9px] uppercase tracking-wide">app</span>
            )}
          </div>
        )}
        <div
          onContextMenu={(e) => {
            e.preventDefault();
            onOpenPicker();
          }}
          className={`relative rounded-2xl text-sm whitespace-pre-wrap break-words ${
            imageOnly ? 'p-0.5' : 'px-3 py-1.5'
          } ${
            m.fromMe
              ? 'bg-primary text-primary-foreground rounded-br-md'
              : 'bg-muted text-foreground rounded-bl-md'
          }`}
        >
          {m.forwarded && (
            <div
              className={`mb-1.5 rounded-lg border-l-2 px-2 py-1 text-[11px] ${
                m.fromMe ? 'bg-primary-foreground/15 border-primary-foreground/40' : 'bg-background/60 border-border'
              }`}
            >
              <div className="font-medium opacity-70 flex items-center gap-1">
                <svg className="w-3 h-3" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={2} strokeLinecap="round" strokeLinejoin="round">
                  <polyline points="15 17 20 12 15 7" /><path d="M4 18v-2a4 4 0 0 1 4-4h12" />
                </svg>
                Forwarded{m.forwarded.senderName ? ` from ${m.forwarded.senderName}` : ''}
                {m.forwarded.originalDate ? ` · ${fmtTime(m.forwarded.originalDate)}` : ''}
              </div>
              {m.forwarded.text && (
                <div className="opacity-80 mt-0.5 whitespace-pre-wrap break-words">{m.forwarded.text}</div>
              )}
            </div>
          )}
          {m.replyTo && (
            <div
              className={`mb-1.5 rounded-lg border-l-2 px-2 py-1 text-[11px] ${
                m.fromMe
                  ? 'bg-primary-foreground/15 border-primary-foreground/40'
                  : 'bg-background/60 border-border'
              }`}
            >
              {m.replyTo.senderName && (
                <div className="font-medium opacity-80">{m.replyTo.senderName}</div>
              )}
              <div className="opacity-75 line-clamp-3 whitespace-pre-wrap break-words">
                {m.replyTo.text ?? '(quoted message)'}
              </div>
            </div>
          )}
          {hasBody && <Segments segments={m.segments} fromMe={m.fromMe} onMentionClick={onMentionClick} />}
          {m.cards.map((card, ci) => (
            <AdaptiveCard
              key={card.id ?? ci}
              contentJson={card.contentJson}
              teamsDeepLink={
                m.chatId
                  ? `https://teams.microsoft.com/l/message/${m.chatId}/${m.id}?context=${encodeURIComponent('{"contextType":"chat"}')}`
                  : null
              }
              onOpenUrl={(url) => window.open?.(url, '_blank')}
              onOpenInTeams={(url) => onOpenInTeams(url)}
            />
          ))}
          {m.files.length > 0 && (
            <div className={`flex flex-col gap-1 ${hasBody || m.cards.length ? 'mt-1.5' : ''}`}>
              {m.files.map((f, fi) => (
                <a
                  key={fi}
                  href={f.url ?? undefined}
                  onClick={(e) => e.preventDefault()}
                  title={f.url ?? undefined}
                  className={`flex items-center gap-2 rounded-lg border px-2.5 py-1.5 text-xs ${
                    m.fromMe ? 'border-primary-foreground/30 bg-primary-foreground/10' : 'border-border bg-card text-foreground'
                  } hover:opacity-90`}
                  style={{ whiteSpace: 'normal', textDecoration: 'none' }}
                >
                  <svg className="w-4 h-4 shrink-0 opacity-70" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={2} strokeLinecap="round" strokeLinejoin="round">
                    <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" /><polyline points="14 2 14 8 20 8" />
                  </svg>
                  <span className="truncate min-w-0">{f.name}</span>
                </a>
              ))}
            </div>
          )}
          {m.hostedImages.length > 0 && (
            <div className={`flex flex-col gap-1 ${hasBody ? 'mt-1.5' : ''}`}>
              {m.hostedImages.map((u) => {
                const data = hostedContents[u];
                if (data) {
                  return (
                    <img
                      key={u}
                      src={data}
                      alt=""
                      onLoad={onContentResize}
                      style={{ maxHeight: 320 }}
                      className="rounded-xl max-w-full object-contain"
                    />
                  );
                }
                return (
                  <div
                    key={u}
                    className="rounded-xl bg-background/20 border border-border/40 px-3 py-4 text-[11px] opacity-70 flex items-center gap-2"
                  >
                    {data === null ? '⚠️ image unavailable' : (
                      <>
                        <span className="w-3 h-3 rounded-full border-2 border-current border-t-transparent animate-spin" />
                        loading image…
                      </>
                    )}
                  </div>
                );
              })}
            </div>
          )}
          {!hasContent && <span className="opacity-50 italic">(no content)</span>}
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
          <svg className="w-3.5 h-3.5" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={2} strokeLinecap="round" strokeLinejoin="round">
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
