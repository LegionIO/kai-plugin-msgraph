import React from 'react';

export function ViewRail({
  active,
  chatUnread,
  mailUnread,
  onAction,
}: {
  active: 'teams' | 'mail';
  chatUnread: number;
  mailUnread: number;
  onAction: (action: string, data?: unknown) => void;
}) {
  const go = (view: 'teams' | 'mail') => {
    if (view === active) return;
    onAction('set-view', { view });
  };
  return (
    <div
      className="flex flex-col items-center gap-1 border-r border-border/50 bg-muted/20 py-2 shrink-0"
      style={{ width: 48 }}
    >
      <RailBtn
        active={active === 'teams'}
        title="Teams"
        badge={chatUnread}
        onClick={() => go('teams')}
      >
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={2} strokeLinecap="round" strokeLinejoin="round">
          <path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z" />
          <path d="M8 10h.01M12 10h.01M16 10h.01" />
        </svg>
      </RailBtn>
      <RailBtn
        active={active === 'mail'}
        title="Outlook"
        badge={mailUnread}
        onClick={() => go('mail')}
      >
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={2} strokeLinecap="round" strokeLinejoin="round">
          <rect x="2" y="4" width="20" height="16" rx="2" />
          <path d="m2 7 10 6 10-6" />
        </svg>
      </RailBtn>
    </div>
  );
}

function RailBtn({
  active,
  title,
  badge,
  onClick,
  children,
}: {
  active: boolean;
  title: string;
  badge: number;
  onClick: () => void;
  children: React.ReactNode;
}) {
  return (
    <button
      type="button"
      title={title}
      onClick={onClick}
      style={{ width: 36, height: 36, position: 'relative' }}
      className={`flex items-center justify-center rounded-lg transition-colors ${
        active
          ? 'bg-primary/15 text-primary'
          : 'text-muted-foreground hover:text-foreground hover:bg-muted'
      }`}
    >
      <span className="block w-5 h-5 [&>svg]:w-5 [&>svg]:h-5">{children}</span>
      {badge > 0 && (
        <span
          style={{
            position: 'absolute', top: -2, right: -2,
            minWidth: 15, height: 15, fontSize: 8.5, lineHeight: '15px',
            borderRadius: '999px', padding: '0 3px',
          }}
          className="bg-primary text-primary-foreground text-center font-semibold"
        >
          {badge > 99 ? '99+' : badge}
        </span>
      )}
    </button>
  );
}
