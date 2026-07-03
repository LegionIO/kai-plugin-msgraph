import React from 'react';

export function ViewTabs({
  active,
  onNavigate,
  chatUnread,
  mailUnread,
}: {
  active: 'teams' | 'mail';
  onNavigate: (view: 'teams' | 'mail') => void;
  chatUnread: number;
  mailUnread: number;
}) {
  return (
    <div className="inline-flex items-center rounded-lg border border-border bg-muted/40 p-0.5">
      <Tab
        active={active === 'teams'}
        label="Teams"
        badge={chatUnread}
        onClick={() => active !== 'teams' && onNavigate('teams')}
      />
      <Tab
        active={active === 'mail'}
        label="Outlook"
        badge={mailUnread}
        onClick={() => active !== 'mail' && onNavigate('mail')}
      />
    </div>
  );
}

function Tab({ active, label, badge, onClick }: { active: boolean; label: string; badge: number; onClick: () => void }) {
  return (
    <button
      type="button"
      onClick={onClick}
      className={`px-3 py-1 text-xs font-medium rounded-md transition-colors flex items-center gap-1.5 ${
        active ? 'bg-card text-foreground shadow-sm' : 'text-muted-foreground hover:text-foreground'
      }`}
    >
      {label}
      {badge > 0 && (
        <span
          style={{ minWidth: 16, height: 14, fontSize: 9, lineHeight: '14px' }}
          className="rounded-full bg-primary text-primary-foreground text-center px-1"
        >
          {badge > 99 ? '99+' : badge}
        </span>
      )}
    </button>
  );
}
