import React from 'react';
import { PanelView as TeamsPanel } from './components/PanelView.tsx';
import { MailPanelView } from './components/MailPanelView.tsx';
import { SettingsView } from './components/SettingsView.tsx';
import { ViewRail } from './components/ViewTabs.tsx';
import type { PluginComponentProps } from './hooks.ts';
import type { MsgraphPluginState } from '../shared/types.ts';

function PanelView(props: PluginComponentProps<MsgraphPluginState>) {
  const view = ((props.props as { view?: string } | undefined)?.view === 'mail' ? 'mail' : 'teams') as 'teams' | 'mail';
  const s = props.pluginState ?? ({} as MsgraphPluginState);
  const authed = !!s.auth?.isAuthenticated;
  const chatUnread = (s.chats ?? []).reduce((n, c) => n + (c.unread ? 1 : 0), 0);
  const mailUnread = (s.mailFolders ?? []).find((f) => f.wellKnownName === 'inbox')?.unreadItemCount ?? 0;
  const child = React.createElement(view === 'mail' ? MailPanelView : TeamsPanel, props);
  if (!authed) return child;
  return React.createElement(
    'div',
    { style: { display: 'flex', height: '100%', minHeight: 0 } },
    React.createElement(ViewRail, { active: view, chatUnread, mailUnread }),
    React.createElement('div', { style: { flex: 1, minWidth: 0, minHeight: 0 } }, child),
  );
}

export function register(env: {
  React: unknown;
  registerComponents: (pluginName: string, components: Record<string, unknown>) => void;
}) {
  (globalThis as Record<string, unknown>).React = env.React;
  env.registerComponents('msgraph', {
    PanelView,
    SettingsView,
  });
}
