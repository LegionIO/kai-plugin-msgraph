import { withMfaActivity } from './MfaActivity';
import React from 'react';
import { PanelView as TeamsPanel } from './components/PanelView.tsx';
import { MailPanelView } from './components/MailPanelView.tsx';
import { SettingsView } from './components/SettingsView.tsx';
import { ViewRail } from './components/ViewTabs.tsx';
import type { PluginComponentProps } from './hooks.ts';
import { usePanelHeight } from './hooks.ts';
import type { MsgraphPluginState } from '../shared/types.ts';

function PanelView(props: PluginComponentProps<MsgraphPluginState>) {
  const [panelRef, panelHeight] = usePanelHeight();
  const s = props.pluginState ?? ({} as MsgraphPluginState);
  // state.activeView is the source of truth (defaults to 'teams'); Kai's own
  // panel selection (props.view) is ignored so alphabetical panel-sort in the
  // host can't land us on Outlook when the user last picked Teams.
  const view: 'teams' | 'mail' = s.activeView ?? 'teams';
  const authed = !!s.auth?.isAuthenticated;
  const chatUnread = (s.chats ?? []).reduce((n, c) => n + (c.unread ? 1 : 0), 0);
  const mailUnread = (s.mailFolders ?? []).find((f) => f.wellKnownName === 'inbox')?.unreadItemCount ?? 0;
  const child = view === 'mail'
    ? React.createElement(MailPanelView, props)
    : React.createElement(TeamsPanel, { ...props, panelHeight });
  return React.createElement(
    'div',
    { ref: panelRef, style: { display: 'flex', height: panelHeight, minHeight: 0, overflow: 'hidden' } },
    authed && React.createElement(ViewRail, { active: view, chatUnread, mailUnread, onAction: props.onAction }),
    React.createElement('div', { style: { flex: 1, minWidth: 0, minHeight: 0 } }, child),
  );
}

export function register(env: {
  React: unknown;
  registerComponents: (pluginName: string, components: Record<string, unknown>) => void;
}) {
  (globalThis as Record<string, unknown>).React = env.React;
  env.registerComponents('msgraph', {
    PanelView: withMfaActivity(PanelView),
    SettingsView: withMfaActivity(SettingsView, true),
  });
}
