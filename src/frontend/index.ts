import React from 'react';
import { PanelView as TeamsPanel } from './components/PanelView.tsx';
import { MailPanelView } from './components/MailPanelView.tsx';
import { SettingsView } from './components/SettingsView.tsx';
import type { PluginComponentProps } from './hooks.ts';
import type { MsgraphPluginState } from '../shared/types.ts';

function PanelView(props: PluginComponentProps<MsgraphPluginState>) {
  const view = (props.props as { view?: string } | undefined)?.view;
  return React.createElement(view === 'mail' ? MailPanelView : TeamsPanel, props);
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
