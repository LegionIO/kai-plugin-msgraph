import { PanelView } from './components/PanelView.tsx';
import { SettingsView } from './components/SettingsView.tsx';

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
