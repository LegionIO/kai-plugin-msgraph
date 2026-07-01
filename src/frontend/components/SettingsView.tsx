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
  createGroupChat: 'create-group-chat — new group + optional first message',
};

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
