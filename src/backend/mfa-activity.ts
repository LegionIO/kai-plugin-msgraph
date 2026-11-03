import { readComputerIdleSeconds } from './computer-idle.js';

export type MfaActivityPolicy = { source: 'computer' | 'plugin'; idleMinutes: number; allowAgent: boolean };
export const DEFAULT_MFA_ACTIVITY_POLICY: MfaActivityPolicy = { source: 'computer', idleMinutes: 5, allowAgent: true };

export function normalizeMfaActivityPolicy(value: any): MfaActivityPolicy {
  return {
    source: value?.source === 'plugin' ? 'plugin' : 'computer',
    idleMinutes: typeof value?.idleMinutes === 'number' && Number.isFinite(value.idleMinutes)
      ? Math.min(60, Math.max(1, value.idleMinutes)) : 5,
    allowAgent: value?.allowAgent !== false,
  };
}

/** Wrap the plugin API once, so every Azure redirect shares the same policy and lifecycle. */
export function createMfaActivityGuard<T extends Record<string, any>>(
  original: T,
  idleSeconds: () => Promise<number | null> = readComputerIdleSeconds,
) {
  let disposed = false;
  const uiActivity = new Map<string, number>();
  const agents = new Set<{ signal?: AbortSignal }>();
  const windows = new Set<{ abort: AbortController; close?: () => unknown; cleanup: () => void }>();
  const waiting = new Set<AbortController>();
  let deferred = false;
  // After an unattended challenge, do not replay credentials on every timer tick.
  // Keep this across plugin restarts; successful authentication clears its partition.
  const blocked = new Set<string>(Array.isArray(original.config.getPluginData().mfaPendingPartitions)
    ? original.config.getPluginData().mfaPendingPartitions.filter((key: unknown) => typeof key === 'string') : []);
  let idleRead: Promise<number | null> | undefined;
  let idleReadPending = false;
  let idleReadAt = -Infinity;
  const policy = () => normalizeMfaActivityPolicy(original.config.getPluginData().mfaActivityPolicy);
  deferred = blocked.size > 0;
  const publish = () => original.state.set('mfaActivityPaused', waiting.size > 0 || deferred);
  const saveBlocked = () => { original.config.setPluginData('mfaPendingPartitions', [...blocked]); publish(); };
  const active = async () => {
    if (disposed) return false;
    const prefs = policy();
    if (prefs.allowAgent && [...agents].some(agent => !agent.signal?.aborted)) return true;
    const now = Date.now();
    for (const [id, at] of uiActivity) {
      if (now - at >= prefs.idleMinutes * 60_000 || now < at) uiActivity.delete(id);
    }
    if (uiActivity.size) return true;
    if (prefs.source === 'plugin') return false;
    if (!idleRead || (!idleReadPending && (now - idleReadAt >= 1_000 || now < idleReadAt))) {
      idleReadAt = now;
      idleReadPending = true;
      idleRead = Promise.resolve().then(idleSeconds).catch(() => null).finally(() => { idleReadPending = false; });
    }
    const idle = await idleRead;
    return !disposed && idle !== null && Number.isFinite(idle) && idle >= 0 && idle < prefs.idleMinutes * 60;
  };
  original.state.set('mfaActivityPolicy', policy());
  publish();

  const api = {
    ...original,
    onAction(target: string, handler: (...args: any[]) => any) {
      return original.onAction(target, (action: string, data: any) => {
        if (action === 'mfa-user-activity') {
          if (typeof data?.source === 'string' && data.source.length < 100) {
            if (data.active === true) uiActivity.set(data.source, Date.now());
            else uiActivity.delete(data.source);
          }
          return;
        }
        if (action === 'save-mfa-activity-policy') {
          const next = normalizeMfaActivityPolicy(data);
          original.config.setPluginData('mfaActivityPolicy', next);
          original.state.set('mfaActivityPolicy', next);
          idleReadAt = -Infinity;
          return;
        }
        return handler(action, data);
      });
    },
    tools: { ...original.tools, register(definitions: any[]) {
      return original.tools.register(definitions.map(tool => ({ ...tool, execute: async (input: unknown, context: any) => {
        const agent = { signal: context?.abortSignal as AbortSignal | undefined };
        // Headless automations can use tokens, but cannot authorize MFA prompts.
        const interactive = !!context?.toolCallId && context?.isHeadless !== true;
        if (interactive) agents.add(agent);
        try { return await tool.execute(input, context); }
        finally { agents.delete(agent); }
      } })));
    } },
    auth: { ...original.auth, async openAuthWindow(options: any) {
      const key = options.partition || new URL(options.url).origin;
      if (disposed) return { success: false, error: 'Plugin stopped.' };
      if (blocked.has(key) && !await active()) {
        deferred = true;
        publish();
        return { success: false, error: 'MFA paused while inactive. Return to the computer and sign in again.' };
      }
      let revealTimer: ReturnType<typeof setTimeout> | undefined;
      let observerTimer: ReturnType<typeof setInterval> | undefined;
      const window = { abort: new AbortController(), close: undefined as (() => unknown) | undefined,
        cleanup: () => { clearTimeout(revealTimer); clearInterval(observerTimer); } };
      windows.add(window);
      const waitForMfaActivity = async (): Promise<boolean> => {
        if (!await active()) {
          waiting.add(window.abort);
          if (!blocked.has(key)) { blocked.add(key); saveBlocked(); }
          else publish();
        }
        try {
          while (!window.abort.signal.aborted && !disposed) {
            if (await active()) { deferred = false; return !window.abort.signal.aborted; }
            await new Promise<void>(resolve => {
              const finish = () => { clearTimeout(timer); window.abort.signal.removeEventListener('abort', finish); resolve(); };
              const timer = setTimeout(finish, 1_000);
              window.abort.signal.addEventListener('abort', finish, { once: true });
              if (window.abort.signal.aborted) finish();
            });
          }
          return false;
        } finally { waiting.delete(window.abort); publish(); }
      };
      try {
        const result = await original.auth.openAuthWindow({
          ...options,
          showOnCreate: options.showOnCreate === false ? false : await active(),
          // Host timers cannot check presence. Reveal through the guarded helper instead.
          showAfterMs: undefined,
          onReady: (helpers: any) => {
            window.close = () => helpers.close();
            if (disposed || window.abort.signal.aborted) { void helpers.close(); return; }
            const show = () => { void active().then(present => {
              if (present && !window.abort.signal.aborted) return helpers.show();
            }).catch(() => {}); };
            // Cookie-only refreshes may reach MFA without running form automation.
            // Record that boundary too, so they cannot trigger repeated pushes.
            let observing = false;
            const observe = async () => {
              if (observing || window.abort.signal.aborted || waiting.has(window.abort)) return;
              observing = true;
              try {
                const url = new URL(await helpers.getURL());
                if (url.protocol !== 'https:' || (url.port && url.port !== '443') ||
                  !['login.microsoftonline.com', 'login.live.com'].includes(url.hostname)) return;
                const mfa = await helpers.executeJavaScript(`(() => {
                  if (location.protocol !== 'https:' || !['login.microsoftonline.com', 'login.live.com'].includes(location.hostname)) return false;
                  return ['#idDiv_SAOTCS_Proofs', '#idDiv_SAOTCAS_Description', '#idTxtBx_SAOTCC_OTC'].some(selector => {
                    const element = document.querySelector(selector); return element && element.offsetParent !== null;
                  });
                })()`);
                if (mfa && !await active() && !window.abort.signal.aborted) {
                  deferred = true;
                  if (!blocked.has(key)) { blocked.add(key); saveBlocked(); }
                  else publish();
                }
              } catch { /* Window closed or navigation replaced the document. */ }
              finally { observing = false; }
            };
            observerTimer = setInterval(() => { void observe(); }, 1_000);
            void observe();
            if (options.showAfterMs) revealTimer = setTimeout(show, options.showAfterMs);
            return options.onReady?.({ ...helpers, show, waitForMfaActivity });
          },
        });
        if (result.success && blocked.delete(key)) { deferred = false; saveBlocked(); }
        return result;
      } finally {
        if (waiting.has(window.abort)) deferred = true;
        window.abort.abort();
        window.cleanup();
        windows.delete(window);
      }
    } },
  } as T;
  return { api, isActive: active, dispose() {
    disposed = true;
    uiActivity.clear(); agents.clear();
    for (const window of windows) {
      window.abort.abort();
      window.cleanup();
      void Promise.resolve().then(() => window.close?.()).catch(() => {});
    }
    waiting.clear();
  } };
}
