import type { PluginAPI } from '../shared/types.js';

type Logger = PluginAPI['log'];

let logger: Logger = {
  info: (...a) => console.log('[msgraph]', ...a),
  warn: (...a) => console.warn('[msgraph]', ...a),
  error: (...a) => console.error('[msgraph]', ...a),
};

export function setLogger(l: Logger): void {
  logger = l;
}

export function getLogger(): Logger {
  return logger;
}
