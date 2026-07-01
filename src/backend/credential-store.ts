/**
 * Credential storage with OS keychain primary + AES-256-GCM fallback.
 * Adapted from kai-plugin-pim/src/backend/credential-store.ts.
 */

import { createCipheriv, createDecipheriv, randomBytes, pbkdf2Sync, createHash } from 'crypto';
import { readFileSync, writeFileSync, existsSync, unlinkSync, mkdirSync, lstatSync, chmodSync } from 'fs';
import { join, dirname } from 'path';
import { hostname, userInfo, homedir } from 'os';
import type { PluginAPI, CredentialStatus } from '../shared/types.js';
import { getLogger } from './logger-singleton.js';

const INSTALL_KEY_PATH = join(homedir(), '.kai', 'msgraph.key');

interface SafeStorageCreds {
  username: string;
  encryptedPassword: string;
  method: 'safeStorage';
}

interface AesCredentials {
  username: string;
  iv: string;
  authTag: string;
  ciphertext: string;
  method: 'aes256gcm';
}

function getOrCreateInstallKey(): Buffer {
  let st;
  try { st = lstatSync(INSTALL_KEY_PATH); } catch { st = null; }
  if (st) {
    if (!st.isFile()) throw new Error(`Install key at ${INSTALL_KEY_PATH} is not a regular file`);
    if ((st.mode & 0o077) !== 0) chmodSync(INSTALL_KEY_PATH, 0o600);
    const hex = readFileSync(INSTALL_KEY_PATH, 'utf-8').trim();
    if (!/^[0-9a-fA-F]{64}$/.test(hex)) throw new Error(`Install key at ${INSTALL_KEY_PATH} is malformed`);
    return Buffer.from(hex, 'hex');
  }
  const key = randomBytes(32);
  mkdirSync(dirname(INSTALL_KEY_PATH), { recursive: true });
  writeFileSync(INSTALL_KEY_PATH, key.toString('hex'), { encoding: 'utf-8', mode: 0o600 });
  return key;
}

function machineSalt(): Buffer {
  const material = `${hostname()}:${userInfo().username}:${homedir()}`;
  return createHash('sha256').update(material).digest();
}

function deriveAesKey(): Buffer {
  return pbkdf2Sync(getOrCreateInstallKey(), machineSalt(), 100_000, 32, 'sha512');
}

function decryptAes(stored: AesCredentials, key: Buffer): string {
  const decipher = createDecipheriv('aes-256-gcm', key, Buffer.from(stored.iv, 'base64'));
  decipher.setAuthTag(Buffer.from(stored.authTag, 'base64'));
  return Buffer.concat([decipher.update(Buffer.from(stored.ciphertext, 'base64')), decipher.final()]).toString('utf8');
}

function encryptAes(password: string, key: Buffer): Omit<AesCredentials, 'username'> {
  const iv = randomBytes(12);
  const cipher = createCipheriv('aes-256-gcm', key, iv);
  const encrypted = Buffer.concat([cipher.update(password, 'utf8'), cipher.final()]);
  return {
    iv: iv.toString('base64'),
    authTag: cipher.getAuthTag().toString('base64'),
    ciphertext: encrypted.toString('base64'),
    method: 'aes256gcm',
  };
}

function getCredFilePath(pluginDir: string): string {
  return join(pluginDir, 'credentials.enc.json');
}

export function saveCredentials(api: PluginAPI, username: string, password: string): void {
  try {
    if (api.safeStorage.isEncryptionAvailable()) {
      const encrypted = api.safeStorage.encryptString(password);
      const creds: SafeStorageCreds = { username, encryptedPassword: encrypted, method: 'safeStorage' };
      api.config.setPluginData('encryptedCredentials', creds);
      getLogger().info(`Saved credentials via OS keychain for ${username}`);
      return;
    }
  } catch { /* fall through */ }

  const creds: AesCredentials = { username, ...encryptAes(password, deriveAesKey()) };
  const credPath = getCredFilePath(api.pluginDir);
  writeFileSync(credPath, JSON.stringify(creds, null, 2), { encoding: 'utf-8', mode: 0o600 });
  try { chmodSync(credPath, 0o600); } catch { /* best effort */ }
  getLogger().info(`Saved credentials via AES-256-GCM fallback for ${username}`);
}

export function getCredentials(api: PluginAPI): { username: string; password: string } | null {
  const storedConfig = api.config.getPluginData().encryptedCredentials as SafeStorageCreds | null | undefined;
  if (storedConfig?.method === 'safeStorage') {
    try {
      const password = api.safeStorage.decryptString(storedConfig.encryptedPassword);
      return { username: storedConfig.username, password };
    } catch (err) {
      getLogger().error(`Failed to decrypt safeStorage credentials: ${err}`);
      return null;
    }
  }

  const filePath = getCredFilePath(api.pluginDir);
  if (!existsSync(filePath)) return null;
  try {
    const stored: AesCredentials = JSON.parse(readFileSync(filePath, 'utf-8'));
    if (stored.method !== 'aes256gcm') return null;
    const password = decryptAes(stored, deriveAesKey());
    return { username: stored.username, password };
  } catch (err) {
    getLogger().error(`Failed to decrypt AES credentials: ${err}`);
    return null;
  }
}

export function hasCredentials(api: PluginAPI): boolean {
  const storedConfig = api.config.getPluginData().encryptedCredentials as SafeStorageCreds | null | undefined;
  if (storedConfig?.method === 'safeStorage') return true;
  return existsSync(getCredFilePath(api.pluginDir));
}

export function getStoredUsername(api: PluginAPI): string | null {
  const storedConfig = api.config.getPluginData().encryptedCredentials as SafeStorageCreds | null | undefined;
  if (storedConfig?.username) return storedConfig.username;
  const filePath = getCredFilePath(api.pluginDir);
  if (!existsSync(filePath)) return null;
  try {
    const stored = JSON.parse(readFileSync(filePath, 'utf-8'));
    return stored.username ?? null;
  } catch {
    return null;
  }
}

export function clearCredentials(api: PluginAPI): void {
  api.config.setPluginData('encryptedCredentials', null);
  const filePath = getCredFilePath(api.pluginDir);
  if (existsSync(filePath)) unlinkSync(filePath);
  getLogger().info('Cleared stored credentials');
}

export function getEncryptionMethod(api: PluginAPI): CredentialStatus['encryptionMethod'] {
  try {
    if (api.safeStorage.isEncryptionAvailable()) return 'os-keychain';
  } catch { /* not available */ }
  return 'aes256gcm';
}

export function getCredentialStatus(api: PluginAPI): CredentialStatus {
  return {
    hasCredentials: hasCredentials(api),
    username: getStoredUsername(api),
    encryptionMethod: getEncryptionMethod(api),
  };
}
