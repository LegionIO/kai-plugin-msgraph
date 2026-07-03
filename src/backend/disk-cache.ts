import { mkdirSync, readFileSync, writeFileSync, existsSync, rmSync } from 'fs';
import { readFile } from 'fs/promises';
import { join } from 'path';
import { homedir } from 'os';
import { getLogger } from './logger-singleton.js';

export interface CacheEntry<V> {
  v: V;
  at: number;
}

/**
 * Simple JSON-file-backed key→value cache with debounced writes.
 * One instance per namespace; all namespaces live under
 * ~/.kai/plugin-caches/{pluginName}/{namespace}.json
 */
export class DiskCache<V> {
  private map = new Map<string, CacheEntry<V>>();
  private file: string;
  private writeTimer: ReturnType<typeof setTimeout> | null = null;
  private dirty = false;

  constructor(
    pluginName: string,
    namespace: string,
    private readonly opts: {
      hardTtlMs: number;
      maxEntries?: number;
      /** Load synchronously in the constructor. Only for tiny caches. */
      sync?: boolean;
    },
    private readonly onLoaded?: () => void,
  ) {
    const dir = join(homedir(), '.kai', 'plugin-caches', pluginName);
    mkdirSync(dir, { recursive: true });
    this.file = join(dir, `${namespace}.json`);
    if (opts.sync) this.hydrate(this.readSync());
    else void this.loadAsync();
  }

  private readSync(): string | null {
    try { return existsSync(this.file) ? readFileSync(this.file, 'utf-8') : null; }
    catch { return null; }
  }

  private async loadAsync(): Promise<void> {
    try {
      const txt = await readFile(this.file, 'utf-8').catch(() => null);
      this.hydrate(txt);
    } finally {
      this.onLoaded?.();
    }
  }

  private hydrate(txt: string | null): void {
    if (!txt) return;
    try {
      const raw = JSON.parse(txt) as Record<string, CacheEntry<V>>;
      const now = Date.now();
      for (const [k, e] of Object.entries(raw)) {
        if (e && typeof e.at === 'number' && now - e.at < this.opts.hardTtlMs && !this.map.has(k)) {
          this.map.set(k, e);
        }
      }
      getLogger().info(`disk-cache[${this.file}]: loaded ${this.map.size} entries`);
    } catch (err) {
      getLogger().warn(`disk-cache load failed (${this.file}): ${err}`);
    }
  }

  get(key: string): CacheEntry<V> | undefined {
    const e = this.map.get(key);
    if (!e) return undefined;
    if (Date.now() - e.at >= this.opts.hardTtlMs) {
      this.map.delete(key);
      this.scheduleWrite();
      return undefined;
    }
    return e;
  }

  set(key: string, value: V): void {
    this.map.set(key, { v: value, at: Date.now() });
    if (this.opts.maxEntries && this.map.size > this.opts.maxEntries) {
      // Evict oldest.
      let oldestK: string | null = null;
      let oldestAt = Infinity;
      for (const [k, e] of this.map) {
        if (e.at < oldestAt) { oldestAt = e.at; oldestK = k; }
      }
      if (oldestK) this.map.delete(oldestK);
    }
    this.scheduleWrite();
  }

  has(key: string): boolean {
    return this.get(key) !== undefined;
  }

  entries(): IterableIterator<[string, CacheEntry<V>]> {
    return this.map.entries();
  }

  clear(): void {
    this.map.clear();
    this.dirty = true;
    this.flush();
    try { rmSync(this.file, { force: true }); } catch { /* ignore */ }
  }

  private scheduleWrite(): void {
    this.dirty = true;
    if (this.writeTimer) return;
    this.writeTimer = setTimeout(() => {
      this.writeTimer = null;
      this.flush();
    }, 500);
  }

  dispose(): void {
    if (this.writeTimer) { clearTimeout(this.writeTimer); this.writeTimer = null; }
    this.flush();
  }

  flush(): void {
    if (!this.dirty) return;
    this.dirty = false;
    try {
      const obj: Record<string, CacheEntry<V>> = {};
      for (const [k, e] of this.map) obj[k] = e;
      writeFileSync(this.file, JSON.stringify(obj), 'utf-8');
    } catch (err) {
      getLogger().warn(`disk-cache write failed (${this.file}): ${err}`);
    }
  }
}
