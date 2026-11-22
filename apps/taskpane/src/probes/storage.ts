// 存储共享探针（Storage Probe）
// 在每个宿主中各运行一次：写入本宿主标记，再读取所有宿主的标记。
// 若在 Word 中能看到 PowerPoint 写入的标记，说明 Key 可跨宿主共享。

import type { HostName } from './env';

const LS_PREFIX = 'keynib.probe.';
const DB_NAME = 'keynib-probe';
const HOSTS: HostName[] = ['PowerPoint', 'Word', 'Excel'];

export interface StorageProbeResult {
  localStorage: { writable: boolean; seen: Record<string, string | null>; error?: string };
  indexedDB: { writable: boolean; seen: Record<string, string | null>; error?: string };
}

function openDb(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, 1);
    req.onupgradeneeded = () => req.result.createObjectStore('marks');
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

function idbPut(db: IDBDatabase, key: string, value: string): Promise<void> {
  return new Promise((resolve, reject) => {
    const tx = db.transaction('marks', 'readwrite');
    tx.objectStore('marks').put(value, key);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
}

function idbGet(db: IDBDatabase, key: string): Promise<string | null> {
  return new Promise((resolve, reject) => {
    const req = db.transaction('marks').objectStore('marks').get(key);
    req.onsuccess = () => resolve((req.result as string) ?? null);
    req.onerror = () => reject(req.error);
  });
}

export async function runStorageProbe(host: HostName): Promise<StorageProbeResult> {
  const stamp = new Date().toISOString();
  const result: StorageProbeResult = {
    localStorage: { writable: false, seen: {} },
    indexedDB: { writable: false, seen: {} },
  };

  try {
    localStorage.setItem(LS_PREFIX + host, stamp);
    result.localStorage.writable = true;
    for (const h of HOSTS) result.localStorage.seen[h] = localStorage.getItem(LS_PREFIX + h);
  } catch (e: any) {
    result.localStorage.error = String(e?.message ?? e);
  }

  try {
    const db = await openDb();
    await idbPut(db, host, stamp);
    result.indexedDB.writable = true;
    for (const h of HOSTS) result.indexedDB.seen[h] = await idbGet(db, h);
    db.close();
  } catch (e: any) {
    result.indexedDB.error = String(e?.message ?? e);
  }

  return result;
}
