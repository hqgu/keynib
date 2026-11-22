// 会话历史：IndexedDB，按文档分组（v0.2 §7.2）
// docId：文档已保存时取 URL 的哈希；未保存的新文档归入"未保存的文档"

import type { ChatMessage } from '@keynib/providers';

export interface Session {
  id: string;
  docId: string;
  docName: string;
  host: string;
  title: string;
  providerId: string;
  model: string;
  createdAt: number;
  updatedAt: number;
  messages: ChatMessage[]; // 存储时去掉图片，避免数据库膨胀
  usage: { inputTokens: number; outputTokens: number };
}

const DB = 'keynib';
const STORE = 'sessions';

function open(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB, 1);
    req.onupgradeneeded = () => {
      const s = req.result.createObjectStore(STORE, { keyPath: 'id' });
      s.createIndex('docId', 'docId');
      s.createIndex('updatedAt', 'updatedAt');
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

async function tx<T>(mode: IDBTransactionMode, fn: (s: IDBObjectStore) => IDBRequest<T>): Promise<T> {
  const db = await open();
  return new Promise((resolve, reject) => {
    const req = fn(db.transaction(STORE, mode).objectStore(STORE));
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  }).finally(() => db.close()) as Promise<T>;
}

function stripImages(messages: ChatMessage[]): ChatMessage[] {
  return messages.map((m) => ('images' in m && m.images?.length ? { ...m, images: undefined, content: m.content + '\n[图片未保存]' } : m));
}

export async function saveSession(s: Session) {
  await tx('readwrite', (st) => st.put({ ...s, messages: stripImages(s.messages) }));
}

export async function listSessions(): Promise<Session[]> {
  const all = await tx<Session[]>('readonly', (st) => st.getAll());
  return all.sort((a, b) => b.updatedAt - a.updatedAt);
}

export async function deleteSession(id: string) {
  await tx('readwrite', (st) => st.delete(id));
}

export async function clearSessions() {
  await tx('readwrite', (st) => st.clear());
}

export async function currentDoc(): Promise<{ docId: string; docName: string }> {
  const url = Office.context?.document?.url;
  if (!url) return { docId: 'unsaved', docName: '未保存的文档' };
  const digest = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(url));
  const hex = [...new Uint8Array(digest)].slice(0, 8).map((b) => b.toString(16).padStart(2, '0')).join('');
  const name = decodeURIComponent(url.split(/[\\/]/).pop() ?? url);
  return { docId: hex, docName: name };
}

export function newSessionId() {
  return crypto.randomUUID();
}

export function sessionToMarkdown(s: Session): string {
  const lines = [`# ${s.title}`, '', `- 文档：${s.docName}`, `- 模型：${s.providerId} / ${s.model}`, `- 时间：${new Date(s.createdAt).toLocaleString()}`, ''];
  for (const m of s.messages) {
    if (m.role === 'user') lines.push(`## 用户`, '', m.content, '');
    else if (m.role === 'assistant') {
      if (m.content) lines.push(`## 妙笔`, '', m.content, '');
      for (const c of m.toolCalls ?? []) lines.push(`- 工具 \`${c.name}\` ${c.arguments}`);
    } else lines.push(`  - 结果：${m.content.slice(0, 200)}`);
  }
  return lines.join('\n');
}
