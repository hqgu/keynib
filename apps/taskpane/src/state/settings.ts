// 设置存储：厂商配置、当前模型、执行模式、Key 存储方式
// Key 存储（v0.2 §8.1）：'local' 记住在本机（localStorage，防不住本机恶意软件）；'session' 仅存内存

import { useSyncExternalStore } from 'react';
import { PROVIDER_PRESETS, type Protocol } from '@keynib/providers';
import type { RunMode } from '@keynib/core';
import { DEFAULT_PROFILE, type DeckProfile } from '../hosts/powerpoint/profile';
import type { SlashCommand } from './commands';

export interface ProviderConfig {
  apiKey: string;
  baseUrl: string;
  model: string;
  models: string[];
  vision: boolean;
  name?: string; // 自定义端点的名称
  protocol?: Protocol; // 覆盖厂商默认协议（所选接口地址为 Anthropic 兼容，或自定义端点选择的格式）
}

export interface Settings {
  activeProviderId: string | null;
  providers: Record<string, ProviderConfig>;
  keyStorage: 'local' | 'session';
  mode: RunMode;
  maxSteps: number;
  profiles: DeckProfile[]; // 演示规范档案
  activeProfileId: string | null;
  pubmed: boolean; // 允许联网到 PubMed 检索 / 核对文献（只发送检索词）
  customCommands: SlashCommand[]; // 自定义快捷指令
}

export const CUSTOM_ID = 'custom';
const STORAGE_KEY = 'keynib.settings.v1';

const defaults: Settings = {
  activeProviderId: null,
  providers: {},
  keyStorage: 'local',
  mode: 'auto',
  maxSteps: 25,
  profiles: [DEFAULT_PROFILE],
  activeProfileId: null,
  pubmed: false,
  customCommands: [],
};

export function activeProfile(s: Settings): DeckProfile | null {
  return s.profiles.find((p) => p.id === s.activeProfileId) ?? null;
}

export function providerMeta(id: string): { name: string; vendor: string; protocol: Protocol; keyUrl?: string } {
  if (id === CUSTOM_ID) return { name: '自定义', vendor: 'OpenAI 兼容端点', protocol: 'openai' };
  const p = PROVIDER_PRESETS.find((x) => x.id === id)!;
  return { name: p.name, vendor: p.vendor, protocol: p.protocol, keyUrl: p.keyUrl };
}

export function emptyConfig(id: string): ProviderConfig {
  const preset = PROVIDER_PRESETS.find((p) => p.id === id);
  return { apiKey: '', baseUrl: preset?.baseUrls[0].url ?? 'http://localhost:11434/v1', model: '', models: [], vision: false };
}

function load(): Settings {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (raw) return { ...defaults, ...JSON.parse(raw) };
  } catch {
    /* 存储不可用时用默认值 */
  }
  return structuredClone(defaults);
}

let state: Settings = load();
const listeners = new Set<() => void>();

function persist() {
  const toSave: Settings =
    state.keyStorage === 'session'
      ? {
          ...state,
          providers: Object.fromEntries(Object.entries(state.providers).map(([k, v]) => [k, { ...v, apiKey: '' }])),
        }
      : state;
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(toSave));
  } catch {
    /* 忽略 */
  }
}

export function getSettings() {
  return state;
}

export function updateSettings(patch: Partial<Settings> | ((s: Settings) => Settings)) {
  state = typeof patch === 'function' ? patch(state) : { ...state, ...patch };
  persist();
  listeners.forEach((l) => l());
}

export function updateProvider(id: string, patch: Partial<ProviderConfig>) {
  updateSettings((s) => ({
    ...s,
    providers: { ...s.providers, [id]: { ...(s.providers[id] ?? emptyConfig(id)), ...patch } },
  }));
}

export function clearAllKeys() {
  updateSettings((s) => ({
    ...s,
    providers: Object.fromEntries(Object.entries(s.providers).map(([k, v]) => [k, { ...v, apiKey: '' }])),
  }));
}

export function useSettings(): Settings {
  return useSyncExternalStore(
    (l) => {
      listeners.add(l);
      return () => listeners.delete(l);
    },
    () => state,
  );
}

// 当前可用的连接（未配置完整时返回 null）
export function activeConnection(s: Settings) {
  const id = s.activeProviderId;
  if (!id) return null;
  const cfg = s.providers[id];
  if (!cfg?.apiKey || !cfg.model) return null;
  return { id, cfg, meta: providerMeta(id), conn: connectionOf(id, cfg) };
}

export function connectionOf(id: string, cfg: ProviderConfig) {
  return { protocol: cfg.protocol ?? providerMeta(id).protocol, baseUrl: cfg.baseUrl, apiKey: cfg.apiKey.trim() };
}

// ———— 配置码（Config Code）：在 PPT / Word / Excel 之间迁移设置（M0 发现三者存储不共享） ————
// 格式：KEYNIB1.<plain|enc>.<base64>
// 含 Key 时必须口令加密：PBKDF2(SHA-256, 200k) → AES-GCM-256

const enc = new TextEncoder();
const dec = new TextDecoder();
const b64 = (u: Uint8Array) => btoa(String.fromCharCode(...u));
const unb64 = (s: string) => Uint8Array.from(atob(s), (c) => c.charCodeAt(0));

async function deriveKey(pass: string, salt: Uint8Array) {
  const base = await crypto.subtle.importKey('raw', enc.encode(pass), 'PBKDF2', false, ['deriveKey']);
  return crypto.subtle.deriveKey(
    { name: 'PBKDF2', salt: salt as BufferSource, iterations: 200000, hash: 'SHA-256' },
    base,
    { name: 'AES-GCM', length: 256 },
    false,
    ['encrypt', 'decrypt'],
  );
}

export async function exportConfigCode(includeKeys: boolean, passphrase?: string): Promise<string> {
  const payload: Partial<Settings> = {
    ...state,
    providers: Object.fromEntries(
      Object.entries(state.providers).map(([k, v]) => [k, includeKeys ? v : { ...v, apiKey: '' }]),
    ),
  };
  const data = enc.encode(JSON.stringify(payload));
  if (!includeKeys) return `KEYNIB1.plain.${b64(data)}`;
  if (!passphrase) throw new Error('包含 Key 时必须设置口令');
  const salt = crypto.getRandomValues(new Uint8Array(16));
  const iv = crypto.getRandomValues(new Uint8Array(12));
  const key = await deriveKey(passphrase, salt);
  const cipher = new Uint8Array(await crypto.subtle.encrypt({ name: 'AES-GCM', iv }, key, data));
  return `KEYNIB1.enc.${b64(salt)}.${b64(iv)}.${b64(cipher)}`;
}

export function configCodeNeedsPassphrase(code: string) {
  return code.trim().startsWith('KEYNIB1.enc.');
}

export async function importConfigCode(code: string, passphrase?: string): Promise<void> {
  const parts = code.trim().split('.');
  if (parts[0] !== 'KEYNIB1') throw new Error('不是有效的配置码');
  let json: string;
  if (parts[1] === 'plain') {
    json = dec.decode(unb64(parts[2]));
  } else if (parts[1] === 'enc') {
    if (!passphrase) throw new Error('请输入口令');
    try {
      const key = await deriveKey(passphrase, unb64(parts[2]));
      json = dec.decode(await crypto.subtle.decrypt({ name: 'AES-GCM', iv: unb64(parts[3]) as BufferSource }, key, unb64(parts[4]) as BufferSource));
    } catch {
      throw new Error('口令错误或配置码已损坏');
    }
  } else {
    throw new Error('不是有效的配置码');
  }
  const incoming = JSON.parse(json) as Settings;
  // 合并：导入的非空 Key 覆盖本地，空 Key 保留本地已有的
  updateSettings((s) => ({
    ...s,
    ...incoming,
    providers: Object.fromEntries(
      [...new Set([...Object.keys(s.providers), ...Object.keys(incoming.providers ?? {})])].map((k) => {
        const a = s.providers[k];
        const b = incoming.providers?.[k];
        return [k, { ...(a ?? emptyConfig(k)), ...(b ?? {}), apiKey: b?.apiKey || a?.apiKey || '' }];
      }),
    ),
  }));
}
