// 界面多语言：以中文原文为键，英文见 en.ts；缺翻译时回退中文。
// 语言设置：auto（跟随 Office 显示语言，中文环境显示中文，其他显示英文）/ zh / en。
// 写给模型的内容（工具说明、提示词）不翻译——模型会按用户提问的语言回复。

import { EN } from './en.ts';

export type Lang = 'zh' | 'en';
export type LangSetting = 'auto' | Lang;

let current: Lang = 'zh';

export function resolveLang(setting: LangSetting | undefined): Lang {
  if (setting === 'zh' || setting === 'en') return setting;
  const office = (globalThis as any).Office?.context?.displayLanguage as string | undefined;
  const nav = (globalThis as any).navigator?.language as string | undefined;
  return /^zh/i.test(office || nav || 'zh') ? 'zh' : 'en';
}

export function setLang(l: Lang) {
  current = l;
}

export function getLang(): Lang {
  return current;
}

// t('第 {n} 页', { n: 3 })
export function t(zh: string, params?: Record<string, string | number>): string {
  let s = current === 'en' ? EN[zh] ?? zh : zh;
  if (params) for (const [k, v] of Object.entries(params)) s = s.split(`{${k}}`).join(String(v));
  return s;
}

// 标记不翻译的文字（写给模型的提示词 / 事件说明，或本身双语的标签）：多语言检查会跳过
export const m = (s: string) => s;
