// 多语言检查：①界面文件中所有中文必须经 t() 包裹；②所有 t() 的键都有英文翻译；③插值参数一致
// 工具文件只检查 label / summary（description 等写给模型的内容不翻译）。
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { EN } from '../src/i18n/en.ts';
import { setLang, t } from '../src/i18n/index.ts';

const SRC = path.resolve(import.meta.dirname, '../src');
const UI_FILES = ['App.tsx', 'views/ChatView.tsx', 'views/SettingsView.tsx', 'views/HistoryView.tsx', 'views/ProfileSection.tsx', 'views/CommandSection.tsx', 'components/Cards.tsx', 'components/SlashMenu.tsx'];
const TOOL_FILES = ['hosts/powerpoint/tools.ts', 'hosts/word/tools.ts', 'hosts/excel/tools.ts'];
const CJK = /[一-龥]/;

// 去掉注释
const strip = (src: string) => src.replace(/\/\*[\s\S]*?\*\//g, '').replace(/(^|[^:'"`])\/\/.*$/gm, '$1').replace(/\{\/\*[\s\S]*?\*\/\}/g, '');

// 找出未经 t() 包裹的中文：字符串字面量或 JSX 文本
function untranslated(src: string): string[] {
  const out: string[] = [];
  const s = strip(src);
  const lit = /((?:\bt|\bm)\(\s*)?(['"`])((?:\\.|(?!\2)[^\\])*?)\2/g;
  let m: RegExpExecArray | null;
  // 模板字符串中 ${…} 内部（如 ${t('…')}）不算未翻译
  while ((m = lit.exec(s))) if (!m[1] && CJK.test(m[3].replace(/\$\{[^}]*\}/g, ''))) out.push(m[3]);
  const jsx = />([^<>{}]*[一-龥][^<>{}]*)</g;
  while ((m = jsx.exec(s))) out.push(m[1].trim());
  return out;
}

function keysOf(src: string): string[] {
  const out: string[] = [];
  const re = /\bt\(\s*(['"`])((?:\\.|(?!\1)[^\\])*?)\1/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(src))) out.push(m[2].replace(/\\(['"`\\])/g, '$1'));
  return out;
}

test('界面文件中的中文都经 t() 包裹', () => {
  const bad: string[] = [];
  for (const f of UI_FILES) for (const x of untranslated(fs.readFileSync(path.join(SRC, f), 'utf8'))) bad.push(`${f}: ${x}`);
  for (const f of TOOL_FILES) {
    const s = strip(fs.readFileSync(path.join(SRC, f), 'utf8'));
    for (const line of s.split('\n')) if (/\b(label|summary):/.test(line)) for (const x of untranslated(line)) bad.push(`${f}: ${x}`);
  }
  assert.deepEqual(bad, []);
});

test('所有 t() 的键都有英文翻译，且插值参数一致', () => {
  const missing: string[] = [];
  const mismatch: string[] = [];
  const files = [...UI_FILES, ...TOOL_FILES, 'state/commands.ts', 'views/Diagnostics.tsx'];
  for (const f of files) {
    for (const k of keysOf(fs.readFileSync(path.join(SRC, f), 'utf8'))) {
      if (k.includes('${')) mismatch.push(`${f}: 键中不能含 \${}：${k}`);
      else if (!(k in EN)) missing.push(`${f}: ${k}`);
      else {
        const p = (x: string) => [...x.matchAll(/\{(\w+)\}/g)].map((m) => m[1]).sort().join(',');
        if (p(k) !== p(EN[k])) mismatch.push(`${f}: 参数不一致：${k}`);
      }
    }
  }
  assert.deepEqual(missing, []);
  assert.deepEqual(mismatch, []);
});

test('英文译文不含中文', () => {
  assert.deepEqual(Object.entries(EN).filter(([, v]) => CJK.test(v)).map(([k]) => k), []);
});

test('t() 切换语言与插值', () => {
  setLang('en');
  assert.equal(t('共 {n} 页', { n: 3 }), '3 slides');
  setLang('zh');
  assert.equal(t('共 {n} 页', { n: 3 }), '共 3 页');
});
