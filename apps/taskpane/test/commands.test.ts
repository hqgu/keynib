// 快捷指令单元测试
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { availableCommands, filterCommands, lintPrompt, slashQuery, type SlashCommand } from '../src/state/commands.ts';

const ctx = { host: 'PowerPoint', hasProfile: true, pubmed: true };

test('只有以 / 开头且未输入空白时才触发', () => {
  assert.equal(slashQuery('/'), '');
  assert.equal(slashQuery('/rs'), 'rs');
  assert.equal(slashQuery('/润色'), '润色');
  assert.equal(slashQuery('/rs 第三页'), null);
  assert.equal(slashQuery('请 /rs'), null);
  assert.equal(slashQuery(''), null);
});

test('按名称、拼音首字母、英文别名、描述筛选', () => {
  const all = availableCommands([], ctx);
  assert.equal(filterCommands(all, '润')[0].id, 'polish');
  assert.equal(filterCommands(all, 'rs')[0].id, 'polish');
  assert.equal(filterCommands(all, 'polish')[0].id, 'polish');
  assert.equal(filterCommands(all, 'gf')[0].id, 'lint');
  assert.ok(filterCommands(all, 'PubMed').some((c) => c.id === 'cite'));
  assert.equal(filterCommands(all, 'zzzz').length, 0);
  assert.equal(filterCommands(all, '').length, all.length);
});

test('按宿主与前置条件过滤', () => {
  const word = availableCommands([], { host: 'Word', hasProfile: true, pubmed: false });
  assert.ok(word.some((c) => c.id === 'polish'));
  assert.ok(word.some((c) => c.id === 'academic'));
  assert.ok(!availableCommands([], ctx).some((c) => c.id === 'academic')); // Word 专用
  assert.ok(!word.some((c) => c.id === 'outline')); // PPT 专用
  assert.ok(!word.some((c) => c.id === 'cite')); // 需要 PubMed
  const noProfile = availableCommands([], { host: 'PowerPoint', hasProfile: false, pubmed: true });
  assert.ok(!noProfile.some((c) => c.id === 'lint'));
});

test('自定义指令同名覆盖内置，并追加到列表', () => {
  const mine: SlashCommand[] = [
    { id: 'u1', name: '润色', aliases: ['rs'], description: '我的润色', prompt: '按我的风格润色', autoSend: true },
    { id: 'u2', name: '课堂提问', aliases: ['tw'], description: '每页加一个思考题', prompt: '为每页加一个思考题', autoSend: true },
  ];
  const all = availableCommands(mine, ctx);
  assert.equal(all.filter((c) => c.name === '润色').length, 1);
  assert.equal(all.find((c) => c.name === '润色')!.prompt, '按我的风格润色');
  assert.equal(filterCommands(all, 'tw')[0].id, 'u2');
});

test('规范检查指令随 PubMed 开关包含引用核对', () => {
  assert.match(lintPrompt(true), /verify_citation/);
  assert.doesNotMatch(lintPrompt(false), /verify_citation/);
});
