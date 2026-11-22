// Word 差异算法单元测试
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { alignSentences, escapeWordSearch, hunkCandidates, minimalHunk, splitSentences } from '../src/hosts/word/diff.ts';

const P = '本研究采用随机对照试验（RCT）设计。受试者按 1:1 比例随机分配至干预组和对照组。主要结局为 12 周时的收缩压变化。';

test('按句切分（不切英文小数点）', () => {
  assert.deepEqual(splitSentences(P), ['本研究采用随机对照试验（RCT）设计。', '受试者按 1:1 比例随机分配至干预组和对照组。', '主要结局为 12 周时的收缩压变化。']);
  assert.deepEqual(splitSentences('OR = 1.5; 95% CI 1.2-1.9'), ['OR = 1.5;', ' 95% CI 1.2-1.9']);
  assert.deepEqual(splitSentences('没有句号'), ['没有句号']);
});

test('句子对齐：只改中间一句', () => {
  const o = splitSentences(P);
  const n = [...o];
  n[1] = '受试者按 1:1 比例经中心随机系统分配至干预组或对照组。';
  assert.deepEqual(alignSentences(o, n).map((x) => x.type), ['keep', 'modify', 'keep']);
});

test('句子对齐：插入、删除', () => {
  const o = ['A。', 'B。', 'C。'];
  assert.deepEqual(alignSentences(o, ['A。', 'X。', 'B。', 'C。']), [
    { type: 'keep', old: 0 },
    { type: 'insert', after: 0, text: 'X。' },
    { type: 'keep', old: 1 },
    { type: 'keep', old: 2 },
  ]);
  assert.deepEqual(alignSentences(o, ['A。', 'C。']), [{ type: 'keep', old: 0 }, { type: 'delete', old: 1 }, { type: 'keep', old: 2 }]);
  assert.deepEqual(alignSentences(o, ['X。', 'A。', 'B。', 'C。'])[0], { type: 'insert', after: -1, text: 'X。' });
});

test('最小差异片段', () => {
  assert.deepEqual(minimalHunk('随机分配至干预组和对照组。', '经中心随机系统分配至干预组或对照组。'), {
    prefix: '',
    suffix: '对照组。',
    oldMid: '随机分配至干预组和',
    newMid: '经中心随机系统分配至干预组或',
  });
  assert.deepEqual(minimalHunk('干预组和对照组', '干预组与对照组'), { prefix: '干预组', suffix: '对照组', oldMid: '和', newMid: '与' });
  assert.deepEqual(minimalHunk('收缩压变化。', '收缩压的变化。'), { prefix: '收缩压', suffix: '变化。', oldMid: '', newMid: '的' });
});

test('检索候选：逐步带上下文；纯插入也有可检索文本', () => {
  const c = hunkCandidates(minimalHunk('干预组和对照组和其他', '干预组与对照组和其他'));
  assert.equal(c[0].find, '和');
  assert.equal(c[1].find, '组和对');
  assert.equal(c[1].replace, '组与对');
  const ins = hunkCandidates(minimalHunk('收缩压变化。', '收缩压的变化。'));
  assert.equal(ins[0].find, '压变');
  assert.equal(ins[0].replace, '压的变');
});

test('转义 ^', () => {
  assert.equal(escapeWordSearch('a^b'), 'a^^b');
});
