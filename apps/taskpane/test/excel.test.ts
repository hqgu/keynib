// Excel 纯逻辑单元测试
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { cellCount, colLetter, describeNumbers, findErrors, profileColumns, quoteSheet } from '../src/hosts/excel/stats.ts';
import { SEED_HEADER, seedTrial } from '../src/hosts/excel/seed.ts';

test('列字母与单元格计数', () => {
  assert.deepEqual([0, 25, 26, 27, 701, 702].map(colLetter), ['A', 'Z', 'AA', 'AB', 'ZZ', 'AAA']);
  assert.equal(cellCount('A1:B3'), 6);
  assert.equal(cellCount("'数据'!$C$2:$C$4"), 3);
  assert.equal(cellCount('F7'), 1);
});

test('表名引用', () => {
  assert.equal(quoteSheet('Sheet1'), 'Sheet1');
  assert.equal(quoteSheet('描述统计'), "'描述统计'");
  assert.equal(quoteSheet('Data 1'), "'Data 1'");
  assert.equal(quoteSheet("Bob's"), "'Bob''s'");
});

test('分位数与 Excel 一致（MEDIAN / QUARTILE.INC / STDEV.S）', () => {
  const d = describeNumbers([1, 2, 3, 4, 5, 6, 7, 8]);
  assert.equal(d.median, 4.5);
  assert.equal(d.q1, 2.75);
  assert.equal(d.q3, 6.25);
  assert.ok(Math.abs(d.sd! - 2.449489743) < 1e-9);
});

test('列类型推断：数字 / 日期 / 文本 / 缺失', () => {
  const p = profileColumns(['组', '年龄', '日期', '空'], [['A', 50, 45000, ''], ['B', '', 45001, ''], ['A', 60, 45002, '']], [['General', 'General', 'yyyy-mm-dd', 'General']]);
  assert.deepEqual(p.map((c) => c.type), ['text', 'number', 'date', 'empty']);
  assert.equal(p[1].missing, 1);
  assert.equal(p[1].mean, 55);
  assert.deepEqual(p[0].levels, [{ value: 'A', count: 2 }, { value: 'B', count: 1 }]);
});

test('错误单元格定位', () => {
  assert.deepEqual(findErrors([['x', '#DIV/0!'], ['#REF!', 1]], { row: 1, col: 2 }), ['D2: #DIV/0!', 'C3: #REF!']);
});

test('种子数据固定可复现，含缺失', () => {
  const a = seedTrial();
  assert.equal(a.length, 200);
  assert.deepEqual(a, seedTrial());
  assert.equal(SEED_HEADER.length, a[0].length);
  assert.ok(a.some((r) => r[5] === ''));
});
