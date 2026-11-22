// 规范检查单元测试
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { DEFAULT_PROFILE as P, isValidShortCitation, lintSlide, profileFontFor, profilePrompt, type LintShape } from '../src/hosts/powerpoint/profile.ts';

const shape = (o: Partial<LintShape>): LintShape => ({ id: '1', name: 'TextBox 1', type: 'TextBox', left: 60, top: 140, width: 840, height: 300, ...o });

test('合规页面没有错误', () => {
  const issues = lintSlide(2, [
    shape({ id: '2', name: 'Title 1', text: '盲法', font: { name: '微软雅黑', size: 36, bold: true }, top: 30, height: 80 }),
    shape({ id: '3', name: 'Content Placeholder 2', text: '单盲\n双盲\n三盲', font: { name: '微软雅黑', size: 20 } }),
  ], P);
  assert.deepEqual(issues.filter((i) => i.severity !== 'info'), []);
});

test('检出：西文字体、小于最小字号、标题字号与粗细', () => {
  const issues = lintSlide(3, [
    shape({ id: '2', name: 'Title 1', text: '混杂', font: { name: 'Aptos Display', size: 44, bold: false }, top: 30, height: 80 }),
    shape({ id: '3', name: 'Content Placeholder 2', text: '定义', font: { name: '微软雅黑', size: 14 } }),
  ], P);
  const rules = issues.map((i) => `${i.shapeId}:${i.rule}`).sort();
  assert.deepEqual(rules, ['2:font', '2:titleBold', '2:titleSize', '3:minSize']);
});

test('副标题不算标题', () => {
  const issues = lintSlide(1, [shape({ name: 'Subtitle 2', text: 'RCT', font: { name: '微软雅黑', size: 20, bold: false } })], P);
  assert.equal(issues.some((i) => i.rule === 'titleBold'), false);
});

test('超出页面', () => {
  const issues = lintSlide(2, [shape({ left: 700, width: 400, text: 'x', font: { name: '微软雅黑', size: 20 } })], P);
  assert.ok(issues.some((i) => i.rule === 'offSlide'));
});

test('页脚引用格式', () => {
  assert.equal(isValidShortCitation('Muller CJ. Int J Epidemiol. 2014; 43 (3): 962-970.'), true);
  assert.equal(isValidShortCitation('Schulz KF, et al. Lancet. 2002; 359: 614-618.'), true);
  assert.equal(isValidShortCitation('Muller, C. J. (2014). Int J Epidemiol, 43(3), 962–970.'), false);
  const issues = lintSlide(2, [shape({ top: 500, height: 30, text: 'Muller, C. J. (2014). Int J Epidemiol, 43(3), 962–970.', font: { name: '微软雅黑', size: 16 } })], P);
  assert.ok(issues.some((i) => i.rule === 'citation'));
});

test('留白较多提示（仅非标题页）', () => {
  const s = [shape({ name: 'Title 1', text: 'A', top: 30, height: 60, font: { name: '微软雅黑', size: 36, bold: true } }), shape({ text: 'b', top: 100, height: 60, font: { name: '微软雅黑', size: 20 } })];
  assert.ok(lintSlide(2, s, P).some((i) => i.rule === 'whitespace'));
  assert.equal(lintSlide(1, s, P).some((i) => i.rule === 'whitespace'), false);
});

test('apply_profile 的修正规则', () => {
  assert.deepEqual(profileFontFor(shape({ name: 'Title 1', font: { size: 44 } }), P, { setBodySize: false }), { name: '微软雅黑', size: 36, bold: true });
  assert.deepEqual(profileFontFor(shape({ font: { size: 12 } }), P, { setBodySize: false }), { name: '微软雅黑', size: 16 });
  assert.deepEqual(profileFontFor(shape({ font: { size: 28 } }), P, { setBodySize: true }), { name: '微软雅黑', size: 20 });
  assert.deepEqual(profileFontFor(shape({ font: { size: 28 } }), P, { setBodySize: false }), { name: '微软雅黑', size: undefined });
});

test('规范提示词包含关键约束', () => {
  const t = profilePrompt(P);
  assert.match(t, /标题 36 磅加粗/);
  assert.match(t, /不小于 16 磅/);
  assert.match(t, /Muller CJ/);
  assert.match(t, /不要编造/);
});

test('开启 PubMed 时要求引用来自检索结果', () => {
  assert.match(profilePrompt(P, { pubmed: true }), /pubmed_search/);
  assert.doesNotMatch(profilePrompt(P), /pubmed_search/);
});

test('“[待核实]”标记不算格式错误，单独提示', () => {
  const issues = lintSlide(3, [shape({ top: 500, height: 30, text: 'Zhang W. Lancet. 2019; 393 (10170): 1234-1240. [待核实]', font: { name: '微软雅黑', size: 16 } })], P);
  const c = issues.filter((i) => i.rule === 'citation');
  assert.deepEqual(c.map((i) => i.severity), ['info']);
});
