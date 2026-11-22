// PubMed 核对单元测试（假 fetch，不联网）
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { expandPages, formatShortCite, parseShortCite, verifyCitation } from '../src/pubmed.ts';

test('页码展开', () => {
  assert.equal(expandPages('962-70'), '962-970');
  assert.equal(expandPages('144-6'), '144-146');
  assert.equal(expandPages('1348-53'), '1348-1353');
  assert.equal(expandPages('109-112'), '109-112');
  assert.equal(expandPages('c332'), 'c332');
});

test('短引格式化（无期号时省略括号）', () => {
  const base = { pmid: '1', firstAuthor: 'Muller CJ', journal: 'Int J Epidemiol', year: '2014', volume: '43', issue: '3', pages: '962-970', title: 't' };
  assert.equal(formatShortCite(base), 'Muller CJ. Int J Epidemiol. 2014; 43 (3): 962-970.');
  assert.equal(formatShortCite({ ...base, issue: '' }), 'Muller CJ. Int J Epidemiol. 2014; 43: 962-970.');
});

test('解析短引', () => {
  assert.deepEqual(parseShortCite('Muller CJ. Int J Epidemiol. 2014; 43 (3): 962-970.'), {
    author: 'Muller', initials: 'CJ', journal: 'Int J Epidemiol', year: '2014', volume: '43', issue: '3', firstPage: '962',
  });
  assert.equal(parseShortCite('Ranganathan P, et al. Perspect Clin Res. 2016; 7 (3): 144-146.')?.firstPage, '144');
  assert.equal(parseShortCite('Schulz KF. BMJ. 2010; 340: c332.')?.firstPage, 'c332');
  assert.equal(parseShortCite('Muller, C. J. (2014). Int J Epidemiol'), null);
});

// 模拟 E-utilities
function mockPubmed(searchResults: string[][], summary: Record<string, any>) {
  const terms: string[] = [];
  globalThis.fetch = (async (url: string) => {
    const u = new URL(url);
    if (u.pathname.endsWith('esearch.fcgi')) {
      terms.push(u.searchParams.get('term')!);
      return new Response(JSON.stringify({ esearchresult: { idlist: searchResults.shift() ?? [] } }));
    }
    const ids = u.searchParams.get('id')!.split(',');
    return new Response(JSON.stringify({ result: { uids: ids, ...Object.fromEntries(ids.map((i) => [i, summary[i]])) } }));
  }) as typeof fetch;
  return terms;
}

const GUPTA = { sortfirstauthor: 'Gupta SK', source: 'Perspect Clin Res', pubdate: '2011 Jul', volume: '2', issue: '3', pages: '109-12', title: 'Intention-to-treat concept: A review.' };

test('核对：完全一致 → verified', async () => {
  const terms = mockPubmed([['21897887']], { '21897887': GUPTA });
  const r = await verifyCitation('Gupta SK. Perspect Clin Res. 2011; 2 (3): 109-112.');
  assert.equal(r.status, 'verified');
  assert.equal(r.record?.pmid, '21897887');
  assert.equal(terms[0], 'Gupta SK[au] AND 2011[dp] AND 2[vi] AND 109[pg]');
});

test('核对：文献存在但写法有误 → corrected 并给出正确写法', async () => {
  mockPubmed([['21897887']], { '21897887': GUPTA });
  const r = await verifyCitation('Gupta SK. Perspect Clin Res. 2011; 2 (4): 109-112.');
  assert.equal(r.status, 'corrected');
  assert.equal(r.record?.shortCite, 'Gupta SK. Perspect Clin Res. 2011; 2 (3): 109-112.');
});

test('核对：查不到 → not_found，并给出候选', async () => {
  mockPubmed([[], ['21897887']], { '21897887': GUPTA });
  const r = await verifyCitation('Gupta SK. Perspect Clin Res. 2011; 9 (9): 999-1000.');
  assert.equal(r.status, 'not_found');
  assert.equal(r.candidates?.length, 1);
});
