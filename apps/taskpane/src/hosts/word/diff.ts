// Word 段落修改的差异计算（纯逻辑，可单元测试）
// 目标：让修订只标出真正改动的片段，而不是整段删除 + 整段插入。
// 1) 按句切分（与 Word getTextRanges 用同一组句末标点）；2) 句子按 LCS 对齐；3) 句内取公共前后缀，得到最小差异片段。

// 不含英文句点，避免把 1.5、e.g. 等切开
export const SENTENCE_MARKS = ['。', '！', '？', '；', '!', '?', ';'];

export function splitSentences(text: string): string[] {
  const out: string[] = [];
  let cur = '';
  for (const ch of text) {
    cur += ch;
    if (SENTENCE_MARKS.includes(ch)) {
      out.push(cur);
      cur = '';
    }
  }
  if (cur) out.push(cur);
  return out;
}

export type SentenceOp =
  | { type: 'keep'; old: number }
  | { type: 'modify'; old: number; text: string }
  | { type: 'delete'; old: number }
  | { type: 'insert'; after: number; text: string }; // after = -1 表示插在最前

// 句子级对齐：LCS 锚定相同句子，锚点之间按位置配对为 modify，多余的为 delete / insert
export function alignSentences(oldS: string[], newS: string[]): SentenceOp[] {
  const n = oldS.length;
  const m = newS.length;
  const dp = Array.from({ length: n + 1 }, () => new Array<number>(m + 1).fill(0));
  for (let i = n - 1; i >= 0; i--)
    for (let j = m - 1; j >= 0; j--) dp[i][j] = oldS[i] === newS[j] ? dp[i + 1][j + 1] + 1 : Math.max(dp[i + 1][j], dp[i][j + 1]);

  const anchors: [number, number][] = [];
  for (let i = 0, j = 0; i < n && j < m; ) {
    if (oldS[i] === newS[j]) {
      anchors.push([i, j]);
      i++;
      j++;
    } else if (dp[i + 1][j] >= dp[i][j + 1]) i++;
    else j++;
  }
  anchors.push([n, m]);

  const ops: SentenceOp[] = [];
  let pi = 0;
  let pj = 0;
  for (const [ai, aj] of anchors) {
    const olds = ai - pi;
    const news = aj - pj;
    const pairs = Math.min(olds, news);
    for (let k = 0; k < pairs; k++) ops.push({ type: 'modify', old: pi + k, text: newS[pj + k] });
    for (let k = pairs; k < olds; k++) ops.push({ type: 'delete', old: pi + k });
    let after = pi + pairs - 1; // 插在最后一个被修改的旧句之后；没有则在锚点前一句之后
    if (pairs === 0) after = pi - 1;
    for (let k = pairs; k < news; k++) ops.push({ type: 'insert', after, text: newS[pj + k] });
    if (ai < n) ops.push({ type: 'keep', old: ai });
    pi = ai + 1;
    pj = aj + 1;
  }
  return ops;
}

export interface Hunk {
  prefix: string; // 公共前缀
  suffix: string; // 公共后缀
  oldMid: string; // 旧句中被替换的部分（可能为空 = 纯插入）
  newMid: string; // 新内容（可能为空 = 纯删除）
}

export function minimalHunk(a: string, b: string): Hunk {
  const A = [...a];
  const B = [...b];
  let p = 0;
  while (p < A.length && p < B.length && A[p] === B[p]) p++;
  let s = 0;
  while (s < A.length - p && s < B.length - p && A[A.length - 1 - s] === B[B.length - 1 - s]) s++;
  return {
    prefix: A.slice(0, p).join(''),
    suffix: s ? A.slice(A.length - s).join('') : '',
    oldMid: A.slice(p, A.length - s).join(''),
    newMid: B.slice(p, B.length - s).join(''),
  };
}

// 在句内定位时的检索候选：差异片段不唯一时逐步带上前后文，直到唯一（由调用方验证唯一性）
export function hunkCandidates(h: Hunk, maxContext = 12): { find: string; replace: string }[] {
  const P = [...h.prefix];
  const S = [...h.suffix];
  const out: { find: string; replace: string }[] = [];
  for (let k = 0; k <= maxContext; k++) {
    const pre = P.slice(Math.max(0, P.length - k)).join('');
    const suf = S.slice(0, k).join('');
    const find = pre + h.oldMid + suf;
    if (!find) continue;
    if ([...find].length > 255) break;
    out.push({ find, replace: pre + h.newMid + suf });
    if (k >= P.length && k >= S.length) break;
  }
  return out;
}

// Word 检索中 ^ 是特殊字符
export function escapeWordSearch(s: string): string {
  return s.replace(/\^/g, '^^');
}
