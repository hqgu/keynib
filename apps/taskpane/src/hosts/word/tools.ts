// Word 工具集（M2）。所有写入在修订模式下进行，用户可在 Word"审阅"中逐条处理。
// 段落用 1 起的序号 paragraph 指代（含空段落与表格内段落）；修改前须带上该段开头文字 oldStart 作校验，防止按过时序号误改。

import type { ToolDef } from '@keynib/core';
import { alignSentences, escapeWordSearch, hunkCandidates, minimalHunk, SENTENCE_MARKS, splitSentences } from './diff';

const json = (v: unknown) => JSON.stringify(v);
const clip = (s: string, n: number) => (s.length > n ? s.slice(0, n) + '…' : s);
const stripCr = (s: string) => s.replace(/\r$/, '');

async function paragraphs(ctx: Word.RequestContext) {
  const ps = ctx.document.body.paragraphs;
  ps.load('items/text,items/style,items/isListItem,items/tableNestingLevel');
  await ctx.sync();
  return ps.items;
}

function check(items: Word.Paragraph[], n: number, oldStart?: string) {
  const p = items[n - 1];
  if (!p) throw new Error(`第 ${n} 段不存在（共 ${items.length} 段）`);
  if (oldStart !== undefined) {
    const want = oldStart.trim().slice(0, 15);
    if (want && !p.text.trim().startsWith(want)) {
      throw new Error(`第 ${n} 段的开头不是「${want}」而是「${clip(p.text.trim(), 20)}」，文档可能已变化，请重新读取后再改`);
    }
  }
  return p;
}

// 在 range 内把 find 唯一命中处替换为 replace；不唯一或检索失败返回 false
async function replaceUnique(ctx: Word.RequestContext, range: Word.Range, find: string, replace: string): Promise<boolean> {
  try {
    const hits = range.search(escapeWordSearch(find), { matchCase: true });
    hits.load('items');
    await ctx.sync();
    if (hits.items.length !== 1) return false;
    if (replace) hits.items[0].insertText(replace, Word.InsertLocation.replace);
    else hits.items[0].delete();
    await ctx.sync();
    return true;
  } catch {
    return false;
  }
}

// 把一段改成 newText，尽量只修订改动的片段。返回采用的粒度统计。
async function editParagraph(ctx: Word.RequestContext, p: Word.Paragraph, newText: string) {
  const stats = { phrase: 0, sentence: 0, paragraph: 0 };
  const oldText = stripCr(p.text);
  if (oldText === newText) return stats;

  const ranges = p.getTextRanges(SENTENCE_MARKS, false);
  ranges.load('items/text');
  await ctx.sync();
  const oldS = ranges.items.map((r) => stripCr(r.text));
  // Word 的切分与我们不一致时，直接整段替换
  if (oldS.join('') !== oldText || splitSentences(oldText).length !== oldS.length) {
    p.insertText(newText, Word.InsertLocation.replace);
    await ctx.sync();
    stats.paragraph++;
    return stats;
  }

  const ops = alignSentences(oldS, splitSentences(newText));
  // 从后往前执行，前面句子的位置不受影响
  for (const op of [...ops].reverse()) {
    if (op.type === 'keep') continue;
    if (op.type === 'delete') {
      if (!(await replaceUnique(ctx, ranges.items[op.old], oldS[op.old], ''))) {
        p.insertText(newText, Word.InsertLocation.replace);
        await ctx.sync();
        return { phrase: 0, sentence: 0, paragraph: 1 };
      }
      stats.sentence++;
    } else if (op.type === 'insert') {
      // 最后一句的范围包含段落标记，插到它"之后"会跑到下一段，所以改为插在段末
      if (op.after === ranges.items.length - 1) p.insertText(op.text, Word.InsertLocation.end);
      else if (op.after >= 0) ranges.items[op.after].insertText(op.text, Word.InsertLocation.after);
      else p.insertText(op.text, Word.InsertLocation.start);
      await ctx.sync();
      stats.sentence++;
    } else {
      const range = ranges.items[op.old];
      let done = false;
      for (const c of hunkCandidates(minimalHunk(oldS[op.old], op.text))) {
        if (await replaceUnique(ctx, range, c.find, c.replace)) {
          done = true;
          stats.phrase++;
          break;
        }
      }
      if (!done && (await replaceUnique(ctx, range, oldS[op.old], op.text))) {
        done = true;
        stats.sentence++;
      }
      if (!done) {
        p.insertText(newText, Word.InsertLocation.replace);
        await ctx.sync();
        return { phrase: 0, sentence: 0, paragraph: 1 };
      }
    }
  }
  return stats;
}

const STYLES = ['Normal', 'Heading1', 'Heading2', 'Heading3', 'Heading4', 'Title', 'Subtitle', 'Quote', 'ListParagraph', 'Caption'];

export function createWordTools(): ToolDef[] {
  return [
    {
      name: 'get_outline',
      label: '读取大纲',
      permission: 'R',
      description: '读取文档概况：段落总数、字数，以及各级标题及其段落序号。先调用它了解结构。',
      parameters: { type: 'object', properties: {} },
      run: () =>
        Word.run(async (ctx) => {
          const items = await paragraphs(ctx);
          const headings = items
            .map((p, i) => ({ paragraph: i + 1, style: p.style, text: clip(p.text.trim(), 60) }))
            .filter((x) => /heading|标题|title/i.test(x.style));
          const chars = items.reduce((n, p) => n + p.text.length, 0);
          return { text: json({ paragraphs: items.length, chars, headings }), summary: `${items.length} 段，${headings.length} 个标题` };
        }),
    },
    {
      name: 'read_paragraphs',
      label: '读取段落',
      permission: 'R',
      description: '读取第 from 到 to 段（含）的序号、样式和全文。一次最多 60 段；长文档请分批读取。修改前必须先读取。',
      parameters: { type: 'object', properties: { from: { type: 'integer' }, to: { type: 'integer' } }, required: ['from', 'to'] },
      run: ({ from, to }: { from: number; to: number }) =>
        Word.run(async (ctx) => {
          const items = await paragraphs(ctx);
          const a = Math.max(1, from);
          const b = Math.min(items.length, to, a + 59);
          const out = [];
          for (let i = a; i <= b; i++) {
            const p = items[i - 1];
            out.push({ paragraph: i, style: p.style, inTable: p.tableNestingLevel > 0 || undefined, list: p.isListItem || undefined, text: stripCr(p.text) });
          }
          return { text: json({ total: items.length, paragraphs: out }), summary: `第 ${a}–${b} 段` };
        }),
    },
    {
      name: 'get_selection',
      label: '读取当前选择',
      permission: 'R',
      description: '获取用户当前选中的文字及其所在段落序号。用户说"这段""选中的"时调用。没有选中时返回光标所在段落。',
      parameters: { type: 'object', properties: {} },
      run: () =>
        Word.run(async (ctx) => {
          const sel = ctx.document.getSelection();
          sel.load('text');
          const first = sel.paragraphs.getFirst();
          const last = sel.paragraphs.getLast();
          const items = ctx.document.body.paragraphs;
          items.load('items');
          await ctx.sync();
          const cmpFirst = items.items.map((p) => p.getRange('Whole').compareLocationWith(first.getRange('Whole')));
          const cmpLast = items.items.map((p) => p.getRange('Whole').compareLocationWith(last.getRange('Whole')));
          await ctx.sync();
          const a = cmpFirst.findIndex((c) => c.value === 'Equal') + 1;
          const b = cmpLast.findIndex((c) => c.value === 'Equal') + 1;
          return {
            text: json({ selectedText: clip(sel.text, 2000), paragraphs: a ? [a, b || a] : [] }),
            summary: sel.text ? `选中 ${sel.text.length} 字（第 ${a}–${b} 段）` : `光标在第 ${a} 段`,
          };
        }),
    },
    {
      name: 'search_text',
      label: '查找',
      permission: 'R',
      description: '在全文查找文字（≤ 255 字，区分大小写），返回命中的段落序号和上下文。',
      parameters: { type: 'object', properties: { query: { type: 'string' } }, required: ['query'] },
      run: ({ query }: { query: string }) =>
        Word.run(async (ctx) => {
          if ([...query].length > 255) throw new Error('查找内容不能超过 255 字');
          const items = await paragraphs(ctx);
          const hits = ctx.document.body.search(escapeWordSearch(query), { matchCase: true });
          hits.load('items');
          await ctx.sync();
          const firsts = hits.items.slice(0, 30).map((h) => h.paragraphs.getFirst());
          firsts.forEach((f) => f.load('text'));
          await ctx.sync();
          const out = firsts.map((f) => {
            const idx = items.findIndex((p) => p.text === f.text) + 1;
            const t = stripCr(f.text);
            const at = t.indexOf(query);
            return { paragraph: idx || null, context: clip(t.slice(Math.max(0, at - 30), at + query.length + 30), 100) };
          });
          return { text: json({ count: hits.items.length, hits: out }), summary: `${hits.items.length} 处` };
        }),
    },
    {
      name: 'edit_paragraph',
      label: '修改段落',
      permission: 'W',
      description:
        '把第 paragraph 段改为 newText（给出修改后的完整段落文字）。系统会自动只把真正改动的词句标为修订。oldStart 填该段当前开头的 5–15 个字用于校验。只改需要改的地方，保持原文其余部分一字不差。',
      parameters: {
        type: 'object',
        properties: { paragraph: { type: 'integer' }, oldStart: { type: 'string' }, newText: { type: 'string' } },
        required: ['paragraph', 'oldStart', 'newText'],
      },
      run: ({ paragraph, oldStart, newText }: { paragraph: number; oldStart: string; newText: string }) =>
        Word.run(async (ctx) => {
          const items = await paragraphs(ctx);
          const p = check(items, paragraph, oldStart);
          if (p.tableNestingLevel > 0) throw new Error('表格内的段落暂不支持用此工具修改');
          const stats = await editParagraph(ctx, p, newText.replace(/\r?\n/g, ' ').trim());
          const how = stats.paragraph ? '整段' : stats.sentence ? `${stats.phrase} 处词句、${stats.sentence} 句` : `${stats.phrase} 处词句`;
          return { text: json({ ok: true, granularity: stats }), summary: `第 ${paragraph} 段：${stats.phrase + stats.sentence + stats.paragraph ? how : '无变化'}` };
        }),
    },
    {
      name: 'insert_paragraphs',
      label: '插入段落',
      permission: 'W',
      description: '在第 after 段之后插入若干段（after=0 表示插在文首）。style 可选：' + STYLES.join(' / '),
      parameters: {
        type: 'object',
        properties: {
          after: { type: 'integer' },
          oldStart: { type: 'string', description: '第 after 段开头文字，用于校验（after=0 时可省略）' },
          paragraphs: {
            type: 'array',
            items: { type: 'object', properties: { text: { type: 'string' }, style: { type: 'string', enum: STYLES } }, required: ['text'] },
          },
        },
        required: ['after', 'paragraphs'],
      },
      run: ({ after, oldStart, paragraphs: add }: { after: number; oldStart?: string; paragraphs: { text: string; style?: string }[] }) =>
        Word.run(async (ctx) => {
          const items = await paragraphs(ctx);
          let anchor: Word.Paragraph;
          let loc: 'After' | 'Before' = 'After';
          if (after <= 0) {
            anchor = items[0];
            loc = 'Before';
          } else anchor = check(items, after, oldStart);
          const inserted: Word.Paragraph[] = [];
          for (const x of loc === 'Before' ? [...add].reverse() : add) {
            const np = (inserted.length && loc === 'After' ? inserted[inserted.length - 1] : anchor).insertParagraph(x.text, loc);
            if (x.style) np.styleBuiltIn = x.style as Word.BuiltInStyleName;
            inserted.push(np);
          }
          await ctx.sync();
          return { text: json({ ok: true, inserted: add.length }), summary: `第 ${after} 段后插入 ${add.length} 段` };
        }),
    },
    {
      name: 'set_style',
      label: '设置样式',
      permission: 'W',
      description: '设置第 paragraph 段的内置样式：' + STYLES.join(' / '),
      parameters: {
        type: 'object',
        properties: { paragraph: { type: 'integer' }, oldStart: { type: 'string' }, style: { type: 'string', enum: STYLES } },
        required: ['paragraph', 'oldStart', 'style'],
      },
      run: ({ paragraph, oldStart, style }: { paragraph: number; oldStart: string; style: string }) =>
        Word.run(async (ctx) => {
          const p = check(await paragraphs(ctx), paragraph, oldStart);
          p.styleBuiltIn = style as Word.BuiltInStyleName;
          await ctx.sync();
          return { text: 'ok', summary: `第 ${paragraph} 段 → ${style}` };
        }),
    },
    {
      name: 'insert_table',
      label: '插入表格',
      permission: 'W',
      description: '在第 after 段之后插入表格，values 为二维字符串数组，第一行为表头。',
      parameters: {
        type: 'object',
        properties: {
          after: { type: 'integer' },
          oldStart: { type: 'string' },
          values: { type: 'array', items: { type: 'array', items: { type: 'string' } } },
        },
        required: ['after', 'oldStart', 'values'],
      },
      run: ({ after, oldStart, values }: { after: number; oldStart: string; values: string[][] }) =>
        Word.run(async (ctx) => {
          const p = check(await paragraphs(ctx), after, oldStart);
          const cols = Math.max(...values.map((r) => r.length));
          const v = values.map((r) => [...r, ...Array(cols - r.length).fill('')]);
          p.insertTable(v.length, cols, Word.InsertLocation.after, v);
          await ctx.sync();
          return { text: 'ok', summary: `${v.length}×${cols} 表格` };
        }),
    },
    {
      name: 'add_comment',
      label: '添加批注',
      permission: 'W',
      description: '在第 paragraph 段（或其中 quote 文字处）添加批注。仅在用户要求批注、或需要向作者提问而不宜直接改动时使用。',
      parameters: {
        type: 'object',
        properties: { paragraph: { type: 'integer' }, oldStart: { type: 'string' }, quote: { type: 'string' }, comment: { type: 'string' } },
        required: ['paragraph', 'oldStart', 'comment'],
      },
      run: ({ paragraph, oldStart, quote, comment }: { paragraph: number; oldStart: string; quote?: string; comment: string }) =>
        Word.run(async (ctx) => {
          const p = check(await paragraphs(ctx), paragraph, oldStart);
          let target: Word.Range = p.getRange('Content');
          if (quote && [...quote].length <= 255) {
            const hits = p.search(escapeWordSearch(quote), { matchCase: true });
            hits.load('items');
            await ctx.sync();
            if (hits.items[0]) target = hits.items[0];
          }
          target.insertComment(comment);
          await ctx.sync();
          return { text: 'ok', summary: `第 ${paragraph} 段：批注` };
        }),
    },
    {
      name: 'delete_paragraphs',
      label: '删除段落',
      permission: 'D',
      description: '删除第 from 到 to 段（含）。会记为修订；需要用户确认。',
      parameters: {
        type: 'object',
        properties: { from: { type: 'integer' }, to: { type: 'integer' }, oldStart: { type: 'string', description: '第 from 段开头文字' } },
        required: ['from', 'to', 'oldStart'],
      },
      describe: (a: { from: number; to: number }) => (a.from === a.to ? `删除第 ${a.from} 段` : `删除第 ${a.from}–${a.to} 段`),
      run: ({ from, to, oldStart }: { from: number; to: number; oldStart: string }) =>
        Word.run(async (ctx) => {
          const items = await paragraphs(ctx);
          check(items, from, oldStart);
          const last = check(items, to);
          items[from - 1].getRange('Whole').expandTo(last.getRange('Whole')).delete();
          await ctx.sync();
          return { text: 'ok', summary: `删除第 ${from}–${to} 段` };
        }),
    },
  ];
}

// 每轮开始时提供给模型的文档概况
export async function wordContext(): Promise<string> {
  return Word.run(async (ctx) => {
    const ps = ctx.document.body.paragraphs;
    ps.load('items/text,items/style');
    const sel = ctx.document.getSelection();
    sel.load('text');
    ctx.document.load('changeTrackingMode');
    await ctx.sync();
    const headings = ps.items.filter((p) => /heading|标题|title/i.test(p.style)).length;
    const chars = ps.items.reduce((n, p) => n + p.text.length, 0);
    return `文档共 ${ps.items.length} 段、约 ${chars} 字、${headings} 个标题；用户当前${sel.text ? `选中了 ${sel.text.length} 字` : '未选中文字'}。`;
  });
}
