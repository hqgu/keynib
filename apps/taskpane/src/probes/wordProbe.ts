// Word 修订探针（M2 前置实测）：在空白文档中运行，结束时恢复原修订模式。
// 回答：修订后读到的文本是否混入已删除文字？search 是否命中已删除文字？
// TrackedChange 的作者 / 时间 / 类型？中文能否按句切分？样式与表格是否记为修订？拒绝后能否精确还原？

const PARA = '本研究采用随机对照试验（RCT）设计。受试者按 1:1 比例随机分配至干预组和对照组。主要结局为 12 周时的收缩压变化，次要结局包括不良事件发生率。';
const NEW_SENTENCE = '受试者按 1:1 比例经中心随机系统分配至干预组或对照组。';

export async function runWordProbe() {
  const out: Record<string, unknown> = {};
  let originalMode: Word.ChangeTrackingMode | string = 'Off';
  await Word.run(async (ctx) => {
    const doc = ctx.document;
    doc.load('changeTrackingMode');
    await ctx.sync();
    originalMode = doc.changeTrackingMode;
    out.originalMode = originalMode;
    if (!Office.context.requirements.isSetSupported('WordApi', '1.6')) throw new Error('需要 WordApi 1.6');

    // 1. 不开修订写入种子内容
    doc.changeTrackingMode = Word.ChangeTrackingMode.off;
    const body = doc.body;
    const h = body.insertParagraph('研究方法', Word.InsertLocation.end);
    h.styleBuiltIn = Word.BuiltInStyleName.heading1;
    const p = body.insertParagraph(PARA, Word.InsertLocation.end);
    await ctx.sync();

    // 2. 已有修订（用户自己的）
    const before = body.getTrackedChanges();
    before.load('items/type,items/text,items/author,items/date');
    await ctx.sync();
    out.trackedBefore = before.items.length;

    try {
      // 3. 开修订，按句切分，替换第 2 句
      doc.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
      const sentences = p.getTextRanges(['。'], false);
      sentences.load('items/text');
      await ctx.sync();
      out.sentences = sentences.items.map((s) => s.text);
      sentences.items[1].insertText(NEW_SENTENCE, Word.InsertLocation.replace);
      await ctx.sync();

      // 4. 修订后读到的文本
      p.load('text');
      const cur = p.getRange().getReviewedText(Word.ChangeTrackingVersion.current);
      const orig = p.getRange().getReviewedText(Word.ChangeTrackingVersion.original);
      await ctx.sync();
      out.paragraphTextAfter = p.text;
      out.reviewedCurrent = cur.value;
      out.reviewedOriginal = orig.value;
      out.textIncludesDeleted = p.text.includes('随机分配至干预组和对照组');

      // 5. search 是否命中已删除文字；特殊字符；超长
      const hitDeleted = body.search('随机分配至干预组和对照组');
      const hitNew = body.search('中心随机系统');
      hitDeleted.load('items');
      hitNew.load('items');
      await ctx.sync();
      out.searchHitsDeleted = hitDeleted.items.length;
      out.searchHitsNew = hitNew.items.length;
      try {
        const long = body.search('研'.repeat(300));
        long.load('items');
        await ctx.sync();
        out.search300 = `ok ${long.items.length}`;
      } catch (e: any) {
        out.search300 = `报错：${e.code ?? e.message}`;
      }
      try {
        const q = body.search('1:1 比例?');
        q.load('items');
        await ctx.sync();
        out.searchQuestionMark = q.items.length;
      } catch (e: any) {
        out.searchQuestionMark = `报错：${e.code ?? e.message}`;
      }

      // 6. 样式变化、插表格是否记为修订
      const afterText = body.getTrackedChanges();
      afterText.load('items');
      await ctx.sync();
      const nText = afterText.items.length;
      h.styleBuiltIn = Word.BuiltInStyleName.heading2;
      await ctx.sync();
      const afterStyle = body.getTrackedChanges();
      afterStyle.load('items');
      await ctx.sync();
      p.insertTable(2, 2, Word.InsertLocation.after, [['组别', 'n'], ['干预组', '60']]);
      await ctx.sync();
      const afterTable = body.getTrackedChanges();
      afterTable.load('items/type,items/text,items/author,items/date');
      await ctx.sync();
      out.trackedCounts = { afterTextEdit: nText, afterStyleChange: afterStyle.items.length, afterInsertTable: afterTable.items.length };
      out.trackedItems = afterTable.items.map((t) => ({ type: t.type, text: t.text.slice(0, 40), author: t.author, date: String(t.date) }));

      // 7. 拒绝本次新增的修订（before 为空时即全部），检查是否精确还原
      afterTable.rejectAll();
      await ctx.sync();
      p.load('text');
      const remain = body.getTrackedChanges();
      remain.load('items');
      const tables = body.tables;
      tables.load('items');
      h.load('styleBuiltIn');
      await ctx.sync();
      out.afterReject = {
        paragraphEqualsOriginal: p.text === PARA,
        paragraphText: p.text,
        trackedRemaining: remain.items.length,
        tables: tables.items.length,
        headingStyle: h.styleBuiltIn,
      };
    } finally {
      doc.changeTrackingMode = originalMode as Word.ChangeTrackingMode;
      await ctx.sync();
      doc.load('changeTrackingMode');
      await ctx.sync();
      out.modeRestored = doc.changeTrackingMode;
    }
  });
  return out;
}
