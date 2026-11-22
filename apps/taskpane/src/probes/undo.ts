// Ctrl+Z / ⌘Z 撤销探针（Undo Probe）
// 1) writeMarker：用 Office.js 写入一个可识别的标记；
// 2) 用户回到文档按 ⌘Z / Ctrl+Z；
// 3) checkMarker：标记已消失 → Office.js 写入可被原生撤销。
// removeMarker 用于清理残留标记。

import type { HostName } from './env';

export const UNDO_MARK = 'KEYNIB-UNDO-PROBE';
const EXCEL_CELL = 'ZZ9999'; // 选偏远单元格，写入前确认为空，避免覆盖用户数据

let pptShapeRef: { slideId: string; shapeId: string } | null = null;

export async function writeMarker(host: HostName): Promise<string> {
  switch (host) {
    case 'PowerPoint':
      return PowerPoint.run(async (ctx) => {
        const sel = ctx.presentation.getSelectedSlides();
        sel.load('items/id');
        await ctx.sync();
        const slide = sel.items[0] ?? ctx.presentation.slides.getItemAt(0);
        slide.load('id');
        const shape = slide.shapes.addTextBox(UNDO_MARK, { left: 20, top: 20, width: 300, height: 40 });
        shape.load('id');
        await ctx.sync();
        pptShapeRef = { slideId: slide.id, shapeId: shape.id };
        return '已在当前幻灯片左上角加入标记文本框';
      });
    case 'Word':
      return Word.run(async (ctx) => {
        ctx.document.body.insertParagraph(UNDO_MARK, Word.InsertLocation.end);
        await ctx.sync();
        return '已在文末加入标记段落';
      });
    case 'Excel':
      return Excel.run(async (ctx) => {
        const cell = ctx.workbook.worksheets.getActiveWorksheet().getRange(EXCEL_CELL);
        cell.load('values');
        await ctx.sync();
        if (cell.values[0][0] !== '') throw new Error(`${EXCEL_CELL} 非空，已取消以免覆盖数据`);
        cell.values = [[UNDO_MARK]];
        await ctx.sync();
        return `已在当前工作表 ${EXCEL_CELL} 写入标记`;
      });
    default:
      throw new Error('未知宿主');
  }
}

// 返回 true = 标记仍在（撤销无效）；false = 标记已消失（撤销有效）
export async function markerExists(host: HostName): Promise<boolean> {
  switch (host) {
    case 'PowerPoint':
      return PowerPoint.run(async (ctx) => {
        if (!pptShapeRef) return false;
        const shape = ctx.presentation.slides.getItem(pptShapeRef.slideId).shapes.getItemOrNullObject(pptShapeRef.shapeId);
        shape.load('id');
        await ctx.sync();
        return !shape.isNullObject;
      });
    case 'Word':
      return Word.run(async (ctx) => {
        const hits = ctx.document.body.search(UNDO_MARK);
        hits.load('items');
        await ctx.sync();
        return hits.items.length > 0;
      });
    case 'Excel':
      return Excel.run(async (ctx) => {
        const cell = ctx.workbook.worksheets.getActiveWorksheet().getRange(EXCEL_CELL);
        cell.load('values');
        await ctx.sync();
        return cell.values[0][0] === UNDO_MARK;
      });
    default:
      return false;
  }
}

export async function removeMarker(host: HostName): Promise<void> {
  switch (host) {
    case 'PowerPoint':
      await PowerPoint.run(async (ctx) => {
        if (!pptShapeRef) return;
        const shape = ctx.presentation.slides.getItem(pptShapeRef.slideId).shapes.getItemOrNullObject(pptShapeRef.shapeId);
        await ctx.sync();
        if (!shape.isNullObject) shape.delete();
        await ctx.sync();
        pptShapeRef = null;
      });
      return;
    case 'Word':
      await Word.run(async (ctx) => {
        const hits = ctx.document.body.search(UNDO_MARK);
        hits.load('items');
        await ctx.sync();
        hits.items.forEach((r) => r.paragraphs.getFirst().delete());
        await ctx.sync();
      });
      return;
    case 'Excel':
      await Excel.run(async (ctx) => {
        const cell = ctx.workbook.worksheets.getActiveWorksheet().getRange(EXCEL_CELL);
        cell.load('values');
        await ctx.sync();
        if (cell.values[0][0] === UNDO_MARK) cell.clear();
        await ctx.sync();
      });
  }
}
