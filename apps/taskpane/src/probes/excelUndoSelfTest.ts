// Excel "撤销本轮"自测：不调用模型，按工具的真实调用方式模拟一轮，再撤销，比较前后指纹。
// 每轮：新建输出表（含引用 Sheet1 的公式）→ 加图表 → 加透视表 → 覆盖 Sheet1 的 3 个单元格 → 撤销。
// 指纹 = 工作表名列表 + 每表已用区域的 formulas 与 numberFormat + 图表数。

import { ExcelChangeset } from '../hosts/excel/changeset';
import { SEED_HEADER, seedTrial } from '../hosts/excel/seed';

async function fingerprint(): Promise<string> {
  return Excel.run(async (ctx) => {
    const sheets = ctx.workbook.worksheets;
    sheets.load('items/name');
    await ctx.sync();
    const parts: string[] = [];
    for (const ws of sheets.items) {
      const used = ws.getUsedRangeOrNullObject();
      used.load('address,formulas,numberFormat');
      ws.charts.load('items');
      await ctx.sync();
      parts.push(`${ws.name}|${ws.charts.items.length}|${used.isNullObject ? '' : used.address + JSON.stringify(used.formulas) + JSON.stringify(used.numberFormat)}`);
    }
    return parts.join('\n');
  });
}

async function oneTurn(dataSheet: string) {
  const cs = new ExcelChangeset();
  const out = `自测输出${Date.now() % 10000}`;
  await Excel.run(async (ctx) => {
    const ws = ctx.workbook.worksheets.add(out);
    await ctx.sync();
    cs.noteSheet(out);
    const ref = `'${dataSheet}'`;
    ws.getRange('A1:B3').formulas = [['统计量', '12周收缩压'], ['均数', `=AVERAGE(${ref}!F2:F21)`], ['n', `=COUNT(${ref}!F2:F21)`]];
    ws.charts.add(Excel.ChartType.boxwhisker, ctx.workbook.worksheets.getItem(dataSheet).getRange('F1:F21'));
    const pt = ws.pivotTables.add(`自测透视${Date.now() % 10000}`, ctx.workbook.worksheets.getItem(dataSheet).getRange('A1:F21'), ws.getRange('D1'));
    pt.rowHierarchies.add(pt.hierarchies.getItem('组别'));
    pt.dataHierarchies.add(pt.hierarchies.getItem('年龄'));
    await ctx.sync();
    // 覆盖原表 3 个单元格（D 级工具的做法：先快照再写）
    await cs.snapshot(ctx, dataSheet, 'C2:C4');
    ctx.workbook.worksheets.getItem(dataSheet).getRange('C2:C4').values = [[999], [998], [997]];
    await ctx.sync();
  });
  await cs.undo();
}

export async function runExcelUndoSelfTest(times = 20) {
  const DATA = '自测数据';
  await Excel.run(async (ctx) => {
    ctx.workbook.worksheets.getItemOrNullObject(DATA).delete();
    await ctx.sync().catch(() => {});
    const ws = ctx.workbook.worksheets.add(DATA);
    ws.getRange('A1:F21').values = [SEED_HEADER, ...seedTrial(20, 7)];
    ws.getRange('C2:C4').numberFormat = [['0.0'], ['0.0'], ['0.0']];
    await ctx.sync();
  });
  let passed = 0;
  const details: string[] = [];
  try {
    for (let i = 0; i < times; i++) {
      const before = await fingerprint();
      await oneTurn(DATA);
      const after = await fingerprint();
      if (before === after) passed++;
      else details.push(`第 ${i + 1} 次不一致`);
    }
  } finally {
    await Excel.run(async (ctx) => {
      ctx.workbook.worksheets.getItem(DATA).delete();
      await ctx.sync();
    });
  }
  return { ok: passed === times, passed, runs: times, details: details.slice(0, 5) };
}
