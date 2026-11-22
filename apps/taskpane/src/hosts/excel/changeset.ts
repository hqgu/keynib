// Excel "撤销本轮"
// 设计上让撤销简单：结果、图表、透视表都放在本轮新建的工作表 → 撤销时删除这些表。
// 覆盖原有工作表的单元格是 D 级操作，执行前对目标区域的 formulas + numberFormat 做快照（≤ 50k 单元格）。
// M3 不在原有工作表上建表格、条件格式或改单元格格式（这些需要格式快照，未实测）。

import { cellCount } from './stats';

export const SNAPSHOT_LIMIT = 50000;

interface RangeSnapshot {
  sheet: string;
  address: string;
  formulas: any[][];
  numberFormat: any[][];
}

export class ExcelChangeset {
  readonly kind = 'excel';
  private createdSheets: string[] = [];
  private snapshots: RangeSnapshot[] = [];
  undone = false;

  get isEmpty() {
    return this.createdSheets.length === 0 && this.snapshots.length === 0;
  }

  isCreated(sheet: string) {
    return this.createdSheets.includes(sheet);
  }

  noteSheet(name: string) {
    this.createdSheets.push(name);
  }

  summary() {
    return {
      createdSheets: [...this.createdSheets],
      overwritten: this.snapshots.map((s) => `${s.sheet}!${s.address}`),
    };
  }

  // 覆盖前调用；在 Excel.run 内由调用方传入 ctx
  async snapshot(ctx: Excel.RequestContext, sheet: string, address: string) {
    if (cellCount(address) > SNAPSHOT_LIMIT) throw new Error(`覆盖范围超过 ${SNAPSHOT_LIMIT} 个单元格，无法保证可撤销，已拒绝`);
    const r = ctx.workbook.worksheets.getItem(sheet).getRange(address);
    r.load('formulas,numberFormat,address');
    await ctx.sync();
    this.snapshots.push({ sheet, address: r.address.split('!').pop()!, formulas: r.formulas, numberFormat: r.numberFormat });
  }

  async undo(): Promise<void> {
    if (this.undone) return;
    await Excel.run(async (ctx) => {
      // 先还原覆盖（后覆盖的先还原），再删除新建的表
      for (const s of [...this.snapshots].reverse()) {
        const ws = ctx.workbook.worksheets.getItemOrNullObject(s.sheet);
        await ctx.sync();
        if (ws.isNullObject) continue;
        const r = ws.getRange(s.address);
        r.numberFormat = s.numberFormat;
        r.formulas = s.formulas;
      }
      await ctx.sync();
      for (const name of [...this.createdSheets].reverse()) {
        const ws = ctx.workbook.worksheets.getItemOrNullObject(name);
        await ctx.sync();
        if (!ws.isNullObject) ws.delete();
      }
      await ctx.sync();
    });
    this.undone = true;
  }
}
