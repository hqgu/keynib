// Excel 工具集（M3：数据分析 + 图表）
// 规则：结果 / 图表 / 透视表只能写到本轮新建的工作表（create_output_sheet）；
// 覆盖原有工作表用 overwrite_range（D 级，需用户确认，执行前快照）。统计量必须写成 Excel 公式。

import type { ToolDef } from '@keynib/core';
import type { ExcelChangeset } from './changeset';
import { cellCount, colLetter, findErrors, profileColumns, quoteSheet } from './stats';

const json = (v: unknown) => JSON.stringify(v);

const CHART_TYPES = ['ColumnClustered', 'BarClustered', 'Line', 'LineMarkers', 'XYScatter', 'Pie', 'Histogram', 'BoxWhisker', 'ColumnStacked', 'Area'];
const AGG = ['Sum', 'Count', 'Average', 'Max', 'Min', 'StandardDeviation'];

function parseTopLeft(address: string) {
  const a = address.split('!').pop()!.replace(/\$/g, '');
  const m = /^([A-Z]+)(\d+)/i.exec(a)!;
  const col = [...m[1].toUpperCase()].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) - 1;
  return { row: Number(m[2]) - 1, col };
}

// 返回实际表名：模型漏填 sheet 且本轮只新建了一张表时，自动使用它
function requireCreated(cs: ExcelChangeset, sheet: string | undefined): string {
  if (!sheet) {
    const created = cs.summary().createdSheets;
    if (created.length === 1) return created[0];
    throw new Error(created.length ? `缺少 sheet 参数，请指定：${created.join('、')}` : '缺少 sheet 参数；请先用 create_output_sheet 新建结果表');
  }
  if (!cs.isCreated(sheet)) {
    throw new Error(`「${sheet}」不是本轮新建的输出表。结果、图表、透视表请写到 create_output_sheet 新建的表；确需修改原有表请用 overwrite_range（需用户确认）`);
  }
  return sheet;
}

// 写入后回读，返回计算出的值与错误
async function writeAndCheck(ctx: Excel.RequestContext, ws: Excel.Worksheet, address: string, formulas: unknown[][], numberFormat?: string[][]) {
  const rows = formulas.length;
  const cols = Math.max(...formulas.map((r) => r.length));
  const grid = formulas.map((r) => [...r, ...Array(cols - r.length).fill('')]);
  const tl = parseTopLeft(address);
  const range = ws.getRangeByIndexes(tl.row, tl.col, rows, cols);
  range.formulas = grid as any;
  // 数字格式尺寸与数据不一致时按行列补齐（缺的用 General），避免整批写入失败
  if (numberFormat?.length) {
    range.numberFormat = Array.from({ length: rows }, (_, i) =>
      Array.from({ length: cols }, (_, j) => numberFormat[i]?.[j] ?? numberFormat[i]?.[0] ?? 'General'),
    ) as any;
  }
  range.load('values,address');
  await ctx.sync();
  const errors = findErrors(range.values, tl);
  const formulaCells = grid.flat().filter((v) => typeof v === 'string' && v.startsWith('=')).length;
  // FILTER 溢出结果里出现 0：多半是源数据的空白被当成 0（实测：箱线图因此混入缺失值）
  const warnings: string[] = [];
  const filters: { cell: string; spill: Excel.Range }[] = [];
  grid.forEach((row, i) =>
    row.forEach((v, j) => {
      if (typeof v === 'string' && /FILTER\(/i.test(v) && !/<>""/.test(v)) {
        const cell = ws.getRangeByIndexes(tl.row + i, tl.col + j, 1, 1);
        const spill = cell.getSpillingToRangeOrNullObject();
        spill.load('values,address');
        filters.push({ cell: `${colLetter(tl.col + j)}${tl.row + i + 1}`, spill });
      }
    }),
  );
  if (filters.length) {
    await ctx.sync();
    for (const f of filters) {
      if (f.spill.isNullObject) continue;
      const zeros = f.spill.values.flat().filter((x) => x === 0).length;
      if (zeros) warnings.push(`${f.cell} 的 FILTER 结果有 ${zeros} 个 0：若源列有空白，空白被当成了 0，请在条件中加入 (数值列<>"") 排除缺失`);
    }
  }
  return { address: range.address, values: range.values, errors, warnings, formulaCells };
}

export function createExcelTools(getChangeset: () => ExcelChangeset): ToolDef[] {
  return [
    {
      name: 'get_workbook_summary',
      label: '浏览工作簿',
      permission: 'R',
      description: '列出所有工作表：已用区域、表头（首行）、行数，以及各列推断类型。分析前先调用。',
      parameters: { type: 'object', properties: {} },
      run: () =>
        Excel.run(async (ctx) => {
          const sheets = ctx.workbook.worksheets;
          sheets.load('items/name');
          await ctx.sync();
          const out = [];
          for (const ws of sheets.items) {
            const used = ws.getUsedRangeOrNullObject(true);
            used.load('address,rowCount,columnCount');
            await ctx.sync();
            if (used.isNullObject) {
              out.push({ sheet: ws.name, empty: true });
              continue;
            }
            const head = used.getRow(0);
            head.load('values');
            // 抽前 200 行推断列类型（只有表头时不抽样）
            const sample = used.rowCount > 1 ? used.getOffsetRange(1, 0).getResizedRange(Math.min(used.rowCount, 200) - used.rowCount - 1, 0) : null;
            sample?.load('values,numberFormat');
            await ctx.sync();
            const header = head.values[0].map(String);
            const cols = profileColumns(header, sample?.values ?? [], sample?.numberFormat).map((c) => `${c.column}:${c.name}(${c.type})`);
            out.push({ sheet: ws.name, used: used.address.split('!').pop(), rows: used.rowCount, columns: cols });
          }
          return { text: json(out), summary: `${out.length} 个工作表` };
        }),
    },
    {
      name: 'read_range',
      label: '读取数据',
      permission: 'R',
      description: '读取区域的显示文本（日期按格式显示），最多返回 maxRows 行（默认 30，最多 200）。分析大表请用 profile_data，而不是读全部数据。',
      parameters: {
        type: 'object',
        properties: { sheet: { type: 'string' }, address: { type: 'string', description: '如 A1:F31' }, maxRows: { type: 'integer' } },
        required: ['sheet', 'address'],
      },
      run: ({ sheet, address, maxRows }: { sheet: string; address: string; maxRows?: number }) =>
        Excel.run(async (ctx) => {
          const r = ctx.workbook.worksheets.getItem(sheet).getRange(address);
          r.load('text,rowCount,address');
          await ctx.sync();
          const lim = Math.min(maxRows ?? 30, 200);
          const rows = r.text.slice(0, lim).map((row) => row.join('\t'));
          return {
            text: json({ address: r.address, totalRows: r.rowCount, returned: rows.length, tsv: rows.join('\n') }),
            summary: `${sheet}!${address}（${Math.min(r.rowCount, lim)} 行）`,
          };
        }),
    },
    {
      name: 'profile_data',
      label: '数据画像',
      permission: 'R',
      description:
        '对一个数据区（首行为表头）逐列计算：类型、非空 n、缺失数、均值、标准差、分位数、最小 / 最大值，文本列给频数。仅供你了解数据；写入结果表的统计量必须用 Excel 公式计算。',
      parameters: { type: 'object', properties: { sheet: { type: 'string' }, address: { type: 'string', description: '省略则用已用区域' } }, required: ['sheet'] },
      run: ({ sheet, address }: { sheet: string; address?: string }) =>
        Excel.run(async (ctx) => {
          const ws = ctx.workbook.worksheets.getItem(sheet);
          const r = address ? ws.getRange(address) : ws.getUsedRange(true);
          r.load('values,numberFormat,address,rowCount');
          await ctx.sync();
          const [header, ...rows] = r.values;
          const profiles = profileColumns(header.map(String), rows, r.numberFormat.slice(1));
          return { text: json({ address: r.address, rows: r.rowCount - 1, columns: profiles }), summary: `${profiles.length} 列 × ${r.rowCount - 1} 行` };
        }),
    },
    {
      name: 'create_output_sheet',
      label: '新建结果表',
      permission: 'W',
      description: '新建一个工作表存放分析结果、图表或透视表（重名时自动加序号）。返回实际表名。所有输出都写到这里。',
      parameters: { type: 'object', properties: { name: { type: 'string', description: '如 描述统计' } }, required: ['name'] },
      run: ({ name }: { name: string }) =>
        Excel.run(async (ctx) => {
          const sheets = ctx.workbook.worksheets;
          sheets.load('items/name');
          await ctx.sync();
          const base = (name || '分析结果').replace(/[\\/?*[\]:]/g, '').slice(0, 28) || '分析结果';
          let final = base;
          for (let i = 2; sheets.items.some((s) => s.name === final); i++) final = `${base}${i}`;
          const ws = sheets.add(final);
          ws.activate();
          await ctx.sync();
          getChangeset().noteSheet(final);
          return { text: json({ sheet: final, refPrefix: quoteSheet(final) }), summary: final };
        }),
    },
    {
      name: 'write_range',
      label: '写入结果',
      permission: 'W',
      description:
        '向本轮新建的结果表写入二维数组（从 address 左上角开始）。文字直接写；统计量必须写公式，如 "=AVERAGE(Sheet1!F2:F201)"，引用中文或含空格的表名要加单引号。可选 numberFormat（同尺寸，如 "0.00"）。返回计算结果和错误单元格（#DIV/0! 等），有错误请修正。',
      parameters: {
        type: 'object',
        properties: {
          sheet: { type: 'string' },
          address: { type: 'string', description: '左上角，如 A1' },
          values: { type: 'array', items: { type: 'array', items: {} } },
          numberFormat: { type: 'array', items: { type: 'array', items: { type: 'string' } } },
        },
        required: ['sheet', 'address', 'values'],
      },
      run: ({ sheet, address, values, numberFormat }: { sheet: string; address: string; values: unknown[][]; numberFormat?: string[][] }) =>
        Excel.run(async (ctx) => {
          const ws = ctx.workbook.worksheets.getItem(requireCreated(getChangeset(), sheet));
          const r = await writeAndCheck(ctx, ws, address, values, numberFormat);
          ws.getUsedRange().format.autofitColumns();
          await ctx.sync();
          return {
            text: json({ address: r.address, computed: r.values, errors: r.errors, warnings: r.warnings }),
            summary: `${r.address.split('!').pop()}，${r.formulaCells} 个公式${r.errors.length ? `，${r.errors.length} 个错误` : ''}${r.warnings.length ? `，${r.warnings.length} 个警告` : ''}`,
          };
        }),
    },
    {
      name: 'format_output',
      label: '设置格式',
      permission: 'W',
      description: '设置结果表中某区域的格式：加粗、数字格式（如 "0.0"、"0.0%"）、背景色、边框。',
      parameters: {
        type: 'object',
        properties: {
          sheet: { type: 'string' },
          address: { type: 'string' },
          bold: { type: 'boolean' },
          numberFormat: { type: 'string' },
          fill: { type: 'string', description: '#RRGGBB' },
          border: { type: 'boolean' },
        },
        required: ['sheet', 'address'],
      },
      run: (a: { sheet: string; address: string; bold?: boolean; numberFormat?: string; fill?: string; border?: boolean }) =>
        Excel.run(async (ctx) => {
          const r = ctx.workbook.worksheets.getItem(requireCreated(getChangeset(), a.sheet)).getRange(a.address);
          if (a.bold !== undefined) r.format.font.bold = a.bold;
          if (a.fill) r.format.fill.color = a.fill;
          if (a.numberFormat) {
            r.load('rowCount,columnCount');
            await ctx.sync();
            r.numberFormat = Array.from({ length: r.rowCount }, () => Array(r.columnCount).fill(a.numberFormat));
          }
          if (a.border) {
            for (const e of ['EdgeTop', 'EdgeBottom', 'EdgeLeft', 'EdgeRight', 'InsideHorizontal', 'InsideVertical'] as const) {
              r.format.borders.getItem(e).style = 'Continuous';
            }
          }
          await ctx.sync();
          return { text: 'ok', summary: `${a.sheet}!${a.address}` };
        }),
    },
    {
      name: 'create_chart',
      label: '创建图表',
      permission: 'W',
      description:
        `在结果表上创建图表，数据可来自任何表。type：${CHART_TYPES.join(' / ')}。dataRange 含表头；比较两组请先在结果表用公式算好汇总再作图，或用 BoxWhisker / Histogram 直接作分布图。anchor 为图表左上角单元格。`,
      parameters: {
        type: 'object',
        properties: {
          sheet: { type: 'string', description: '放图表的结果表' },
          dataSheet: { type: 'string' },
          dataRange: { type: 'string' },
          type: { type: 'string', enum: CHART_TYPES },
          title: { type: 'string' },
          xTitle: { type: 'string' },
          yTitle: { type: 'string' },
          seriesBy: { type: 'string', enum: ['Auto', 'Columns', 'Rows'] },
          anchor: { type: 'string', description: '如 H2' },
        },
        required: ['sheet', 'dataSheet', 'dataRange', 'type', 'title'],
      },
      run: (a: { sheet: string; dataSheet: string; dataRange: string; type: string; title: string; xTitle?: string; yTitle?: string; seriesBy?: string; anchor?: string }) =>
        Excel.run(async (ctx) => {
          const ws = ctx.workbook.worksheets.getItem(requireCreated(getChangeset(), a.sheet));
          const src = ctx.workbook.worksheets.getItem(a.dataSheet).getRange(a.dataRange);
          const chart = ws.charts.add(a.type as Excel.ChartType, src, (a.seriesBy ?? 'Auto') as Excel.ChartSeriesBy);
          chart.title.text = a.title;
          const hasAxes = !['Pie'].includes(a.type);
          if (hasAxes && a.xTitle) {
            chart.axes.categoryAxis.title.text = a.xTitle;
            chart.axes.categoryAxis.title.visible = true;
          }
          if (hasAxes && a.yTitle) {
            chart.axes.valueAxis.title.text = a.yTitle;
            chart.axes.valueAxis.title.visible = true;
          }
          if (a.anchor) chart.setPosition(a.anchor);
          chart.width = 480;
          chart.height = 300;
          await ctx.sync();
          return { text: 'ok', summary: `${a.type}：${a.title}` };
        }),
    },
    {
      name: 'create_pivot',
      label: '数据透视表',
      permission: 'W',
      description: `在结果表创建数据透视表。rows / columns 为字段名（表头），values 为 [{field, summarizeBy}]，summarizeBy：${AGG.join(' / ')}。`,
      parameters: {
        type: 'object',
        properties: {
          sheet: { type: 'string' },
          dataSheet: { type: 'string' },
          dataRange: { type: 'string', description: '含表头，如 A1:F201' },
          anchor: { type: 'string', description: '如 A1' },
          rows: { type: 'array', items: { type: 'string' } },
          columns: { type: 'array', items: { type: 'string' } },
          values: { type: 'array', items: { type: 'object', properties: { field: { type: 'string' }, summarizeBy: { type: 'string', enum: AGG } }, required: ['field'] } },
        },
        required: ['sheet', 'dataSheet', 'dataRange', 'rows', 'values'],
      },
      run: (a: { sheet: string; dataSheet: string; dataRange: string; anchor?: string; rows: string[]; columns?: string[]; values: { field: string; summarizeBy?: string }[] }) =>
        Excel.run(async (ctx) => {
          const ws = ctx.workbook.worksheets.getItem(requireCreated(getChangeset(), a.sheet));
          const name = `透视${Date.now() % 100000}`;
          const pt = ws.pivotTables.add(name, ctx.workbook.worksheets.getItem(a.dataSheet).getRange(a.dataRange), ws.getRange(a.anchor ?? 'A1'));
          for (const r of a.rows) pt.rowHierarchies.add(pt.hierarchies.getItem(r));
          for (const c of a.columns ?? []) pt.columnHierarchies.add(pt.hierarchies.getItem(c));
          for (const v of a.values) {
            const dh = pt.dataHierarchies.add(pt.hierarchies.getItem(v.field));
            if (v.summarizeBy) dh.summarizeBy = v.summarizeBy as Excel.AggregationFunction;
          }
          await ctx.sync();
          const body = pt.layout.getRange();
          body.load('address,values');
          await ctx.sync();
          return { text: json({ address: body.address, values: body.values.slice(0, 30) }), summary: `${a.rows.join('×')} 透视` };
        }),
    },
    {
      name: 'overwrite_range',
      label: '修改原表',
      permission: 'D',
      description: '覆盖原有工作表的单元格（如补一列计算值、修正数据）。需要用户确认；执行前自动快照，可撤销。能放到结果表的内容不要用它。',
      parameters: {
        type: 'object',
        properties: { sheet: { type: 'string' }, address: { type: 'string', description: '左上角或完整区域' }, values: { type: 'array', items: { type: 'array', items: {} } } },
        required: ['sheet', 'address', 'values'],
      },
      describe: (a: { sheet: string; address: string; values: unknown[][] }) => {
        const n = a.values.length * Math.max(...a.values.map((r) => r.length));
        return `覆盖原表「${a.sheet}」从 ${a.address.split(':')[0]} 开始的 ${n} 个单元格（可撤销）`;
      },
      run: ({ sheet, address, values }: { sheet: string; address: string; values: unknown[][] }) =>
        Excel.run(async (ctx) => {
          const cs = getChangeset();
          const ws = ctx.workbook.worksheets.getItem(sheet);
          const tl = parseTopLeft(address);
          const cols = Math.max(...values.map((r) => r.length));
          const target = ws.getRangeByIndexes(tl.row, tl.col, values.length, cols);
          target.load('address');
          await ctx.sync();
          const addr = target.address.split('!').pop()!;
          if (cellCount(addr) && !cs.isCreated(sheet)) await cs.snapshot(ctx, sheet, addr);
          const r = await writeAndCheck(ctx, ws, addr, values);
          return { text: json({ address: r.address, errors: r.errors }), summary: `${sheet}!${addr}` };
        }),
    },
  ];
}

export async function excelContext(): Promise<string> {
  return Excel.run(async (ctx) => {
    const sheets = ctx.workbook.worksheets;
    sheets.load('items/name');
    const active = ctx.workbook.worksheets.getActiveWorksheet();
    active.load('name');
    const sel = ctx.workbook.getSelectedRange();
    sel.load('address');
    await ctx.sync();
    return `工作簿有 ${sheets.items.length} 个工作表（${sheets.items.map((s) => s.name).join('、')}）；当前工作表「${active.name}」，选中 ${sel.address.split('!').pop()}。`;
  });
}
