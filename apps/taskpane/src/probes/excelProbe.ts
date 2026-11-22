// Excel 探针（M3 前置实测）：在空白工作簿中新建临时工作表测试，结束后删除。
// 回答：直方图 / 箱线图能否创建？数据透视表能否创建？valueTypes 对数字 / 文本数字 / 空白 / 日期返回什么？
// 删除被引用的工作表后，其他表的公式是否变成 #REF!？

export async function runExcelProbe() {
  const out: Record<string, unknown> = {};
  const DATA = '探针数据';
  const OUT = '探针输出';
  await Excel.run(async (ctx) => {
    const wb = ctx.workbook;
    for (const n of [DATA, OUT]) wb.worksheets.getItemOrNullObject(n).delete();
    await ctx.sync().catch(() => {});

    const data = wb.worksheets.add(DATA);
    data.getRange('A1:D7').values = [
      ['组别', '年龄', '收缩压', '入组日期'],
      ['干预', 52, 138, 45000],
      ['对照', 61, 145, 45001],
      ['干预', '47', 132, 45002], // 文本形式的数字
      ['对照', 58, '', 45003], // 空白
      ['干预', 66, 150, 45004],
      ['对照', 49, 128, 45005],
    ];
    data.getRange('D2:D7').numberFormat = [['yyyy-mm-dd'], ['yyyy-mm-dd'], ['yyyy-mm-dd'], ['yyyy-mm-dd'], ['yyyy-mm-dd'], ['yyyy-mm-dd']];
    await ctx.sync();

    // 1. 值类型
    const r = data.getRange('A4:D5');
    r.load('values,valueTypes,numberFormat,text');
    await ctx.sync();
    out.valueTypes = { values: r.values, valueTypes: r.valueTypes, numberFormat: r.numberFormat, text: r.text };

    // 2. 输出表 + 跨表公式
    const o = wb.worksheets.add(OUT);
    o.getRange('A1:B3').formulas = [
      ['统计量', '收缩压'],
      ['均值', `=AVERAGE('${DATA}'!C2:C7)`],
      ['n', `=COUNT('${DATA}'!C2:C7)`],
    ];
    await ctx.sync();
    const fr = o.getRange('B2:B3');
    fr.load('values');
    await ctx.sync();
    out.crossSheetFormulas = fr.values;

    // 3. 图表类型
    const charts: Record<string, string> = {};
    for (const t of ['ColumnClustered', 'Histogram', 'BoxWhisker', 'XYScatter']) {
      try {
        const c = o.charts.add(t as Excel.ChartType, data.getRange(t === 'XYScatter' ? 'B1:C7' : 'C1:C7'), Excel.ChartSeriesBy.auto);
        c.title.text = t;
        c.load('chartType');
        await ctx.sync();
        charts[t] = `ok（${c.chartType}）`;
      } catch (e: any) {
        charts[t] = `失败：${e.code ?? e.message}`;
      }
    }
    out.charts = charts;

    // 4. 数据透视表
    try {
      const pt = o.pivotTables.add('探针透视', data.getRange('A1:C7'), o.getRange('E1'));
      pt.rowHierarchies.add(pt.hierarchies.getItem('组别'));
      const dh = pt.dataHierarchies.add(pt.hierarchies.getItem('收缩压'));
      dh.summarizeBy = Excel.AggregationFunction.average;
      await ctx.sync();
      const body = pt.layout.getRange();
      body.load('values');
      await ctx.sync();
      out.pivot = { ok: true, values: body.values };
    } catch (e: any) {
      out.pivot = `失败：${e.code ?? e.message}`;
    }

    // 5. 数据表引用输出表，然后删除输出表
    data.getRange('F1').formulas = [[`='${OUT}'!B2`]];
    await ctx.sync();
    o.delete();
    await ctx.sync();
    const f = data.getRange('F1');
    f.load('values,formulas');
    const dataAfter = data.getRange('A1:D7');
    dataAfter.load('values');
    await ctx.sync();
    out.afterDeleteOutput = { refCell: f.values[0][0], refFormula: f.formulas[0][0], dataIntact: dataAfter.values[3][1] === '47' || dataAfter.values[3][1] === 47 };

    // 清理
    data.delete();
    await ctx.sync();
    const left = wb.worksheets;
    left.load('items/name');
    await ctx.sync();
    out.sheetsLeft = left.items.map((s) => s.name);
  });
  return out;
}
