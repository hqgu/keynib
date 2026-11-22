// PPT 能力逐项探针：每个接口单独调用、单独记录成败（含 Office 的 debugInfo），用于定位网页版的 GeneralException。
// 会在末尾临时新建幻灯片，测完删除。

export interface StepResult {
  step: string;
  ok: boolean;
  detail?: string;
}

function describeError(e: any): string {
  const info = e?.debugInfo ?? {};
  return [e?.code, e?.message, info.errorLocation && `at ${info.errorLocation}`, info.statement && `stmt ${String(info.statement).slice(0, 120)}`]
    .filter(Boolean)
    .join(' | ');
}

export async function runPptCapabilityProbe(): Promise<StepResult[]> {
  const out: StepResult[] = [];
  const created: string[] = [];
  let layoutId: string | undefined;
  let exported: string | undefined;

  const step = async (name: string, fn: (ctx: PowerPoint.RequestContext) => Promise<string | void>) => {
    try {
      const detail = await PowerPoint.run(fn);
      out.push({ step: name, ok: true, detail: detail || undefined });
    } catch (e) {
      out.push({ step: name, ok: false, detail: describeError(e) });
    }
  };

  const newest = async (ctx: PowerPoint.RequestContext) => {
    const slides = ctx.presentation.slides;
    slides.load('items/id');
    await ctx.sync();
    return slides.items[slides.items.length - 1];
  };

  await step('读取幻灯片列表 slides.load', async (ctx) => {
    const s = ctx.presentation.slides;
    s.load('items/id');
    await ctx.sync();
    return `${s.items.length} 页`;
  });
  await step('读取选中页 getSelectedSlides', async (ctx) => {
    const s = ctx.presentation.getSelectedSlides();
    s.load('items/id');
    await ctx.sync();
    return `${s.items.length} 页`;
  });
  await step('读取母版 slideMasters（只取 id）', async (ctx) => {
    const m = ctx.presentation.slideMasters;
    m.load('items/id');
    await ctx.sync();
    return `${m.items.length} 个`;
  });
  await step('读取母版名称 slideMasters name', async (ctx) => {
    const m = ctx.presentation.slideMasters;
    m.load('items/id,items/name');
    await ctx.sync();
    return m.items.map((x) => x.name).join(', ');
  });
  await step('读取版式 layouts', async (ctx) => {
    const m = ctx.presentation.slideMasters.getItemAt(0);
    m.layouts.load('items/id,items/name');
    await ctx.sync();
    layoutId = m.layouts.items[1]?.id ?? m.layouts.items[0]?.id;
    return `${m.layouts.items.length} 个：${m.layouts.items.map((x) => x.name).slice(0, 4).join(', ')}`;
  });
  await step('读取当前页版式 slide.layout', async (ctx) => {
    const s = ctx.presentation.slides.getItemAt(0);
    s.layout.load('name');
    await ctx.sync();
    return s.layout.name;
  });
  await step('新建页 slides.add()', async (ctx) => {
    ctx.presentation.slides.add();
    await ctx.sync();
    created.push((await newest(ctx)).id);
  });
  await step('新建页（指定版式）slides.add({layoutId})', async (ctx) => {
    if (!layoutId) return '跳过：没有拿到版式 ID';
    ctx.presentation.slides.add({ layoutId });
    await ctx.sync();
    created.push((await newest(ctx)).id);
  });
  await step('加文本框 addTextBox', async (ctx) => {
    const s = await newest(ctx);
    s.shapes.addTextBox('Keynib probe', { left: 40, top: 40, width: 300, height: 40 });
    await ctx.sync();
  });
  await step('加文本框（按序号定位 getItemAt）', async (ctx) => {
    const s = ctx.presentation.slides;
    s.load('items');
    await ctx.sync();
    ctx.presentation.slides.getItemAt(s.items.length - 1).shapes.addTextBox('Keynib probe 2', { left: 40, top: 260, width: 300, height: 40 });
    await ctx.sync();
  });
  await step('加形状 addGeometricShape', async (ctx) => {
    const s = await newest(ctx);
    s.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, { left: 40, top: 100, width: 120, height: 60 });
    await ctx.sync();
  });
  await step('加表格 addTable', async (ctx) => {
    const s = await newest(ctx);
    s.shapes.addTable(2, 2, { values: [['a', 'b'], ['c', 'd']], left: 40, top: 180, width: 300 });
    await ctx.sync();
  });
  await step('截图 getImageAsBase64', async (ctx) => {
    const s = await newest(ctx);
    const img = s.getImageAsBase64({ width: 480 });
    await ctx.sync();
    return `${Math.round(img.value.length / 1024)} KB`;
  });
  await step('导出单页 exportAsBase64', async (ctx) => {
    const s = await newest(ctx);
    const b = s.exportAsBase64();
    await ctx.sync();
    exported = b.value;
    return `${Math.round(b.value.length / 1024)} KB`;
  });
  await step('插回 insertSlidesFromBase64', async (ctx) => {
    if (!exported) return '跳过：没有导出结果';
    const s = await newest(ctx);
    ctx.presentation.insertSlidesFromBase64(exported, { formatting: PowerPoint.InsertSlideFormatting.keepSourceFormatting, targetSlideId: s.id });
    await ctx.sync();
    created.push((await newest(ctx)).id);
  });
  await step('移动页 moveTo', async (ctx) => {
    const s = await newest(ctx);
    s.moveTo(0);
    await ctx.sync();
    s.load('index');
    await ctx.sync();
  });
  await step('删除测试页 delete', async (ctx) => {
    for (const id of created) {
      const s = ctx.presentation.slides.getItemOrNullObject(id);
      await ctx.sync();
      if (!s.isNullObject) s.delete();
    }
    await ctx.sync();
    return `${created.length} 页`;
  });
  return out;
}
