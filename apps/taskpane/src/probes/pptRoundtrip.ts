// PPT 导出 / 插入往返探针（Export–Insert Roundtrip Probe）
// 验证"撤销本轮"的核心路线：slide.exportAsBase64()（1.8）→ insertSlidesFromBase64（1.2）。
// 步骤：在末尾新建测试页并加文本框 → 导出 → 插回到测试页之后 → 比较两页形状 → 清理。

export interface ShapeSnapshot {
  name: string;
  type: string;
  left: number;
  top: number;
  width: number;
  height: number;
  text?: string;
  fontName?: string;
  fontSize?: number;
}

export interface RoundtripResult {
  ok: boolean;
  base64Length?: number;
  originalId?: string;
  copyId?: string;
  idChanged?: boolean;
  original?: ShapeSnapshot[];
  copy?: ShapeSnapshot[];
  differences: string[];
  error?: string;
}

const PROBE_TEXT = 'Keynib 往返探针 · 微软雅黑 24';

async function snapshotShapes(ctx: PowerPoint.RequestContext, slide: PowerPoint.Slide): Promise<ShapeSnapshot[]> {
  const shapes = slide.shapes;
  shapes.load('items/name,items/type,items/left,items/top,items/width,items/height');
  await ctx.sync();
  const out: ShapeSnapshot[] = [];
  for (const s of shapes.items) {
    const snap: ShapeSnapshot = {
      name: s.name,
      type: String(s.type),
      left: Math.round(s.left),
      top: Math.round(s.top),
      width: Math.round(s.width),
      height: Math.round(s.height),
    };
    // 非文本形状读取 textFrame 会报错，逐个容错
    try {
      const tr = s.textFrame.textRange;
      tr.load('text,font/name,font/size');
      await ctx.sync();
      snap.text = tr.text;
      snap.fontName = tr.font.name ?? undefined;
      snap.fontSize = tr.font.size ?? undefined;
    } catch {
      /* 无文本框 */
    }
    out.push(snap);
  }
  return out;
}

function diffShapes(a: ShapeSnapshot[], b: ShapeSnapshot[]): string[] {
  const diffs: string[] = [];
  if (a.length !== b.length) diffs.push(`形状数量不同：${a.length} → ${b.length}`);
  const n = Math.min(a.length, b.length);
  for (let i = 0; i < n; i++) {
    for (const k of Object.keys(a[i]) as (keyof ShapeSnapshot)[]) {
      if (k === 'name') continue; // 名称可能被重编号，不计入差异
      if (a[i][k] !== b[i][k]) diffs.push(`形状 ${i + 1} 的 ${k}：${a[i][k]} → ${b[i][k]}`);
    }
  }
  return diffs;
}

export async function runPptRoundtrip(cleanup = true): Promise<RoundtripResult> {
  if (!Office.context.requirements.isSetSupported('PowerPointApi', '1.8')) {
    return { ok: false, differences: [], error: '需要 PowerPointApi 1.8，当前版本不支持' };
  }
  try {
    return await PowerPoint.run(async (ctx) => {
      const slides = ctx.presentation.slides;

      // 1. 在末尾新建测试页并加文本框
      slides.add();
      await ctx.sync();
      slides.load('items/id');
      await ctx.sync();
      const testIndex = slides.items.length - 1;
      const test = slides.items[testIndex];
      const tb = test.shapes.addTextBox(PROBE_TEXT, { left: 60, top: 60, width: 480, height: 60 });
      tb.textFrame.textRange.font.name = '微软雅黑';
      tb.textFrame.textRange.font.size = 24;
      tb.fill.setSolidColor('#E8F3F1');
      await ctx.sync();

      // 2. 导出单页
      const b64 = test.exportAsBase64();
      await ctx.sync();

      // 3. 插回到测试页之后
      ctx.presentation.insertSlidesFromBase64(b64.value, {
        formatting: PowerPoint.InsertSlideFormatting.keepSourceFormatting,
        targetSlideId: test.id,
      });
      await ctx.sync();
      slides.load('items/id');
      await ctx.sync();
      const copy = slides.items[testIndex + 1];

      // 4. 比较
      const original = await snapshotShapes(ctx, test);
      const copied = await snapshotShapes(ctx, copy);
      const differences = diffShapes(original, copied);
      const result: RoundtripResult = {
        ok: differences.length === 0,
        base64Length: b64.value.length,
        originalId: test.id,
        copyId: copy.id,
        idChanged: test.id !== copy.id,
        original,
        copy: copied,
        differences,
      };

      // 5. 清理
      if (cleanup) {
        copy.delete();
        test.delete();
        await ctx.sync();
      }
      return result;
    });
  } catch (e: any) {
    return { ok: false, differences: [], error: `${e?.code ?? ''} ${e?.message ?? e}`.trim() };
  }
}
