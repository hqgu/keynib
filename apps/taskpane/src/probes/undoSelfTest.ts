// "撤销本轮"自测：不调用模型，按工具的真实调用方式模拟一轮改动，再撤销，比较前后指纹。
// 一轮包含：改第 1 页（覆盖"插回到开头"路径）→ 在第 2 位新建页 → 把第 3 页移到第 1 位 → 删除最后一页（原有页）→ 撤销。
// 指纹 = 页数 + 每页形状文本（按顺序）+ 母版数量（检查插回是否产生重复母版）。

import { PptChangeset, type InsertFormatting } from '../hosts/powerpoint/changeset';
import { readShapes } from '../hosts/powerpoint/tools';

interface Fingerprint {
  count: number;
  masters: number;
  slides: string[];
}

async function fingerprint(): Promise<Fingerprint> {
  return PowerPoint.run(async (ctx) => {
    const slides = ctx.presentation.slides;
    slides.load('items/id');
    const masters = ctx.presentation.slideMasters;
    masters.load('items/id');
    await ctx.sync();
    const out: string[] = [];
    for (let i = 0; i < slides.items.length; i++) {
      const shapes = await readShapes(ctx, slides.getItemAt(i)); // 按序号取页（网页版按 ID 查找新页会暂时失败）
      out.push(shapes.map((x) => `${x.type}:${x.text ?? ''}:${x.left},${x.top}`).join('|'));
    }
    return { count: slides.items.length, masters: masters.items.length, slides: out };
  });
}

function compare(a: Fingerprint, b: Fingerprint): string[] {
  const d: string[] = [];
  if (a.count !== b.count) d.push(`页数 ${a.count} → ${b.count}`);
  if (a.masters !== b.masters) d.push(`母版 ${a.masters} → ${b.masters}`);
  for (let i = 0; i < Math.max(a.slides.length, b.slides.length); i++) {
    if (a.slides[i] !== b.slides[i]) d.push(`第 ${i + 1} 页内容不一致`);
  }
  return d;
}

// 保证至少 3 页可供测试；返回为测试临时加的页 ID（测完删除）
async function ensureSlides(): Promise<string[]> {
  return PowerPoint.run(async (ctx) => {
    const slides = ctx.presentation.slides;
    slides.load('items/id');
    await ctx.sync();
    const added: string[] = [];
    for (let i = slides.items.length; i < 3; i++) {
      slides.add();
      await ctx.sync();
      slides.load('items/id');
      await ctx.sync();
      const s = slides.getItemAt(slides.items.length - 1);
      s.load('id');
      s.shapes.addTextBox(`自测页 ${String.fromCharCode(65 + i)}`, { left: 80, top: 80, width: 400, height: 60 });
      await ctx.sync();
      added.push(s.id);
    }
    return added;
  });
}

async function removeSlides(ids: string[]) {
  if (!ids.length) return;
  await PowerPoint.run(async (ctx) => {
    for (const id of ids) {
      const slides = ctx.presentation.slides;
      slides.load('items/id');
      await ctx.sync();
      const i = slides.items.findIndex((x) => x.id === id);
      if (i >= 0) {
        slides.getItemAt(i).delete();
        await ctx.sync();
      }
    }
  });
}

// 出错时标明阶段与 Office 的定位信息（GeneralException 本身没有细节）
export function describeOfficeError(e: any): string {
  const loc = e?.debugInfo?.errorLocation;
  return `${e?.code ?? ''} ${e?.message ?? e}${loc ? ` [${loc}]` : ''}`.trim();
}

async function phase<T>(name: string, fn: () => Promise<T>): Promise<T> {
  try {
    return await fn();
  } catch (e) {
    throw new Error(`${name}：${describeOfficeError(e)}`);
  }
}

// 模拟一轮，完全按工具的调用顺序：写前快照 → 执行；中途失败时尽量撤销，避免残留测试页
async function oneTurn(formatting: InsertFormatting): Promise<void> {
  const cs = new PptChangeset(formatting);
  try {
    await oneTurnSteps(cs);
  } catch (e) {
    await cs.undo().catch(() => {});
    throw e;
  }
}

async function oneTurnSteps(cs: PptChangeset): Promise<void> {

  // 1. 改第 1 页（加文本框，相当于 add_textbox / set_text）
  await phase('快照第 1 页', () => cs.beforeWrite([1]));
  await phase('改第 1 页', () => PowerPoint.run(async (ctx) => {
    const s = ctx.presentation.slides.getItemAt(0);
    s.shapes.addTextBox('UNDO-SELFTEST', { left: 20, top: 20, width: 200, height: 40 });
    await ctx.sync();
  }));

  // 2. 在第 2 位新建页（相当于 add_slide position=2）
  await phase('记录原顺序', () => cs.ensureOrder());
  await phase('新建并移到第 2 位', () => PowerPoint.run(async (ctx) => {
    const slides = ctx.presentation.slides;
    slides.add();
    await ctx.sync();
    slides.load('items/id');
    await ctx.sync();
    const s = slides.getItemAt(slides.items.length - 1);
    s.load('id');
    await ctx.sync();
    cs.noteCreated(s.id);
    s.moveTo(1);
    await ctx.sync();
  }));

  // 3. 把第 3 页移到第 1 位（相当于 move_slide）
  await phase('快照第 3 页', () => cs.beforeWrite([3]));
  await phase('移动第 3 页', () => PowerPoint.run(async (ctx) => {
    ctx.presentation.slides.getItemAt(2).moveTo(0);
    await ctx.sync();
  }));

  // 4. 删除最后一页（原有页，相当于 delete_slide）
  let last = 0;
  await PowerPoint.run(async (ctx) => {
    const slides = ctx.presentation.slides;
    slides.load('items');
    await ctx.sync();
    last = slides.items.length;
  });
  await phase('快照最后一页', () => cs.beforeWrite([last]));
  await phase('删除最后一页', () => PowerPoint.run(async (ctx) => {
    ctx.presentation.slides.getItemAt(last - 1).delete();
    await ctx.sync();
  }));

  // 5. 撤销
  await phase('撤销', () => cs.undo());
}

export interface UndoSelfTestResult {
  ok: boolean;
  formatting?: InsertFormatting;
  runs: number;
  passed: number;
  details: string[];
}

export async function runUndoSelfTest(times: number, onProgress?: (n: number) => void): Promise<UndoSelfTestResult> {
  if (!Office.context.requirements.isSetSupported('PowerPointApi', '1.8')) {
    return { ok: false, runs: 0, passed: 0, details: ['需要 PowerPointApi 1.8'] };
  }
  const temp = await phase('准备测试页', ensureSlides);
  try {
    // 先各试一次两种插回方式，选不产生差异（尤其不重复母版）的那种
    let chosen: InsertFormatting | undefined;
    const details: string[] = [];
    for (const f of ['KeepSourceFormatting', 'UseDestinationTheme'] as InsertFormatting[]) {
      const before = await phase('读取指纹', fingerprint);
      await oneTurn(f);
      const diff = compare(before, await phase('读取指纹', fingerprint));
      details.push(`${f}：${diff.length ? diff.join('；') : '无差异'}`);
      if (!diff.length) {
        chosen = f;
        break;
      }
    }
    if (!chosen) return { ok: false, runs: 0, passed: 0, details };

    let passed = 0;
    for (let i = 0; i < times; i++) {
      const before = await phase('读取指纹', fingerprint);
      await oneTurn(chosen);
      const diff = compare(before, await phase('读取指纹', fingerprint));
      if (diff.length) details.push(`第 ${i + 1} 次：${diff.join('；')}`);
      else passed++;
      onProgress?.(i + 1);
    }
    return { ok: passed === times, formatting: chosen, runs: times, passed, details };
  } finally {
    await removeSlides(temp);
  }
}
