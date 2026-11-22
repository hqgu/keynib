// PowerPoint 工具集（M1）。全部基于 M0 核实过的 PowerPointApi ≤ 1.8。
// 约定：幻灯片用 1 起的序号 slide 指代；形状用 read_slide 返回的 shapeId 指代；单位为磅（pt）。

import type { ToolDef } from '@keynib/core';
import { m, t } from '../../i18n';
import type { PptChangeset } from './changeset';
import { lintSlide, profileFontFor, type DeckProfile, type LintIssue } from './profile';

const TEXT_TYPES = new Set(['GeometricShape', 'TextBox', 'Placeholder']);

async function getSlides(ctx: PowerPoint.RequestContext) {
  const slides = ctx.presentation.slides;
  slides.load('items/id');
  await ctx.sync();
  return slides.items;
}

async function slideAt(ctx: PowerPoint.RequestContext, n: number) {
  const items = await getSlides(ctx);
  const s = items[n - 1];
  if (!s) throw new Error(`第 ${n} 页不存在（共 ${items.length} 页）`);
  return s;
}

interface ShapeInfo {
  id: string;
  name: string;
  type: string;
  left: number;
  top: number;
  width: number;
  height: number;
  text?: string;
  font?: { name?: string | null; size?: number | null; bold?: boolean | null; color?: string | null };
  table?: string[][];
}

const r1 = (n: number) => Math.round(n * 10) / 10;

export async function readShapes(ctx: PowerPoint.RequestContext, slide: PowerPoint.Slide, withText = true): Promise<ShapeInfo[]> {
  const shapes = slide.shapes;
  shapes.load('items/id,items/name,items/type,items/left,items/top,items/width,items/height');
  await ctx.sync();
  const infos: ShapeInfo[] = shapes.items.map((s) => ({
    id: s.id,
    name: s.name,
    type: String(s.type),
    left: r1(s.left),
    top: r1(s.top),
    width: r1(s.width),
    height: r1(s.height),
  }));
  if (!withText) return infos;

  // 先整批读取文本，失败再逐个容错（个别占位符没有文本框）
  const readOne = async (i: number) => {
    const s = shapes.items[i];
    if (infos[i].type === 'Table') {
      const t = s.getTable();
      t.load('values');
      await ctx.sync();
      infos[i].table = t.values;
      return;
    }
    if (!TEXT_TYPES.has(infos[i].type)) return;
    const tr = s.textFrame.textRange;
    tr.load('text,font/name,font/size,font/bold,font/color');
    await ctx.sync();
    infos[i].text = tr.text;
    infos[i].font = { name: tr.font.name, size: tr.font.size, bold: tr.font.bold, color: tr.font.color };
  };
  for (let i = 0; i < infos.length; i++) {
    try {
      await readOne(i);
    } catch {
      /* 该形状不支持文本 */
    }
  }
  return infos;
}

const json = (v: unknown) => JSON.stringify(v);

const fontSchema = {
  type: 'object',
  description: '字体设置，均可选',
  properties: {
    name: { type: 'string', description: '字体名，如 微软雅黑' },
    size: { type: 'number', description: '字号（磅）' },
    bold: { type: 'boolean' },
    italic: { type: 'boolean' },
    color: { type: 'string', description: '#RRGGBB' },
  },
};
const boxSchema = {
  left: { type: 'number' },
  top: { type: 'number' },
  width: { type: 'number' },
  height: { type: 'number' },
};

interface FontArgs {
  name?: string;
  size?: number;
  bold?: boolean;
  italic?: boolean;
  color?: string;
}

function applyFont(font: PowerPoint.ShapeFont, f?: FontArgs) {
  if (!f) return;
  if (f.name) font.name = f.name;
  if (f.size) font.size = f.size;
  if (f.bold !== undefined) font.bold = f.bold;
  if (f.italic !== undefined) font.italic = f.italic;
  if (f.color) font.color = f.color;
}

export function createPptTools(getChangeset: () => PptChangeset, getProfile: () => DeckProfile | null = () => null): ToolDef[] {
  return [
    {
      name: 'lint_deck',
      label: t('规范检查'),
      permission: 'R',
      description: '按当前演示规范检查幻灯片：字体、最小字号、标题字号与加粗、超出页面、页脚引用格式、留白。slides 省略则检查全部。',
      parameters: { type: 'object', properties: { slides: { type: 'array', items: { type: 'integer' }, description: '要检查的页序号，省略为全部' } } },
      run: ({ slides }: { slides?: number[] }) =>
        PowerPoint.run(async (ctx) => {
          const profile = getProfile();
          if (!profile) return { text: m('当前未选择演示规范，无法检查。请让用户在输入框旁选择规范。'), summary: t('未选择规范') };
          const all = await getSlides(ctx);
          const nums = slides?.length ? slides : all.map((_, i) => i + 1);
          const issues: LintIssue[] = [];
          for (const n of nums) {
            if (!all[n - 1]) continue;
            issues.push(...lintSlide(n, await readShapes(ctx, all[n - 1]), profile));
          }
          const errors = issues.filter((i) => i.severity === 'error').length;
          const warns = issues.filter((i) => i.severity === 'warn').length;
          return {
            text: json({ profile: profile.name, checked: nums.length, errors, warns, issues }),
            summary: issues.length
              ? t('{e} 个错误，{w} 个警告，{i} 条提示', { e: errors, w: warns, i: issues.length - errors - warns })
              : t('{n} 页全部符合', { n: nums.length }),
          };
        }),
    },
    {
      name: 'apply_profile',
      label: t('套用规范'),
      permission: 'W',
      description: '把当前演示规范的字体、标题字号与加粗、最小字号一次性套用到指定页的所有文字（比逐个 format_shape 快）。setBodySize=true 时正文统一为规范正文字号。slides 省略为全部。',
      parameters: {
        type: 'object',
        properties: { slides: { type: 'array', items: { type: 'integer' } }, setBodySize: { type: 'boolean' } },
      },
      targets: (a: { slides?: number[] }) => (a.slides?.length ? a.slides : 'all'),
      run: ({ slides, setBodySize }: { slides?: number[]; setBodySize?: boolean }) =>
        PowerPoint.run(async (ctx) => {
          const profile = getProfile();
          if (!profile) throw new Error('当前未选择演示规范');
          const all = await getSlides(ctx);
          const nums = slides?.length ? slides : all.map((_, i) => i + 1);
          let changed = 0;
          for (const n of nums) {
            const slide = all[n - 1];
            if (!slide) continue;
            for (const sh of await readShapes(ctx, slide)) {
              if (!sh.text?.trim()) continue;
              const f = profileFontFor(sh, profile, { setBodySize: !!setBodySize });
              const font = slide.shapes.getItem(sh.id).textFrame.textRange.font;
              font.name = f.name;
              if (f.size) font.size = f.size;
              if (f.bold) font.bold = true;
              changed++;
            }
          }
          await ctx.sync();
          return { text: json({ slides: nums.length, shapes: changed }), summary: t('{n} 页 {c} 处文字', { n: nums.length, c: changed }) };
        }),
    },
    {
      name: 'list_slides',
      label: t('浏览幻灯片'),
      permission: 'R',
      description: '列出所有幻灯片：序号、版式、标题及形状数量。改动前先调用以了解整体结构。',
      parameters: { type: 'object', properties: {} },
      run: () =>
        PowerPoint.run(async (ctx) => {
          const slides = await getSlides(ctx);
          slides.forEach((s) => s.layout.load('name'));
          await ctx.sync();
          const out = [];
          for (let i = 0; i < slides.length; i++) {
            const shapes = await readShapes(ctx, slides[i]);
            const titled = shapes.find((s) => /title|标题/i.test(s.name) && s.text) ?? shapes.find((s) => s.text);
            out.push({ slide: i + 1, layout: slides[i].layout.name, title: titled?.text?.slice(0, 60) ?? '', shapes: shapes.length });
          }
          return { text: json(out), summary: t('共 {n} 页', { n: slides.length }) };
        }),
    },
    {
      name: 'read_slide',
      label: t('读取幻灯片'),
      permission: 'R',
      description: '读取一页的全部形状：shapeId、名称、类型、位置尺寸（磅）、文本、字体、表格内容。修改形状前必须先读取以获得 shapeId。',
      parameters: { type: 'object', properties: { slide: { type: 'integer', description: '序号，从 1 开始' } }, required: ['slide'] },
      run: ({ slide }: { slide: number }) =>
        PowerPoint.run(async (ctx) => {
          const s = await slideAt(ctx, slide);
          s.layout.load('name');
          await ctx.sync();
          const shapes = await readShapes(ctx, s);
          return { text: json({ slide, layout: s.layout.name, shapes }), summary: t('第 {n} 页，{s} 个形状', { n: slide, s: shapes.length }) };
        }),
    },
    {
      name: 'get_selection',
      label: t('读取当前选择'),
      permission: 'R',
      description: '获取用户当前选中的幻灯片序号和形状 shapeId。用户说"这页""选中的"时调用。',
      parameters: { type: 'object', properties: {} },
      run: () =>
        PowerPoint.run(async (ctx) => {
          const all = await getSlides(ctx);
          const sel = ctx.presentation.getSelectedSlides();
          sel.load('items/id');
          const shapes = ctx.presentation.getSelectedShapes();
          shapes.load('items/id,items/name');
          await ctx.sync();
          const slideNums = sel.items.map((s) => all.findIndex((a) => a.id === s.id) + 1);
          return {
            text: json({ slides: slideNums, shapes: shapes.items.map((s) => ({ id: s.id, name: s.name })) }),
            summary: t('选中第 {list} 页', { list: slideNums.join(', ') || '-' }),
          };
        }),
    },
    {
      name: 'get_layouts',
      label: t('读取版式'),
      permission: 'R',
      description: '列出母版及其版式（layoutId、名称、占位符）。新建幻灯片前调用，优先选用合适版式而不是空白页。',
      parameters: { type: 'object', properties: {} },
      run: () =>
        PowerPoint.run(async (ctx) => {
          const masters = ctx.presentation.slideMasters;
          masters.load('items/id,items/name');
          await ctx.sync();
          const out = [];
          for (const m of masters.items) {
            m.layouts.load('items/id,items/name');
            await ctx.sync();
            const layouts = [];
            for (const l of m.layouts.items) {
              l.shapes.load('items/name');
              await ctx.sync();
              layouts.push({ layoutId: l.id, name: l.name, placeholders: l.shapes.items.map((s) => s.name) });
            }
            out.push({ masterId: m.id, name: m.name, layouts });
          }
          return { text: json(out), summary: t('{n} 个版式', { n: out.reduce((n, x) => n + x.layouts.length, 0) }) };
        }),
    },
    {
      name: 'screenshot_slide',
      label: t('查看幻灯片截图'),
      permission: 'R',
      needsVision: true,
      description: '获取一页的渲染截图，用于检查排版：文字溢出、重叠、留白、对齐。完成一批修改后建议调用。',
      parameters: { type: 'object', properties: { slide: { type: 'integer' } }, required: ['slide'] },
      run: ({ slide }: { slide: number }) =>
        PowerPoint.run(async (ctx) => {
          const s = await slideAt(ctx, slide);
          const img = s.getImageAsBase64({ width: 960 });
          await ctx.sync();
          return { text: m('第 {n} 页截图如下。').replace('{n}', String(slide)), images: [{ mediaType: 'image/png', data: img.value }], summary: t('第 {n} 页截图', { n: slide }) };
        }),
    },
    {
      name: 'add_slide',
      label: t('新建幻灯片'),
      permission: 'W',
      description: '新建一页。可指定 layoutId（来自 get_layouts）与插入位置 position（新页的序号，默认放在最后）。返回新页的序号和占位符 shapeId，之后用 set_text 填写。',
      parameters: {
        type: 'object',
        properties: { layoutId: { type: 'string' }, position: { type: 'integer', description: '新页序号，从 1 开始' } },
      },
      run: async ({ layoutId, position }: { layoutId?: string; position?: number }) => {
        const cs = getChangeset();
        await cs.ensureOrder();
        return PowerPoint.run(async (ctx) => {
          ctx.presentation.slides.add(layoutId ? { layoutId } : undefined);
          await ctx.sync();
          let items = await getSlides(ctx);
          let slide = items[items.length - 1];
          cs.noteCreated(slide.id);
          if (position && position >= 1 && position < items.length) {
            slide.moveTo(position - 1);
            await ctx.sync();
            items = await getSlides(ctx);
          }
          const n = items.findIndex((s) => s.id === slide.id) + 1;
          slide = items[n - 1];
          const shapes = await readShapes(ctx, slide, false);
          return {
            text: json({ slide: n, placeholders: shapes.map(({ id, name, left, top, width, height }) => ({ id, name, left, top, width, height })) }),
            summary: t('新建第 {n} 页', { n }),
          };
        });
      },
    },
    {
      name: 'set_text',
      label: t('修改文字'),
      permission: 'W',
      description: '替换某个形状（文本框 / 占位符 / 形状）的全部文字。多段用 \\n 分隔。可同时设置字体。',
      parameters: {
        type: 'object',
        properties: { slide: { type: 'integer' }, shapeId: { type: 'string' }, text: { type: 'string' }, font: fontSchema },
        required: ['slide', 'shapeId', 'text'],
      },
      targets: (a: { slide: number }) => [a.slide],
      run: ({ slide, shapeId, text, font }: { slide: number; shapeId: string; text: string; font?: FontArgs }) =>
        PowerPoint.run(async (ctx) => {
          const s = await slideAt(ctx, slide);
          const tr = s.shapes.getItem(shapeId).textFrame.textRange;
          tr.text = text;
          applyFont(tr.font, font);
          await ctx.sync();
          return { text: 'ok', summary: t('第 {n} 页：{text}', { n: slide, text: text.replace(/\n/g, ' ').slice(0, 30) }) };
        }),
    },
    {
      name: 'format_shape',
      label: t('设置格式'),
      permission: 'W',
      description: '设置形状的字体（作用于全部文字）、填充色、位置和尺寸。只传需要修改的字段。',
      parameters: {
        type: 'object',
        properties: {
          slide: { type: 'integer' },
          shapeId: { type: 'string' },
          font: fontSchema,
          fill: { type: 'string', description: '填充色 #RRGGBB，或 "none" 清除' },
          ...boxSchema,
        },
        required: ['slide', 'shapeId'],
      },
      targets: (a: { slide: number }) => [a.slide],
      run: (a: { slide: number; shapeId: string; font?: FontArgs; fill?: string; left?: number; top?: number; width?: number; height?: number }) =>
        PowerPoint.run(async (ctx) => {
          const s = await slideAt(ctx, a.slide);
          const shape = s.shapes.getItem(a.shapeId);
          if (a.font) applyFont(shape.textFrame.textRange.font, a.font);
          if (a.fill === 'none') shape.fill.clear();
          else if (a.fill) shape.fill.setSolidColor(a.fill);
          if (a.left !== undefined) shape.left = a.left;
          if (a.top !== undefined) shape.top = a.top;
          if (a.width !== undefined) shape.width = a.width;
          if (a.height !== undefined) shape.height = a.height;
          await ctx.sync();
          return { text: 'ok', summary: t('第 {n} 页：调整格式', { n: a.slide }) };
        }),
    },
    {
      name: 'add_textbox',
      label: t('添加文本框'),
      permission: 'W',
      description: '在指定页添加文本框。位置尺寸单位为磅；16:9 页面通常为 960×540。',
      parameters: {
        type: 'object',
        properties: { slide: { type: 'integer' }, text: { type: 'string' }, font: fontSchema, ...boxSchema },
        required: ['slide', 'text', 'left', 'top', 'width', 'height'],
      },
      targets: (a: { slide: number }) => [a.slide],
      run: (a: { slide: number; text: string; font?: FontArgs; left: number; top: number; width: number; height: number }) =>
        PowerPoint.run(async (ctx) => {
          const s = await slideAt(ctx, a.slide);
          const tb = s.shapes.addTextBox(a.text, { left: a.left, top: a.top, width: a.width, height: a.height });
          applyFont(tb.textFrame.textRange.font, a.font);
          tb.load('id');
          await ctx.sync();
          return { text: json({ shapeId: tb.id }), summary: t('第 {n} 页：加文本框', { n: a.slide }) };
        }),
    },
    {
      name: 'add_shape',
      label: t('添加形状'),
      permission: 'W',
      description: '添加几何形状（如 Rectangle、RoundRectangle、Ellipse、RightArrow、Chevron、Pentagon），可带文字和填充色。',
      parameters: {
        type: 'object',
        properties: {
          slide: { type: 'integer' },
          shape: { type: 'string', description: 'GeometricShapeType，如 Rectangle' },
          text: { type: 'string' },
          fill: { type: 'string', description: '#RRGGBB' },
          font: fontSchema,
          ...boxSchema,
        },
        required: ['slide', 'shape', 'left', 'top', 'width', 'height'],
      },
      targets: (a: { slide: number }) => [a.slide],
      run: (a: { slide: number; shape: string; text?: string; fill?: string; font?: FontArgs; left: number; top: number; width: number; height: number }) =>
        PowerPoint.run(async (ctx) => {
          const s = await slideAt(ctx, a.slide);
          const sh = s.shapes.addGeometricShape(a.shape as PowerPoint.GeometricShapeType, {
            left: a.left,
            top: a.top,
            width: a.width,
            height: a.height,
          });
          if (a.fill) sh.fill.setSolidColor(a.fill);
          if (a.text) {
            sh.textFrame.textRange.text = a.text;
            applyFont(sh.textFrame.textRange.font, a.font);
          }
          sh.load('id');
          await ctx.sync();
          return { text: json({ shapeId: sh.id }), summary: t('第 {n} 页：加 {shape}', { n: a.slide, shape: a.shape }) };
        }),
    },
    {
      name: 'add_table',
      label: t('添加表格'),
      permission: 'W',
      description: '在指定页添加表格。values 为二维字符串数组，第一行通常是表头。',
      parameters: {
        type: 'object',
        properties: {
          slide: { type: 'integer' },
          values: { type: 'array', items: { type: 'array', items: { type: 'string' } } },
          fontSize: { type: 'number' },
          ...boxSchema,
        },
        required: ['slide', 'values', 'left', 'top', 'width'],
      },
      targets: (a: { slide: number }) => [a.slide],
      run: (a: { slide: number; values: string[][]; fontSize?: number; left: number; top: number; width: number; height?: number }) =>
        PowerPoint.run(async (ctx) => {
          const rows = a.values.length;
          const cols = Math.max(...a.values.map((r) => r.length));
          const values = a.values.map((r) => [...r, ...Array(cols - r.length).fill('')]);
          const s = await slideAt(ctx, a.slide);
          const sh = s.shapes.addTable(rows, cols, {
            values,
            left: a.left,
            top: a.top,
            width: a.width,
            height: a.height,
            uniformCellProperties: a.fontSize ? { font: { size: a.fontSize } } : undefined,
          });
          sh.load('id');
          await ctx.sync();
          return { text: json({ shapeId: sh.id }), summary: t('第 {n} 页：{r}×{c} 表格', { n: a.slide, r: rows, c: cols }) };
        }),
    },
    {
      name: 'move_slide',
      label: t('移动幻灯片'),
      permission: 'W',
      description: '把第 slide 页移动到第 to 页的位置。',
      parameters: { type: 'object', properties: { slide: { type: 'integer' }, to: { type: 'integer' } }, required: ['slide', 'to'] },
      targets: (a: { slide: number }) => [a.slide],
      run: ({ slide, to }: { slide: number; to: number }) =>
        PowerPoint.run(async (ctx) => {
          const s = await slideAt(ctx, slide);
          s.moveTo(to - 1);
          await ctx.sync();
          return { text: 'ok', summary: t('第 {n} 页 → 第 {to} 页', { n: slide, to }) };
        }),
    },
    {
      name: 'delete_shape',
      label: t('删除形状'),
      permission: 'D',
      description: '删除某页上的一个形状。需要用户确认。',
      parameters: { type: 'object', properties: { slide: { type: 'integer' }, shapeId: { type: 'string' } }, required: ['slide', 'shapeId'] },
      describe: (a: { slide: number; shapeId: string }) => t('删除第 {n} 页上的形状 {id}', { n: a.slide, id: a.shapeId }),
      targets: (a: { slide: number }) => [a.slide],
      run: ({ slide, shapeId }: { slide: number; shapeId: string }) =>
        PowerPoint.run(async (ctx) => {
          const s = await slideAt(ctx, slide);
          s.shapes.getItem(shapeId).delete();
          await ctx.sync();
          return { text: 'ok', summary: t('第 {n} 页：删除形状', { n: slide }) };
        }),
    },
    {
      name: 'delete_slide',
      label: t('删除幻灯片'),
      permission: 'D',
      description: '删除一页。需要用户确认。',
      parameters: { type: 'object', properties: { slide: { type: 'integer' } }, required: ['slide'] },
      describe: (a: { slide: number }) => t('删除第 {n} 页', { n: a.slide }),
      targets: (a: { slide: number }) => [a.slide],
      run: ({ slide }: { slide: number }) =>
        PowerPoint.run(async (ctx) => {
          const s = await slideAt(ctx, slide);
          s.delete();
          await ctx.sync();
          return { text: 'ok', summary: t('删除第 {n} 页', { n: slide }) };
        }),
    },
  ];
}

// 每轮开始时提供给模型的文档概况
export async function pptContext(): Promise<string> {
  return PowerPoint.run(async (ctx) => {
    const slides = await getSlides(ctx);
    const sel = ctx.presentation.getSelectedSlides();
    sel.load('items/id');
    await ctx.sync();
    const current = sel.items[0] ? slides.findIndex((s) => s.id === sel.items[0].id) + 1 : 0;
    return `演示文稿共 ${slides.length} 页；用户当前所在：第 ${current || '-'} 页。`;
  });
}
