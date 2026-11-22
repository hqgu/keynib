// 面板端的开发测试通道客户端（仅开发模式加载）
// 只提供固定命令：自测、发消息、撤销最近一轮、只读查看幻灯片、读取（脱敏）设置。
// 不能执行任意代码；确认卡片（删除等破坏性操作）只能由用户在面板上点击，通道不代答。

import { runUndoSelfTest } from '../probes/undoSelfTest';
import { activeConnection, getSettings, updateProvider, updateSettings } from '../state/settings';
import { PROVIDER_PRESETS } from '@keynib/providers';
import { testConnection } from '@keynib/providers';
import { readShapes } from '../hosts/powerpoint/tools';
import { runWordProbe } from '../probes/wordProbe';
import { runExcelProbe } from '../probes/excelProbe';
import { runPptCapabilityProbe } from '../probes/pptCapabilityProbe';
import { runExcelUndoSelfTest } from '../probes/excelUndoSelfTest';
import { SEED_HEADER, seedTrial } from '../hosts/excel/seed';

type Handler = (payload: any) => Promise<unknown> | unknown;

export const devApi: Record<string, Handler> = {
  ping: () => ({ host: Office.context?.host, time: new Date().toISOString(), officeLang: Office.context?.displayLanguage, contentLang: (Office.context as any)?.contentLanguage, navLang: navigator.language, navLangs: navigator.languages }),
  selftest: (times?: number) => runUndoSelfTest(times ?? 20),
  // 只读：整套幻灯片概况（每页形状类型、文本、字体、字号）
  deck: () =>
    PowerPoint.run(async (ctx) => {
      const slides = ctx.presentation.slides;
      slides.load('items/id');
      const masters = ctx.presentation.slideMasters;
      masters.load('items/id');
      await ctx.sync();
      const out = [];
      for (let i = 0; i < slides.items.length; i++) {
        const shapes = await readShapes(ctx, slides.items[i]);
        out.push({
          slide: i + 1,
          shapes: shapes.map((s) => ({ type: s.type, text: s.text?.slice(0, 80), font: s.font?.name, size: s.font?.size, box: [s.left, s.top, s.width, s.height] })),
        });
      }
      return { count: slides.items.length, masters: masters.items.length, slides: out };
    }),
  // 用当前选中的厂商做三步连接测试（Key 只在面板内使用）
  testConn: async () => {
    const a = activeConnection(getSettings());
    if (!a) return '未配置当前模型';
    return { provider: a.id, model: a.cfg.model, steps: await testConnection(a.conn, a.cfg.model) };
  },
  // 只读：某页截图（PNG base64），用于人工核对版面
  shot: (n: number) =>
    PowerPoint.run(async (ctx) => {
      const s = ctx.presentation.slides.getItemAt((n ?? 1) - 1);
      const img = s.getImageAsBase64({ width: 960 });
      await ctx.sync();
      return { png: img.value };
    }),
  // 切换当前使用的厂商（必须是已填好 Key 和模型的）
  useProvider: (id: string) => {
    const cfg = getSettings().providers[id];
    if (!cfg?.apiKey || !cfg.model) return `${id} 尚未配置完整`;
    updateSettings({ activeProviderId: id });
    return { active: id, model: cfg.model };
  },
  // 切换某厂商的预置接口地址（同时切换协议），如 DeepSeek 的 Anthropic 兼容接口
  setEndpoint: ({ id, label }: { id: string; label: string }) => {
    const opt = PROVIDER_PRESETS.find((p) => p.id === id)?.baseUrls.find((b) => b.label === label);
    if (!opt) return '没有这个接口选项';
    updateProvider(id, { baseUrl: opt.url, protocol: opt.protocol });
    return opt;
  },
  useProfile: (id: string | null) => {
    updateSettings({ activeProfileId: id });
    return { activeProfileId: id };
  },
  setPubmed: (on: boolean) => {
    updateSettings({ pubmed: !!on });
    return { pubmed: !!on };
  },
  // 只读：Word 文档概况（段落数、是否已保存、前几段、修订模式）
  doc: () =>
    Word.run(async (ctx) => {
      const ps = ctx.document.body.paragraphs;
      ps.load('items/text,items/style');
      ctx.document.load('changeTrackingMode');
      await ctx.sync();
      return {
        url: Office.context.document.url || '(未保存)',
        paragraphs: ps.items.length,
        trackingMode: ctx.document.changeTrackingMode,
        first: ps.items.slice(0, 5).map((p) => `[${p.style}] ${p.text.slice(0, 60)}`),
      };
    }),
  wordProbe: () => runWordProbe(),
  excelProbe: () => runExcelProbe(),
  pptCaps: () => runPptCapabilityProbe(),
  excelSelftest: (n?: number) => runExcelUndoSelfTest(n ?? 20),
  // 测试用：把固定种子的 200 例试验数据写入 Sheet1（仅当 Sheet1 为空）
  excelSeed: () =>
    Excel.run(async (ctx) => {
      const ws = ctx.workbook.worksheets.getItem('Sheet1');
      const used = ws.getUsedRangeOrNullObject(true);
      await ctx.sync();
      if (!used.isNullObject) return 'Sheet1 非空，未写入';
      const rows = seedTrial();
      ws.getRange(`A1:F${rows.length + 1}`).values = [SEED_HEADER, ...rows];
      ws.getRange('A1:F1').format.font.bold = true;
      await ctx.sync();
      return `已写入 ${rows.length} 行`;
    }),
  // 只读：某表已用区域的 formulas 与 values（用于核对结果）
  excelSheet: (name: string) =>
    Excel.run(async (ctx) => {
      const used = ctx.workbook.worksheets.getItem(name).getUsedRangeOrNullObject(true);
      used.load('address,formulas,values');
      const charts = ctx.workbook.worksheets.getItem(name).charts;
      charts.load('items/name,items/chartType');
      await ctx.sync();
      return used.isNullObject ? { empty: true } : { address: used.address, formulas: used.formulas, values: used.values, charts: charts.items.map((c) => c.chartType) };
    }),
  // 只读：Excel 工作簿概况
  book: () =>
    Excel.run(async (ctx) => {
      const sheets = ctx.workbook.worksheets;
      sheets.load('items/name');
      await ctx.sync();
      const out = [];
      for (const ws of sheets.items) {
        const used = ws.getUsedRangeOrNullObject(true);
        used.load('address,rowCount,columnCount');
        const charts = ws.charts;
        charts.load('items/name');
        const tables = ws.tables;
        tables.load('items/name');
        await ctx.sync();
        out.push({ sheet: ws.name, used: used.isNullObject ? null : `${used.address}（${used.rowCount}×${used.columnCount}）`, charts: charts.items.length, tables: tables.items.length });
      }
      return { url: Office.context.document.url || '(未保存)', sheets: out };
    }),
  // 只读：Word 当前文本（不含已删除文字）、原始文本、修订列表、修订模式
  wordState: () =>
    Word.run(async (ctx) => {
      const body = ctx.document.body;
      const cur = body.getRange('Whole').getReviewedText(Word.ChangeTrackingVersion.current);
      const tc = body.getTrackedChanges();
      tc.load('items/type,items/text');
      ctx.document.load('changeTrackingMode');
      const ps = body.paragraphs;
      ps.load('items/text');
      await ctx.sync();
      return {
        mode: ctx.document.changeTrackingMode,
        paragraphs: ps.items.map((p, i) => `${i + 1}: ${p.text}`),
        current: cur.value,
        tracked: tc.items.map((t) => `${t.type}: ${t.text}`),
      };
    }),
  // 测试用：拒绝文本为空的 Deleted 修订（修复前留下的残留）
  wordRejectEmptyDeleted: () =>
    Word.run(async (ctx) => {
      const tc = ctx.document.body.getTrackedChanges();
      tc.load('items/type,items/text');
      await ctx.sync();
      const targets = tc.items.filter((t) => t.type === 'Deleted' && !t.text.trim());
      targets.forEach((t) => t.reject());
      await ctx.sync();
      return `rejected ${targets.length}`;
    }),
  // 测试用：模拟"用户自己已有的一条修订"（开修订插入一句，再恢复原修订模式）
  wordSeedRevision: (text: string) =>
    Word.run(async (ctx) => {
      ctx.document.load('changeTrackingMode');
      await ctx.sync();
      const mode = ctx.document.changeTrackingMode;
      ctx.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
      ctx.document.body.insertParagraph(text || '（这是用户自己已有的一条修订）', Word.InsertLocation.end);
      ctx.document.changeTrackingMode = mode;
      await ctx.sync();
      return 'ok';
    }),
  setLanguage: (l: 'auto' | 'zh' | 'en') => {
    updateSettings({ language: l });
    return { language: l };
  },
  // 只读：面板当前显示的文字（检查翻译）
  uiText: () => document.body.innerText,
  settings: () => {
    const s = getSettings();
    return {
      ...s,
      providers: Object.fromEntries(
        Object.entries(s.providers).map(([k, v]) => [k, { ...v, apiKey: v.apiKey ? `(已填，${v.apiKey.length} 位)` : '' }]),
      ),
    };
  },
};

// 递归剔除可能出现的 Key 字段，并截断超长字符串（截图 base64 等）
function sanitize(v: unknown, depth = 0): unknown {
  if (depth > 8) return '…';
  if (typeof v === 'string') return v.length > 4000 && !/^iVBOR/.test(v) ? v.slice(0, 4000) + `…(${v.length})` : v;
  if (Array.isArray(v)) return v.map((x) => sanitize(x, depth + 1));
  if (v && typeof v === 'object') {
    return Object.fromEntries(
      Object.entries(v).map(([k, x]) => [k, /apikey|authorization|x-api-key/i.test(k) ? '(隐藏)' : sanitize(x, depth + 1)]),
    );
  }
  return v;
}

export function startDevBridge() {
  const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms));
  (async () => {
    for (;;) {
      try {
        // 按"宿主@平台"路由，区分桌面版与网页版（如 PowerPoint@Mac、PowerPoint@OfficeOnline）
        const host = encodeURIComponent(`${Office.context?.host ?? 'Unknown'}@${Office.context?.platform ?? 'Unknown'}`);
        const res = await fetch(`/__bridge/next?host=${host}`, { cache: 'no-store' });
        if (res.status !== 200) continue;
        const cmd = await res.json();
        let out: { ok: boolean; value?: unknown; error?: string };
        try {
          const h = devApi[cmd.type];
          if (!h) throw new Error(`未知命令：${cmd.type}（可用：${Object.keys(devApi).join(', ')}）`);
          out = { ok: true, value: sanitize(await h(cmd.payload)) };
        } catch (e: any) {
          out = { ok: false, error: String(e?.stack ?? e?.message ?? e) };
        }
        await fetch('/__bridge/result', { method: 'POST', body: JSON.stringify({ id: cmd.id, ...out }) });
      } catch {
        await sleep(2000);
      }
    }
  })();
}
