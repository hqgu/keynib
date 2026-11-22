// 演示规范档案（Deck Style Profile）与规范检查（Deck Lint）
// 纯逻辑，不依赖 Office.js，可直接单元测试。

export interface DeckProfile {
  id: string;
  name: string;
  fonts: { title: string; body: string; allowed: string[] }; // allowed：允许出现的其他字体（如英文代码字体）
  sizes: { title: number; titleBold: boolean; body: number; min: number };
  density: 'full' | 'normal'; // full：版面饱满，留白多时提示
  citation: { enabled: boolean; example: string }; // 页脚学术短引
  noReferencesSlide: boolean;
  notes: string; // 其他自由规则，原样交给模型
}

export const DEFAULT_PROFILE: DeckProfile = {
  id: 'lecture',
  name: '讲课规范',
  fonts: { title: '微软雅黑', body: '微软雅黑', allowed: ['微软雅黑 Light', 'Microsoft YaHei', 'Microsoft YaHei Light', 'Consolas'] },
  sizes: { title: 36, titleBold: true, body: 20, min: 16 },
  density: 'full',
  citation: { enabled: true, example: 'Muller CJ. Int J Epidemiol. 2014; 43 (3): 962-970.' },
  noReferencesSlide: true,
  notes: '受众兼顾临床与流行病学统计本科至博士；概念与术语首次出现给出中英文及缩写。',
};

// 规范 → 系统提示词片段
export function profilePrompt(p: DeckProfile, opts: { pubmed?: boolean } = {}): string {
  const lines = [
    `# 演示规范「${p.name}」（必须遵守）`,
    `- 字体：标题 ${p.fonts.title}，正文 ${p.fonts.body}。新增或修改文字时显式设置字体。`,
    `- 字号：标题 ${p.sizes.title} 磅${p.sizes.titleBold ? '加粗' : ''}；正文一般 ${p.sizes.body} 磅；任何文字不小于 ${p.sizes.min} 磅。内容放不下时精简文字或拆页，不要缩小到 ${p.sizes.min} 磅以下。`,
    p.density === 'full' ? '- 版面饱满：避免大片留白，可适当增加要点说明、示意形状或表格，但不要拥挤或溢出。' : '',
    p.citation.enabled
      ? opts.pubmed
        ? `- 引用：在相应页的页脚用学术短引格式标注来源，例如「${p.citation.example}」。引用必须来自 pubmed_search 的结果并原样使用其 shortCite；已有引用用 verify_citation 核对，corrected 的按返回写法更正，not_found 的换成检索到的真实文献或标"[待核实]"。不得凭记忆编造。`
        : `- 引用：在相应页的页脚用学术短引格式标注来源，例如「${p.citation.example}」。只引用你确定存在的文献；不确定时写"[待核实]"，不要编造。`
      : '',
    p.noReferencesSlide ? '- 不要在末尾单独设置参考文献页。' : '',
    p.notes ? `- 其他要求：${p.notes}` : '',
    '- 完成修改后调用 lint_deck 检查本规范，修正发现的问题；字体字号的批量修正可用 apply_profile。',
  ];
  return lines.filter(Boolean).join('\n');
}

// —— 规范检查 ——

export interface LintShape {
  id: string;
  name: string;
  type: string;
  left: number;
  top: number;
  width: number;
  height: number;
  text?: string;
  font?: { name?: string | null; size?: number | null; bold?: boolean | null };
}

export interface LintIssue {
  slide: number;
  shapeId?: string;
  rule: 'font' | 'minSize' | 'titleSize' | 'titleBold' | 'bodySize' | 'offSlide' | 'citation' | 'whitespace' | 'referencesSlide';
  severity: 'error' | 'warn' | 'info';
  message: string;
}

export const SLIDE_W = 960;
export const SLIDE_H = 540;

export function isTitleShape(s: LintShape): boolean {
  return /^(title|标题)\b/i.test(s.name.trim()) || /^(center|ctr)?\s*title/i.test(s.name.trim());
}

// 学术短引：作者 缩写. 期刊. 年; 卷 (期): 页码.
const CITATION_RE = /^[A-Z][A-Za-z'’-]+(?: [A-Z][A-Za-z'’-]+)* [A-Z]{1,3}(?:,? et al)?\. .+\. (?:19|20)\d{2}; ?\d+(?: ?\(\w+\))?: ?[\w]+(?:[-–][\w]+)?\.?$/;

export function looksLikeCitation(text: string): boolean {
  return /(19|20)\d{2}/.test(text) && /[.;:]/.test(text) && text.length < 200;
}

export function isValidShortCitation(text: string): boolean {
  return text
    .split(/\n|(?<=\.)\s{2,}/)
    .map((t) => t.trim())
    .filter(Boolean)
    .every((t) => CITATION_RE.test(t));
}

const fontOk = (name: string | null | undefined, p: DeckProfile, isTitle: boolean) => {
  if (!name) return true; // 混合字体时 API 返回 null，交给截图判断
  const allowed = [isTitle ? p.fonts.title : p.fonts.body, p.fonts.title, p.fonts.body, ...p.fonts.allowed].map((x) => x.toLowerCase());
  return allowed.includes(name.toLowerCase());
};

export function lintSlide(slide: number, shapes: LintShape[], p: DeckProfile): LintIssue[] {
  const issues: LintIssue[] = [];
  let textArea = 0;
  for (const s of shapes) {
    const hasText = !!s.text?.trim();
    const title = isTitleShape(s);
    const label = `${title ? '标题' : '文字'}「${(s.text ?? s.name).replace(/\s+/g, ' ').slice(0, 16)}」`;

    // 超出页面（容差 2 磅）
    if (s.left < -2 || s.top < -2 || s.left + s.width > SLIDE_W + 2 || s.top + s.height > SLIDE_H + 2) {
      issues.push({ slide, shapeId: s.id, rule: 'offSlide', severity: 'error', message: `${label} 超出页面范围` });
    }
    if (!hasText) continue;
    textArea += Math.max(0, s.width) * Math.max(0, s.height);

    const f = s.font ?? {};
    if (!fontOk(f.name, p, title)) {
      issues.push({ slide, shapeId: s.id, rule: 'font', severity: 'error', message: `${label} 字体为 ${f.name}，应为 ${title ? p.fonts.title : p.fonts.body}` });
    }
    if (f.size != null && f.size < p.sizes.min) {
      issues.push({ slide, shapeId: s.id, rule: 'minSize', severity: 'error', message: `${label} 字号 ${f.size}，小于最小 ${p.sizes.min}` });
    }
    if (title) {
      if (f.size != null && f.size !== p.sizes.title) {
        issues.push({ slide, shapeId: s.id, rule: 'titleSize', severity: 'warn', message: `${label} 字号 ${f.size}，规范为 ${p.sizes.title}` });
      }
      if (p.sizes.titleBold && f.bold === false) {
        issues.push({ slide, shapeId: s.id, rule: 'titleBold', severity: 'warn', message: `${label} 未加粗` });
      }
    } else if (f.size != null && f.size > p.sizes.body + 4 && s.top < SLIDE_H * 0.85) {
      issues.push({ slide, shapeId: s.id, rule: 'bodySize', severity: 'info', message: `${label} 字号 ${f.size}，规范正文一般为 ${p.sizes.body}` });
    }

    // 页脚引用格式
    if (p.citation.enabled && s.top > SLIDE_H * 0.8 && looksLikeCitation(s.text!)) {
      // "[待核实]"是有意的标记：去掉后再校验格式，并单独提示
      const pendingMark = /\s*\[待核实\]/g;
      if (pendingMark.test(s.text!)) {
        issues.push({ slide, shapeId: s.id, rule: 'citation', severity: 'info', message: '有标注"[待核实]"的引用，需要人工确认来源' });
      }
      if (!isValidShortCitation(s.text!.replace(pendingMark, ''))) {
        issues.push({ slide, shapeId: s.id, rule: 'citation', severity: 'warn', message: `页脚引用格式不符，应形如「${p.citation.example}」` });
      }
    }
    if (p.noReferencesSlide && title && /^(参考文献|references?|bibliography)$/i.test(s.text!.trim())) {
      issues.push({ slide, shapeId: s.id, rule: 'referencesSlide', severity: 'warn', message: '规范要求不单独设置参考文献页' });
    }
  }
  // 版面饱满：文字区域占比过低时提示（标题页除外）
  if (p.density === 'full' && slide > 1 && textArea > 0 && textArea / (SLIDE_W * SLIDE_H) < 0.35) {
    const usedBottom = Math.max(...shapes.filter((s) => s.text?.trim()).map((s) => s.top + s.height));
    if (usedBottom < SLIDE_H * 0.6) {
      issues.push({ slide, rule: 'whitespace', severity: 'info', message: '下方留白较多，可补充内容或调整版式（建议截图确认）' });
    }
  }
  return issues;
}

// apply_profile 的确定性修正：返回每个形状应设置的字体
export function profileFontFor(s: LintShape, p: DeckProfile, opts: { setBodySize: boolean }) {
  const title = isTitleShape(s);
  const size = s.font?.size ?? undefined;
  if (title) return { name: p.fonts.title, size: p.sizes.title, bold: p.sizes.titleBold || undefined };
  let newSize: number | undefined;
  if (opts.setBodySize && s.top < SLIDE_H * 0.85) newSize = p.sizes.body;
  else if (size != null && size < p.sizes.min) newSize = p.sizes.min;
  return { name: p.fonts.body, size: newSize };
}
