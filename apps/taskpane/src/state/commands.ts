// 快捷指令（Slash Commands）：在输入框开头输入 "/" 调出
// 纯逻辑，不依赖 Office.js / React，可单元测试。

export interface SlashCommand {
  id: string;
  name: string; // 显示名，如 "润色"
  aliases: string[]; // 拼音首字母 / 英文，方便不切输入法筛选
  description: string;
  prompt: string;
  autoSend: boolean; // true：选中即发送；false：填入输入框供补充
  hosts?: string[]; // 仅在这些宿主显示；省略为全部
  requires?: 'profile' | 'pubmed';
  builtin?: boolean;
}

export function lintPrompt(pubmed: boolean, lang: 'zh' | 'en' = 'zh'): string {
  if (lang === 'en') {
    return (
      'Check the whole deck against the current style profile. Fix only the formatting issues it reports (font, size, bold, off-slide, citation format); do not add, delete or rewrite content. For hints such as "too much white space", just list suggestions. ' +
      (pubmed ? 'Verify each existing footer citation with verify_citation: correct wrong ones per PubMed; mark ones that cannot be found as "[to verify]" and tell me—do not replace them with other papers. ' : '') +
      'Re-check after fixing, then list what you changed.'
    );
  }
  return (
    '按当前演示规范检查整套幻灯片，只修正检查发现的格式问题（字体、字号、加粗、超出页面、引用格式），不要新增、删除或改写内容；“留白较多”等提示只需列出建议，不要动手。' +
    (pubmed ? '已有的页脚引用逐条用 verify_citation 核对：写法有误的按 PubMed 更正，查不到的标注“[待核实]”并告诉我，不要自行替换成别的文献。' : '') +
    '修正后再检查一遍，最后列出改了什么。'
  );
}

const PPT = ['PowerPoint'];
const SEL = '当前选中的内容（没有选中时：PPT 为当前页，Word 为光标所在段落）';
const WORD = ['Word'];
const XL = ['Excel'];

// 英文界面下的名称、说明与指令（发给模型的指令也用英文，模型会用英文回复）
const SEL_EN = 'the current selection (if nothing is selected: the current slide in PowerPoint, the paragraph at the cursor in Word)';
const EN_CMDS: Record<string, { name: string; description: string; prompt: string }> = {
  outline: { name: 'Outline to slides', description: 'Paste an outline; one slide per heading', prompt: 'Create slides from the outline below, one slide per top-level heading, using suitable layouts:\n' },
  polish: { name: 'Polish', description: 'Polish the wording of the selection', prompt: `Polish the wording of ${SEL_EN}: more accurate, concise and professional; keep the meaning and the number of points; do not change formatting.` },
  shorten: { name: 'Shorten', description: 'Remove redundancy and tighten', prompt: `Shorten ${SEL_EN}: remove redundant modifiers and repetition while keeping all key information; do not change formatting.` },
  academic: { name: 'Academic polish', description: 'Polish for academic writing', prompt: `Polish ${SEL_EN} for academic writing: precise and consistent terminology, rigorous sentences, no colloquialisms; keep all data, statistics and citations; make only necessary edits.` },
  logic: { name: 'Check logic', description: 'Find contradictions and inconsistencies; add comments', prompt: 'Read the whole document and find contradictions, inconsistent numbers or sample sizes, inconsistent terminology, and conclusions not supported by results. Do not edit; add comments at the relevant places with suggestions, then summarize them.' },
  translate: { name: 'Translate', description: 'Translate the selection (Chinese ↔ English)', prompt: `Translate ${SEL_EN} into Chinese (or into English if it is Chinese), with accurate terminology, replacing the original.` },
  terms: { name: 'Bilingual terms', description: 'Add English names and abbreviations at first mention', prompt: 'Check the whole document; at the first mention of each technical term add the English full name and abbreviation, e.g. "随机对照试验（Randomized Controlled Trial, RCT）". Only add annotations; change nothing else.' },
  layout: { name: 'Check layout', description: 'Fix overflow, overlap and alignment on this slide', prompt: 'Check the layout of the current slide—text overflow, overlapping elements, alignment and white space—fix any problems, then explain what you changed.' },
  table: { name: 'Comparison table', description: 'Insert a comparison table slide after this one', prompt: 'Insert a comparison table slide after the current slide, comparing: ' },
  summary: { name: 'Summary slide', description: 'Add a summary slide at the end', prompt: 'Read the whole deck and append a "Summary" slide with 3–5 key points.' },
  describe: { name: 'Descriptive stats', description: 'Descriptive statistics with Excel formulas', prompt: 'Run descriptive statistics on the current data: for numeric variables report n, missing, mean ± SD, median (IQR), min/max; for categorical variables report counts and percentages. Put the results on a new "Descriptive statistics" sheet, all computed with formulas.' },
  compare: { name: 'Group comparison', description: 'Compare variables by a grouping variable', prompt: 'Compare the variables by the following grouping variable (new results sheet, formulas only, plus grouped box plots): ' },
  chart: { name: 'Chart', description: 'Draw a suitable chart for given variables', prompt: 'Draw suitable charts for the following variables (on a new results sheet, with clear titles and axis units): ' },
  pivot: { name: 'Pivot table', description: 'Create a pivot table for given fields', prompt: 'Create a pivot table: rows = , values = (average / count)' },
  lint: { name: 'Style check', description: 'Check and fix formatting against the style profile', prompt: '' },
  cite: { name: 'Verify citations', description: 'Verify existing citations on PubMed', prompt: 'Find all citations in the document and verify each with verify_citation: correct wrong ones per PubMed; mark ones that cannot be found as "[to verify]" and list them—do not replace them with other papers.' },
};

export function builtinCommands(opts: { pubmed: boolean; lang?: 'zh' | 'en' }): SlashCommand[] {
  const c = (o: Omit<SlashCommand, 'builtin'>): SlashCommand => {
    const en = opts.lang === 'en' ? EN_CMDS[o.id] : undefined;
    return en ? { ...o, ...en, prompt: o.id === 'lint' ? lintPrompt(opts.pubmed, 'en') : en.prompt, builtin: true } : { ...o, builtin: true };
  };
  return [
    c({ id: 'outline', name: '按大纲生成', aliases: ['dg', 'outline'], hosts: PPT, autoSend: false,
      description: '粘贴大纲，每个标题生成一页', prompt: '按下面的大纲生成幻灯片，每个一级标题一页，使用合适的版式：\n' }),
    c({ id: 'polish', name: '润色', aliases: ['rs', 'polish'], autoSend: true,
      description: `润色${SEL}的文字`, prompt: `润色${SEL}的文字：更准确、简洁、专业，保持原意和要点数量，不改格式。` }),
    c({ id: 'shorten', name: '精简', aliases: ['jj', 'short'], autoSend: true,
      description: '去掉冗余，压缩篇幅', prompt: `精简${SEL}：去掉冗余修饰和重复表达，保持专业准确和全部关键信息，不改格式。` }),
    c({ id: 'academic', name: '学术润色', aliases: ['xs', 'academic'], hosts: WORD, autoSend: true,
      description: '按学术论文规范润色选中段落', prompt: `按学术论文写作规范润色${SEL}：术语准确统一、句式严谨、去口语化，中文用全角标点，数字与单位间保留空格；保留全部数据、统计量和引用，只做必要修改。` }),
    c({ id: 'logic', name: '检查逻辑', aliases: ['lj', 'logic'], hosts: WORD, autoSend: true,
      description: '找出前后矛盾、数据不一致之处并批注', prompt: '通读全文，找出前后矛盾、数字或样本量不一致、术语不统一、结论与结果不符之处；不要直接改动，在相应位置用批注指出并给出修改建议，最后汇总列出。' }),
    c({ id: 'translate', name: '翻译', aliases: ['fy', 'translate'], autoSend: false,
      description: '中英互译选中内容', prompt: `把${SEL}翻译成英文（若原文是英文则译成中文），术语准确，直接替换原文。` }),
    c({ id: 'terms', name: '术语中英对照', aliases: ['sy', 'terms'], autoSend: true,
      description: '专业术语首次出现补英文及缩写', prompt: '检查全文，专业术语首次出现时补充英文全称及缩写，如“随机对照试验（Randomized Controlled Trial, RCT）”。只加注释，不改其他内容。' }),
    c({ id: 'layout', name: '检查排版', aliases: ['pb', 'layout'], hosts: PPT, autoSend: true,
      description: '检查当前页溢出、重叠、对齐并修正', prompt: '检查当前页的排版：文字溢出、元素重叠、对齐与留白，发现问题直接修正，最后说明改了什么。' }),
    c({ id: 'table', name: '加对比表', aliases: ['dbb', 'table'], hosts: PPT, autoSend: false,
      description: '在当前页后插入一页对比表', prompt: '在当前页后面插入一页对比表，比较：' }),
    c({ id: 'summary', name: '小结页', aliases: ['xj', 'summary'], hosts: PPT, autoSend: true,
      description: '根据全部内容生成一页小结', prompt: '通读整套幻灯片，在最后追加一页“小结”，用 3–5 条要点概括核心内容。' }),
    c({ id: 'describe', name: '描述统计', aliases: ['ms', 'describe', 'stats'], hosts: XL, autoSend: true,
      description: '对当前数据做描述统计（公式计算）', prompt: '对当前数据做描述统计：数值变量报告 n、缺失、均数 ± 标准差、中位数（四分位数）、最小 / 最大值，分类变量报告频数与百分比；结果放到新建的“描述统计”表，全部用公式计算。' }),
    c({ id: 'compare', name: '分组比较', aliases: ['fz', 'compare', 'group'], hosts: XL, autoSend: false,
      description: '按分组变量比较各指标', prompt: '按以下分组变量比较各指标（新建结果表，公式计算，并画分组箱线图）：' }),
    c({ id: 'chart', name: '画图', aliases: ['ht', 'chart', 'plot'], hosts: XL, autoSend: false,
      description: '为指定变量画合适的图', prompt: '为以下变量画合适的图（放在新建结果表，写清标题和坐标轴单位）：' }),
    c({ id: 'pivot', name: '数据透视', aliases: ['ts', 'pivot'], hosts: XL, autoSend: false,
      description: '按指定字段生成数据透视表', prompt: '生成数据透视表：行 = ，值 = （平均值 / 计数）' }),
    c({ id: 'lint', name: '规范检查', aliases: ['gf', 'lint', 'check'], hosts: PPT, autoSend: true, requires: 'profile',
      description: '按演示规范检查并修正格式', prompt: lintPrompt(opts.pubmed) }),
    c({ id: 'cite', name: '核对引用', aliases: ['hd', 'cite', 'verify'], autoSend: true, requires: 'pubmed',
      description: '在 PubMed 逐条核对已有引用', prompt: '找出文档中所有文献引用，逐条用 verify_citation 核对：写法有误的按 PubMed 更正，查不到的标注“[待核实]”并列出，不要替换成别的文献。' }),
  ];
}

// 合并内置与自定义（同名时自定义覆盖），按宿主和前置条件过滤
export function availableCommands(
  custom: SlashCommand[],
  ctx: { host: string; hasProfile: boolean; pubmed: boolean; lang?: 'zh' | 'en' },
): SlashCommand[] {
  const custNames = new Set(custom.map((c) => c.name));
  return [...builtinCommands({ pubmed: ctx.pubmed, lang: ctx.lang }).filter((b) => !custNames.has(b.name)), ...custom].filter(
    (c) =>
      (!c.hosts?.length || c.hosts.includes(ctx.host)) &&
      (c.requires !== 'profile' || ctx.hasProfile) &&
      (c.requires !== 'pubmed' || ctx.pubmed),
  );
}

// 输入框内容是否处于"正在输入指令"状态：以 / 开头且尚未输入空白
export function slashQuery(input: string): string | null {
  const m = /^\/(\S*)$/.exec(input);
  return m ? m[1] : null;
}

// 筛选：名称包含 > 别名前缀 > 描述包含；空查询返回全部
export function filterCommands(list: SlashCommand[], q: string): SlashCommand[] {
  const k = q.trim().toLowerCase();
  if (!k) return list;
  const score = (c: SlashCommand) =>
    c.name.toLowerCase().startsWith(k) ? 0 :
    c.name.toLowerCase().includes(k) ? 1 :
    c.aliases.some((a) => a.toLowerCase().startsWith(k)) ? 2 :
    c.description.toLowerCase().includes(k) ? 3 : -1;
  return list
    .map((c) => [c, score(c)] as const)
    .filter(([, s]) => s >= 0)
    .sort((a, b) => a[1] - b[1])
    .map(([c]) => c);
}
