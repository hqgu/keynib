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

export function lintPrompt(pubmed: boolean): string {
  return (
    '按当前演示规范检查整套幻灯片，只修正检查发现的格式问题（字体、字号、加粗、超出页面、引用格式），不要新增、删除或改写内容；“留白较多”等提示只需列出建议，不要动手。' +
    (pubmed ? '已有的页脚引用逐条用 verify_citation 核对：写法有误的按 PubMed 更正，查不到的标注“[待核实]”并告诉我，不要自行替换成别的文献。' : '') +
    '修正后再检查一遍，最后列出改了什么。'
  );
}

const PPT = ['PowerPoint'];
const SEL = '当前选中的内容（没有选中时：PPT 为当前页，Word 为光标所在段落）';
const WORD = ['Word'];

export function builtinCommands(opts: { pubmed: boolean }): SlashCommand[] {
  const c = (o: Omit<SlashCommand, 'builtin'>): SlashCommand => ({ ...o, builtin: true });
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
    c({ id: 'lint', name: '规范检查', aliases: ['gf', 'lint', 'check'], hosts: PPT, autoSend: true, requires: 'profile',
      description: '按演示规范检查并修正格式', prompt: lintPrompt(opts.pubmed) }),
    c({ id: 'cite', name: '核对引用', aliases: ['hd', 'cite', 'verify'], autoSend: true, requires: 'pubmed',
      description: '在 PubMed 逐条核对已有引用', prompt: '找出文档中所有文献引用，逐条用 verify_citation 核对：写法有误的按 PubMed 更正，查不到的标注“[待核实]”并列出，不要替换成别的文献。' }),
  ];
}

// 合并内置与自定义（同名时自定义覆盖），按宿主和前置条件过滤
export function availableCommands(
  custom: SlashCommand[],
  ctx: { host: string; hasProfile: boolean; pubmed: boolean },
): SlashCommand[] {
  const custNames = new Set(custom.map((c) => c.name));
  return [...builtinCommands({ pubmed: ctx.pubmed }).filter((b) => !custNames.has(b.name)), ...custom].filter(
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
