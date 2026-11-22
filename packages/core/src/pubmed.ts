// PubMed 文献检索与引用核对（NCBI E-utilities，允许浏览器直连）
// 只发送检索词（作者、期刊、年份、主题词），不发送文档内容。
// NCBI 无 API Key 时限 3 次/秒，这里串行并间隔 ≥ 400ms。

import type { ToolDef } from './agent.ts';

const EUTILS = 'https://eutils.ncbi.nlm.nih.gov/entrez/eutils';

// 界面摘要语言（由界面设置同步；给模型的 note 保持中文）
let uiLang: 'zh' | 'en' = 'zh';
export function setCoreLang(l: 'zh' | 'en') {
  uiLang = l;
}
const L = (zh: string, en: string) => (uiLang === 'en' ? en : zh);

export interface PubmedRecord {
  pmid: string;
  firstAuthor: string; // 如 "Muller CJ"
  journal: string; // ISO 缩写，如 "Int J Epidemiol"
  year: string;
  volume: string;
  issue: string;
  pages: string; // 已展开，如 "962-970"
  title: string;
  doi?: string;
  shortCite: string; // 学术短引：Muller CJ. Int J Epidemiol. 2014; 43 (3): 962-970.
}

// "962-70" → "962-970"；"144-6" → "144-146"；"c332"、"e123" 原样
export function expandPages(p: string): string {
  const m = /^(\d+)[-–](\d+)$/.exec(p.trim());
  if (!m) return p.trim();
  const [, a, b] = m;
  return b.length < a.length ? `${a}-${a.slice(0, a.length - b.length)}${b}` : `${a}-${b}`;
}

export function formatShortCite(r: Omit<PubmedRecord, 'shortCite'>): string {
  const issue = r.issue ? ` (${r.issue})` : '';
  const pages = r.pages ? `: ${r.pages}` : '';
  return `${r.firstAuthor}. ${r.journal}. ${r.year}; ${r.volume}${issue}${pages}.`;
}

export interface ParsedCite {
  author: string; // 姓
  initials: string;
  journal: string;
  year: string;
  volume: string;
  issue?: string;
  firstPage: string;
}

const CITE_RE =
  /^(?<author>[A-Z][A-Za-z'’-]+(?: [A-Z][A-Za-z'’-]+)*) (?<init>[A-Z]{1,3})(?:,? et al)?\. (?<journal>.+?)\. (?<year>(?:19|20)\d{2}); ?(?<vol>\d+)(?: ?\((?<iss>[\w-]+)\))?: ?(?<pages>\w+)(?:[-–]\w+)?\.?$/;

export function parseShortCite(text: string): ParsedCite | null {
  const m = CITE_RE.exec(text.trim());
  if (!m?.groups) return null;
  const g = m.groups;
  return { author: g.author, initials: g.init, journal: g.journal, year: g.year, volume: g.vol, issue: g.iss, firstPage: g.pages };
}

const norm = (s: string) => s.replace(/[–—]/g, '-').replace(/\s+/g, '').replace(/\.$/, '').toLowerCase();

// —— 网络层（可注入 fetch 以便测试） ——

let last = 0;
async function throttle() {
  const wait = last + 400 - Date.now();
  if (wait > 0) await new Promise((r) => setTimeout(r, wait));
  last = Date.now();
}

async function getJson(url: string): Promise<any> {
  await throttle();
  const res = await fetch(url);
  if (!res.ok) throw new Error(`PubMed 返回 ${res.status}`);
  return res.json();
}

export async function esearch(term: string, max = 5): Promise<string[]> {
  const url = `${EUTILS}/esearch.fcgi?db=pubmed&retmode=json&tool=keynib&retmax=${max}&term=${encodeURIComponent(term)}`;
  const j = await getJson(url);
  return j?.esearchresult?.idlist ?? [];
}

export async function esummary(ids: string[]): Promise<PubmedRecord[]> {
  if (!ids.length) return [];
  const j = await getJson(`${EUTILS}/esummary.fcgi?db=pubmed&retmode=json&tool=keynib&id=${ids.join(',')}`);
  const out: PubmedRecord[] = [];
  for (const id of j?.result?.uids ?? []) {
    const d = j.result[id];
    const base = {
      pmid: id,
      firstAuthor: d.sortfirstauthor || d.authors?.[0]?.name || '',
      journal: d.source ?? '',
      year: String(d.pubdate ?? '').slice(0, 4),
      volume: d.volume ?? '',
      issue: d.issue ?? '',
      pages: expandPages(d.pages ?? ''),
      title: String(d.title ?? '').replace(/\.$/, ''),
      doi: (d.articleids ?? []).find((a: any) => a.idtype === 'doi')?.value,
    };
    out.push({ ...base, shortCite: formatShortCite(base) });
  }
  return out;
}

export type VerifyStatus = 'verified' | 'corrected' | 'not_found' | 'unparseable';

export interface VerifyResult {
  status: VerifyStatus;
  input: string;
  record?: PubmedRecord;
  candidates?: PubmedRecord[];
  note: string;
}

export async function verifyCitation(text: string): Promise<VerifyResult> {
  const c = parseShortCite(text);
  if (!c) return { status: 'unparseable', input: text, note: '无法解析为学术短引格式（作者 缩写. 期刊. 年; 卷 (期): 页码.）' };

  // 精确：作者 + 年 + 卷 + 起始页
  let ids = await esearch(`${c.author} ${c.initials}[au] AND ${c.year}[dp] AND ${c.volume}[vi] AND ${c.firstPage}[pg]`, 3);
  if (ids.length === 1) {
    const [rec] = await esummary(ids);
    const same = norm(rec.shortCite) === norm(text);
    return {
      status: same ? 'verified' : 'corrected',
      input: text,
      record: rec,
      note: same ? '已在 PubMed 核实' : `文献存在，但写法与 PubMed 不一致，应为：${rec.shortCite}`,
    };
  }
  // 宽松：作者 + 年 + 期刊，给出候选
  ids = await esearch(`${c.author} ${c.initials}[au] AND ${c.year}[dp] AND "${c.journal}"[ta]`, 5);
  const candidates = await esummary(ids);
  return {
    status: 'not_found',
    input: text,
    candidates,
    note: candidates.length ? '未找到完全匹配的文献；以下为同作者同年同期刊的候选，请核对' : 'PubMed 中未找到该文献，可能不存在或信息有误',
  };
}

export async function searchPubmed(query: string, max = 5): Promise<PubmedRecord[]> {
  return esummary(await esearch(query, Math.min(Math.max(max, 1), 10)));
}

// —— 工具定义（宿主无关，PPT / Word 共用） ——

export function pubmedTools(): ToolDef[] {
  return [
    {
      name: 'pubmed_search',
      label: '检索 PubMed',
      permission: 'R',
      description:
        '在 PubMed 检索文献，返回 PMID、标题和已排好的学术短引 shortCite。query 支持 PubMed 语法，如 "allocation concealment[ti] AND randomized" 或 "Schulz KF[au] AND 2002[dp]"。需要引用文献时先检索，并原样使用返回的 shortCite。',
      parameters: {
        type: 'object',
        properties: { query: { type: 'string' }, max: { type: 'integer', description: '返回条数，默认 5，最多 10' } },
        required: ['query'],
      },
      run: async ({ query, max }: { query: string; max?: number }) => {
        const recs = await searchPubmed(query, max ?? 5);
        return {
          text: JSON.stringify(recs.map(({ pmid, title, shortCite, doi }) => ({ pmid, title, shortCite, doi }))),
          summary: L(`「${query.slice(0, 30)}」${recs.length} 篇`, `"${query.slice(0, 30)}": ${recs.length} results`),
        };
      },
    },
    {
      name: 'verify_citation',
      label: '核对引用',
      permission: 'R',
      description:
        '在 PubMed 核对一条学术短引是否真实存在、写法是否准确。status：verified=真实且准确；corrected=真实但写法有误（用返回的 shortCite 替换）；not_found=查不到（改为检索到的真实文献，或标注"[待核实]"）。',
      parameters: { type: 'object', properties: { citation: { type: 'string' } }, required: ['citation'] },
      run: async ({ citation }: { citation: string }) => {
        const r = await verifyCitation(citation);
        const label = { verified: L('✓ 已核实', '✓ Verified'), corrected: L('需更正', 'Needs correction'), not_found: L('✗ 查不到', '✗ Not found'), unparseable: L('格式无法解析', 'Unparseable') }[r.status];
        return {
          text: JSON.stringify({
            status: r.status,
            note: r.note,
            pmid: r.record?.pmid,
            shortCite: r.record?.shortCite,
            title: r.record?.title,
            candidates: r.candidates?.map(({ pmid, shortCite, title }) => ({ pmid, shortCite, title })),
          }),
          summary: `${label}${r.record ? ` PMID ${r.record.pmid}` : ''}`,
        };
      },
    },
  ];
}
