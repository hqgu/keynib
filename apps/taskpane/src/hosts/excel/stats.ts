// Excel 数据画像与地址工具（纯逻辑，可单元测试）
// 画像只给模型提供上下文；写到结果表的统计量必须是 Excel 公式。

export type ColType = 'number' | 'date' | 'text' | 'empty' | 'mixed';

export interface ColumnProfile {
  name: string;
  column: string; // 列字母
  type: ColType;
  n: number; // 非空个数
  missing: number;
  mean?: number;
  sd?: number; // 样本标准差
  min?: number;
  q1?: number;
  median?: number;
  q3?: number;
  max?: number;
  levels?: { value: string; count: number }[]; // 文本列的频数（前 10）
  distinct?: number;
}

export function colLetter(i: number): string {
  let s = '';
  for (let n = i + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
}

// 公式中引用工作表名：含空格、中文或特殊字符时加单引号，内部单引号双写
export function quoteSheet(name: string): string {
  return /^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) ? name : `'${name.replace(/'/g, "''")}'`;
}

const isDateFormat = (f: string) => /[ymd]/i.test(f.replace(/"[^"]*"/g, '')) && !/^general$/i.test(f) && !/0\.?0*%?$/.test(f);

// 线性插值分位数（与 Excel QUARTILE.INC / MEDIAN 一致）
export function quantile(sorted: number[], p: number): number {
  if (!sorted.length) return NaN;
  const h = (sorted.length - 1) * p;
  const lo = Math.floor(h);
  return sorted[lo] + (h - lo) * ((sorted[Math.min(lo + 1, sorted.length - 1)] ?? sorted[lo]) - sorted[lo]);
}

export function describeNumbers(xs: number[]) {
  const n = xs.length;
  if (!n) return {};
  const s = [...xs].sort((a, b) => a - b);
  const mean = xs.reduce((a, b) => a + b, 0) / n;
  const sd = n > 1 ? Math.sqrt(xs.reduce((a, b) => a + (b - mean) ** 2, 0) / (n - 1)) : undefined;
  return { mean, sd, min: s[0], q1: quantile(s, 0.25), median: quantile(s, 0.5), q3: quantile(s, 0.75), max: s[n - 1] };
}

const r4 = (x: number | undefined) => (x === undefined || Number.isNaN(x) ? undefined : Math.round(x * 1e4) / 1e4);

// values / numberFormat 为数据区（不含表头）；header 为表头
export function profileColumns(header: string[], values: unknown[][], numberFormat?: string[][]): ColumnProfile[] {
  return header.map((name, j) => {
    const col = values.map((r) => r[j]);
    const fmt = numberFormat?.[0]?.[j] ?? 'General';
    const nonEmpty = col.filter((v) => v !== '' && v !== null && v !== undefined);
    const nums = nonEmpty.filter((v): v is number => typeof v === 'number');
    const missing = col.length - nonEmpty.length;
    const base = { name: String(name || colLetter(j)), column: colLetter(j), n: nonEmpty.length, missing };
    if (!nonEmpty.length) return { ...base, type: 'empty' as const };
    if (nums.length === nonEmpty.length) {
      if (isDateFormat(fmt)) {
        const d = describeNumbers(nums);
        return { ...base, type: 'date' as const, min: d.min, max: d.max };
      }
      const d = describeNumbers(nums);
      return { ...base, type: 'number' as const, mean: r4(d.mean), sd: r4(d.sd), min: d.min, q1: r4(d.q1), median: r4(d.median), q3: r4(d.q3), max: d.max };
    }
    const counts = new Map<string, number>();
    for (const v of nonEmpty) counts.set(String(v), (counts.get(String(v)) ?? 0) + 1);
    const levels = [...counts.entries()].sort((a, b) => b[1] - a[1]).slice(0, 10).map(([value, count]) => ({ value, count }));
    return { ...base, type: nums.length ? ('mixed' as const) : ('text' as const), levels, distinct: counts.size };
  });
}

// 单元格计数（A1:B3 → 6），用于快照上限
export function cellCount(address: string): number {
  const a = address.split('!').pop()!.replace(/\$/g, '');
  const m = /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/i.exec(a);
  if (!m) return Infinity;
  const col = (s: string) => [...s.toUpperCase()].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0);
  if (!m[3]) return 1;
  return (Math.abs(col(m[3]) - col(m[1])) + 1) * (Math.abs(Number(m[4]) - Number(m[2])) + 1);
}

export const EXCEL_ERRORS = ['#DIV/0!', '#NAME?', '#REF!', '#VALUE!', '#N/A', '#NUM!', '#NULL!', '#SPILL!', '#CALC!'];

export function findErrors(values: unknown[][], topLeft: { row: number; col: number }): string[] {
  const out: string[] = [];
  values.forEach((r, i) =>
    r.forEach((v, j) => {
      if (typeof v === 'string' && EXCEL_ERRORS.includes(v)) out.push(`${colLetter(topLeft.col + j)}${topLeft.row + i + 1}: ${v}`);
    }),
  );
  return out;
}
