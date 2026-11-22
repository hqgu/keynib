// 测试数据：固定种子的两组平行 RCT（200 例），供真机验收与 Node 计算标准答案共用

function mulberry32(seed: number) {
  return () => {
    seed |= 0;
    seed = (seed + 0x6d2b79f5) | 0;
    let t = Math.imul(seed ^ (seed >>> 15), 1 | seed);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

export const SEED_HEADER = ['编号', '组别', '年龄', '性别', '基线收缩压', '12周收缩压'];

export function seedTrial(n = 200, seed = 20261122): (string | number)[][] {
  const rnd = mulberry32(seed);
  const normal = () => {
    const u = 1 - rnd();
    const v = rnd();
    return Math.sqrt(-2 * Math.log(u)) * Math.cos(2 * Math.PI * v);
  };
  const rows: (string | number)[][] = [];
  for (let i = 1; i <= n; i++) {
    const group = rnd() < 0.5 ? '干预' : '对照';
    const age = Math.round(55 + 10 * normal());
    const sex = rnd() < 0.48 ? '女' : '男';
    const base = Math.round(150 + 12 * normal());
    const drop = group === '干预' ? 12 : 4;
    const w12 = rnd() < 0.06 ? '' : Math.round(base - drop + 8 * normal()); // 约 6% 失访
    rows.push([`P${String(i).padStart(3, '0')}`, group, age, sex, base, w12]);
  }
  return rows;
}
