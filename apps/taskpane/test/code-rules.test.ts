// 代码规则检查（防回归）
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';

const SRC = path.resolve(import.meta.dirname, '../src');

function files(dir: string): string[] {
  return fs.readdirSync(dir, { withFileTypes: true }).flatMap((d) => {
    const p = path.join(dir, d.name);
    return d.isDirectory() ? files(p) : /\.tsx?$/.test(d.name) ? [p] : [];
  });
}

test('useEffect / useLayoutEffect 不得用表达式体（返回值会被当作清理函数）', () => {
  // 实测：新版 Chrome 中 scrollIntoView() 返回 Promise，表达式体 effect 导致面板崩溃
  const bad = files(SRC).flatMap((f) =>
    fs
      .readFileSync(f, 'utf8')
      .split('\n')
      .map((line, i) => (/use(Layout)?Effect\(\s*\(\)\s*=>\s*[^{\s]/.test(line) ? `${path.relative(SRC, f)}:${i + 1}` : ''))
      .filter(Boolean),
  );
  assert.deepEqual(bad, []);
});
