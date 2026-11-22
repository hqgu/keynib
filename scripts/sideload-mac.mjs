// Mac 侧载：把开发清单复制到 PowerPoint / Word / Excel 的 wef 目录。
// 用法：npm run sideload:mac       移除：npm run sideload:mac -- --remove
// 复制后需完全退出并重启对应 Office 应用，在"开始 → 加载项"中可见。
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';

const manifest = path.resolve('manifest/manifest.dev.xml');
const fileName = 'keynib.manifest.dev.xml';
const apps = ['Powerpoint', 'Word', 'Excel'];
const remove = process.argv.includes('--remove');

for (const app of apps) {
  const dir = path.join(os.homedir(), 'Library/Containers', `com.microsoft.${app}`, 'Data/Documents/wef');
  const target = path.join(dir, fileName);
  if (remove) {
    if (fs.existsSync(target)) fs.rmSync(target);
    console.log(`已移除 ${app}`);
    continue;
  }
  fs.mkdirSync(dir, { recursive: true });
  fs.copyFileSync(manifest, target);
  console.log(`已侧载到 ${app}: ${target}`);
}
