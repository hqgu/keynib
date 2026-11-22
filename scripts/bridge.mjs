// 终端端：向正在运行的面板发送测试命令
// 用法：
//   node scripts/bridge.mjs status
//   node scripts/bridge.mjs <type> [payload-json]
//   node scripts/bridge.mjs --host Word ping              # 指定目标程序（默认 PowerPoint）
// 需要先 `npm run dev`，并在 Office 中打开面板。
import fs from 'node:fs';
import https from 'node:https';
import os from 'node:os';
import path from 'node:path';

const ca = fs.readFileSync(path.join(os.homedir(), '.office-addin-dev-certs/ca.crt'));
// --host Word / Excel 指定目标程序，默认 PowerPoint
const args = process.argv.slice(2);
let host = 'PowerPoint';
const hi = args.indexOf('--host');
if (hi >= 0) {
  host = args[hi + 1];
  args.splice(hi, 2);
}
const [type, raw] = args;
if (!type) {
  console.error('用法：node scripts/bridge.mjs <status|ping|selftest|deck|chat|undoLast|chatState|newChat|settings> [payload]');
  process.exit(1);
}

let payload;
if (raw !== undefined) {
  try {
    payload = JSON.parse(raw);
  } catch {
    payload = raw;
  }
}

function request(method, p, body) {
  return new Promise((resolve, reject) => {
    const req = https.request(
      { host: 'localhost', port: 3000, path: p, method, ca, headers: { 'content-type': 'application/json' }, timeout: 15 * 60 * 1000 },
      (res) => {
        let b = '';
        res.on('data', (c) => (b += c));
        res.on('end', () => resolve(b ? JSON.parse(b) : null));
      },
    );
    req.on('error', reject);
    if (body) req.write(JSON.stringify(body));
    req.end();
  });
}

const out = type === 'status' ? await request('GET', '/__bridge/status') : await request('POST', '/__bridge/cmd', { type, host, payload });
console.log(JSON.stringify(out, null, 2));
