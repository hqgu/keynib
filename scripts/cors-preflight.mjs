// 服务器侧 CORS 预检：模拟浏览器 OPTIONS 预检与带错误 Key 的实际请求，
// 检查 Access-Control-Allow-Origin 是否放行加载项来源。
// 注意：这只验证服务器响应头，Office WebView 内的最终结论以面板内探针为准。
// 用法：npm run probe:cors [-- https://your-origin]
const origin = process.argv[2] ?? 'https://localhost:3000';

const targets = [
  ['anthropic', 'https://api.anthropic.com/v1/models', { 'x-api-key': 'bad', 'anthropic-version': '2023-06-01', 'anthropic-dangerous-direct-browser-access': 'true' }],
  ['openai', 'https://api.openai.com/v1/models', { Authorization: 'Bearer bad' }],
  ['gemini(openai)', 'https://generativelanguage.googleapis.com/v1beta/openai/models', { Authorization: 'Bearer bad' }],
  ['gemini(native)', 'https://generativelanguage.googleapis.com/v1beta/models', { 'x-goog-api-key': 'bad' }],
  ['xai', 'https://api.x.ai/v1/models', { Authorization: 'Bearer bad' }],
  ['deepseek', 'https://api.deepseek.com/models', { Authorization: 'Bearer bad' }],
  ['glm(cn)', 'https://open.bigmodel.cn/api/paas/v4/models', { Authorization: 'Bearer bad' }],
  ['glm(intl)', 'https://api.z.ai/api/paas/v4/models', { Authorization: 'Bearer bad' }],
  ['minimax(cn)', 'https://api.minimaxi.com/v1/models', { Authorization: 'Bearer bad' }],
  ['minimax(intl)', 'https://api.minimax.io/v1/models', { Authorization: 'Bearer bad' }],
];

const allowed = (acao) => acao === '*' || acao === origin;

for (const [name, url, headers] of targets) {
  try {
    const pre = await fetch(url, {
      method: 'OPTIONS',
      headers: {
        Origin: origin,
        'Access-Control-Request-Method': 'POST',
        'Access-Control-Request-Headers': [...Object.keys(headers), 'content-type'].join(',').toLowerCase(),
      },
      signal: AbortSignal.timeout(15000),
    });
    const get = await fetch(url, { headers: { Origin: origin, ...headers }, signal: AbortSignal.timeout(15000) });
    const ok = allowed(pre.headers.get('access-control-allow-origin')) && allowed(get.headers.get('access-control-allow-origin'));
    console.log(`${ok ? '✓' : '✗'} ${name.padEnd(15)} 预检 ${pre.status} / 请求 ${get.status}  ACAO=${get.headers.get('access-control-allow-origin')}`);
  } catch (e) {
    console.log(`✗ ${name.padEnd(15)} 不可达：${e.cause?.code ?? e.message}`);
  }
}
