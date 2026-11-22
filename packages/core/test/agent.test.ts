// Agent 循环单元测试：用假 fetch 模拟 OpenAI 兼容流式响应
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { runAgent, type ToolDef } from '../src/agent.ts';

const sse = (...objs: unknown[]) => objs.map((o) => `data: ${JSON.stringify(o)}\n\n`).join('') + 'data: [DONE]\n\n';
const toolCallTurn = (id: string, name: string, args: string) =>
  sse({ choices: [{ delta: { tool_calls: [{ index: 0, id, function: { name, arguments: args } }] }, finish_reason: 'tool_calls' }] });
const textTurn = (text: string) => sse({ choices: [{ delta: { content: text }, finish_reason: 'stop' }] });

function mockFetch(turns: string[]) {
  const bodies: any[] = [];
  globalThis.fetch = (async (_url: string, init: RequestInit) => {
    bodies.push(JSON.parse(String(init.body)));
    const t = turns.shift();
    if (!t) throw new Error('意外的额外请求');
    return new Response(t, { status: 200 });
  }) as typeof fetch;
  return bodies;
}

const conn = { protocol: 'openai' as const, baseUrl: 'https://x.test/v1', apiKey: 'k' };

function tools(log: string[]): ToolDef[] {
  return [
    { name: 'list_slides', label: '列出幻灯片', description: '', parameters: { type: 'object' }, permission: 'R', run: async () => ({ text: '[{"index":1}]' }) },
    { name: 'delete_slide', label: '删除幻灯片', description: '', parameters: { type: 'object' }, permission: 'D', describe: (a: any) => `删除第 ${a.slide} 页`, run: async () => { log.push('deleted'); return { text: 'ok' }; } },
    { name: 'set_text', label: '改文字', description: '', parameters: { type: 'object' }, permission: 'W', run: async () => { log.push('set'); return { text: 'ok' }; } },
  ];
}

test('工具调用 → 回填结果 → 最终回答', async () => {
  const bodies = mockFetch([toolCallTurn('c1', 'list_slides', '{}'), textTurn('共 1 页')]);
  const log: string[] = [];
  const r = await runAgent({ conn, model: 'm', system: 's', history: [{ role: 'user', content: '几页？' }], tools: tools(log), vision: false, mode: 'auto' });
  assert.equal(r.stoppedBy, 'done');
  assert.deepEqual(r.messages.map((m) => m.role), ['assistant', 'tool', 'assistant']);
  // 第二次请求里必须带上 tool 结果
  assert.equal(bodies[1].messages.at(-1).role, 'tool');
  assert.equal(bodies[1].messages.at(-1).content, '[{"index":1}]');
});

test('D 级操作被拒绝时不执行，并告知模型', async () => {
  mockFetch([toolCallTurn('c1', 'delete_slide', '{"slide":3}'), textTurn('好的，不删')]);
  const log: string[] = [];
  let asked = '';
  const r = await runAgent({
    conn, model: 'm', system: 's', history: [{ role: 'user', content: '删第三页' }], tools: tools(log), vision: false, mode: 'auto',
    hooks: { confirm: async (_c, _t, _a, desc) => { asked = desc; return false; } },
  });
  assert.equal(asked, '删除第 3 页');
  assert.deepEqual(log, []);
  const toolMsg = r.messages[1] as any;
  assert.equal(toolMsg.isError, true);
  assert.match(toolMsg.content, /拒绝/);
});

test('自动模式下 W 级不确认，但会先调用 beforeWrite（快照）', async () => {
  mockFetch([toolCallTurn('c1', 'set_text', '{}'), textTurn('改好了')]);
  const log: string[] = [];
  await runAgent({
    conn, model: 'm', system: 's', history: [{ role: 'user', content: '改' }], tools: tools(log), vision: false, mode: 'auto',
    hooks: { confirm: async () => { log.push('asked'); return true; }, beforeWrite: async () => { log.push('snapshot'); } },
  });
  assert.deepEqual(log, ['snapshot', 'set']);
});

test('参数非法与未知工具返回错误而不是崩溃', async () => {
  mockFetch([toolCallTurn('c1', 'set_text', '{bad'), toolCallTurn('c2', 'nope', '{}'), textTurn('完成')]);
  const r = await runAgent({ conn, model: 'm', system: 's', history: [{ role: 'user', content: 'x' }], tools: tools([]), vision: false, mode: 'auto' });
  assert.equal(r.stoppedBy, 'done');
  assert.match((r.messages[1] as any).content, /合法 JSON/);
  assert.match((r.messages[3] as any).content, /未知工具/);
});

test('401 错误给出可读提示', async () => {
  globalThis.fetch = (async () => new Response('{"error":{"message":"invalid key"}}', { status: 401 })) as typeof fetch;
  const r = await runAgent({ conn, model: 'm', system: 's', history: [{ role: 'user', content: 'x' }], tools: [], vision: false, mode: 'auto' });
  assert.equal(r.stoppedBy, 'error');
  assert.match(r.error!, /Key 无效.*401.*invalid key/);
});

test('截图只保留最近一步的（同批全部保留）', async () => {
  const { keepRecentImages } = await import('../src/agent.ts');
  const img = { mediaType: 'image/png', data: 'x' };
  const tool = (i: number) => ({ role: 'tool' as const, toolCallId: String(i), name: 's', content: `p${i}`, images: [img] });
  const msgs = [tool(1), { role: 'assistant' as const, content: '' }, tool(2), tool(3), tool(4)];
  const out = keepRecentImages(msgs, 8) as any[];
  assert.equal(out[0].images, undefined);
  assert.match(out[0].content, /已省略/);
  assert.deepEqual([out[2], out[3], out[4]].map((m) => m.images.length), [1, 1, 1]);
  const capped = keepRecentImages(msgs, 2) as any[];
  assert.equal(capped[2].images, undefined); // 超出上限时从较早的开始省略
});

test('无参数工具发给 OpenAI 兼容端点时省略 parameters', async () => {
  const bodies = mockFetch([textTurn('ok')]);
  await runAgent({ conn, model: 'm', system: 's', history: [{ role: 'user', content: 'x' }], tools: tools([]), vision: false, mode: 'auto' });
  const fns = bodies[0].tools.map((t: any) => t.function);
  assert.equal('parameters' in fns.find((f: any) => f.name === 'list_slides'), false);
});

test('一步里多个并行调用失败只算一次，不会立即暂停', async () => {
  const many = sse({ choices: [{ delta: { tool_calls: [0, 1, 2, 3].map((i) => ({ index: i, id: 'e' + i, function: { name: 'nope', arguments: '{}' } })) }, finish_reason: 'tool_calls' }] });
  mockFetch([many, textTurn('我换个做法')]);
  const r = await runAgent({ conn, model: 'm', system: 's', history: [{ role: 'user', content: 'x' }], tools: tools([]), vision: false, mode: 'auto' });
  assert.equal(r.stoppedBy, 'done');
});

test('连续 3 步全部失败才暂停', async () => {
  mockFetch([toolCallTurn('a', 'nope', '{}'), toolCallTurn('b', 'nope', '{}'), toolCallTurn('c', 'nope', '{}')]);
  const r = await runAgent({ conn, model: 'm', system: 's', history: [{ role: 'user', content: 'x' }], tools: tools([]), vision: false, mode: 'auto' });
  assert.equal(r.stoppedBy, 'errors');
});
