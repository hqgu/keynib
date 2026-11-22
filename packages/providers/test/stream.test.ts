// 流式解析与消息转换的单元测试（node --test，Node ≥ 22 原生支持 TS 类型擦除）
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readSSE } from '../src/http.ts';
import { parseAnthropicStream, toAnthropicMessages } from '../src/anthropic.ts';
import { parseOpenAIStream, toOpenAIMessages } from '../src/openai.ts';

// 把字符串分片模拟成网络响应
function fakeResponse(chunks: string[]): Response {
  const enc = new TextEncoder();
  const stream = new ReadableStream({
    start(c) {
      for (const ch of chunks) c.enqueue(enc.encode(ch));
      c.close();
    },
  });
  return new Response(stream);
}

test('SSE 解析：跨分片、CRLF、多行 data', async () => {
  const res = fakeResponse(['event: a\r\ndata: {"x":', '1}\r\n\r\ndata: line1\ndata: line2\n\n', 'data: [DONE]\n\n']);
  const out = [];
  for await (const e of readSSE(res)) out.push(e);
  assert.deepEqual(out, [
    { event: 'a', data: '{"x":1}' },
    { event: undefined, data: 'line1\nline2' },
    { event: undefined, data: '[DONE]' },
  ]);
});

test('OpenAI 流：文本、思考、分片工具调用、usage', async () => {
  const ev = (o: unknown) => `data: ${JSON.stringify(o)}\n\n`;
  const res = fakeResponse([
    ev({ choices: [{ delta: { reasoning_content: '想一下' } }] }),
    ev({ choices: [{ delta: { content: '好的，' } }] }),
    ev({ choices: [{ delta: { tool_calls: [{ index: 0, id: 'c1', function: { name: 'read_slide', arguments: '{"sli' } }] } }] }),
    ev({ choices: [{ delta: { tool_calls: [{ index: 0, function: { arguments: 'de":2}' } }] } }] }),
    ev({ choices: [{ delta: { tool_calls: [{ index: 1, id: 'c2', function: { name: 'list_slides', arguments: '{}' } }] } }] }),
    ev({ choices: [{ delta: {}, finish_reason: 'tool_calls' }], usage: { prompt_tokens: 10, completion_tokens: 5 } }),
    'data: [DONE]\n\n',
  ]);
  let names: string[] = [];
  const r = await parseOpenAIStream(readSSE(res), { onToolCallStart: (c) => names.push(c.name) });
  assert.equal(r.content, '好的，');
  assert.equal(r.reasoning, '想一下');
  assert.deepEqual(r.toolCalls.map((c) => [c.id, c.name, c.arguments]), [
    ['c1', 'read_slide', '{"slide":2}'],
    ['c2', 'list_slides', '{}'],
  ]);
  assert.deepEqual(names, ['read_slide', 'list_slides']);
  assert.deepEqual(r.usage, { inputTokens: 10, outputTokens: 5 });
  assert.equal(r.stopReason, 'tool_calls');
});

test('OpenAI 流：<think> 标签跨分片拆分', async () => {
  const ev = (s: string) => `data: ${JSON.stringify({ choices: [{ delta: { content: s } }] })}\n\n`;
  const res = fakeResponse([ev('<thi'), ev('nk>推理'), ev('中</th'), ev('ink>答案')]);
  const r = await parseOpenAIStream(readSSE(res));
  assert.equal(r.reasoning, '推理中');
  assert.equal(r.content, '答案');
});

test('Anthropic 流：文本 + 工具调用 + usage', async () => {
  const ev = (o: any) => `event: ${o.type}\ndata: ${JSON.stringify(o)}\n\n`;
  const res = fakeResponse([
    ev({ type: 'message_start', message: { usage: { input_tokens: 20 } } }),
    ev({ type: 'content_block_start', index: 0, content_block: { type: 'text', text: '' } }),
    ev({ type: 'content_block_delta', index: 0, delta: { type: 'text_delta', text: '我来看看' } }),
    ev({ type: 'content_block_start', index: 1, content_block: { type: 'tool_use', id: 'tu1', name: 'list_slides', input: {} } }),
    ev({ type: 'content_block_delta', index: 1, delta: { type: 'input_json_delta', partial_json: '{"a"' } }),
    ev({ type: 'content_block_delta', index: 1, delta: { type: 'input_json_delta', partial_json: ':1}' } }),
    ev({ type: 'message_delta', delta: { stop_reason: 'tool_use' }, usage: { output_tokens: 7 } }),
    ev({ type: 'message_stop' }),
  ]);
  const r = await parseAnthropicStream(readSSE(res));
  assert.equal(r.content, '我来看看');
  assert.deepEqual(r.toolCalls, [{ id: 'tu1', name: 'list_slides', arguments: '{"a":1}' }]);
  assert.deepEqual(r.usage, { inputTokens: 20, outputTokens: 7 });
  assert.equal(r.stopReason, 'tool_use');
});

test('消息转换：连续工具结果合并、图片放位', () => {
  const msgs = [
    { role: 'user' as const, content: '做 PPT' },
    {
      role: 'assistant' as const,
      content: '',
      toolCalls: [
        { id: 'a', name: 'x', arguments: '{}' },
        { id: 'b', name: 'y', arguments: '{"k":1}' },
      ],
    },
    { role: 'tool' as const, toolCallId: 'a', name: 'x', content: 'r1' },
    { role: 'tool' as const, toolCallId: 'b', name: 'y', content: 'r2', images: [{ mediaType: 'image/png', data: 'AAA' }] },
  ];
  const an = toAnthropicMessages(msgs);
  assert.equal(an.length, 3);
  assert.equal(an[2].content.length, 2); // 两个 tool_result 合并在一条 user 消息里
  assert.equal(an[2].content[1].content[1].type, 'image');

  const oa = toOpenAIMessages('sys', msgs);
  assert.deepEqual(oa.map((m: any) => m.role), ['system', 'user', 'assistant', 'tool', 'tool', 'user']);
  assert.equal(oa[2].tool_calls[1].function.arguments, '{"k":1}');
  assert.equal(oa[5].content[1].type, 'image_url');
});

test('Gemini thought_signature：解析时保留，回传时放回 tool_call', async () => {
  const sig = { google: { thought_signature: 'SIG123' } };
  const res = fakeResponse([
    `data: ${JSON.stringify({ choices: [{ delta: { tool_calls: [{ index: 0, id: 'g1', type: 'function', function: { name: 'list_slides', arguments: '{}' }, extra_content: sig }] }, finish_reason: 'tool_calls' }] })}\n\n`,
    'data: [DONE]\n\n',
  ]);
  const r = await parseOpenAIStream(readSSE(res));
  assert.deepEqual(r.toolCalls[0].extra, { extra_content: sig });
  const out = toOpenAIMessages('s', [{ role: 'assistant', content: '', toolCalls: r.toolCalls }]);
  assert.deepEqual(out[1].tool_calls[0].extra_content, sig);
});
