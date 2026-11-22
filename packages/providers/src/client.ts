// 对外统一入口：chat / listModels / testConnection

import { anthropicChat, anthropicHeaders } from './anthropic.ts';
import { fetchWithRetry, friendlyError, getProviderLang, type FetchOptions } from './http.ts';

const L = (zh: string, en: string) => (getProviderLang() === 'en' ? en : zh);
import { openaiChat } from './openai.ts';
import type { ChatRequest, ChatResult, ProviderConnection, StreamHandlers } from './types.ts';

export function chat(
  conn: ProviderConnection,
  req: ChatRequest,
  handlers?: StreamHandlers,
  opts?: FetchOptions,
): Promise<ChatResult> {
  // Gemini 在 M1 走 OpenAI 兼容端点
  return conn.protocol === 'anthropic' ? anthropicChat(conn, req, handlers, opts) : openaiChat(conn, req, handlers, opts);
}

export async function listModels(conn: ProviderConnection, opts: FetchOptions = {}): Promise<string[]> {
  const headers: Record<string, string> =
    conn.protocol === 'anthropic' ? anthropicHeaders(conn.apiKey) : { Authorization: `Bearer ${conn.apiKey}` };
  const res = await fetchWithRetry(conn.baseUrl.replace(/\/$/, '') + '/models', { headers }, { maxRetries: 1, ...opts });
  const body = await res.json();
  const ids: string[] = (body?.data ?? body?.models ?? [])
    .map((m: any) => String(m.id ?? m.name ?? '').replace(/^models\//, ''))
    .filter(Boolean);
  return [...new Set(ids)].sort();
}

// 按模型名猜测是否支持看图（用户可在设置里手动改）
export function guessVision(model: string): boolean {
  const m = model.toLowerCase();
  if (/deepseek|glm-4(\.\d+)?(-air|-flash)?$|minimax-m1|text|embedding|reasoner/.test(m)) return false;
  return /claude|gpt-4o|gpt-4\.1|gpt-5|o3|o4|gemini|grok-(2-vision|4)|glm-.*v|vision|vl|minimax/.test(m);
}

export interface TestStep {
  name: '模型列表' | '流式输出' | '工具调用';
  ok: boolean;
  ms: number;
  detail: string;
}

// 三步测试连接：拉取模型 → 极小流式请求 → 一次工具调用
export async function testConnection(conn: ProviderConnection, model: string, signal?: AbortSignal): Promise<TestStep[]> {
  const steps: TestStep[] = [];
  const opts: FetchOptions = { signal, firstByteTimeoutMs: 30000, maxRetries: 0 };

  let t = performance.now();
  try {
    const models = await listModels(conn, opts);
    steps.push({ name: '模型列表', ok: true, ms: ms(t), detail: L(`共 ${models.length} 个模型`, `${models.length} models`) });
  } catch (e) {
    // 部分端点不提供模型列表，不阻断后续测试
    steps.push({ name: '模型列表', ok: false, ms: ms(t), detail: friendlyError(e) });
  }

  t = performance.now();
  try {
    const r = await chat(
      conn,
      { model, system: 'You are a connectivity test.', messages: [{ role: 'user', content: 'Reply with: OK' }] },
      {},
      opts,
    );
    steps.push({ name: '流式输出', ok: true, ms: ms(t), detail: `${L('回复：', 'Reply: ')}${(r.content || r.reasoning).trim().slice(0, 40) || '(empty)'}` });
  } catch (e) {
    steps.push({ name: '流式输出', ok: false, ms: ms(t), detail: friendlyError(e) });
    return steps; // 流式都不通，工具调用不用测了
  }

  t = performance.now();
  try {
    const r = await chat(
      conn,
      {
        model,
        system: 'You are a connectivity test. Always call the provided tool.',
        messages: [{ role: 'user', content: 'Call get_time with timezone "UTC".' }],
        // 同时带一个无参数工具：部分厂商（如 Gemini）对空参数对象很挑剔，真实 PPT 工具里就有这种
        tools: [
          {
            name: 'get_time',
            description: 'Get current time in a timezone',
            parameters: { type: 'object', properties: { timezone: { type: 'string' } }, required: ['timezone'] },
          },
          { name: 'list_items', description: 'List items. Takes no arguments.', parameters: { type: 'object', properties: {} } },
        ],
      },
      {},
      opts,
    );
    const call = r.toolCalls[0];
    steps.push(
      call
        ? { name: '工具调用', ok: true, ms: ms(t), detail: `${call.name}(${call.arguments})` }
        : { name: '工具调用', ok: false, ms: ms(t), detail: L('模型没有调用工具，可能不支持工具调用', 'The model did not call the tool; it may not support tool calling') },
    );
  } catch (e) {
    steps.push({ name: '工具调用', ok: false, ms: ms(t), detail: friendlyError(e) });
  }
  return steps;
}

const ms = (t: number) => Math.round(performance.now() - t);
