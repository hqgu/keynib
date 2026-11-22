// OpenAI Chat Completions 兼容协议适配器（OpenAI / xAI / DeepSeek / GLM / MiniMax / Gemini 兼容端点 / 自定义）

import { fetchWithRetry, readSSE, type FetchOptions } from './http.ts';
import type { ChatMessage, ChatRequest, ChatResult, ProviderConnection, StreamHandlers, ToolCall } from './types.ts';

function dataUrl(img: { mediaType: string; data: string }) {
  return `data:${img.mediaType};base64,${img.data}`;
}

// 统一消息 → OpenAI messages。
// tool 消息不能带图片：图片放到紧随其后的一条 user 消息里。
export function toOpenAIMessages(system: string, messages: ChatMessage[]): any[] {
  const out: any[] = [{ role: 'system', content: system }];
  let pendingImages: { mediaType: string; data: string }[] = [];
  const flushImages = () => {
    if (!pendingImages.length) return;
    out.push({
      role: 'user',
      content: [
        { type: 'text', text: '（以上工具调用返回的截图）' },
        ...pendingImages.map((img) => ({ type: 'image_url', image_url: { url: dataUrl(img) } })),
      ],
    });
    pendingImages = [];
  };
  for (const m of messages) {
    if (m.role !== 'tool') flushImages();
    if (m.role === 'user') {
      out.push(
        m.images?.length
          ? {
              role: 'user',
              content: [
                { type: 'text', text: m.content },
                ...m.images.map((img) => ({ type: 'image_url', image_url: { url: dataUrl(img) } })),
              ],
            }
          : { role: 'user', content: m.content },
      );
    } else if (m.role === 'assistant') {
      const msg: any = { role: 'assistant', content: m.content || null };
      if (m.toolCalls?.length) {
        msg.tool_calls = m.toolCalls.map((c) => ({
          id: c.id,
          type: 'function',
          function: { name: c.name, arguments: c.arguments || '{}' },
          ...(c.extra ?? {}),
        }));
      }
      out.push(msg);
    } else {
      out.push({ role: 'tool', tool_call_id: m.toolCallId, content: m.content || ' ' });
      if (m.images?.length) pendingImages.push(...m.images);
    }
  }
  flushImages();
  return out;
}

export async function openaiChat(
  conn: ProviderConnection,
  req: ChatRequest,
  handlers: StreamHandlers = {},
  opts: FetchOptions = {},
): Promise<ChatResult> {
  const body: any = {
    model: req.model,
    messages: toOpenAIMessages(req.system, req.messages),
    stream: true,
    max_tokens: req.maxTokens,
    temperature: req.temperature,
  };
  if (req.tools?.length) {
    body.tools = req.tools.map((t) => {
      // 无参数工具省略 parameters：Gemini 等兼容端点会拒绝空的 properties
      const props = (t.parameters as any)?.properties;
      const hasParams = props && Object.keys(props).length > 0;
      return { type: 'function', function: { name: t.name, description: t.description, ...(hasParams ? { parameters: t.parameters } : {}) } };
    });
  }
  const res = await fetchWithRetry(
    conn.baseUrl.replace(/\/$/, '') + '/chat/completions',
    {
      method: 'POST',
      headers: { 'content-type': 'application/json', Authorization: `Bearer ${conn.apiKey}` },
      body: JSON.stringify(body),
    },
    opts,
  );
  return parseOpenAIStream(readSSE(res, opts.signal), handlers);
}

// 部分模型（如 MiniMax、部分开源模型）把思考过程写在 <think>…</think> 里，这里拆出来
class ThinkSplitter {
  private inThink = false;
  private carry = '';
  push(chunk: string, onText: (s: string) => void, onThink: (s: string) => void) {
    let s = this.carry + chunk;
    this.carry = '';
    while (s) {
      const tag = this.inThink ? '</think>' : '<think>';
      const i = s.indexOf(tag);
      if (i >= 0) {
        (this.inThink ? onThink : onText)(s.slice(0, i));
        s = s.slice(i + tag.length);
        this.inThink = !this.inThink;
        continue;
      }
      // 末尾可能是半个标签，留到下一块再判断
      const keep = partialSuffix(s, tag);
      (this.inThink ? onThink : onText)(s.slice(0, s.length - keep));
      this.carry = s.slice(s.length - keep);
      break;
    }
  }
}

function partialSuffix(s: string, tag: string): number {
  for (let k = Math.min(tag.length - 1, s.length); k > 0; k--) {
    if (tag.startsWith(s.slice(s.length - k))) return k;
  }
  return 0;
}

export async function parseOpenAIStream(
  events: AsyncIterable<{ event?: string; data: string }>,
  handlers: StreamHandlers = {},
): Promise<ChatResult> {
  const result: ChatResult = { content: '', reasoning: '', toolCalls: [], usage: {} };
  // 以 index 归并分片；个别兼容端点不给 index，则按 id 归并
  const byKey: Record<string, ToolCall> = {};
  let lastKey = '0';
  const splitter = new ThinkSplitter();
  const emitText = (s: string) => {
    if (!s) return;
    result.content += s;
    handlers.onText?.(s);
  };
  const emitThink = (s: string) => {
    if (!s) return;
    result.reasoning += s;
    handlers.onReasoning?.(s);
  };

  for await (const { data } of events) {
    if (data === '[DONE]') break;
    let chunk: any;
    try {
      chunk = JSON.parse(data);
    } catch {
      continue;
    }
    if (chunk.error) throw new Error(chunk.error.message ?? '流式响应出错');
    if (chunk.usage) {
      result.usage.inputTokens = chunk.usage.prompt_tokens ?? result.usage.inputTokens;
      result.usage.outputTokens = chunk.usage.completion_tokens ?? result.usage.outputTokens;
    }
    const choice = chunk.choices?.[0];
    if (!choice) continue;
    const d = choice.delta ?? choice.message ?? {};
    // DeepSeek / GLM 等用 reasoning_content，部分网关用 reasoning
    const r = d.reasoning_content ?? d.reasoning;
    if (typeof r === 'string') emitThink(r);
    if (typeof d.content === 'string') splitter.push(d.content, emitText, emitThink);
    for (const tc of d.tool_calls ?? []) {
      const key = tc.index != null ? String(tc.index) : tc.id ? `id:${tc.id}` : lastKey;
      lastKey = key;
      let call = byKey[key];
      if (!call) {
        call = byKey[key] = { id: tc.id ?? `call_${key}_${Date.now()}`, name: '', arguments: '' };
        result.toolCalls.push(call);
      }
      if (tc.id) call.id = tc.id;
      if (tc.function?.name) {
        const first = !call.name;
        call.name += tc.function.name;
        if (first) handlers.onToolCallStart?.({ id: call.id, name: call.name });
      }
      if (tc.function?.arguments) call.arguments += tc.function.arguments;
      // 原样保留厂商扩展字段（如 Gemini 的 extra_content），回传时放回 tool_call 上
      if (tc.extra_content) call.extra = { ...(call.extra ?? {}), extra_content: tc.extra_content };
    }
    if (choice.finish_reason) result.stopReason = choice.finish_reason;
  }
  return result;
}
