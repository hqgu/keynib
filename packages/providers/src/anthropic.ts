// Anthropic Messages 协议适配器（流式 + 工具调用）

import { fetchWithRetry, readSSE, type FetchOptions } from './http.ts';
import type { ChatMessage, ChatRequest, ChatResult, ProviderConnection, StreamHandlers, ToolCall } from './types.ts';

export function anthropicHeaders(apiKey: string): Record<string, string> {
  return {
    'content-type': 'application/json',
    'x-api-key': apiKey,
    'anthropic-version': '2023-06-01',
    'anthropic-dangerous-direct-browser-access': 'true',
  };
}

// 统一消息 → Anthropic messages。连续的 tool 结果合并进同一条 user 消息。
export function toAnthropicMessages(messages: ChatMessage[]): any[] {
  const out: any[] = [];
  for (const m of messages) {
    if (m.role === 'user') {
      const content: any[] = (m.images ?? []).map((img) => ({
        type: 'image',
        source: { type: 'base64', media_type: img.mediaType, data: img.data },
      }));
      content.push({ type: 'text', text: m.content || ' ' });
      out.push({ role: 'user', content });
    } else if (m.role === 'assistant') {
      const content: any[] = [];
      if (m.content) content.push({ type: 'text', text: m.content });
      for (const c of m.toolCalls ?? []) {
        content.push({ type: 'tool_use', id: c.id, name: c.name, input: safeJson(c.arguments) });
      }
      if (content.length === 0) content.push({ type: 'text', text: ' ' });
      out.push({ role: 'assistant', content });
    } else {
      const block: any = {
        type: 'tool_result',
        tool_use_id: m.toolCallId,
        is_error: m.isError || undefined,
        content: [
          { type: 'text', text: m.content || ' ' },
          ...(m.images ?? []).map((img) => ({
            type: 'image',
            source: { type: 'base64', media_type: img.mediaType, data: img.data },
          })),
        ],
      };
      const last = out[out.length - 1];
      if (last?.role === 'user' && Array.isArray(last.content) && last.content[0]?.type === 'tool_result') {
        last.content.push(block);
      } else {
        out.push({ role: 'user', content: [block] });
      }
    }
  }
  return out;
}

function safeJson(s: string): unknown {
  try {
    return s ? JSON.parse(s) : {};
  } catch {
    return {};
  }
}

export async function anthropicChat(
  conn: ProviderConnection,
  req: ChatRequest,
  handlers: StreamHandlers = {},
  opts: FetchOptions = {},
): Promise<ChatResult> {
  const body = {
    model: req.model,
    system: req.system,
    max_tokens: req.maxTokens ?? 8192,
    temperature: req.temperature,
    stream: true,
    messages: toAnthropicMessages(req.messages),
    tools: req.tools?.length
      ? req.tools.map((t) => ({ name: t.name, description: t.description, input_schema: t.parameters }))
      : undefined,
  };
  const res = await fetchWithRetry(
    conn.baseUrl.replace(/\/$/, '') + '/messages',
    { method: 'POST', headers: anthropicHeaders(conn.apiKey), body: JSON.stringify(body) },
    opts,
  );
  return parseAnthropicStream(readSSE(res, opts.signal), handlers);
}

// 单独导出，便于用录制的事件做单元测试
export async function parseAnthropicStream(
  events: AsyncIterable<{ event?: string; data: string }>,
  handlers: StreamHandlers = {},
): Promise<ChatResult> {
  const result: ChatResult = { content: '', reasoning: '', toolCalls: [], usage: {} };
  const blocks: Record<number, { type: string; call?: ToolCall }> = {};
  for await (const { data } of events) {
    let ev: any;
    try {
      ev = JSON.parse(data);
    } catch {
      continue;
    }
    switch (ev.type) {
      case 'message_start':
        result.usage.inputTokens = ev.message?.usage?.input_tokens;
        break;
      case 'content_block_start': {
        const cb = ev.content_block;
        if (cb.type === 'tool_use') {
          const call: ToolCall = { id: cb.id, name: cb.name, arguments: '' };
          blocks[ev.index] = { type: 'tool_use', call };
          result.toolCalls.push(call);
          handlers.onToolCallStart?.({ id: cb.id, name: cb.name });
        } else {
          blocks[ev.index] = { type: cb.type };
        }
        break;
      }
      case 'content_block_delta': {
        const d = ev.delta;
        if (d.type === 'text_delta') {
          result.content += d.text;
          handlers.onText?.(d.text);
        } else if (d.type === 'thinking_delta') {
          result.reasoning += d.thinking;
          handlers.onReasoning?.(d.thinking);
        } else if (d.type === 'input_json_delta') {
          const b = blocks[ev.index];
          if (b?.call) b.call.arguments += d.partial_json;
        }
        break;
      }
      case 'message_delta':
        result.stopReason = ev.delta?.stop_reason ?? result.stopReason;
        if (ev.usage?.output_tokens != null) result.usage.outputTokens = ev.usage.output_tokens;
        break;
      case 'error':
        throw new Error(ev.error?.message ?? '流式响应出错');
    }
  }
  return result;
}
