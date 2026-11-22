// 统一消息模型：各协议适配器都与这套结构互相转换

export interface ImagePart {
  mediaType: string; // 如 image/png
  data: string; // base64，不含 data: 前缀
}

export interface ToolCall {
  id: string;
  name: string;
  arguments: string; // 原始 JSON 字符串（流式拼接得到）
  // 厂商附带的额外字段，需原样回传。例：Gemini 的 extra_content.google.thought_signature（缺失会 400）
  extra?: Record<string, unknown>;
}

export type ChatMessage =
  | { role: 'user'; content: string; images?: ImagePart[] }
  | { role: 'assistant'; content: string; reasoning?: string; toolCalls?: ToolCall[] }
  | { role: 'tool'; toolCallId: string; name: string; content: string; isError?: boolean; images?: ImagePart[] };

// 工具定义：parameters 为 JSON Schema
export interface ToolSchema {
  name: string;
  description: string;
  parameters: Record<string, unknown>;
}

export interface Usage {
  inputTokens?: number;
  outputTokens?: number;
}

export interface ChatRequest {
  model: string;
  system: string;
  messages: ChatMessage[];
  tools?: ToolSchema[];
  maxTokens?: number;
  temperature?: number;
}

// 流式回调
export interface StreamHandlers {
  onText?: (delta: string) => void;
  onReasoning?: (delta: string) => void;
  onToolCallStart?: (call: { id: string; name: string }) => void;
}

export interface ChatResult {
  content: string;
  reasoning: string;
  toolCalls: ToolCall[];
  usage: Usage;
  stopReason?: string;
}

// 运行时的厂商连接配置
export interface ProviderConnection {
  protocol: 'anthropic' | 'openai' | 'gemini';
  baseUrl: string;
  apiKey: string;
}

export type ProviderErrorKind = 'auth' | 'notFound' | 'rateLimit' | 'badRequest' | 'server' | 'network' | 'timeout' | 'aborted';

export class ProviderError extends Error {
  status?: number;
  kind: ProviderErrorKind;
  constructor(message: string, status?: number, kind: ProviderErrorKind = 'server') {
    super(message);
    this.name = 'ProviderError';
    this.status = status;
    this.kind = kind;
  }
}
