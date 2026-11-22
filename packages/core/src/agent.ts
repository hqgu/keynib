// Agent 循环：流式调用模型 → 权限检查 → 写前快照 → 执行工具 → 回填结果 → 直到完成

import {
  chat,
  friendlyError,
  ProviderError,
  type ChatMessage,
  type ImagePart,
  type ProviderConnection,
  type ToolCall,
  type ToolSchema,
  type Usage,
} from '@keynib/providers';

// 权限等级：R 只读；W 可撤销写入；D 破坏性，始终需要确认
export type Permission = 'R' | 'W' | 'D';

export interface ToolOutput {
  text: string; // 回给模型的结果（JSON 或说明）
  images?: ImagePart[];
  summary?: string; // 给界面看的一句话
}

export interface ToolDef<A = any> {
  name: string;
  label: string; // 中文名，显示在时间线上
  description: string;
  parameters: Record<string, unknown>;
  permission: Permission;
  needsVision?: boolean; // 只有支持看图的模型才提供
  describe?: (args: A) => string; // 确认卡片上的影响说明
  targets?: (args: A) => number[] | 'all'; // 将被修改的对象（PPT：幻灯片序号，从 1 开始；'all' 为全部），供写前快照使用
  run: (args: A) => Promise<ToolOutput>;
}

export type RunMode = 'auto' | 'confirm';

export interface AgentHooks {
  onText?: (delta: string) => void;
  onReasoning?: (delta: string) => void;
  onAssistantMessage?: (msg: ChatMessage) => void; // 一次模型调用结束
  onToolStart?: (call: ToolCall, tool: ToolDef | undefined) => void;
  onToolEnd?: (call: ToolCall, result: ChatMessage & { role: 'tool' }, meta: { summary?: string; ms: number; denied?: boolean }) => void;
  // 返回 true 继续执行
  confirm?: (call: ToolCall, tool: ToolDef, args: unknown, description: string) => Promise<boolean>;
  // 写操作前调用（PPT 快照），抛错则不执行
  beforeWrite?: (tool: ToolDef, args: unknown) => Promise<void>;
  onUsage?: (u: Usage) => void;
}

export interface AgentOptions {
  conn: ProviderConnection;
  model: string;
  system: string;
  history: ChatMessage[]; // 已有对话（含本轮用户消息）
  tools: ToolDef[];
  vision: boolean;
  mode: RunMode;
  maxSteps?: number;
  signal?: AbortSignal;
  hooks?: AgentHooks;
}

export interface AgentResult {
  messages: ChatMessage[]; // 本轮新增的消息
  stoppedBy: 'done' | 'maxSteps' | 'aborted' | 'error' | 'errors';
  error?: string;
}

const MAX_RESULT_CHARS = 12000;

function truncate(s: string) {
  return s.length > MAX_RESULT_CHARS ? s.slice(0, MAX_RESULT_CHARS) + `\n…（已截断，共 ${s.length} 字符；可缩小范围再读）` : s;
}

export async function runAgent(o: AgentOptions): Promise<AgentResult> {
  const { hooks = {}, maxSteps = 20 } = o;
  const tools = o.tools.filter((t) => !t.needsVision || o.vision);
  const schemas: ToolSchema[] = tools.map((t) => ({ name: t.name, description: t.description, parameters: t.parameters }));
  const added: ChatMessage[] = [];
  let consecutiveErrors = 0;

  for (let step = 0; step < maxSteps; step++) {
    let res;
    try {
      res = await chat(
        o.conn,
        { model: o.model, system: o.system, messages: keepRecentImages([...o.history, ...added], 8), tools: schemas },
        { onText: hooks.onText, onReasoning: hooks.onReasoning },
        { signal: o.signal },
      );
    } catch (e) {
      const aborted = e instanceof ProviderError && e.kind === 'aborted';
      return { messages: added, stoppedBy: aborted ? 'aborted' : 'error', error: aborted ? undefined : friendlyError(e) };
    }
    hooks.onUsage?.(res.usage);
    const assistant: ChatMessage = {
      role: 'assistant',
      content: res.content,
      reasoning: res.reasoning || undefined,
      toolCalls: res.toolCalls.length ? res.toolCalls : undefined,
    };
    added.push(assistant);
    hooks.onAssistantMessage?.(assistant);
    if (!res.toolCalls.length) return { messages: added, stoppedBy: 'done' };

    let stepFailures = 0;
    for (const call of res.toolCalls) {
      if (o.signal?.aborted) {
        // 模型要求的每个工具调用都必须有结果，否则下一轮请求会被厂商拒绝
        added.push({ role: 'tool', toolCallId: call.id, name: call.name, content: '用户已停止，未执行', isError: true });
        continue;
      }
      const tool = tools.find((t) => t.name === call.name);
      hooks.onToolStart?.(call, tool);
      const started = performance.now();
      const result = await execute(call, tool, o, hooks);
      added.push(result.msg);
      hooks.onToolEnd?.(call, result.msg, { summary: result.summary, ms: Math.round(performance.now() - started), denied: result.denied });
      if (result.msg.isError && !result.denied) stepFailures++;
    }
    // 按"步"计：一步里的工具调用全部失败才算一次；并行调用中个别失败交给模型自行纠正
    consecutiveErrors = stepFailures > 0 && stepFailures === res.toolCalls.length ? consecutiveErrors + 1 : 0;
    if (o.signal?.aborted) return { messages: added, stoppedBy: 'aborted' };
    if (consecutiveErrors >= 3) return { messages: added, stoppedBy: 'errors', error: '连续 3 步操作全部失败，已暂停。可以补充说明后让我继续。' };
  }
  return { messages: added, stoppedBy: 'maxSteps' };
}

// 截图只保留最近一步（同一批并行工具调用）的，且最多 max 张；更早的替换为占位说明。
// 只保留"最近 n 张"会让模型丢掉同一批里较早的截图而重复截图。
export function keepRecentImages(messages: ChatMessage[], max: number): ChatMessage[] {
  const out = [...messages];
  let i = out.length - 1;
  let budget = max;
  // 末尾连续的 tool 消息属于最近一步
  for (; i >= 0 && out[i].role === 'tool'; i--) {
    const m = out[i] as ChatMessage & { role: 'tool' };
    if (!m.images?.length) continue;
    if (budget >= m.images.length) budget -= m.images.length;
    else out[i] = { ...m, images: undefined, content: m.content + '\n（截图过多，已省略）' };
  }
  for (; i >= 0; i--) {
    const m = out[i];
    if (m.role === 'tool' && m.images?.length) out[i] = { ...m, images: undefined, content: m.content + '\n（较早的截图已省略，需要时可重新截图）' };
  }
  return out;
}

async function execute(
  call: ToolCall,
  tool: ToolDef | undefined,
  o: AgentOptions,
  hooks: AgentHooks,
): Promise<{ msg: ChatMessage & { role: 'tool' }; summary?: string; denied?: boolean }> {
  const base = { role: 'tool' as const, toolCallId: call.id, name: call.name };
  if (!tool) return { msg: { ...base, content: `未知工具：${call.name}`, isError: true } };

  let args: unknown;
  try {
    args = call.arguments ? JSON.parse(call.arguments) : {};
  } catch {
    return { msg: { ...base, content: `参数不是合法 JSON：${call.arguments.slice(0, 200)}`, isError: true } };
  }

  const needConfirm = tool.permission === 'D' || (tool.permission === 'W' && o.mode === 'confirm');
  if (needConfirm && hooks.confirm) {
    const ok = await hooks.confirm(call, tool, args, tool.describe?.(args as any) ?? tool.label);
    if (!ok) {
      return { msg: { ...base, content: '用户拒绝了此操作。请不要重试同一操作，改为询问用户或换一种做法。', isError: true }, denied: true };
    }
  }

  try {
    if (tool.permission !== 'R') await hooks.beforeWrite?.(tool, args);
    const out = await tool.run(args);
    return { msg: { ...base, content: truncate(out.text), images: out.images }, summary: out.summary };
  } catch (e: any) {
    return { msg: { ...base, content: `执行出错：${e?.message ?? e}`, isError: true }, summary: `出错：${e?.message ?? e}` };
  }
}
