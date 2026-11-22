// 对话页：消息流 + 步骤时间线 + 确认卡片 + 本轮变更 + 输入区

import { useEffect, useMemo, useRef, useState, type KeyboardEvent } from 'react';
import { Button, makeStyles, MessageBar, MessageBarBody, Select, Textarea, tokens } from '@fluentui/react-components';
import { SendFilled, StopFilled } from '@fluentui/react-icons';
import { pubmedTools, runAgent, type ToolDef } from '@keynib/core';
import type { ChatMessage } from '@keynib/providers';
import { activeConnection, activeProfile, getSettings, updateSettings, useSettings } from '../state/settings';
import { currentDoc, newSessionId, saveSession, type Session } from '../state/sessions';
import { PptChangeset } from '../hosts/powerpoint/changeset';
import { createPptTools, pptContext } from '../hosts/powerpoint/tools';
import { pptSystemPrompt } from '../hosts/powerpoint/prompt';
import { ChangeCard, ConfirmCard, Reasoning, StepList, type StepView } from '../components/Cards';
import { Markdown } from '../components/Markdown';
import { SlashMenu } from '../components/SlashMenu';
import { availableCommands, filterCommands, lintPrompt, slashQuery, type SlashCommand } from '../state/commands';
import { currentHost } from '../probes/env';

const useStyles = makeStyles({
  root: { display: 'flex', flexDirection: 'column', height: '100%', minHeight: 0 },
  scroll: { flex: 1, overflowY: 'auto', padding: '8px 12px', minHeight: 0 },
  user: {
    alignSelf: 'flex-end',
    backgroundColor: tokens.colorBrandBackground2,
    borderRadius: '10px 10px 2px 10px',
    padding: '6px 10px',
    margin: '8px 0 4px auto',
    maxWidth: '88%',
    width: 'fit-content',
    whiteSpace: 'pre-wrap',
    wordBreak: 'break-word',
  },
  assistant: { margin: '4px 0', fontSize: '14px' },
  composer: { borderTop: `1px solid ${tokens.colorNeutralStroke2}`, padding: '8px 12px', display: 'flex', flexDirection: 'column', gap: '6px', position: 'relative' },
  bar: { display: 'flex', alignItems: 'center', gap: '6px', flexWrap: 'wrap' },
  meta: { fontSize: '11px', color: tokens.colorNeutralForeground3, flex: 1, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' },
  empty: { padding: '16px 4px', display: 'flex', flexDirection: 'column', gap: '8px' },
  suggestion: {
    textAlign: 'left',
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderRadius: '8px',
    padding: '8px 10px',
    backgroundColor: tokens.colorNeutralBackground1,
    cursor: 'pointer',
    fontSize: '13px',
    color: tokens.colorNeutralForeground1,
    ':hover': { backgroundColor: tokens.colorNeutralBackground1Hover },
  },
  notice: { fontSize: '12px', color: tokens.colorNeutralForeground3, textAlign: 'center', margin: '6px 0' },
});

const PPT_SUGGESTIONS = [
  '按下面的大纲生成一套幻灯片，每个一级标题一页：\n',
  '把所有幻灯片的正文统一为微软雅黑 20 号，标题 36 号加粗',
  '检查当前这一页的排版问题并修正',
  '在当前页后面插入一页三列对比表：方案 / 优点 / 缺点',
];

interface Props {
  initialSession?: Session;
  onOpenSettings: () => void;
}

export function ChatView({ initialSession, onOpenSettings }: Props) {
  const s = useStyles();
  const settings = useSettings();
  const active = activeConnection(settings);
  const host = currentHost();
  const isPpt = host === 'PowerPoint';

  const [session, setSession] = useState<Session | null>(initialSession ?? null);
  const [messages, setMessages] = useState<ChatMessage[]>(initialSession?.messages ?? []);
  const [steps, setSteps] = useState<Record<string, StepView>>({});
  const [live, setLive] = useState<{ text: string; reasoning: string } | null>(null);
  const [running, setRunning] = useState(false);
  const [error, setError] = useState('');
  const [confirm, setConfirm] = useState<{ description: string; resolve: (ok: boolean) => void } | null>(null);
  const [changesets, setChangesets] = useState<Record<number, PptChangeset>>({});
  const [, force] = useState(0);
  const [input, setInput] = useState('');
  const [menuIndex, setMenuIndex] = useState(0);
  const [menuClosed, setMenuClosed] = useState(false); // Esc 关闭后，直到输入变化前不再弹出
  const inputRef = useRef<HTMLTextAreaElement>(null);
  const abortRef = useRef<AbortController | null>(null);
  // 下一轮要告诉模型的事件（如用户撤销了某一轮）
  const eventsRef = useRef<string[]>([]);
  const scrollRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    scrollRef.current?.scrollTo({ top: scrollRef.current.scrollHeight });
  }, [messages, live, steps, confirm]);

  // 最近一个有变更的轮次（只允许撤销它）
  const lastChangeTurn = useMemo(() => {
    // 已撤销的轮次跳过，这样撤销后还能继续撤销更早的一轮
    const idx = Object.keys(changesets).map(Number).filter((k) => !changesets[k].isEmpty && !changesets[k].undone).sort((a, b) => b - a);
    return idx[0];
  }, [changesets, messages, running]);

  const toolLabels = useMemo(
    () => Object.fromEntries([...createPptTools(() => new PptChangeset()), ...pubmedTools()].map((t) => [t.name, t.label])),
    [],
  );

  async function send(text: string): Promise<unknown> {
    if (!text.trim() || running) return { error: running ? '正在运行' : '空消息' };
    if (!active) {
      onOpenSettings();
      return { error: '未配置模型' };
    }
    setError('');
    setInput('');
    const userMsg: ChatMessage = { role: 'user', content: text.trim() };
    const history = [...messages, userMsg];
    const turnIndex = history.length - 1;
    setMessages(history);
    setRunning(true);
    const ctrl = new AbortController();
    abortRef.current = ctrl;

    const doc = await currentDoc().catch(() => ({ docId: 'unsaved', docName: '未保存的文档' }));
    const sess: Session =
      session ?? {
        id: newSessionId(),
        docId: doc.docId,
        docName: doc.docName,
        host,
        title: text.trim().replace(/\s+/g, ' ').slice(0, 40),
        providerId: active.id,
        model: active.cfg.model,
        createdAt: Date.now(),
        updatedAt: Date.now(),
        messages: [],
        usage: { inputTokens: 0, outputTokens: 0 },
      };

    const cs = new PptChangeset();
    setChangesets((c) => ({ ...c, [turnIndex]: cs }));
    let tools: ToolDef[] = [];
    let system = '你是"妙笔 Keynib"，一个 Office 助手。当前宿主暂未开放编辑工具（Word 将在 M2、Excel 在 M3 支持），请直接用文字回答。';
    if (isPpt) {
      tools = createPptTools(() => cs, () => activeProfile(getSettings()));
      if (settings.pubmed) tools = [...tools, ...pubmedTools()];
      const ctx = await pptContext().catch(() => '（无法读取文档概况）');
      system = pptSystemPrompt(ctx, active.cfg.vision, eventsRef.current, activeProfile(settings), settings.pubmed);
      eventsRef.current = [];
    }

    let added: ChatMessage[] = [];
    const usage = { ...sess.usage };
    setLive({ text: '', reasoning: '' });
    const result = await runAgent({
      conn: active.conn,
      model: active.cfg.model,
      system,
      history,
      tools,
      vision: active.cfg.vision,
      mode: settings.mode,
      maxSteps: settings.maxSteps,
      signal: ctrl.signal,
      hooks: {
        onText: (d) => setLive((l) => ({ text: (l?.text ?? '') + d, reasoning: l?.reasoning ?? '' })),
        onReasoning: (d) => setLive((l) => ({ text: l?.text ?? '', reasoning: (l?.reasoning ?? '') + d })),
        onAssistantMessage: (m) => {
          added = [...added, m];
          setMessages([...history, ...added]);
          setLive({ text: '', reasoning: '' });
        },
        onToolStart: (call, tool) =>
          setSteps((st) => ({ ...st, [call.id]: { id: call.id, label: tool?.label ?? call.name, status: 'running', args: call.arguments } })),
        onToolEnd: (call, msg, meta) => {
          added = [...added, msg];
          setMessages([...history, ...added]);
          setSteps((st) => ({
            ...st,
            [call.id]: {
              ...st[call.id],
              status: meta.denied ? 'denied' : msg.isError ? 'error' : 'ok',
              summary: meta.summary ?? (msg.isError ? msg.content : undefined),
              ms: meta.ms,
              result: msg.content,
            },
          }));
          force((n) => n + 1); // 变更卡片依赖 changeset 内部状态
        },
        confirm: (_call, _tool, _args, description) => new Promise<boolean>((resolve) => setConfirm({ description, resolve })),
        beforeWrite: async (tool, args) => {
          if (tool.targets) await cs.beforeWrite(tool.targets(args));
        },
        onUsage: (u) => {
          usage.inputTokens += u.inputTokens ?? 0;
          usage.outputTokens += u.outputTokens ?? 0;
        },
      },
    });
    setLive(null);
    setRunning(false);
    setConfirm(null);
    abortRef.current = null;
    if (result.error) setError(result.error);
    else if (result.stoppedBy === 'maxSteps') setError(`已达到单轮最多 ${settings.maxSteps} 步，回复"继续"可接着做。`);

    const finalMessages = [...history, ...result.messages];
    setMessages(finalMessages);
    const saved = { ...sess, messages: finalMessages, usage, updatedAt: Date.now(), model: active.cfg.model, providerId: active.id };
    setSession(saved);
    saveSession(saved).catch(() => {});
    return {
      stoppedBy: result.stoppedBy,
      error: result.error,
      changes: cs.isEmpty ? null : cs.summary(),
      reply: result.messages.filter((m) => m.role === 'assistant').map((m) => m.content).join('\n').trim(),
      steps: result.messages.flatMap((m) => (m.role === 'tool' ? [`${m.isError ? '✗' : '✓'} ${m.name}: ${m.content.slice(0, 160)}`] : [])),
      usage: { input: usage.inputTokens - sess.usage.inputTokens, output: usage.outputTokens - sess.usage.outputTokens },
    };
  }

  // 开发测试通道：发消息、撤销最近一轮、读取状态（每次渲染更新，拿到最新闭包）
  // 注意：确认卡片不经通道，必须由用户在面板上点击
  useEffect(() => {
    if (!import.meta.env.DEV) return;
    import('../dev/bridge').then(({ devApi }) => {
      devApi.chat = (text: string) => send(text);
      devApi.undoLast = async () => {
        if (lastChangeTurn === undefined) return '没有可撤销的轮次';
        await undoTurn(lastChangeTurn);
        return 'ok';
      };
      devApi.slashMenu = (q: string) => filterCommands(commands, q).map((c) => `/${c.name}（${c.autoSend ? '直接执行' : '填入'}）`);
      devApi.chatState = () => ({ running, error, messages: messages.length, waitingForUserConfirm: confirm?.description ?? null });
    });
  });

  async function undoTurn(turnIndex: number) {
    const cs = changesets[turnIndex];
    await cs.undo();
    const req = messages[turnIndex];
    const what = req?.role === 'user' ? `“${req.content.replace(/\s+/g, ' ').slice(0, 30)}”` : '';
    eventsRef.current.push(`用户撤销了请求${what}那一轮的全部修改，文档已恢复到该轮之前的状态。`);
    setChangesets((c) => ({ ...c }));
  }

  function stop() {
    abortRef.current?.abort();
    if (confirm) {
      confirm.resolve(false);
      setConfirm(null);
    }
  }

  // "/" 快捷指令
  const commands = useMemo(
    () => availableCommands(settings.customCommands, { host, hasProfile: !!settings.activeProfileId, pubmed: settings.pubmed }),
    [settings.customCommands, settings.activeProfileId, settings.pubmed, host],
  );
  const query = slashQuery(input);
  const menuItems = query !== null && !menuClosed ? filterCommands(commands, query) : null;

  function pickCommand(c: SlashCommand) {
    setMenuIndex(0);
    if (c.autoSend) {
      send(c.prompt);
      return;
    }
    setInput(c.prompt);
    requestAnimationFrame(() => {
      const el = inputRef.current;
      if (!el) return;
      el.focus();
      el.setSelectionRange(c.prompt.length, c.prompt.length);
    });
  }

  function onKey(e: KeyboardEvent<HTMLTextAreaElement>) {
    if (menuItems && !e.nativeEvent.isComposing) {
      if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
        e.preventDefault();
        const n = menuItems.length || 1;
        setMenuIndex((i) => (i + (e.key === 'ArrowDown' ? 1 : n - 1)) % n);
        return;
      }
      if ((e.key === 'Enter' || e.key === 'Tab') && menuItems[menuIndex]) {
        e.preventDefault();
        pickCommand(menuItems[menuIndex]);
        return;
      }
      if (e.key === 'Escape') {
        e.preventDefault();
        setMenuClosed(true);
        return;
      }
    }
    if (e.key === 'Enter' && !e.shiftKey && !e.nativeEvent.isComposing) {
      e.preventDefault();
      send(input);
    } else if (e.key === 'Escape' && running) {
      stop();
    }
  }

  // 由 tool 消息还原步骤（加载历史会话时没有实时状态）
  const resultById = useMemo(() => {
    const m: Record<string, ChatMessage & { role: 'tool' }> = {};
    for (const msg of messages) if (msg.role === 'tool') m[msg.toolCallId] = msg;
    return m;
  }, [messages]);

  const rendered = [];
  for (let i = 0; i < messages.length; i++) {
    const m = messages[i];
    if (m.role === 'user') {
      rendered.push(<div key={i} className={s.user}>{m.content}</div>);
    } else if (m.role === 'assistant') {
      const stepViews: StepView[] = (m.toolCalls ?? []).map((c) => {
        const r = resultById[c.id];
        return (
          steps[c.id] ?? {
            id: c.id,
            label: toolLabels[c.name] ?? c.name,
            status: r ? (r.isError ? (/拒绝/.test(r.content) ? 'denied' : 'error') : 'ok') : 'error',
            args: c.arguments,
            result: r?.content,
          }
        );
      });
      rendered.push(
        <div key={i} className={s.assistant}>
          {m.reasoning && <Reasoning text={m.reasoning} />}
          {m.content && <Markdown text={m.content} />}
          <StepList steps={stepViews} />
        </div>,
      );
    }
    // 一轮结束（下一条是 user 或到末尾且未运行）时显示变更卡片
    const isTurnEnd = messages[i + 1]?.role === 'user' || (i === messages.length - 1 && !running);
    if (isTurnEnd) {
      let start = i;
      while (start > 0 && messages[start].role !== 'user') start--;
      const cs = changesets[start];
      if (cs && !cs.isEmpty) {
        const sum = cs.summary();
        rendered.push(
          <ChangeCard
            key={`cs${start}`}
            {...sum}
            undone={cs.undone}
            canUndo={start === lastChangeTurn && !running}
            onUndo={() => undoTurn(start)}
          />,
        );
      }
    }
  }

  const totalTokens = (session?.usage.inputTokens ?? 0) + (session?.usage.outputTokens ?? 0);

  return (
    <div className={s.root}>
      <div className={s.scroll} ref={scrollRef}>
        {!isPpt && host !== 'Unknown' && (
          <MessageBar intent="info">
            <MessageBarBody>{host} 的编辑工具将在后续版本开放，目前只能对话。</MessageBarBody>
          </MessageBar>
        )}
        {messages.length === 0 && (
          <div className={s.empty}>
            {!active && (
              <MessageBar intent="warning">
                <MessageBarBody>
                  还没有配置模型。<Button size="small" appearance="transparent" onClick={onOpenSettings}>去设置</Button>
                </MessageBarBody>
              </MessageBar>
            )}
            {isPpt &&
              PPT_SUGGESTIONS.map((t) => (
                <button key={t} className={s.suggestion} onClick={() => setInput(t)}>
                  {t.trim()}
                </button>
              ))}
          </div>
        )}
        {rendered}
        {live && (live.text || live.reasoning) && (
          <div className={s.assistant}>
            <Reasoning text={live.reasoning} live={!live.text} />
            {live.text && <Markdown text={live.text} />}
          </div>
        )}
        {running && !live?.text && !confirm && <div className={s.notice}>妙笔正在工作…</div>}
        {confirm && (
          <ConfirmCard
            description={confirm.description}
            onAnswer={(ok) => {
              confirm.resolve(ok);
              setConfirm(null);
            }}
          />
        )}
        {error && (
          <MessageBar intent="error">
            <MessageBarBody>{error}</MessageBarBody>
          </MessageBar>
        )}
      </div>

      <div className={s.composer}>
        {menuItems && <SlashMenu items={menuItems} index={Math.min(menuIndex, Math.max(menuItems.length - 1, 0))} onHover={setMenuIndex} onPick={pickCommand} />}
        <Textarea
          value={input}
          onChange={(_, d) => {
            setInput(d.value);
            setMenuIndex(0);
            setMenuClosed(false);
          }}
          onKeyDown={onKey}
          placeholder={isPpt ? '告诉妙笔要对幻灯片做什么… 输入 / 调出快捷指令' : '输入问题… 输入 / 调出快捷指令'}
          resize="vertical"
          textarea={{ rows: 3, ref: inputRef }}
          aria-autocomplete="list"
          aria-expanded={!!menuItems}
        />
        <div className={s.bar}>
          <Select
            size="small"
            value={settings.mode}
            onChange={(_, d) => updateSettings({ mode: d.value as 'auto' | 'confirm' })}
            aria-label="执行模式"
          >
            <option value="auto">自动执行</option>
            <option value="confirm">每步确认</option>
          </Select>
          {isPpt && (
            <Select
              size="small"
              value={settings.activeProfileId ?? ''}
              onChange={(_, d) => updateSettings({ activeProfileId: d.value || null })}
              aria-label="演示规范"
              title="演示规范：妙笔会按所选规范写作和检查"
            >
              <option value="">无规范</option>
              {settings.profiles.map((p) => (
                <option key={p.id} value={p.id}>{p.name}</option>
              ))}
            </Select>
          )}
          {isPpt && settings.activeProfileId && !running && (
            <Button size="small" appearance="subtle" onClick={() => send(lintPrompt(settings.pubmed))}>
              规范检查
            </Button>
          )}
          <span className={s.meta}>{totalTokens > 0 ? `≈ ${(totalTokens / 1000).toFixed(1)}k tokens` : ''}</span>
          {running ? (
            <Button size="small" icon={<StopFilled />} onClick={stop}>停止</Button>
          ) : (
            <Button size="small" appearance="primary" icon={<SendFilled />} disabled={!input.trim()} onClick={() => send(input)}>
              发送
            </Button>
          )}
        </div>
      </div>
    </div>
  );
}
