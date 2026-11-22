// 对话流中的卡片：工具步骤、确认、本轮变更、思考过程

import { useState } from 'react';
import { Button, makeStyles, Spinner, tokens } from '@fluentui/react-components';
import {
  ArrowUndoRegular,
  CheckmarkCircleFilled,
  ChevronDownRegular,
  ChevronRightRegular,
  DismissCircleFilled,
  ErrorCircleFilled,
  WarningFilled,
} from '@fluentui/react-icons';

const useStyles = makeStyles({
  steps: {
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderRadius: '8px',
    padding: '4px 8px',
    margin: '4px 0',
    backgroundColor: tokens.colorNeutralBackground2,
  },
  step: { fontSize: '12px', padding: '3px 0' },
  stepHead: { display: 'flex', alignItems: 'center', gap: '6px', cursor: 'pointer', userSelect: 'none' },
  stepLabel: { flexShrink: 0, fontWeight: 600 },
  stepSummary: { color: tokens.colorNeutralForeground3, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', flex: 1 },
  stepMs: { color: tokens.colorNeutralForeground4, flexShrink: 0 },
  detail: {
    fontFamily: tokens.fontFamilyMonospace,
    fontSize: '11px',
    whiteSpace: 'pre-wrap',
    wordBreak: 'break-all',
    maxHeight: '160px',
    overflowY: 'auto',
    backgroundColor: tokens.colorNeutralBackground1,
    padding: '4px 6px',
    borderRadius: '4px',
    margin: '4px 0 2px 20px',
  },
  ok: { color: tokens.colorPaletteGreenForeground1 },
  err: { color: tokens.colorPaletteRedForeground1 },
  warnCard: {
    border: `1px solid ${tokens.colorPaletteMarigoldBorder2}`,
    backgroundColor: tokens.colorPaletteMarigoldBackground1,
    borderRadius: '8px',
    padding: '8px 10px',
    margin: '6px 0',
    fontSize: '13px',
  },
  row: { display: 'flex', gap: '6px', marginTop: '6px', flexWrap: 'wrap', alignItems: 'center' },
  changeCard: {
    border: `1px solid ${tokens.colorBrandStroke2}`,
    borderRadius: '8px',
    padding: '8px 10px',
    margin: '6px 0',
    fontSize: '12px',
    backgroundColor: tokens.colorBrandBackground2,
  },
  reasoning: { fontSize: '12px', color: tokens.colorNeutralForeground3, margin: '2px 0' },
  reasoningBody: { whiteSpace: 'pre-wrap', borderLeft: `2px solid ${tokens.colorNeutralStroke2}`, paddingLeft: '8px', marginTop: '4px', maxHeight: '200px', overflowY: 'auto' },
});

export type StepStatus = 'running' | 'ok' | 'error' | 'denied';

export interface StepView {
  id: string;
  label: string;
  status: StepStatus;
  summary?: string;
  ms?: number;
  args?: string;
  result?: string;
}

export function StepList({ steps }: { steps: StepView[] }) {
  const s = useStyles();
  if (!steps.length) return null;
  return (
    <div className={s.steps}>
      {steps.map((st) => (
        <Step key={st.id} st={st} />
      ))}
    </div>
  );
}

function Step({ st }: { st: StepView }) {
  const s = useStyles();
  const [open, setOpen] = useState(false);
  const icon =
    st.status === 'running' ? <Spinner size="extra-tiny" /> :
    st.status === 'ok' ? <CheckmarkCircleFilled className={s.ok} /> :
    st.status === 'denied' ? <DismissCircleFilled className={s.err} /> :
    <ErrorCircleFilled className={s.err} />;
  return (
    <div className={s.step}>
      <div className={s.stepHead} onClick={() => setOpen(!open)} role="button" aria-expanded={open}>
        {icon}
        <span className={s.stepLabel}>{st.label}</span>
        <span className={s.stepSummary}>{st.status === 'denied' ? '已拒绝' : st.summary}</span>
        {st.ms !== undefined && <span className={s.stepMs}>{(st.ms / 1000).toFixed(1)}s</span>}
        {open ? <ChevronDownRegular /> : <ChevronRightRegular />}
      </div>
      {open && (
        <div className={s.detail}>
          {st.args && `参数：${st.args}\n`}
          {st.result && `结果：${st.result.slice(0, 2000)}`}
        </div>
      )}
    </div>
  );
}

// 破坏性操作的确认：两个按钮同等权重，"跳过"在前并默认聚焦；执行按钮写明动作、红字
// （实测中用户两次在看清内容前点了醒目的主按钮"确认执行"）
export function confirmVerb(description: string): string {
  const m = /^(删除|覆盖|清空|替换)/.exec(description);
  return m ? `确认${m[1]}` : '确认执行';
}

export function ConfirmCard({ description, onAnswer }: { description: string; onAnswer: (ok: boolean) => void }) {
  const s = useStyles();
  return (
    <div className={s.warnCard} role="alertdialog" aria-label="需要你确认">
      <div><WarningFilled /> <b>需要你确认：</b>{description}</div>
      <div className={s.row}>
        <Button size="small" autoFocus onClick={() => onAnswer(false)}>跳过，不执行</Button>
        <Button size="small" className={s.err} onClick={() => onAnswer(true)}>{confirmVerb(description)}</Button>
      </div>
    </div>
  );
}

export function ChangeCard({
  modified,
  created,
  modifiedIndexes,
  label,
  undone,
  canUndo,
  onUndo,
}: {
  modified?: number;
  created?: number;
  modifiedIndexes?: number[];
  label?: string; // 非 PPT 时直接给出描述
  undone: boolean;
  canUndo: boolean;
  onUndo: () => Promise<void>;
}) {
  const s = useStyles();
  const [busy, setBusy] = useState(false);
  const [err, setErr] = useState('');
  return (
    <div className={s.changeCard}>
      <b>本轮变更</b>　
      {label ?? (
        <>
          新增 {created} 页 · 修改 {modified} 页
          {modifiedIndexes && modifiedIndexes.length > 0 && <span>（第 {modifiedIndexes.join('、')} 页）</span>}
        </>
      )}
      <div className={s.row}>
        {undone ? (
          <span>✓ 已撤销</span>
        ) : (
          <Button
            size="small"
            icon={busy ? <Spinner size="extra-tiny" /> : <ArrowUndoRegular />}
            disabled={!canUndo || busy}
            onClick={async () => {
              setBusy(true);
              setErr('');
              try {
                await onUndo();
              } catch (e: any) {
                setErr(e?.message ?? String(e));
              }
              setBusy(false);
            }}
          >
            撤销本轮
          </Button>
        )}
        {!canUndo && !undone && <span>只能撤销最近一轮</span>}
        {err && <span className={s.err}>撤销失败：{err}</span>}
      </div>
      {!undone && canUndo && <div style={{ marginTop: 4, color: tokens.colorNeutralForeground3 }}>撤销会恢复到本轮开始前（新建的结果表会被删除），期间你手动做的修改也会被还原。关闭或刷新面板后无法再点此撤销，届时可按 ⌘Z / Ctrl+Z。</div>}
    </div>
  );
}

export function Reasoning({ text, live }: { text: string; live?: boolean }) {
  const s = useStyles();
  const [open, setOpen] = useState(false);
  if (!text.trim()) return null;
  return (
    <div className={s.reasoning}>
      <span role="button" onClick={() => setOpen(!open)} style={{ cursor: 'pointer' }}>
        {open ? <ChevronDownRegular /> : <ChevronRightRegular />} {live ? '思考中…' : '思考过程'}
      </span>
      {open && <div className={s.reasoningBody}>{text}</div>}
    </div>
  );
}

// Word：本轮修订卡片（修订已在文档中，用户也可在 Word "审阅"里逐条处理）
export function RevisionCard({
  count,
  state,
  canSettle,
  onAccept,
  onReject,
}: {
  count: number;
  state: 'open' | 'pending' | 'accepted' | 'rejected';
  canSettle: boolean;
  onAccept: () => Promise<void>;
  onReject: () => Promise<void>;
}) {
  const s = useStyles();
  const [busy, setBusy] = useState<'' | 'accept' | 'reject'>('');
  const [err, setErr] = useState('');
  const run = async (kind: 'accept' | 'reject', fn: () => Promise<void>) => {
    setBusy(kind);
    setErr('');
    try {
      await fn();
    } catch (e: any) {
      setErr(e?.message ?? String(e));
    }
    setBusy('');
  };
  return (
    <div className={s.changeCard}>
      <b>本轮修订</b>　{count} 处
      <div className={s.row}>
        {state === 'accepted' ? (
          <span>✓ 已全部接受</span>
        ) : state === 'rejected' ? (
          <span>✓ 已全部拒绝，文档已恢复</span>
        ) : (
          <>
            <Button size="small" appearance="primary" disabled={!canSettle || !!busy} icon={busy === 'accept' ? <Spinner size="extra-tiny" /> : undefined} onClick={() => run('accept', onAccept)}>
              全部接受
            </Button>
            <Button size="small" disabled={!canSettle || !!busy} icon={busy === 'reject' ? <Spinner size="extra-tiny" /> : <ArrowUndoRegular />} onClick={() => run('reject', onReject)}>
              全部拒绝
            </Button>
          </>
        )}
        {err && <span className={s.err}>操作失败：{err}</span>}
      </div>
      {(state === 'pending' || state === 'open') && (
        <div style={{ marginTop: 4, color: tokens.colorNeutralForeground3 }}>
          只处理本轮的修订，你原有的修订不受影响。也可以在 Word「审阅」选项卡中逐条接受或拒绝；修订保存在文档里，关闭面板也不会丢。
        </div>
      )}
    </div>
  );
}
