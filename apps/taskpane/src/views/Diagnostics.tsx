// 诊断页（M0 探针）：环境、直连、存储、PPT 往返、撤销；可导出 JSON 报告

import { Fragment, useMemo, useRef, useState } from 'react';
import { copyText } from '../components/copy';
import {
  Badge,
  Button,
  Card,
  CardHeader,
  Input,
  makeStyles,
  MessageBar,
  MessageBarBody,
  Select,
  Spinner,
  Text,
  Textarea,
  tokens,
} from '@fluentui/react-components';
import { PROVIDER_PRESETS, probeProvider, type ProbeResult } from '@keynib/providers';
import { collectEnv, type EnvInfo } from '../probes/env';
import { runStorageProbe, type StorageProbeResult } from '../probes/storage';
import { runPptRoundtrip, type RoundtripResult } from '../probes/pptRoundtrip';
import { markerExists, removeMarker, writeMarker } from '../probes/undo';
import { runUndoSelfTest, type UndoSelfTestResult } from '../probes/undoSelfTest';
import { runPptCapabilityProbe, type StepResult } from '../probes/pptCapabilityProbe';
import { setUndoFormatting } from '../hosts/powerpoint/changeset';

const useStyles = makeStyles({
  root: {
    display: 'flex',
    flexDirection: 'column',
    gap: '12px',
    padding: '12px',
    fontFamily: `${tokens.fontFamilyBase}, "PingFang SC", "Microsoft YaHei", sans-serif`,
  },
  title: { display: 'flex', alignItems: 'baseline', gap: '8px' },
  brand: { fontSize: '18px', fontWeight: 600, color: '#1F6F68' },
  row: { display: 'flex', gap: '6px', alignItems: 'center', flexWrap: 'wrap' },
  kv: { display: 'grid', gridTemplateColumns: 'auto 1fr', columnGap: '8px', rowGap: '2px', fontSize: '12px' },
  provider: {
    display: 'grid',
    gridTemplateColumns: '1fr auto',
    gap: '4px',
    padding: '6px 0',
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    fontSize: '12px',
  },
  mono: { fontFamily: tokens.fontFamilyMonospace, fontSize: '11px', wordBreak: 'break-all' },
});

type Status = 'idle' | 'running' | 'done';

export function Diagnostics() {
  const s = useStyles();
  const env: EnvInfo | null = useMemo(() => (Office.context?.requirements ? collectEnv() : null), []);
  const host = env?.host ?? 'Unknown';

  // —— CORS 探针 ——
  const [baseUrlChoice, setBaseUrlChoice] = useState<Record<string, string>>(
    Object.fromEntries(PROVIDER_PRESETS.map((p) => [p.id, p.baseUrls[0].url])),
  );
  const [keys, setKeys] = useState<Record<string, string>>({}); // 仅存内存，不持久化
  const [cors, setCors] = useState<Record<string, ProbeResult>>({});
  const [corsStatus, setCorsStatus] = useState<Status>('idle');

  async function runCors() {
    setCorsStatus('running');
    const results = await Promise.all(
      PROVIDER_PRESETS.map((p) => probeProvider(p, baseUrlChoice[p.id], keys[p.id]?.trim() || undefined)),
    );
    setCors(Object.fromEntries(results.map((r) => [r.providerId, r])));
    setCorsStatus('done');
  }

  // —— 存储 / 往返 / 撤销探针 ——
  const [storage, setStorage] = useState<StorageProbeResult | null>(null);
  const [roundtrip, setRoundtrip] = useState<RoundtripResult | null>(null);
  const [caps, setCaps] = useState<StepResult[] | null>(null);
  const [capsRunning, setCapsRunning] = useState(false);
  const [rtStatus, setRtStatus] = useState<Status>('idle');
  const [undoMsg, setUndoMsg] = useState<string>('');
  const [undoResult, setUndoResult] = useState<'reverted' | 'kept' | null>(null);

  const reportRef = useRef<HTMLTextAreaElement>(null);
  const [copyMsg, setCopyMsg] = useState('');
  const [selfTest, setSelfTest] = useState<UndoSelfTestResult | null>(null);
  const [selfTestProgress, setSelfTestProgress] = useState<number | null>(null);

  async function runSelfTest() {
    setSelfTest(null);
    setSelfTestProgress(0);
    try {
      const r = await runUndoSelfTest(20, setSelfTestProgress);
      if (r.formatting) setUndoFormatting(r.formatting);
      setSelfTest(r);
    } catch (e: any) {
      setSelfTest({ ok: false, runs: 0, passed: 0, details: [`出错：${e?.message ?? e}`] });
    }
    setSelfTestProgress(null);
  }

  async function runRoundtrip() {
    setRtStatus('running');
    setRoundtrip(await runPptRoundtrip(true));
    setRtStatus('done');
  }

  async function undoStep(step: 'write' | 'check' | 'clean') {
    try {
      if (step === 'write') {
        setUndoResult(null);
        setUndoMsg((await writeMarker(host)) + '。现在点一下文档，按 ⌘Z / Ctrl+Z，再回来点"检查"。');
      } else if (step === 'check') {
        const exists = await markerExists(host);
        setUndoResult(exists ? 'kept' : 'reverted');
        setUndoMsg(exists ? '标记仍在：Office.js 写入未被原生撤销。' : '标记已消失：Office.js 写入可被原生撤销。');
      } else {
        await removeMarker(host);
        setUndoMsg('已清理标记。');
      }
    } catch (e: any) {
      setUndoMsg(`出错：${e?.message ?? e}`);
    }
  }

  const report = JSON.stringify(
    {
      generatedAt: new Date().toISOString(),
      env,
      cors: Object.values(cors),
      storage,
      pptCapabilities: caps,
      pptRoundtrip: roundtrip,
      undoSelfTest: selfTest,
      undo: undoResult ? { host, result: undoResult } : null,
    },
    null,
    2,
  );

  return (
    <div className={s.root}>

      {!env && (
        <MessageBar intent="warning">
          <MessageBarBody>未在 Office 中运行，只能使用"浏览器直连探针"。</MessageBarBody>
        </MessageBar>
      )}

      {env && (
        <Card>
          <CardHeader header={<Text weight="semibold">环境与要求集</Text>} />
          <div className={s.kv}>
            <span>宿主</span><b>{env.host}</b>
            <span>平台</span><span>{env.platform}</span>
            <span>版本</span><span>{env.officeVersion}</span>
            <span>语言</span><span>{env.displayLanguage}</span>
            <span>主题</span><span>{env.theme ? `${env.theme.body}（${env.theme.isDark ? '深色' : '浅色'}）` : '无法读取'}</span>
            {Object.entries(env.requirementSets).map(([k, v]) => (
              <Fragment key={k}><span>{k}</span><b>{v}</b></Fragment>
            ))}
          </div>
        </Card>
      )}

      <Card>
        <CardHeader
          header={<Text weight="semibold">浏览器直连探针（CORS）</Text>}
          description={<Text size={200}>不填 Key 只测跨域；填 Key 会额外拉取模型列表。Key 只留在内存，不保存。</Text>}
        />
        {PROVIDER_PRESETS.map((p) => {
          const r = cors[p.id];
          return (
            <div key={p.id} className={s.provider}>
              <div>
                <b>{p.name}</b> <span style={{ color: tokens.colorNeutralForeground3 }}>{p.vendor}</span>
              </div>
              <div>
                {!r ? (
                  <Badge appearance="outline">未测</Badge>
                ) : r.reachable ? (
                  <Badge color="success">直连 ✓ {r.status} · {r.ms}ms</Badge>
                ) : (
                  <Badge color="danger">不可达</Badge>
                )}
              </div>
              {p.baseUrls.length > 1 && (
                <Select
                  size="small"
                  value={baseUrlChoice[p.id]}
                  onChange={(_, d) => setBaseUrlChoice({ ...baseUrlChoice, [p.id]: d.value })}
                >
                  {p.baseUrls.map((b) => (
                    <option key={b.url} value={b.url}>{b.label}</option>
                  ))}
                </Select>
              )}
              <Input
                size="small"
                type="password"
                placeholder="API Key（可选）"
                value={keys[p.id] ?? ''}
                onChange={(_, d) => setKeys({ ...keys, [p.id]: d.value })}
                style={{ gridColumn: '1 / -1' }}
              />
              {r?.modelCount !== undefined && (
                <span className={s.mono} style={{ gridColumn: '1 / -1' }}>
                  模型 {r.modelCount} 个：{r.sampleModels?.join(', ')}
                </span>
              )}
              {r?.error && <span className={s.mono} style={{ gridColumn: '1 / -1' }}>{r.error}</span>}
            </div>
          );
        })}
        <div className={s.row}>
          <Button appearance="primary" onClick={runCors} disabled={corsStatus === 'running'}>
            {corsStatus === 'running' ? <Spinner size="tiny" /> : '全部测试'}
          </Button>
        </div>
      </Card>

      {env && (
        <Card>
          <CardHeader
            header={<Text weight="semibold">存储共享探针</Text>}
            description={<Text size={200}>在 PPT、Word、Excel 中各点一次，看能否读到其他宿主写入的标记。</Text>}
          />
          <Button onClick={async () => setStorage(await runStorageProbe(host))}>写入并读取</Button>
          {storage && (
            <div className={s.kv}>
              {(['localStorage', 'indexedDB'] as const).map((kind) => (
                <Fragment key={kind}>
                  <b>{kind}</b>
                  <span>
                    {storage[kind].error
                      ? `出错：${storage[kind].error}`
                      : Object.entries(storage[kind].seen)
                          .map(([h, v]) => `${h}:${v ? '✓' : '—'}`)
                          .join('  ')}
                  </span>
                </Fragment>
              ))}
            </div>
          )}
        </Card>
      )}

      {host === 'PowerPoint' && (
        <Card>
          <CardHeader
            header={<Text weight="semibold">PPT 能力逐项探针</Text>}
            description={<Text size={200}>逐个调用新建页、版式、表格、截图、导出、插回、移动等接口，分别记录成败；会临时新建几页，测完删除。</Text>}
          />
          <Button
            appearance="primary"
            disabled={capsRunning}
            onClick={async () => {
              setCapsRunning(true);
              setCaps(await runPptCapabilityProbe());
              setCapsRunning(false);
            }}
          >
            {capsRunning ? <Spinner size="tiny" /> : '逐项测试'}
          </Button>
          {caps && (
            <div className={s.kv}>
              {caps.map((c) => (
                <Fragment key={c.step}>
                  <span>{c.ok ? '✅' : '❌'} {c.step}</span>
                  <span className={s.mono}>{c.detail ?? ''}</span>
                </Fragment>
              ))}
            </div>
          )}
        </Card>
      )}

      {host === 'PowerPoint' && (
        <Card>
          <CardHeader
            header={<Text weight="semibold">“撤销本轮”自测（20 次）</Text>}
            description={<Text size={200}>请在空白演示文稿中运行。模拟改第 1 页、插页、移页、删页后撤销，比较前后是否完全一致。约需 1–2 分钟，期间幻灯片会闪动。</Text>}
          />
          <Button appearance="primary" onClick={runSelfTest} disabled={selfTestProgress !== null}>
            {selfTestProgress !== null ? `运行中… ${selfTestProgress}/20` : '开始自测'}
          </Button>
          {selfTest && (
            <MessageBar intent={selfTest.ok ? 'success' : 'error'}>
              <MessageBarBody>
                {selfTest.ok ? `通过 ${selfTest.passed}/${selfTest.runs}（插回方式：${selfTest.formatting}，已设为默认）` : `未通过：${selfTest.passed}/${selfTest.runs}`}
                {selfTest.details.length > 0 && <div className={s.mono}>{selfTest.details.join('\n')}</div>}
              </MessageBarBody>
            </MessageBar>
          )}
        </Card>
      )}

      {host === 'PowerPoint' && (
        <Card>
          <CardHeader
            header={<Text weight="semibold">PPT 导出 / 插入往返探针</Text>}
            description={<Text size={200}>在末尾临时新建测试页 → 导出 → 插回 → 比较 → 删除两页。</Text>}
          />
          <Button onClick={runRoundtrip} disabled={rtStatus === 'running'}>
            {rtStatus === 'running' ? <Spinner size="tiny" /> : '运行往返测试'}
          </Button>
          {roundtrip && (
            <MessageBar intent={roundtrip.ok ? 'success' : 'error'}>
              <MessageBarBody>
                {roundtrip.error
                  ? roundtrip.error
                  : roundtrip.ok
                    ? `无损 ✓（${roundtrip.original?.length} 个形状；导出 ${Math.round((roundtrip.base64Length ?? 0) / 1024)} KB；ID ${roundtrip.idChanged ? '会变化' : '不变'}）`
                    : `有差异：${roundtrip.differences.join('；')}`}
              </MessageBarBody>
            </MessageBar>
          )}
        </Card>
      )}

      {env && (
        <Card>
          <CardHeader
            header={<Text weight="semibold">⌘Z / Ctrl+Z 撤销探针</Text>}
            description={<Text size={200}>写入标记 → 回到文档按撤销 → 回来检查。</Text>}
          />
          <div className={s.row}>
            <Button onClick={() => undoStep('write')}>1. 写入标记</Button>
            <Button onClick={() => undoStep('check')}>2. 检查</Button>
            <Button appearance="subtle" onClick={() => undoStep('clean')}>清理</Button>
          </div>
          {undoMsg && <Text size={200}>{undoMsg}</Text>}
        </Card>
      )}

      <Card>
        <CardHeader header={<Text weight="semibold">探针报告（JSON）</Text>} />
        <Textarea value={report} readOnly resize="vertical" textarea={{ className: s.mono, rows: 8, ref: reportRef }} />
        <div className={s.row}>
          <Button onClick={async () => setCopyMsg((await copyText(report, reportRef.current)) ? '已复制 ✓' : '浏览器不允许自动复制：报告已全选，请按 ⌘C / Ctrl+C')}>复制报告</Button>
          {copyMsg && <Text size={200}>{copyMsg}</Text>}
        </div>
      </Card>
    </div>
  );
}
