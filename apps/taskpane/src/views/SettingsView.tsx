// 设置页：厂商列表 → 详情（Key、Base URL、拉取模型、测试连接）；配置码；隐私与安全

import { useState } from 'react';
import {
  Badge,
  Button,
  Field,
  Input,
  Link,
  makeStyles,
  MessageBar,
  MessageBarBody,
  Radio,
  RadioGroup,
  Select,
  Spinner,
  Switch,
  Text,
  Textarea,
  tokens,
} from '@fluentui/react-components';
import { ArrowLeftRegular, ChevronRightRegular, EyeOffRegular, EyeRegular } from '@fluentui/react-icons';
import { PROVIDER_PRESETS, friendlyError, guessVision, listModels, testConnection, type TestStep } from '@keynib/providers';
import {
  clearAllKeys,
  configCodeNeedsPassphrase,
  connectionOf,
  CUSTOM_ID,
  emptyConfig,
  exportConfigCode,
  importConfigCode,
  providerMeta,
  updateProvider,
  updateSettings,
  useSettings,
} from '../state/settings';
import { clearSessions } from '../state/sessions';
import { ProfileSection } from './ProfileSection';
import { copyText } from '../components/copy';
import { m, t, type LangSetting } from '../i18n';
import { CommandSection } from './CommandSection';

const useStyles = makeStyles({
  root: { padding: '12px', display: 'flex', flexDirection: 'column', gap: '14px', overflowY: 'auto', height: '100%', boxSizing: 'border-box' },
  section: { display: 'flex', flexDirection: 'column', gap: '8px' },
  h: { fontSize: '13px', fontWeight: 600, color: tokens.colorNeutralForeground2 },
  item: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
    padding: '8px 10px',
    borderRadius: '8px',
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    backgroundColor: tokens.colorNeutralBackground1,
    cursor: 'pointer',
    textAlign: 'left',
    color: tokens.colorNeutralForeground1,
    ':hover': { backgroundColor: tokens.colorNeutralBackground1Hover },
  },
  mono: {
    width: '28px',
    height: '28px',
    borderRadius: '6px',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    fontWeight: 700,
    fontSize: '12px',
    color: '#fff',
    backgroundColor: '#1F6F68',
    flexShrink: 0,
  },
  grow: { flex: 1, minWidth: 0 },
  sub: { fontSize: '11px', color: tokens.colorNeutralForeground3 },
  row: { display: 'flex', gap: '6px', alignItems: 'center', flexWrap: 'wrap' },
  step: { fontSize: '12px', display: 'flex', gap: '6px' },
  ok: { color: tokens.colorPaletteGreenForeground1 },
  err: { color: tokens.colorPaletteRedForeground1 },
  code: { fontFamily: tokens.fontFamilyMonospace, fontSize: '11px' },
});

declare const __BUILD_ID__: string;

const ALL_IDS = [...PROVIDER_PRESETS.map((p) => p.id), CUSTOM_ID];

export function SettingsView({ onOpenDiagnostics, welcome }: { onOpenDiagnostics: () => void; welcome?: boolean }) {
  const s = useStyles();
  const settings = useSettings();
  const [detail, setDetail] = useState<string | null>(null);

  if (detail) return <ProviderDetail id={detail} onBack={() => setDetail(null)} />;

  return (
    <div className={s.root}>
      {welcome && (
        <MessageBar intent="info">
          <MessageBarBody>
            {t('欢迎使用妙笔！先选一家模型厂商并填入你自己的 API Key。已在其他 Office 程序里配置过？用下方"配置码"一键导入。')}
          </MessageBarBody>
        </MessageBar>
      )}

      <div className={s.section}>
        <span className={s.h}>{m('语言 / Language')}</span>
        <Select value={settings.language} onChange={(_, d) => updateSettings({ language: d.value as LangSetting })}>
          <option value="auto">{t('自动（跟随 Office）')}</option>
          <option value="zh">{m('中文')}</option>
          <option value="en">English</option>
        </Select>
      </div>

      <div className={s.section}>
        <span className={s.h}>{t('模型与 Key')}</span>
        {ALL_IDS.map((id) => {
          const meta = providerMeta(id);
          const cfg = settings.providers[id];
          const configured = !!cfg?.apiKey && !!cfg.model;
          return (
            <button key={id} className={s.item} onClick={() => setDetail(id)}>
              <span className={s.mono} aria-hidden>{meta.name.slice(0, 2)}</span>
              <span className={s.grow}>
                <div>{id === CUSTOM_ID && cfg?.name ? cfg.name : meta.name}</div>
                <div className={s.sub}>{configured ? cfg!.model : t(meta.vendor)}</div>
              </span>
              {settings.activeProviderId === id && configured && <Badge color="brand">{t('使用中')}</Badge>}
              {configured && settings.activeProviderId !== id && <Badge appearance="outline">{t('已配置')}</Badge>}
              <ChevronRightRegular />
            </button>
          );
        })}
      </div>

      <ProfileSection />

      <CommandSection />

      <ConfigCodeSection />

      <div className={s.section}>
        <span className={s.h}>{t('隐私与安全')}</span>
        <RadioGroup
          value={settings.keyStorage}
          onChange={(_, d) => updateSettings({ keyStorage: d.value as 'local' | 'session' })}
        >
          <Radio value="local" label={t('记住在本机（方便；无法防住本机恶意软件）')} />
          <Radio value="session" label={t('仅本次使用（关闭面板即清除，每次需重新填写）')} />
        </RadioGroup>
        <Text size={200}>{t('你的文档内容和 Key 只从本机直接发往你选择的厂商，妙笔没有服务器，也不收集任何数据。')}</Text>
        <Switch
          checked={settings.pubmed}
          onChange={(_, d) => updateSettings({ pubmed: d.checked })}
          label={t('允许联网核对文献（PubMed）：只发送检索词（作者、期刊、年份、主题词），不发送文档内容')}
        />
        <div className={s.row}>
          <ConfirmButton label={t('清除所有 Key')} onConfirm={clearAllKeys} />
          <ConfirmButton label={t('清除会话历史')} onConfirm={() => clearSessions()} />
        </div>
      </div>

      <div className={s.section}>
        <span className={s.h}>{t('高级')}</span>
        <Field label={t('单轮最多步数')}>
          <Input
            type="number"
            value={String(settings.maxSteps)}
            onChange={(_, d) => updateSettings({ maxSteps: Math.max(1, Math.min(100, Number(d.value) || 25)) })}
          />
        </Field>
        <Link onClick={onOpenDiagnostics}>{t('诊断工具（环境、直连、往返测试）')}</Link>
        <Text size={200}>{t('妙笔 Keynib v0.1.0（预览版）· MIT 开源')} · build {typeof __BUILD_ID__ === 'string' ? __BUILD_ID__ : 'dev'}</Text>
      </div>
    </div>
  );
}

function ProviderDetail({ id, onBack }: { id: string; onBack: () => void }) {
  const s = useStyles();
  const settings = useSettings();
  const meta = providerMeta(id);
  const preset = PROVIDER_PRESETS.find((p) => p.id === id);
  const cfg = settings.providers[id] ?? emptyConfig(id);
  const [showKey, setShowKey] = useState(false);
  const [loading, setLoading] = useState(false);
  const [modelErr, setModelErr] = useState('');
  const [testing, setTesting] = useState(false);
  const [steps, setSteps] = useState<TestStep[] | null>(null);
  const conn = connectionOf(id, cfg);

  async function fetchModels() {
    setLoading(true);
    setModelErr('');
    try {
      const models = await listModels(conn);
      const model = cfg.model || models[0] || '';
      updateProvider(id, { models, model, vision: model ? guessVision(model) : cfg.vision });
    } catch (e) {
      setModelErr(friendlyError(e));
    }
    setLoading(false);
  }

  async function test() {
    setTesting(true);
    setSteps(null);
    setSteps(await testConnection(conn, cfg.model));
    setTesting(false);
  }

  return (
    <div className={s.root}>
      <div className={s.row}>
        <Button appearance="subtle" icon={<ArrowLeftRegular />} onClick={onBack} aria-label={t('返回')} />
        <Text weight="semibold" size={400}>{meta.name}</Text>
        <Text size={200}>{meta.vendor}</Text>
      </div>

      {id === CUSTOM_ID && (
        <Field label={t('名称')}>
          <Input value={cfg.name ?? ''} placeholder={t('如：本地 Ollama / 公司网关')} onChange={(_, d) => updateProvider(id, { name: d.value })} />
        </Field>
      )}

      <Field
        label="API Key"
        hint={meta.keyUrl ? <Link href={meta.keyUrl} target="_blank">{t('去 {vendor} 获取 Key ↗', { vendor: meta.vendor })}</Link> : t('本地模型可随便填，如 ollama')}
      >
        <Input
          type={showKey ? 'text' : 'password'}
          value={cfg.apiKey}
          onChange={(_, d) => updateProvider(id, { apiKey: d.value })}
          contentAfter={
            <Button size="small" appearance="transparent" icon={showKey ? <EyeOffRegular /> : <EyeRegular />} onClick={() => setShowKey(!showKey)} aria-label={t('显示 Key')} />
          }
        />
      </Field>

      <Field label={t('接口地址（Base URL）')} hint={preset?.notes ? t(preset.notes) : undefined}>
        {preset && preset.baseUrls.length > 1 && (
          <Select
            value={preset.baseUrls.some((b) => b.url === cfg.baseUrl) ? cfg.baseUrl : ''}
            onChange={(_, d) => {
              const opt = preset.baseUrls.find((b) => b.url === d.value);
              if (opt) updateProvider(id, { baseUrl: opt.url, protocol: opt.protocol, models: [] });
            }}
          >
            {preset.baseUrls.map((b) => (
              <option key={b.url} value={b.url}>{b.label}：{b.url}</option>
            ))}
            <option value="">{t('自定义')}</option>
          </Select>
        )}
        <Input value={cfg.baseUrl} onChange={(_, d) => updateProvider(id, { baseUrl: d.value.trim() })} />
      </Field>

      {id === CUSTOM_ID && (
        <Field label={t('接口格式')}>
          <Select value={cfg.protocol ?? 'openai'} onChange={(_, d) => updateProvider(id, { protocol: d.value as 'openai' | 'anthropic' })}>
            <option value="openai">{t('OpenAI 兼容（/chat/completions）')}</option>
            <option value="anthropic">{t('Anthropic 兼容（/messages）')}</option>
          </Select>
        </Field>
      )}

      <Field label={t('模型')} validationMessage={modelErr || undefined}>
        <div className={s.row}>
          {cfg.models.length > 0 && (
            <Select
              value={cfg.models.includes(cfg.model) ? cfg.model : ''}
              onChange={(_, d) => d.value && updateProvider(id, { model: d.value, vision: guessVision(d.value) })}
              style={{ flex: 1, minWidth: 0 }}
            >
              <option value="">{t('（手动填写）')}</option>
              {cfg.models.map((m) => (
                <option key={m} value={m}>{m}</option>
              ))}
            </Select>
          )}
          <Button size="small" disabled={!cfg.apiKey || loading} onClick={fetchModels} icon={loading ? <Spinner size="extra-tiny" /> : undefined}>
            {t('拉取模型')}
          </Button>
        </div>
        <Input value={cfg.model} placeholder={t('模型名，如 deepseek-chat')} onChange={(_, d) => updateProvider(id, { model: d.value.trim() })} />
      </Field>

      <Switch
        checked={cfg.vision}
        onChange={(_, d) => updateProvider(id, { vision: d.checked })}
        label={t('模型支持看图（开启后可截图检查排版）')}
      />

      <div className={s.row}>
        <Button disabled={!cfg.apiKey || !cfg.model || testing} onClick={test} icon={testing ? <Spinner size="extra-tiny" /> : undefined}>
          {t('测试连接')}
        </Button>
        <Button
          appearance="primary"
          disabled={!cfg.apiKey || !cfg.model}
          onClick={() => {
            updateSettings({ activeProviderId: id });
            onBack();
          }}
        >
          {settings.activeProviderId === id ? t('使用中') : t('设为当前使用')}
        </Button>
      </div>

      {steps && (
        <div className={s.section}>
          {steps.map((st) => (
            <div key={st.name} className={s.step}>
              <span className={st.ok ? s.ok : s.err}>{st.ok ? '✓' : '✗'}</span>
              <b>{t(st.name)}</b>
              <span>{st.detail}</span>
              <span className={s.sub}>{st.ms}ms</span>
            </div>
          ))}
          {steps.length === 3 && steps[1].ok && !steps[2].ok && (
            <Text size={200} className={s.err}>{t('该模型无法调用工具，只能聊天，不能编辑文档。请换一个支持工具调用的模型。')}</Text>
          )}
        </div>
      )}
    </div>
  );
}

function ConfigCodeSection() {
  const s = useStyles();
  const [includeKeys, setIncludeKeys] = useState(true);
  const [pass, setPass] = useState('');
  const [code, setCode] = useState('');
  const [importCode, setImportCode] = useState('');
  const [importPass, setImportPass] = useState('');
  const [msg, setMsg] = useState<{ ok: boolean; text: string } | null>(null);

  async function doExport() {
    setMsg(null);
    try {
      const c = await exportConfigCode(includeKeys, pass);
      setCode(c);
      const copied = await copyText(c);
      setMsg({
        ok: true,
        text: copied
          ? t('配置码已复制。到另一个 Office 程序的设置页粘贴导入。')
          : t('配置码已生成。浏览器不允许自动复制，请点下方文本框、全选后按 ⌘C / Ctrl+C。'),
      });
    } catch (e: any) {
      setMsg({ ok: false, text: e.message });
    }
  }

  async function doImport() {
    setMsg(null);
    try {
      await importConfigCode(importCode, importPass);
      setImportCode('');
      setImportPass('');
      setMsg({ ok: true, text: t('导入成功。') });
    } catch (e: any) {
      setMsg({ ok: false, text: e.message });
    }
  }

  return (
    <div className={s.section}>
      <span className={s.h}>{t('配置码（在 PowerPoint / Word / Excel 之间同步设置）')}</span>
      <Switch checked={includeKeys} onChange={(_, d) => setIncludeKeys(d.checked)} label={t('包含 API Key（需设置口令加密）')} />
      {includeKeys && <Input type="password" placeholder={t('设置口令（导入时需要）')} value={pass} onChange={(_, d) => setPass(d.value)} />}
      <Button size="small" onClick={doExport} disabled={includeKeys && pass.length < 4}>{t('生成配置码')}</Button>
      {code && <Textarea readOnly value={code} textarea={{ className: s.code, rows: 3 }} />}
      <Textarea placeholder={t('粘贴配置码以导入')} value={importCode} onChange={(_, d) => setImportCode(d.value)} textarea={{ className: s.code, rows: 2 }} />
      {configCodeNeedsPassphrase(importCode) && (
        <Input type="password" placeholder={t('输入口令')} value={importPass} onChange={(_, d) => setImportPass(d.value)} />
      )}
      <Button size="small" onClick={doImport} disabled={!importCode.trim()}>{t('导入')}</Button>
      {msg && (
        <MessageBar intent={msg.ok ? 'success' : 'error'}>
          <MessageBarBody>{msg.text}</MessageBarBody>
        </MessageBar>
      )}
    </div>
  );
}

// Office 加载项里不能用 window.confirm，改为"再点一次确认"
function ConfirmButton({ label, onConfirm }: { label: string; onConfirm: () => unknown }) {
  const [armed, setArmed] = useState(false);
  const [done, setDone] = useState(false);
  return (
    <Button
      size="small"
      appearance={armed ? 'primary' : 'secondary'}
      onClick={async () => {
        if (!armed) {
          setArmed(true);
          setTimeout(() => setArmed(false), 4000);
          return;
        }
        await onConfirm();
        setArmed(false);
        setDone(true);
        setTimeout(() => setDone(false), 2000);
      }}
    >
      {done ? t('已清除') : armed ? t('再点一次确认：{label}', { label }) : label}
    </Button>
  );
}
