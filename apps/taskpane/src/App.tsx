// 应用外壳：顶栏（品牌 · 模型胶囊 · 新对话 · 历史 · 设置）+ 页面切换

import { useEffect, useState } from 'react';
import { Button, makeStyles, tokens, Tooltip } from '@fluentui/react-components';
import { AddRegular, ArrowLeftRegular, HistoryRegular, SettingsRegular } from '@fluentui/react-icons';
import { t } from './i18n';
import { activeConnection, CUSTOM_ID, providerMeta, useSettings } from './state/settings';
import type { Session } from './state/sessions';
import { ChatView } from './views/ChatView';
import { SettingsView } from './views/SettingsView';
import { HistoryView } from './views/HistoryView';
import { Diagnostics } from './views/Diagnostics';

const useStyles = makeStyles({
  root: {
    display: 'flex',
    flexDirection: 'column',
    height: '100vh',
    fontFamily: `${tokens.fontFamilyBase}, "PingFang SC", "Microsoft YaHei", sans-serif`,
    backgroundColor: tokens.colorNeutralBackground1,
  },
  header: {
    display: 'flex',
    alignItems: 'center',
    gap: '4px',
    padding: '6px 8px',
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    flexShrink: 0,
  },
  brand: { fontWeight: 700, color: '#1F6F68', fontSize: '15px', marginRight: '4px', whiteSpace: 'nowrap' },
  pill: {
    flex: 1,
    minWidth: 0,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderRadius: '999px',
    padding: '2px 10px',
    fontSize: '12px',
    background: tokens.colorNeutralBackground2,
    color: tokens.colorNeutralForeground1,
    cursor: 'pointer',
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
    textAlign: 'left',
  },
  body: { flex: 1, minHeight: 0 },
});

type View = 'chat' | 'history' | 'settings' | 'diagnostics';

export function App() {
  const s = useStyles();
  const settings = useSettings();
  const active = activeConnection(settings);
  const [view, setView] = useState<View>(active ? 'chat' : 'settings');
  const [chatKey, setChatKey] = useState(0); // 变化即重开对话
  const [openSession, setOpenSession] = useState<Session | undefined>();

  // 开发测试通道：新对话
  useEffect(() => {
    if (!import.meta.env.DEV) return;
    import('./dev/bridge').then(({ devApi }) => {
      devApi.view = (v: View) => {
        setView(v);
        return v;
      };
      devApi.newChat = () => {
        setOpenSession(undefined);
        setChatKey((k) => k + 1);
        setView('chat');
        return 'ok';
      };
    });
  }, []);

  const pillText = active
    ? `${active.id === CUSTOM_ID && active.cfg.name ? active.cfg.name : providerMeta(active.id).name} · ${active.cfg.model}`
    : t('未配置模型');

  return (
    <div className={s.root}>
      <div className={s.header}>
        {view === 'chat' ? (
          <span className={s.brand}>{t('妙笔')}</span>
        ) : (
          <Button appearance="subtle" size="small" icon={<ArrowLeftRegular />} onClick={() => setView(view === 'diagnostics' ? 'settings' : 'chat')} aria-label={t('返回')} />
        )}
        <button className={s.pill} onClick={() => setView('settings')} title={t('切换模型')}>
          {pillText}
        </button>
        <Tooltip content={t('新对话')} relationship="label">
          <Button
            appearance="subtle"
            size="small"
            icon={<AddRegular />}
            onClick={() => {
              setOpenSession(undefined);
              setChatKey((k) => k + 1);
              setView('chat');
            }}
          />
        </Tooltip>
        <Tooltip content={t('历史')} relationship="label">
          <Button appearance="subtle" size="small" icon={<HistoryRegular />} onClick={() => setView('history')} />
        </Tooltip>
        <Tooltip content={t('设置')} relationship="label">
          <Button appearance="subtle" size="small" icon={<SettingsRegular />} onClick={() => setView('settings')} />
        </Tooltip>
      </div>
      <div className={s.body}>
        {/* 对话页常驻，切走时隐藏而不卸载，避免中断正在运行的任务 */}
        <div style={{ display: view === 'chat' ? 'block' : 'none', height: '100%' }}>
          <ChatView key={chatKey} initialSession={openSession} onOpenSettings={() => setView('settings')} />
        </div>
        {view === 'history' && (
          <HistoryView
            onOpen={(x) => {
              setOpenSession(x);
              setChatKey((k) => k + 1);
              setView('chat');
            }}
          />
        )}
        {view === 'settings' && <SettingsView welcome={!active} onOpenDiagnostics={() => setView('diagnostics')} />}
        {view === 'diagnostics' && (
          <div style={{ height: '100%', overflowY: 'auto' }}>
            <Diagnostics />
          </div>
        )}
      </div>
    </div>
  );
}
