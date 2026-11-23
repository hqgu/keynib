// 历史页：当前文档 / 全部；搜索；打开继续；导出 Markdown；删除

import { useEffect, useMemo, useState } from 'react';
import { Button, Input, makeStyles, Tab, TabList, Text, Textarea, tokens } from '@fluentui/react-components';
import { CopyRegular, DeleteRegular, SearchRegular } from '@fluentui/react-icons';
import { t } from '../i18n';
import { copyText } from '../components/copy';
import { currentDoc, deleteSession, listSessions, sessionToMarkdown, type Session } from '../state/sessions';

const useStyles = makeStyles({
  root: { padding: '12px', display: 'flex', flexDirection: 'column', gap: '10px', height: '100%', boxSizing: 'border-box', overflowY: 'auto' },
  group: { fontSize: '12px', fontWeight: 600, color: tokens.colorNeutralForeground3, marginTop: '6px' },
  item: {
    display: 'flex',
    alignItems: 'center',
    gap: '4px',
    padding: '6px 8px',
    borderRadius: '8px',
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    backgroundColor: tokens.colorNeutralBackground1,
  },
  main: { flex: 1, minWidth: 0, cursor: 'pointer' },
  title: { overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', fontSize: '13px' },
  sub: { fontSize: '11px', color: tokens.colorNeutralForeground3, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' },
});

function dayLabel(time: number) {
  const d = new Date(time);
  const today = new Date();
  const diff = Math.floor((new Date(today.toDateString()).getTime() - new Date(d.toDateString()).getTime()) / 86400000);
  return diff === 0 ? t('今天') : diff === 1 ? t('昨天') : diff < 7 ? t('{n} 天前', { n: diff }) : d.toLocaleDateString();
}

export function HistoryView({ onOpen }: { onOpen: (s: Session) => void }) {
  const s = useStyles();
  const [all, setAll] = useState<Session[]>([]);
  const [docId, setDocId] = useState('');
  const [scope, setScope] = useState<'doc' | 'all'>('doc');
  const [q, setQ] = useState('');
  // 复制结果：text 为空表示已复制；非空表示需要用户手动复制
  const [manual, setManual] = useState<{ id: string; text: string } | null>(null);

  const reload = () => listSessions().then(setAll).catch(() => setAll([]));
  useEffect(() => {
    reload();
    currentDoc().then((d) => setDocId(d.docId)).catch(() => {});
  }, []);

  const shown = useMemo(() => {
    const kw = q.trim().toLowerCase();
    return all.filter(
      (x) =>
        (scope === 'all' || x.docId === docId) &&
        (!kw || x.title.toLowerCase().includes(kw) || x.messages.some((m) => m.role === 'user' && m.content.toLowerCase().includes(kw))),
    );
  }, [all, scope, q, docId]);

  let lastGroup = '';
  return (
    <div className={s.root}>
      <TabList selectedValue={scope} onTabSelect={(_, d) => setScope(d.value as 'doc' | 'all')} size="small">
        <Tab value="doc">{t('当前文档')}</Tab>
        <Tab value="all">{t('全部')}</Tab>
      </TabList>
      <Input contentBefore={<SearchRegular />} placeholder={t('搜索会话')} value={q} onChange={(_, d) => setQ(d.value)} />
      {shown.length === 0 && <Text size={200}>{scope === 'doc' ? t('这份文档还没有会话。') : t('还没有会话。')}</Text>}
      {shown.map((x) => {
        const g = dayLabel(x.updatedAt);
        const header = g !== lastGroup ? <div className={s.group}>{g}</div> : null;
        lastGroup = g;
        return (
          <div key={x.id}>
            {header}
            <div className={s.item}>
              <div className={s.main} onClick={() => onOpen(x)} role="button">
                <div className={s.title}>{x.title}</div>
                <div className={s.sub}>
                  {scope === 'all' ? `${x.host} · ${x.docName} · ` : ''}
                  {x.model} · {t('{n} 条提问', { n: x.messages.filter((m) => m.role === 'user').length })}
                </div>
              </div>
              <Button
                size="small"
                appearance="subtle"
                icon={<CopyRegular />}
                aria-label={t('复制为 Markdown')}
                title={t('复制为 Markdown')}
                onClick={async () => {
                  const md = sessionToMarkdown(x);
                  setManual((await copyText(md)) ? { id: x.id, text: '' } : { id: x.id, text: md });
                }}
              />
              <Button
                size="small"
                appearance="subtle"
                icon={<DeleteRegular />}
                aria-label={t('删除')}
                onClick={async () => {
                  await deleteSession(x.id);
                  reload();
                }}
              />
            </div>
            {manual?.id === x.id &&
              (manual.text ? (
                <Textarea
                  readOnly
                  value={manual.text}
                  textarea={{ rows: 4, onFocus: (e) => e.currentTarget.select() }}
                  placeholder={t('浏览器不允许自动复制，请全选后按 ⌘C / Ctrl+C')}
                  aria-label={t('浏览器不允许自动复制，请全选后按 ⌘C / Ctrl+C')}
                />
              ) : (
                <Text size={200}>{t('已复制 ✓')}</Text>
              ))}
          </div>
        );
      })}
    </div>
  );
}
