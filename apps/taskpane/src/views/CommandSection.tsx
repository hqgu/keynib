// 设置页：快捷指令（查看内置、增删改自定义）

import { useState } from 'react';
import { Button, Field, Input, makeStyles, Switch, Text, Textarea, tokens } from '@fluentui/react-components';
import { AddRegular, DeleteRegular, EditRegular } from '@fluentui/react-icons';
import { updateSettings, useSettings } from '../state/settings';
import { getLang, t } from '../i18n';
import { builtinCommands, type SlashCommand } from '../state/commands';

const useStyles = makeStyles({
  section: { display: 'flex', flexDirection: 'column', gap: '6px' },
  h: { fontSize: '13px', fontWeight: 600, color: tokens.colorNeutralForeground2 },
  row: { display: 'flex', gap: '4px', alignItems: 'center' },
  name: { flex: 1, minWidth: 0, fontSize: '13px', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' },
  sub: { color: tokens.colorNeutralForeground3, fontSize: '11px', marginLeft: '4px' },
  builtin: { fontSize: '12px', color: tokens.colorNeutralForeground3, lineHeight: '1.7' },
  editor: {
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderRadius: '8px',
    padding: '10px',
    display: 'flex',
    flexDirection: 'column',
    gap: '8px',
    backgroundColor: tokens.colorNeutralBackground2,
  },
});

export function CommandSection() {
  const s = useStyles();
  const settings = useSettings();
  const [editing, setEditing] = useState<string | null>(null);
  const custom = settings.customCommands;

  const save = (c: SlashCommand) =>
    updateSettings((st) => ({ ...st, customCommands: st.customCommands.map((x) => (x.id === c.id ? c : x)) }));
  const add = () => {
    const c: SlashCommand = { id: crypto.randomUUID(), name: t('新指令'), aliases: [], description: '', prompt: '', autoSend: false };
    updateSettings((st) => ({ ...st, customCommands: [...st.customCommands, c] }));
    setEditing(c.id);
  };
  const remove = (id: string) => updateSettings((st) => ({ ...st, customCommands: st.customCommands.filter((x) => x.id !== id) }));

  return (
    <div className={s.section}>
      <span className={s.h}>{t('快捷指令')}</span>
      <Text size={200}>{t('在输入框开头输入 / 调出。自定义指令与内置同名时，以自定义为准。')}</Text>
      <div className={s.builtin}>
        {t('内置：')}
        {builtinCommands({ pubmed: settings.pubmed, lang: getLang() }).map((c) => `/${c.name}`).join('  ')}
      </div>
      {custom.map((c) => (
        <div key={c.id}>
          <div className={s.row}>
            <span className={s.name}>
              /{c.name}
              <span className={s.sub}>{c.autoSend ? t('直接执行') : t('填入后补充')}</span>
            </span>
            <Button size="small" appearance="subtle" icon={<EditRegular />} aria-label={t('编辑')} onClick={() => setEditing(editing === c.id ? null : c.id)} />
            <Button size="small" appearance="subtle" icon={<DeleteRegular />} aria-label={t('删除')} onClick={() => remove(c.id)} />
          </div>
          {editing === c.id && <CommandEditor cmd={c} onChange={save} />}
        </div>
      ))}
      <Button size="small" icon={<AddRegular />} onClick={add} style={{ alignSelf: 'flex-start' }}>
        {t('新建指令')}
      </Button>
    </div>
  );
}

function CommandEditor({ cmd: c, onChange }: { cmd: SlashCommand; onChange: (c: SlashCommand) => void }) {
  const s = useStyles();
  return (
    <div className={s.editor}>
      <Field label={t('名称（输入 / 后显示）')}>
        <Input value={c.name} onChange={(_, d) => onChange({ ...c, name: d.value.replace(/\s/g, '') })} />
      </Field>
      <Field label={t('别名（逗号分隔，建议填拼音首字母，如 kt）')}>
        <Input value={c.aliases.join(', ')} onChange={(_, d) => onChange({ ...c, aliases: d.value.split(/[,，]\s*/).filter(Boolean) })} />
      </Field>
      <Field label={t('说明')}>
        <Input value={c.description} onChange={(_, d) => onChange({ ...c, description: d.value })} />
      </Field>
      <Field label={t('指令内容（发给 AI 的话）')}>
        <Textarea value={c.prompt} onChange={(_, d) => onChange({ ...c, prompt: d.value })} resize="vertical" textarea={{ rows: 4 }} />
      </Field>
      <Switch checked={c.autoSend} onChange={(_, d) => onChange({ ...c, autoSend: d.checked })} label={t('选中后直接执行（关闭则填入输入框，补充后再发送）')} />
    </div>
  );
}
