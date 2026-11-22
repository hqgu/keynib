// 设置页：演示规范档案（选择、编辑、复制、删除）

import { useState } from 'react';
import { Button, Field, Input, makeStyles, Radio, RadioGroup, Select, Switch, Text, Textarea, tokens } from '@fluentui/react-components';
import { CopyRegular, DeleteRegular, EditRegular } from '@fluentui/react-icons';
import { updateSettings, useSettings } from '../state/settings';
import { t } from '../i18n';
import { DEFAULT_PROFILE, type DeckProfile } from '../hosts/powerpoint/profile';

const useStyles = makeStyles({
  section: { display: 'flex', flexDirection: 'column', gap: '8px' },
  h: { fontSize: '13px', fontWeight: 600, color: tokens.colorNeutralForeground2 },
  row: { display: 'flex', gap: '6px', alignItems: 'center' },
  grid: { display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '6px' },
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

export function ProfileSection() {
  const s = useStyles();
  const settings = useSettings();
  const [editing, setEditing] = useState<string | null>(null);

  const save = (p: DeckProfile) => updateSettings((st) => ({ ...st, profiles: st.profiles.map((x) => (x.id === p.id ? p : x)) }));
  const duplicate = (p: DeckProfile) => {
    const copy = { ...structuredClone(p), id: crypto.randomUUID(), name: t('{name}（副本）', { name: p.name }) };
    updateSettings((st) => ({ ...st, profiles: [...st.profiles, copy] }));
    setEditing(copy.id);
  };
  const remove = (id: string) =>
    updateSettings((st) => {
      const profiles = st.profiles.filter((x) => x.id !== id);
      return { ...st, profiles: profiles.length ? profiles : [DEFAULT_PROFILE], activeProfileId: st.activeProfileId === id ? null : st.activeProfileId };
    });

  return (
    <div className={s.section}>
      <span className={s.h}>{t('演示规范（PowerPoint）')}</span>
      <Text size={200}>{t('选中的规范会在每轮交给 AI 遵守，并用于"规范检查"。')}</Text>
      <RadioGroup value={settings.activeProfileId ?? ''} onChange={(_, d) => updateSettings({ activeProfileId: d.value || null })}>
        <Radio value="" label={t('不使用规范')} />
        {settings.profiles.map((p) => (
          <div key={p.id}>
            <div className={s.row}>
              <Radio value={p.id} label={p.name} />
              <span style={{ flex: 1 }} />
              <Button size="small" appearance="subtle" icon={<EditRegular />} aria-label={t('编辑')} onClick={() => setEditing(editing === p.id ? null : p.id)} />
              <Button size="small" appearance="subtle" icon={<CopyRegular />} aria-label={t('复制')} onClick={() => duplicate(p)} />
              <Button size="small" appearance="subtle" icon={<DeleteRegular />} aria-label={t('删除')} onClick={() => remove(p.id)} />
            </div>
            {editing === p.id && <ProfileEditor profile={p} onChange={save} />}
          </div>
        ))}
      </RadioGroup>
    </div>
  );
}

function ProfileEditor({ profile: p, onChange }: { profile: DeckProfile; onChange: (p: DeckProfile) => void }) {
  const s = useStyles();
  const num = (v: string, fallback: number) => Math.max(6, Math.min(120, Number(v) || fallback));
  return (
    <div className={s.editor}>
      <Field label={t('名称')}>
        <Input value={p.name} onChange={(_, d) => onChange({ ...p, name: d.value })} />
      </Field>
      <div className={s.grid}>
        <Field label={t('标题字体')}>
          <Input value={p.fonts.title} onChange={(_, d) => onChange({ ...p, fonts: { ...p.fonts, title: d.value } })} />
        </Field>
        <Field label={t('正文字体')}>
          <Input value={p.fonts.body} onChange={(_, d) => onChange({ ...p, fonts: { ...p.fonts, body: d.value } })} />
        </Field>
        <Field label={t('标题字号')}>
          <Input type="number" value={String(p.sizes.title)} onChange={(_, d) => onChange({ ...p, sizes: { ...p.sizes, title: num(d.value, 36) } })} />
        </Field>
        <Field label={t('正文字号')}>
          <Input type="number" value={String(p.sizes.body)} onChange={(_, d) => onChange({ ...p, sizes: { ...p.sizes, body: num(d.value, 20) } })} />
        </Field>
        <Field label={t('最小字号')}>
          <Input type="number" value={String(p.sizes.min)} onChange={(_, d) => onChange({ ...p, sizes: { ...p.sizes, min: num(d.value, 16) } })} />
        </Field>
        <Field label={t('版面')}>
          <Select value={p.density} onChange={(_, d) => onChange({ ...p, density: d.value as DeckProfile['density'] })}>
            <option value="full">{t('饱满（少留白）')}</option>
            <option value="normal">{t('常规')}</option>
          </Select>
        </Field>
      </div>
      <Field label={t('其他允许的字体（逗号分隔）')}>
        <Input
          value={p.fonts.allowed.join(', ')}
          onChange={(_, d) => onChange({ ...p, fonts: { ...p.fonts, allowed: d.value.split(/[,，]\s*/).filter(Boolean) } })}
        />
      </Field>
      <Switch checked={p.sizes.titleBold} onChange={(_, d) => onChange({ ...p, sizes: { ...p.sizes, titleBold: d.checked } })} label={t('标题加粗')} />
      <Switch checked={p.citation.enabled} onChange={(_, d) => onChange({ ...p, citation: { ...p.citation, enabled: d.checked } })} label={t('页脚标注学术短引')} />
      {p.citation.enabled && (
        <Field label={t('引用格式示例')}>
          <Input value={p.citation.example} onChange={(_, d) => onChange({ ...p, citation: { ...p.citation, example: d.value } })} />
        </Field>
      )}
      <Switch checked={p.noReferencesSlide} onChange={(_, d) => onChange({ ...p, noReferencesSlide: d.checked })} label={t('末尾不设参考文献页')} />
      <Field label={t('其他要求（会原样交给 AI）')}>
        <Textarea value={p.notes} onChange={(_, d) => onChange({ ...p, notes: d.value })} resize="vertical" />
      </Field>
      <Text size={200}>{t('说明：字距调整（kern）目前无法通过 Office 接口设置；使用模板请直接用模板新建演示文稿，妙笔会使用其中的版式。')}</Text>
    </div>
  );
}
