// "/" 快捷指令弹出菜单（显示在输入框上方）

import { useEffect, useRef } from 'react';
import { makeStyles, tokens } from '@fluentui/react-components';
import type { SlashCommand } from '../state/commands';
import { t } from '../i18n';

const useStyles = makeStyles({
  menu: {
    position: 'absolute',
    left: '12px',
    right: '12px',
    bottom: '100%',
    marginBottom: '4px',
    maxHeight: '260px',
    overflowY: 'auto',
    backgroundColor: tokens.colorNeutralBackground1,
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    borderRadius: '8px',
    boxShadow: tokens.shadow16,
    padding: '4px',
    zIndex: 10,
  },
  item: { padding: '6px 8px', borderRadius: '6px', cursor: 'pointer', display: 'flex', flexDirection: 'column', gap: '1px' },
  active: { backgroundColor: tokens.colorBrandBackground2 },
  name: { fontSize: '13px', fontWeight: 600 },
  tag: { fontSize: '11px', fontWeight: 400, color: tokens.colorNeutralForeground3, marginLeft: '6px' },
  desc: { fontSize: '11px', color: tokens.colorNeutralForeground3 },
  empty: { padding: '8px', fontSize: '12px', color: tokens.colorNeutralForeground3 },
  hint: { padding: '4px 8px 2px', fontSize: '11px', color: tokens.colorNeutralForeground4, borderTop: `1px solid ${tokens.colorNeutralStroke2}`, marginTop: '4px' },
});

export function SlashMenu({
  items,
  index,
  onHover,
  onPick,
}: {
  items: SlashCommand[];
  index: number;
  onHover: (i: number) => void;
  onPick: (c: SlashCommand) => void;
}) {
  const s = useStyles();
  const activeRef = useRef<HTMLDivElement>(null);
  useEffect(() => activeRef.current?.scrollIntoView({ block: 'nearest' }), [index]);
  return (
    <div className={s.menu} role="listbox" aria-label={t('快捷指令')}>
      {items.length === 0 && <div className={s.empty}>{t('没有匹配的指令（可在设置中添加自定义指令）')}</div>}
      {items.map((c, i) => (
        <div
          key={c.id}
          ref={i === index ? activeRef : undefined}
          role="option"
          aria-selected={i === index}
          className={`${s.item} ${i === index ? s.active : ''}`}
          onMouseEnter={() => onHover(i)}
          // mousedown 而不是 click：避免输入框先失焦
          onMouseDown={(e) => {
            e.preventDefault();
            onPick(c);
          }}
        >
          <span className={s.name}>
            /{c.name}
            <span className={s.tag}>{c.autoSend ? t('直接执行') : t('填入后补充')}{c.builtin ? '' : ` · ${t('自定义')}`}</span>
          </span>
          <span className={s.desc}>{c.description}</span>
        </div>
      ))}
      <div className={s.hint}>{t('↑↓ 选择 · Enter / Tab 确认 · Esc 关闭 · 可输入拼音首字母如 /rs')}</div>
    </div>
  );
}
