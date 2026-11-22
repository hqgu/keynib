// 轻量 Markdown 渲染：只生成 React 文本节点，不使用 innerHTML，天然防 XSS
// 支持：标题、无序 / 有序列表、**粗体**、`代码`、代码块、段落

import { Fragment, type ReactNode } from 'react';
import { makeStyles, tokens } from '@fluentui/react-components';

const useStyles = makeStyles({
  root: { lineHeight: '1.6', wordBreak: 'break-word', '& p': { margin: '0 0 6px' }, '& ul, & ol': { margin: '0 0 6px', paddingLeft: '20px' } },
  h: { fontWeight: 600, margin: '8px 0 4px' },
  code: {
    fontFamily: tokens.fontFamilyMonospace,
    fontSize: '12px',
    backgroundColor: tokens.colorNeutralBackground3,
    borderRadius: '3px',
    padding: '0 3px',
  },
  pre: {
    fontFamily: tokens.fontFamilyMonospace,
    fontSize: '12px',
    backgroundColor: tokens.colorNeutralBackground3,
    borderRadius: '4px',
    padding: '8px',
    overflowX: 'auto',
    margin: '0 0 6px',
    whiteSpace: 'pre-wrap',
  },
});

function inline(text: string, codeClass: string): ReactNode[] {
  const out: ReactNode[] = [];
  const re = /(\*\*[^*]+\*\*|`[^`]+`)/g;
  let last = 0;
  let m: RegExpExecArray | null;
  let k = 0;
  while ((m = re.exec(text))) {
    if (m.index > last) out.push(text.slice(last, m.index));
    const tok = m[0];
    out.push(
      tok.startsWith('**') ? <strong key={k++}>{tok.slice(2, -2)}</strong> : <code key={k++} className={codeClass}>{tok.slice(1, -1)}</code>,
    );
    last = m.index + tok.length;
  }
  if (last < text.length) out.push(text.slice(last));
  return out;
}

export function Markdown({ text }: { text: string }) {
  const s = useStyles();
  const lines = text.replace(/\r\n/g, '\n').split('\n');
  const blocks: ReactNode[] = [];
  let i = 0;
  let key = 0;
  while (i < lines.length) {
    const line = lines[i];
    if (line.startsWith('```')) {
      const body: string[] = [];
      i++;
      while (i < lines.length && !lines[i].startsWith('```')) body.push(lines[i++]);
      i++;
      blocks.push(<pre key={key++} className={s.pre}>{body.join('\n')}</pre>);
    } else if (/^#{1,4}\s/.test(line)) {
      blocks.push(<div key={key++} className={s.h}>{inline(line.replace(/^#+\s/, ''), s.code)}</div>);
      i++;
    } else if (/^\s*([-*]|\d+\.)\s/.test(line)) {
      const ordered = /^\s*\d+\./.test(line);
      const items: string[] = [];
      while (i < lines.length && /^\s*([-*]|\d+\.)\s/.test(lines[i])) items.push(lines[i++].replace(/^\s*([-*]|\d+\.)\s/, ''));
      const List = ordered ? 'ol' : 'ul';
      blocks.push(
        <List key={key++}>
          {items.map((t, j) => (
            <li key={j}>{inline(t, s.code)}</li>
          ))}
        </List>,
      );
    } else if (!line.trim()) {
      i++;
    } else {
      const para: string[] = [];
      while (i < lines.length && lines[i].trim() && !/^(#{1,4}\s|```|\s*([-*]|\d+\.)\s)/.test(lines[i])) para.push(lines[i++]);
      blocks.push(
        <p key={key++}>
          {para.map((t, j) => (
            <Fragment key={j}>
              {j > 0 && <br />}
              {inline(t, s.code)}
            </Fragment>
          ))}
        </p>,
      );
    }
  }
  return <div className={s.root}>{blocks}</div>;
}
