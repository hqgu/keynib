// 错误兜底：界面渲染出错时不再整片空白，而是显示错误信息与"重新加载"按钮，便于用户截图反馈
import { Component, type ReactNode } from 'react';
import { t } from '../i18n';

interface State {
  error: Error | null;
}

export class ErrorBoundary extends Component<{ children: ReactNode }, State> {
  state: State = { error: null };

  static getDerivedStateFromError(error: Error): State {
    return { error };
  }

  render() {
    const { error } = this.state;
    if (!error) return this.props.children;
    const host = (globalThis as any).Office?.context;
    return (
      <div style={{ padding: 16, fontFamily: '-apple-system, "PingFang SC", "Microsoft YaHei", sans-serif', fontSize: 13 }}>
        <h3 style={{ margin: '0 0 8px', color: '#b42318' }}>{t('面板出错了')}</h3>
        <p style={{ margin: '0 0 8px' }}>{t('请截图发给开发者，然后点面板右上角 × 关闭，再重新打开面板。')}</p>
        <pre style={{ whiteSpace: 'pre-wrap', wordBreak: 'break-all', background: '#f6f8fa', padding: 8, borderRadius: 6, fontSize: 11, maxHeight: 300, overflow: 'auto' }}>
          {`${error.name}: ${error.message}\n${(error.stack ?? '').split('\n').slice(0, 8).join('\n')}\n\n${host?.host ?? ''} · ${host?.platform ?? ''} · ${typeof __BUILD_ID__ === 'string' ? __BUILD_ID__ : 'dev'}`}
        </pre>
        {/* 实测：网页版 Excel 中在面板内 location.reload() 后面板卡在空白，关闭重开才恢复，故不再提供"重新加载"按钮 */}
      </div>
    );
  }
}

declare const __BUILD_ID__: string;
