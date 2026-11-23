import React from 'react';
import { createRoot } from 'react-dom/client';
import { FluentProvider, webDarkTheme, webLightTheme } from '@fluentui/react-components';
import { App } from './App';
import { ErrorBoundary } from './components/ErrorBoundary';
import { isDarkColor } from './probes/env';

declare const __BUILD_ID__: string;

// 正式版：绕过缓存检查线上版本，有新版就带上 ?v=<版本> 重新加载（新网址不会命中旧缓存）
async function checkForUpdate() {
  if (import.meta.env.DEV) return;
  try {
    const res = await fetch(`version.json?t=${Date.now()}`, { cache: 'no-store' });
    const { build } = await res.json();
    const url = new URL(location.href);
    if (build && build !== __BUILD_ID__ && url.searchParams.get('v') !== build) {
      url.searchParams.set('v', build);
      location.replace(url.toString());
    }
  } catch {
    /* 离线或检查失败时继续使用当前版本 */
  }
}
checkForUpdate();

// 兜底：10 秒内 Office 仍未初始化面板（例如自动更新重载后未能重新连接），显示提示而不是一片空白
let rendered = false;
setTimeout(() => {
  if (rendered) return;
  const el = document.getElementById('root');
  if (!el) return;
  const reloaded = new URL(location.href).searchParams.has('v');
  el.innerHTML = `<div style="padding:16px;font:13px -apple-system,'PingFang SC','Microsoft YaHei',sans-serif">
    <h3 style="margin:0 0 8px;color:#b42318">面板未能启动 / Pane failed to start</h3>
    <p>Office 未能完成面板初始化。请点面板右上角 × 关闭后重新打开；仍不行请刷新整个网页，再截图反馈。<br>Office did not initialize the pane. Close it (×) and reopen it; if it persists, reload the page and send a screenshot.</p>
    <pre style="background:#f6f8fa;padding:8px;border-radius:6px;font-size:11px;white-space:pre-wrap">build ${__BUILD_ID__} · auto-update reload: ${reloaded ? 'yes' : 'no'}</pre></div>`;
}, 10000);

// 等 Office.js 就绪后再渲染；在普通浏览器打开时 host 为 null，仍可跑 CORS 探针
Office.onReady(() => {
  rendered = true;
  // 开发模式：启动测试通道（生产构建会被摇树移除）
  if (import.meta.env.DEV) import('./dev/bridge').then((m) => m.startDevBridge());

  const bg = (Office.context as any)?.officeTheme?.bodyBackgroundColor as string | undefined;
  const theme = bg && isDarkColor(bg) ? webDarkTheme : webLightTheme;
  createRoot(document.getElementById('root')!).render(
    <React.StrictMode>
      <ErrorBoundary>
        <FluentProvider theme={theme} style={{ height: '100%' }}>
          <App />
        </FluentProvider>
      </ErrorBoundary>
    </React.StrictMode>,
  );
});
