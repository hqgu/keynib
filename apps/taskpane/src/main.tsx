import React from 'react';
import { createRoot } from 'react-dom/client';
import { FluentProvider, webDarkTheme, webLightTheme } from '@fluentui/react-components';
import { App } from './App';
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

// 等 Office.js 就绪后再渲染；在普通浏览器打开时 host 为 null，仍可跑 CORS 探针
Office.onReady(() => {
  // 开发模式：启动测试通道（生产构建会被摇树移除）
  if (import.meta.env.DEV) import('./dev/bridge').then((m) => m.startDevBridge());

  const bg = (Office.context as any)?.officeTheme?.bodyBackgroundColor as string | undefined;
  const theme = bg && isDarkColor(bg) ? webDarkTheme : webLightTheme;
  createRoot(document.getElementById('root')!).render(
    <React.StrictMode>
      <FluentProvider theme={theme} style={{ height: '100%' }}>
        <App />
      </FluentProvider>
    </React.StrictMode>,
  );
});
