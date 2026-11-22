import React from 'react';
import { createRoot } from 'react-dom/client';
import { FluentProvider, webDarkTheme, webLightTheme } from '@fluentui/react-components';
import { App } from './App';
import { isDarkColor } from './probes/env';

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
