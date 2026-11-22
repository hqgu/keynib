import { defineConfig, type Plugin } from 'vite';
import react from '@vitejs/plugin-react';
import { devBridge } from './devBridgePlugin';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';

// 开发证书由 `npm run certs`（office-addin-dev-certs）生成在 ~/.office-addin-dev-certs
const certDir = path.join(os.homedir(), '.office-addin-dev-certs');
const certFile = path.join(certDir, 'localhost.crt');
const keyFile = path.join(certDir, 'localhost.key');
const https = fs.existsSync(certFile) && fs.existsSync(keyFile)
  ? { cert: fs.readFileSync(certFile), key: fs.readFileSync(keyFile) }
  : undefined;

if (!https) {
  console.warn('\n[keynib] 未找到开发证书，Office 无法加载 http 页面。请先运行：npm run certs\n');
}

// 每次构建的版本号：写入 version.json，并注入页面；面板启动时比对，发现新版自动换网址重载（绕过 Pages 10 分钟缓存）
const BUILD_ID = new Date().toISOString().replace(/[-:.TZ]/g, '').slice(0, 14);
const versionFile = (): Plugin => ({
  name: 'keynib-version',
  apply: 'build',
  generateBundle() {
    this.emitFile({ type: 'asset', fileName: 'version.json', source: JSON.stringify({ build: BUILD_ID }) });
  },
});

export default defineConfig({
  define: { __BUILD_ID__: JSON.stringify(BUILD_ID) },
  // 生产构建部署在 GitHub Pages 子路径（/keynib/）；开发时为根路径
  base: process.env.KEYNIB_BASE ?? '/',
  plugins: [react(), devBridge(), versionFile()],
  server: { port: 3000, strictPort: true, https },
  build: {
    rollupOptions: { input: { taskpane: path.resolve(__dirname, 'taskpane.html') } },
  },
});
