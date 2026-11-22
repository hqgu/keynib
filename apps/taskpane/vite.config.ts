import { defineConfig } from 'vite';
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

export default defineConfig({
  // 生产构建部署在 GitHub Pages 子路径（/keynib/）；开发时为根路径
  base: process.env.KEYNIB_BASE ?? '/',
  plugins: [react(), devBridge()],
  server: { port: 3000, strictPort: true, https },
  build: {
    rollupOptions: { input: { taskpane: path.resolve(__dirname, 'taskpane.html') } },
  },
});
