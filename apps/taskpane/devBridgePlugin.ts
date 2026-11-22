// 开发测试通道（Dev Bridge，仅 `vite` 开发服务器生效，生产构建不包含）
// 终端 → POST /__bridge/cmd → 面板长轮询 GET /__bridge/next?host=… 取到命令并执行 → POST /__bridge/result → 终端拿到结果。
// 多个 Office 程序同时打开面板时，按 host@platform 路由（默认 PowerPoint@Mac）。

import type { Plugin } from 'vite';
import type { IncomingMessage, ServerResponse } from 'node:http';

interface Cmd {
  id: string;
  type: string;
  host: string;
  payload?: unknown;
}

export function devBridge(): Plugin {
  const queue: Cmd[] = [];
  const pollers = new Map<string, ServerResponse>(); // host → 等待中的长轮询
  const pending = new Map<string, ServerResponse>();
  const lastSeen = new Map<string, number>();

  const readBody = (req: IncomingMessage) =>
    new Promise<string>((resolve) => {
      let b = '';
      req.on('data', (c) => (b += c));
      req.on('end', () => resolve(b));
    });

  const send = (res: ServerResponse, status: number, body?: unknown) => {
    if (res.writableEnded) return;
    res.statusCode = status;
    res.setHeader('content-type', 'application/json');
    res.end(body === undefined ? '' : JSON.stringify(body));
  };

  const dispatch = () => {
    for (let i = 0; i < queue.length; i++) {
      const poller = pollers.get(queue[i].host);
      if (!poller) continue;
      pollers.delete(queue[i].host);
      send(poller, 200, queue.splice(i, 1)[0]);
      i--;
    }
  };

  return {
    name: 'keynib-dev-bridge',
    apply: 'serve',
    configureServer(server) {
      server.middlewares.use('/__bridge', async (req, res) => {
        const url = new URL(req.url ?? '/', 'https://localhost');
        if (req.method === 'GET' && url.pathname === '/next') {
          const host = url.searchParams.get('host') ?? 'Unknown';
          lastSeen.set(host, Date.now());
          const old = pollers.get(host);
          if (old) send(old, 204);
          pollers.set(host, res);
          const timer = setTimeout(() => {
            if (pollers.get(host) === res) {
              pollers.delete(host);
              send(res, 204);
            }
          }, 25000);
          res.on('close', () => clearTimeout(timer));
          dispatch();
        } else if (req.method === 'POST' && url.pathname === '/result') {
          const body = JSON.parse((await readBody(req)) || '{}');
          const waiter = pending.get(body.id);
          pending.delete(body.id);
          if (waiter) send(waiter, 200, body);
          send(res, 204);
        } else if (req.method === 'POST' && url.pathname === '/cmd') {
          const body = JSON.parse((await readBody(req)) || '{}');
          const cmd: Cmd = { id: Math.random().toString(36).slice(2), type: body.type, host: body.host ?? 'PowerPoint@Mac', payload: body.payload };
          pending.set(cmd.id, res);
          queue.push(cmd);
          dispatch();
        } else if (req.method === 'GET' && url.pathname === '/status') {
          const now = Date.now();
          send(res, 200, Object.fromEntries([...lastSeen].map(([h, t]) => [h, now - t < 30000 ? '已连接' : '未连接'])));
        } else {
          send(res, 404);
        }
      });
    },
  };
}
