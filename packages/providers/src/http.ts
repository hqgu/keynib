// HTTP 容错层：重试退避、超时、错误归类、SSE 解析

import { ProviderError } from './types.ts';

export interface FetchOptions {
  signal?: AbortSignal; // 用户"停止"
  firstByteTimeoutMs?: number; // 发出请求到拿到响应头的超时
  maxRetries?: number;
}

const sleep = (ms: number, signal?: AbortSignal) =>
  new Promise<void>((resolve, reject) => {
    const t = setTimeout(resolve, ms);
    signal?.addEventListener('abort', () => {
      clearTimeout(t);
      reject(new ProviderError('已停止', undefined, 'aborted'));
    });
  });

// 把各家的错误体提炼成一句话
async function readErrorMessage(res: Response): Promise<string> {
  try {
    const text = await res.text();
    try {
      const j = JSON.parse(text);
      return j?.error?.message ?? j?.message ?? j?.error ?? j?.base_resp?.status_msg ?? text.slice(0, 300);
    } catch {
      return text.slice(0, 300);
    }
  } catch {
    return '';
  }
}

export function classify(status: number): ProviderError['kind'] {
  if (status === 401 || status === 403) return 'auth';
  if (status === 404) return 'notFound';
  if (status === 429) return 'rateLimit';
  if (status >= 400 && status < 500) return 'badRequest';
  return 'server';
}

// 错误提示语言（由界面设置同步）
let lang: 'zh' | 'en' = 'zh';
export function setProviderLang(l: 'zh' | 'en') {
  lang = l;
}
export function getProviderLang() {
  return lang;
}

const KIND_HINT_EN: Record<ProviderError['kind'], string> = {
  auth: 'Invalid API key or no permission',
  notFound: 'Endpoint or model not found',
  rateLimit: 'Rate limited or insufficient balance',
  badRequest: 'Request rejected',
  server: 'Provider service error',
  network: 'Network unreachable (check your network or configure a relay)',
  timeout: 'Request timed out',
  aborted: 'Stopped',
};

const KIND_HINT: Record<ProviderError['kind'], string> = {
  auth: 'Key 无效或没有权限',
  notFound: '地址或模型不存在',
  rateLimit: '请求过于频繁或余额不足',
  badRequest: '请求被拒绝',
  server: '厂商服务出错',
  network: '网络不通（检查网络或配置中转）',
  timeout: '请求超时',
  aborted: '已停止',
};

export function friendlyError(e: unknown): string {
  if (e instanceof ProviderError) {
    const hint = (lang === 'en' ? KIND_HINT_EN : KIND_HINT)[e.kind];
    const [l, r, c] = lang === 'en' ? [' (', ')', ': '] : ['（', '）', '：'];
    return e.status ? `${hint}${l}${e.status}${r}${c}${e.message}` : `${hint}${e.message && e.message !== hint ? c + e.message : ''}`;
  }
  return String((e as any)?.message ?? e);
}

// 带重试的 fetch：429 / 5xx / 网络错误按指数退避重试，遵守 Retry-After
export async function fetchWithRetry(url: string, init: RequestInit, opts: FetchOptions = {}): Promise<Response> {
  const { signal, firstByteTimeoutMs = 60000, maxRetries = 3 } = opts;
  let attempt = 0;
  for (;;) {
    const ctrl = new AbortController();
    const onAbort = () => ctrl.abort();
    signal?.addEventListener('abort', onAbort);
    let timedOut = false;
    const timer = setTimeout(() => {
      timedOut = true;
      ctrl.abort();
    }, firstByteTimeoutMs);
    try {
      const res = await fetch(url, { ...init, signal: ctrl.signal });
      clearTimeout(timer);
      if (res.ok) {
        // 响应头已到；把用户的停止信号继续传给响应体读取
        return res;
      }
      const retryable = res.status === 429 || res.status >= 500;
      if (retryable && attempt < maxRetries) {
        const ra = Number(res.headers.get('retry-after'));
        await sleep(ra > 0 ? Math.min(ra * 1000, 30000) : 1000 * 2 ** attempt, signal);
        attempt++;
        continue;
      }
      throw new ProviderError(await readErrorMessage(res), res.status, classify(res.status));
    } catch (e: any) {
      clearTimeout(timer);
      if (e instanceof ProviderError) throw e;
      if (signal?.aborted) throw new ProviderError('已停止', undefined, 'aborted');
      if (timedOut) throw new ProviderError('', undefined, 'timeout');
      if (attempt < 1) {
        attempt++;
        await sleep(1000, signal);
        continue;
      }
      throw new ProviderError(String(e?.message ?? e), undefined, 'network');
    } finally {
      signal?.removeEventListener('abort', onAbort);
    }
  }
}

// 解析 SSE（Server-Sent Events）：逐个产出 {event, data}
export async function* readSSE(
  res: Response,
  signal?: AbortSignal,
  idleTimeoutMs = 120000,
): AsyncGenerator<{ event?: string; data: string }> {
  if (!res.body) throw new ProviderError('响应没有内容', res.status, 'server');
  const reader = res.body.getReader();
  const decoder = new TextDecoder();
  let buffer = '';
  const abort = () => reader.cancel().catch(() => {});
  signal?.addEventListener('abort', abort);
  try {
    for (;;) {
      let idleTimer: ReturnType<typeof setTimeout> | undefined;
      const idle = new Promise<never>((_, reject) => {
        idleTimer = setTimeout(() => {
          abort();
          reject(new ProviderError('', undefined, 'timeout'));
        }, idleTimeoutMs);
      });
      const { done, value } = await Promise.race([reader.read(), idle]);
      clearTimeout(idleTimer);
      if (signal?.aborted) throw new ProviderError('已停止', undefined, 'aborted');
      if (done) break;
      buffer += decoder.decode(value, { stream: true });
      yield* drain(false);
    }
    buffer += decoder.decode();
    yield* drain(true);
  } finally {
    signal?.removeEventListener('abort', abort);
  }

  function* drain(final: boolean): Generator<{ event?: string; data: string }> {
    const normalized = buffer.replace(/\r\n/g, '\n');
    const blocks = normalized.split('\n\n');
    buffer = final ? '' : blocks.pop() ?? '';
    for (const block of blocks) {
      let event: string | undefined;
      const data: string[] = [];
      for (const line of block.split('\n')) {
        if (line.startsWith('event:')) event = line.slice(6).trim();
        else if (line.startsWith('data:')) data.push(line.slice(5).replace(/^ /, ''));
      }
      if (data.length) yield { event, data: data.join('\n') };
    }
  }
}
