// 浏览器直连探针（CORS Probe）
// 原理：带上与真实请求相同的鉴权头访问模型列表。
// - 能读到 HTTP 状态码（哪怕 401）→ 浏览器允许跨域，直连可行；
// - fetch 抛 TypeError → 被 CORS 拦截或网络不通，需要中转。
// 不传 apiKey 时用占位 Key，只测跨域；传入真实 Key 时额外返回模型数量。

import { authHeaders, type ProviderPreset } from './presets.ts';

export interface ProbeResult {
  providerId: string;
  baseUrl: string;
  reachable: boolean; // 浏览器能否拿到响应
  status?: number;
  ms: number;
  modelCount?: number;
  sampleModels?: string[];
  error?: string;
}

const PLACEHOLDER_KEY = 'keynib-probe-placeholder';

export async function probeProvider(
  preset: ProviderPreset,
  baseUrl: string,
  apiKey?: string,
  timeoutMs = 15000,
): Promise<ProbeResult> {
  const started = performance.now();
  const ctrl = new AbortController();
  const timer = setTimeout(() => ctrl.abort(), timeoutMs);
  try {
    const res = await fetch(baseUrl.replace(/\/$/, '') + preset.modelsPath, {
      method: 'GET',
      headers: authHeaders(preset.protocol, apiKey || PLACEHOLDER_KEY),
      signal: ctrl.signal,
    });
    const result: ProbeResult = {
      providerId: preset.id,
      baseUrl,
      reachable: true,
      status: res.status,
      ms: Math.round(performance.now() - started),
    };
    if (apiKey && res.ok) {
      const body = await res.json();
      // OpenAI 兼容与 Anthropic 均返回 { data: [{ id }] }
      const ids: string[] = (body?.data ?? body?.models ?? []).map((m: any) => m.id ?? m.name);
      result.modelCount = ids.length;
      result.sampleModels = ids.slice(0, 5);
    }
    return result;
  } catch (e: any) {
    return {
      providerId: preset.id,
      baseUrl,
      reachable: false,
      ms: Math.round(performance.now() - started),
      error: e?.name === 'AbortError' ? '超时' : `${e?.name ?? 'Error'}: ${e?.message ?? e}`,
    };
  } finally {
    clearTimeout(timer);
  }
}
