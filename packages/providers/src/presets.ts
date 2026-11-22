// 内置厂商预设（Provider Presets）
// protocol 决定使用哪个协议适配器；baseUrls 第一项为默认值，用户可在设置中修改。
// cors 字段记录 M0 探针结论（2026-11-22 预检实测，见 docs/M0-探针报告.md）。

export type Protocol = 'anthropic' | 'openai' | 'gemini';

export interface BaseUrlOption {
  label: string; // 如 "国内" / "国际"
  url: string;
  protocol?: Protocol; // 与厂商默认协议不同时标注，如 DeepSeek 的 Anthropic 兼容接口
}

export interface ProviderPreset {
  id: string;
  name: string; // 界面显示名
  vendor: string; // 厂商名（仅文字，不使用 Logo）
  protocol: Protocol;
  baseUrls: BaseUrlOption[];
  modelsPath: string; // 相对 baseUrl 的模型列表路径
  keyUrl: string; // "获取 Key" 链接
  cors: 'ok' | 'proxy' | 'unknown';
  notes?: string;
}

export const PROVIDER_PRESETS: ProviderPreset[] = [
  {
    id: 'anthropic',
    name: 'Claude',
    vendor: 'Anthropic',
    protocol: 'anthropic',
    baseUrls: [{ label: '默认', url: 'https://api.anthropic.com/v1' }],
    modelsPath: '/models',
    keyUrl: 'https://console.anthropic.com/settings/keys',
    cors: 'ok',
    notes: '浏览器直连需请求头 anthropic-dangerous-direct-browser-access: true',
  },
  {
    id: 'openai',
    name: 'ChatGPT',
    vendor: 'OpenAI',
    protocol: 'openai',
    baseUrls: [{ label: '默认', url: 'https://api.openai.com/v1' }],
    modelsPath: '/models',
    keyUrl: 'https://platform.openai.com/api-keys',
    cors: 'ok',
  },
  {
    id: 'gemini',
    name: 'Gemini',
    vendor: 'Google',
    // M1 先走 OpenAI 兼容端点，M2 再补原生协议
    protocol: 'openai',
    baseUrls: [{ label: 'OpenAI 兼容', url: 'https://generativelanguage.googleapis.com/v1beta/openai' }],
    modelsPath: '/models',
    keyUrl: 'https://aistudio.google.com/apikey',
    cors: 'ok',
  },
  {
    id: 'xai',
    name: 'Grok',
    vendor: 'xAI',
    protocol: 'openai',
    baseUrls: [{ label: '默认', url: 'https://api.x.ai/v1' }],
    modelsPath: '/models',
    keyUrl: 'https://console.x.ai',
    cors: 'ok',
  },
  {
    id: 'deepseek',
    name: 'DeepSeek',
    vendor: 'DeepSeek',
    protocol: 'openai',
    baseUrls: [
      { label: '默认', url: 'https://api.deepseek.com' },
      { label: 'Anthropic 兼容', url: 'https://api.deepseek.com/anthropic/v1', protocol: 'anthropic' },
    ],
    modelsPath: '/models',
    keyUrl: 'https://platform.deepseek.com/api_keys',
    cors: 'ok',
  },
  {
    id: 'glm',
    name: 'GLM',
    vendor: '智谱 Zhipu',
    protocol: 'openai',
    baseUrls: [
      { label: '国内', url: 'https://open.bigmodel.cn/api/paas/v4' },
      { label: '国际', url: 'https://api.z.ai/api/paas/v4' },
    ],
    modelsPath: '/models',
    keyUrl: 'https://open.bigmodel.cn/usercenter/apikeys',
    cors: 'ok',
  },
  {
    id: 'minimax',
    name: 'MiniMax',
    vendor: 'MiniMax',
    protocol: 'openai',
    baseUrls: [
      { label: '国内', url: 'https://api.minimaxi.com/v1' },
      { label: '国际', url: 'https://api.minimax.io/v1' },
    ],
    modelsPath: '/models',
    keyUrl: 'https://platform.minimaxi.com/user-center/basic-information/interface-key',
    cors: 'ok',
  },
];

// 按协议生成鉴权请求头
export function authHeaders(protocol: Protocol, apiKey: string): Record<string, string> {
  switch (protocol) {
    case 'anthropic':
      return {
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01',
        'anthropic-dangerous-direct-browser-access': 'true',
      };
    case 'gemini':
      return { 'x-goog-api-key': apiKey };
    default:
      return { Authorization: `Bearer ${apiKey}` };
  }
}
