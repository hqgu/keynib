// 环境与要求集探针（Requirement Set Probe）

export type HostName = 'PowerPoint' | 'Word' | 'Excel' | 'Unknown';

export function currentHost(): HostName {
  switch (Office.context?.host) {
    case Office.HostType.PowerPoint:
      return 'PowerPoint';
    case Office.HostType.Word:
      return 'Word';
    case Office.HostType.Excel:
      return 'Excel';
    default:
      return 'Unknown';
  }
}

export interface EnvInfo {
  host: HostName;
  platform: string;
  officeVersion: string;
  displayLanguage: string;
  userAgent: string;
  origin: string;
  theme?: { body: string; isDark: boolean };
  requirementSets: Record<string, string>; // 集合名 → 支持的最高版本
}

// 各宿主需检测的集合及待测的最高次版本号
const SETS: Record<HostName, [string, number][]> = {
  PowerPoint: [['PowerPointApi', 10]],
  Word: [['WordApi', 9], ['WordApiDesktop', 3]],
  Excel: [['ExcelApi', 20]],
  Unknown: [],
};

function highestSupported(name: string, maxMinor: number): string {
  for (let minor = maxMinor; minor >= 1; minor--) {
    if (Office.context.requirements.isSetSupported(name, `1.${minor}`)) return `1.${minor}`;
  }
  return '不支持';
}

// 用背景色亮度判断深色主题
export function isDarkColor(hex: string): boolean {
  const m = /^#?([0-9a-f]{6})$/i.exec(hex);
  if (!m) return false;
  const n = parseInt(m[1], 16);
  const lum = 0.299 * (n >> 16) + 0.587 * ((n >> 8) & 255) + 0.114 * (n & 255);
  return lum < 128;
}

export function collectEnv(): EnvInfo {
  const host = currentHost();
  const diag = Office.context.diagnostics;
  const theme = (Office.context as any).officeTheme as Office.OfficeTheme | undefined;
  const requirementSets: Record<string, string> = {};
  for (const [name, max] of SETS[host]) requirementSets[name] = highestSupported(name, max);
  return {
    host,
    platform: String(diag?.platform ?? '未知'),
    officeVersion: diag?.version ?? '未知',
    displayLanguage: Office.context.displayLanguage,
    userAgent: navigator.userAgent,
    origin: location.origin,
    theme: theme?.bodyBackgroundColor
      ? { body: theme.bodyBackgroundColor, isDark: isDarkColor(theme.bodyBackgroundColor) }
      : undefined,
    requirementSets,
  };
}
