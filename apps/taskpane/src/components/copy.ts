// 复制到剪贴板：网页版 Office 的面板框架常被禁止写剪贴板（navigator.clipboard 失败且无提示）。
// 依次尝试 Clipboard API、execCommand('copy')；都失败时返回 false，由调用方提示用户手动 ⌘C。
export async function copyText(text: string, select?: HTMLTextAreaElement | null): Promise<boolean> {
  try {
    await navigator.clipboard.writeText(text);
    return true;
  } catch {
    /* 继续尝试 */
  }
  const ta = select ?? document.createElement('textarea');
  const temp = !select;
  if (temp) {
    ta.value = text;
    ta.style.position = 'fixed';
    ta.style.opacity = '0';
    document.body.appendChild(ta);
  }
  ta.focus();
  ta.select();
  let ok = false;
  try {
    ok = document.execCommand('copy');
  } catch {
    ok = false;
  }
  if (temp) ta.remove();
  return ok;
}
