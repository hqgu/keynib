#!/bin/bash
# 妙笔 Keynib · Mac 一键安装
# 用法：在"终端"中运行
#   curl -fsSL https://hqgu.github.io/keynib/install-mac.sh | bash
# 卸载：
#   curl -fsSL https://hqgu.github.io/keynib/install-mac.sh | bash -s -- --uninstall
#
# 做的事：把加载项清单（几 KB 的 XML）放进 PowerPoint / Word / Excel 各自的加载项文件夹。
# 不需要管理员权限，不安装任何程序。面板本身是托管在 GitHub Pages 上的网页。

set -u
MANIFEST_URL="https://hqgu.github.io/keynib/manifest.xml"
FILE="keynib.manifest.xml"
APPS=("Powerpoint" "Word" "Excel")
NAMES=("PowerPoint" "Word" "Excel")

if [[ "${1:-}" == "--uninstall" ]]; then
  for app in "${APPS[@]}"; do
    rm -f "$HOME/Library/Containers/com.microsoft.$app/Data/Documents/wef/$FILE" 2>/dev/null
  done
  echo "已卸载妙笔 Keynib。请完全退出（⌘Q）并重新打开 Office。"
  echo "Keynib uninstalled. Quit (⌘Q) and reopen Office."
  exit 0
fi

TMP="$(mktemp)"
if ! curl -fsSL "$MANIFEST_URL" -o "$TMP"; then
  echo "下载清单失败，请检查网络后重试。/ Failed to download the manifest."
  exit 1
fi

ok=()
fail=()
for i in "${!APPS[@]}"; do
  app="${APPS[$i]}"
  container="$HOME/Library/Containers/com.microsoft.$app"
  if [[ ! -d "$container" ]]; then
    continue # 没装这个程序（或从未打开过）
  fi
  dir="$container/Data/Documents/wef"
  if mkdir -p "$dir" 2>/dev/null && cp "$TMP" "$dir/$FILE" 2>/dev/null; then
    ok+=("${NAMES[$i]}")
  else
    fail+=("${NAMES[$i]}|$dir")
  fi
done
rm -f "$TMP"

if ((${#ok[@]})); then
  echo "✓ 已安装到：${ok[*]}"
  echo "  请完全退出（⌘Q）并重新打开这些程序，在「开始」选项卡点击「妙笔助手」。"
  echo "  Installed. Quit (⌘Q) and reopen Office, then click Keynib on the Home tab."
fi

if ((${#fail[@]})); then
  echo ""
  echo "⚠ 以下程序因 macOS 隐私保护未能写入，需要手动拖入（约 30 秒）："
  echo "  macOS privacy protection blocked these; drag the file in manually:"
  curl -fsSL "$MANIFEST_URL" -o "$HOME/Downloads/$FILE" && open -R "$HOME/Downloads/$FILE"
  for item in "${fail[@]}"; do
    name="${item%%|*}"
    dir="${item#*|}"
    echo "  - $name：把「下载」里的 $FILE 按住 ⌥ 拖进刚打开的文件夹"
    mkdir -p "$dir" 2>/dev/null
    open "$dir" 2>/dev/null
  done
fi

if ((${#ok[@]} == 0 && ${#fail[@]} == 0)); then
  echo "没有找到 PowerPoint / Word / Excel。请先安装并至少打开一次 Microsoft Office。"
  echo "No Office apps found. Install Microsoft Office and open it once first."
fi
