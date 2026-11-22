# 妙笔 Keynib

自带密钥（BYOK）的 Office AI 助手，支持 PowerPoint、Word、Excel。开源免费（MIT）。
Bring-your-own-key AI assistant for PowerPoint, Word and Excel. Free and open source (MIT).

**安装 / Install：<https://hqgu.github.io/keynib/>** · [隐私说明 / Privacy](https://hqgu.github.io/keynib/privacy.html)

> 测试版。Mac 桌面版已实测；网页版与 Windows 桌面版尚在测试中。
> Preview. Tested on Mac desktop; Office on the web and Windows desktop are still being tested.

## 功能

| 程序 | 能做什么 | 安全网 |
|---|---|---|
| PowerPoint | 按大纲生成幻灯片、改文字与格式、表格、形状；"演示规范"（字体、字号、页脚学术短引）与规范检查；截图自检 | 每轮修改可一键撤销 |
| Word | 学术润色、逻辑检查、插入段落 / 表格、批注 | 所有修改以**修订痕迹**呈现，只标出改动的词句；可按轮接受 / 拒绝 |
| Excel | 数据画像、描述统计、分组比较、图表（含箱线图、直方图）、数据透视表 | 统计量全部写成 Excel 公式；结果放新工作表，撤销即删除；改原表需确认 |

通用：`/` 快捷指令（可自定义）、PubMed 引用检索与核对（可选）、会话历史、配置码（在三个程序间同步设置）、中英双语界面。
删除、覆盖原表等破坏性操作必须由用户在面板上确认。

## 支持的模型服务

内置多家大模型服务的接口预设，并支持任意 OpenAI 兼容接口（含本地模型，如 Ollama）。你填写自己的 API Key，文档内容从本机直接发送到你选择的服务，妙笔没有服务器、不收集数据。

## 开发

```bash
npm install
npm run certs          # 安装本地 HTTPS 开发证书（一次）
npm run dev            # https://localhost:3000
npm run sideload:mac   # 把开发版清单装进 Mac 版 Office
npm test               # 单元测试
npm run typecheck
npm run manifest       # 由 scripts/gen-manifest.mjs 生成开发版 / 正式版清单
```

目录：`apps/taskpane`（面板）、`packages/providers`（模型接口适配）、`packages/core`（Agent 循环、PubMed）、`docs/`（规划与各阶段测试记录）。详见 [docs/开发指南.md](docs/开发指南.md)。

推送到 `main` 后由 GitHub Actions 自动测试、构建并发布到 GitHub Pages。

## License

MIT. 第三方名称与商标说明见 [NOTICE.md](NOTICE.md)。
