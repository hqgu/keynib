# BYOK Office 加载项（Add-in）需求与设计规划 v0.1

## ✅ 已确认决策（2026-11-22）
- 名称：**妙笔 Keynib**（M0 做重名/商标核查）。
- MVP 宿主：**PowerPoint 优先**，其次 Word，最后 Excel。
- 审阅痕迹：**只在 Word 实现**（原生修订）；PPT / Excel 不做审阅层，仅保留"整轮撤销"快照。
- Excel 定位：**数据分析 + 图表制作**。
- CORS：不通的厂商提供**可选自部署中转**（Cloudflare Worker 模板）。
- 发布：**免费 + 开源 MIT**。

## 0. Context（背景）
- 目标：做一个自带密钥（Bring Your Own Key, BYOK）的 Office 加载项（Office Add-in），用户填自己的 API Key，在 Word / Excel / PowerPoint 任务窗格（Task Pane）里用大语言模型（Large Language Model, LLM）编辑文档。
- 硬约束：不走 OAuth 登录（避免订阅账号被封）；必须有审阅痕迹（Tracked Changes）；可查看会话历史；内置 7 家厂商；UI 简洁好看；目标上架微软商店（Microsoft Marketplace / AppSource）。
- 当前目录为空，仅 `readme.md`，属于全新项目（greenfield）。
- 参考调研结论（已核实）：
  - **hewliyang/office-agents**：MIT 许可；pnpm monorepo + Svelte 5 + Vite + TypeBox；三宿主各一个加载项；工具以"读结构 + 截图 + 执行 Office.js（execute_office_js）"为核心；会话存 IndexedDB 并按 `getDocumentId()` 绑定文档；配置存 localStorage；有 CORS proxy 选项；**未见审阅痕迹能力**；作者声明不适合上架。→ 可借鉴架构思想（适配器模式、按文档绑定会话、截图自检），不复制代码。
  - **OpenDocBot**：自定义 "fair-code" 许可（非 OSI，禁止去品牌，>30 用户需商业授权）、已上架 Marketplace、纯前端直连厂商。→ **只看产品形态，不读、不借用其代码**。
  - **Word**：`changeTrackingMode`（WordApi 1.4）可开启修订；`TrackedChange.accept()/reject()`、`getTrackedChanges()`（WordApi 1.6）可逐条接受/拒绝。
  - **Excel / PowerPoint**：Office.js 无修订 API；按已确认决策不做审阅层，仅提供整轮撤销。

## 1. 命名（Naming）
原则：不含 Office / Word / Excel / Copilot / Claude / GPT 等商标词（微软商店认证会拒绝或要求修改；当前目录名 "OfficePlus" 本身有商标风险，**不建议作产品名**）。中英双名，好记、寓意好。

| 候选 | 寓意 | 冲突风险（待核实：AppSource / GitHub / npm / 域名 / 商标） |
|---|---|---|
| **妙笔 Keynib**（推荐） | Key（自带密钥）+ Nib（笔尖）；"妙笔生花" | 新造词，预计冲突低 |
| 墨言 Inkwise | 墨 + 智慧，文书助手 | 中，需查 Inkwise 现有 SaaS |
| 文枢 Docmind | 文档中枢 | 高，Docmind 已有多家使用 |
| 笔匠 Quillsmith | 匠心改稿 | 低-中 |

实施第一步：对选定名做 AppSource 搜索、GitHub/npm 同名、`.app/.com` 域名、中美商标库检索，形成命名核查表。

## 2. 需求（MoSCoW 优先级）
**Must（MVP 必须）**
1. 三宿主统一加载项（Word / Excel / PowerPoint），一套代码。
2. 厂商设置：内置 Claude(Anthropic)、ChatGPT(OpenAI)、Gemini(Google)、Grok(xAI)、DeepSeek、GLM(智谱，国内/国际两端点)、MiniMax(国内/国际两端点) 的名称、Logo、Base URL；用户填 Key；**拉取模型列表**；**一键测试连接**（返回延迟 + 是否支持工具调用）；支持"自定义 OpenAI 兼容端点"。
3. 智能体对话（Agent Loop）：流式输出（Streaming）+ 工具调用（Tool Calling），读取选区/全文，执行编辑。
4. 审阅痕迹：仅 Word（原生修订，见 §5）；PPT/Excel 提供整轮撤销。
5. 会话历史：按文档分组、可搜索、可继续、可删除、可导出 Markdown/JSON。
6. Key 仅存本地，不上传任何自有服务器；无遥测（No Telemetry）。
7. 中英双语 UI，跟随 Office 主题（亮/暗）。
8. 开源 MIT，仓库含 LICENSE、贡献指南；不引入 OpenDocBot 代码。

**Should**
- 快捷指令（Prompt 模板）：润色、翻译、学术改写、表格清洗、生成公式、PPT 美化等，可自定义。
- 每步编辑前"计划预览"（Plan Preview）+ 一键撤销本轮全部修改。
- Token 用量与估算费用显示。
- 附件：拖入 PDF/图片/文本作为上下文（多模态模型）。

**Could**
- 截图自检（PPT 幻灯渲染图回传给视觉模型）。
- 可选 CORS 中转（用户自部署 Cloudflare Worker 模板）。
- 技能（Skills）库：用户上传 SKILL.md 定制工作流。

**Won't（本期不做）**
- OAuth / 订阅账号登录；自有后端账号体系；云同步会话。

## 3. 技术架构
- 栈：TypeScript + React 18 + Fluent UI React v9（与 Office 视觉一致，利于上架审核）+ Vite + Zustand（状态）+ Dexie（IndexedDB 封装）+ Zod（工具参数校验）。
- 清单（Manifest）：用 **XML 清单**（add-in only manifest）以保证 Word/Excel/PPT 在 Mac、Windows、Web 全覆盖；统一 JSON 清单（Unified Manifest）待 PPT/Mac 支持成熟后迁移（待核实当前支持情况）。
- 托管：纯静态站点（GitHub Pages / Cloudflare Pages），HTTPS，无后端。
- 目录结构（pnpm monorepo）：
  ```
  apps/taskpane/            # React UI 入口
  packages/core/            # Agent loop、会话、消息模型（宿主无关）
  packages/providers/       # 厂商适配器 + 预设表
  packages/hosts/word|excel|powerpoint/  # 各宿主工具集（Word 含修订层，PPT/Excel 含撤销快照）
  packages/ui/              # 共用组件（Diff 卡片、模型选择器）
  manifest/                 # manifest.xml(dev/prod)
  docs/                     # 隐私政策、支持页（上架必需）
  ```
- 宿主适配器接口：`HostAdapter { id; getContext(); tools[]; systemPrompt; getDocumentId(); undo: SnapshotLayer; review?: WordRevisionLayer }`。

## 4. 厂商抽象层（Provider Layer）
- 三类协议适配器，统一成内部 `ChatRequest/ChatChunk/ToolCall` 结构：
  1. **Anthropic Messages**：`x-api-key` + `anthropic-version`，浏览器直连需 `anthropic-dangerous-direct-browser-access: true`（待核实官方文档页）。
  2. **OpenAI Chat Completions 兼容**：OpenAI、xAI(`https://api.x.ai/v1`)、DeepSeek(`https://api.deepseek.com`)、GLM(`https://open.bigmodel.cn/api/paas/v4` / `https://api.z.ai/api/paas/v4`)、MiniMax(`https://api.minimaxi.com/v1` / `https://api.minimax.io/v1`)、自定义端点；`Authorization: Bearer`。
  3. **Gemini 原生**（`generativelanguage.googleapis.com/v1beta`，`x-goog-api-key`）；备用其 OpenAI 兼容端点。
- 预设表 `providers.json`：`{id, name, logo, baseUrl[], protocol, modelsEndpoint, defaultModels[], supportsTools, supportsVision}`；拉取失败时回落到内置默认模型列表。
- **CORS（跨域资源共享）风险**：任务窗格是浏览器环境。实施第 1 周先做"直连探针"，逐家实测浏览器直连；不通的厂商 → 提示用户使用可选自部署中转（§2 Could）或标注"需中转"。
- 测试连接：调用 models 接口 + 1 次极小工具调用请求，显示 ✅/⚠️ 及错误原因（401/403/CORS/超时）。

## 5. 审阅痕迹（Tracked Changes）——仅 Word
- **Word**：写入前设 `document.changeTrackingMode = trackAll`（WordApi 1.4），所有 AI 编辑落为原生修订；作者名受限于当前 Office 用户（待核实能否标注 "AI"），同时可选给 AI 编辑加批注（Comment）"AI 修改：原因…"；面板列出本轮修订，调用 `TrackedChange.accept/reject`（WordApi 1.6）逐条/批量处理。低版本降级：只插批注。用户也可直接在 Word 原生"审阅"选项卡接受/拒绝。结束后恢复用户原来的修订模式。
- **PowerPoint / Excel**：不做审阅层。每轮 Agent 运行前对将被修改的对象做轻量快照（PPT：形状文本/位置/格式；Excel：目标区域值+公式+格式），提供"撤销本轮"按钮；消息流中以变更摘要卡片列出"改了哪些幻灯片/区域"，可点击定位。

## 6. Office 工具集（Tools，参数用 Zod/JSON Schema）
- **Word**：`get_selection`、`get_document_outline`、`read_paragraphs(range)`、`search_text`、`replace_text(target, new, reason)`、`insert_paragraphs(after, content, style)`、`apply_style`、`add_comment`、`list_tracked_changes`、`insert_table`。
- **PowerPoint（MVP 重点）**：`list_slides`、`read_slide(shapes+text+layout)`、`get_theme_and_layouts`、`edit_shape_text`、`add_slide(layout)`、`add_textbox/shape/image/table`、`set_shape_format(位置/字体/颜色)`、`delete_shape`、`duplicate/reorder_slides`、`add_speaker_notes`、`screenshot_slide`（渲染图回传视觉模型自检，Should）。典型场景：按大纲生成整套幻灯、统一字体字号、改写要点、补讲者备注、排版美化。
- **Excel（数据分析 + 图表）**：`get_workbook_summary`（各表维度/表头/类型）、`read_range(as csv)`、`profile_data`（描述统计：缺失、均值、分位数、频数）、`write_range(values|formulas)`、`create_table`、`create_pivot`（待核实 API 支持度）、`create_chart(type, range, title, axis)`、`format_chart`、`add_sheet`、`conditional_format`。分析逻辑以"在表中写公式/新表输出结果"为主，便于复核；复杂统计由模型生成步骤说明。
- 原则：优先结构化细粒度工具（Word 可留痕、PPT/Excel 可还原），**不默认开放任意执行 Office.js**；高级模式可开启沙箱化 `run_officejs`，且执行前同样做快照/开修订。

## 7. Agent Loop 与会话
- 循环：用户输入 → 组装上下文（选区/大纲，长文按需检索）→ 流式调用 → 工具调用 → 经快照/修订层执行 → 回填结果 → 直至完成或达到步数上限（默认 20）；可随时"停止"。
- 会话存储（Dexie/IndexedDB）：`sessions{id, docId, host, title, model, createdAt}`、`messages{sessionId, role, content, toolCalls, usage}`、`changesets{sessionId, turn, host, snapshot, status}`。`docId` 用文档 URL/设置存储（`Office.context.document.settings`）写入的 UUID。
- 历史页：按文档/日期分组、搜索、重命名、导出、清空。

## 8. 安全与隐私
- Key 存 localStorage/IndexedDB（WebCrypto 加密 + 本机派生密钥，主要防止明文浏览，非绝对安全，UI 如实告知）；不写日志、不进错误上报；可一键清除。
- 严格 CSP（内容安全策略，Content Security Policy）：仅允许 Office.js CDN 与厂商域名；Markdown 渲染做 XSS 过滤（DOMPurify）。
- 数据流透明：设置页说明"文档内容直接从本机发往你所选厂商"。
- 上架所需：隐私政策页、支持页、条款页（托管在 docs/）。

## 9. UI / 交互设计
- 布局（任务窗格 ~350px）：顶部 = 模型选择胶囊（厂商图标+模型名）+ 历史 + 设置；中部 = 消息流（工具调用折叠卡片、变更摘要卡片（Word：接受/拒绝/全部；PPT/Excel：撤销本轮））；底部 = 输入框 + 快捷指令芯片 + 附件 + 停止按钮 + "上下文：选区/全文"切换。
- 首次引导（Onboarding）3 步：选厂商 → 填 Key 并测试 → 试一个示例指令。
- 视觉：Fluent v9 令牌、跟随 Office 主题、圆角卡片、克制配色；中文字体回落微软雅黑 / PingFang。
- 参考 Claude 官方 Office 加载项的交互理念（对话+预览+确认），**不复刻其视觉与品牌**。

## 10. 合规与上架（Marketplace）
- 名称/图标不含微软与第三方商标；描述中提到厂商名仅作"兼容"说明。
- 需要 API Key 的加载项须在描述中写明，并提供审核用测试说明（审核员可用的说明或演示视频；待核实是否需提供测试 Key）。
- 需通过 Office 加载项验证（`office-addin-manifest validate`），支持 Office on the web + 至少一个桌面平台。
- 上架流程走 Partner Center；需准备截图、视频、隐私政策 URL、支持 URL。

## 11. 里程碑
| 阶段 | 内容 | 产出 |
|---|---|---|
| M0（1 周） | 命名核查、脚手架、清单侧载、7 家 CORS 直连探针 | 可侧载空面板 + 探针报告 |
| M1（2–3 周） | 厂商层 + 设置页（拉模型/测试）+ Agent loop + 会话历史 + **PowerPoint 工具集 + 撤销快照** | PPT MVP |
| M2（2 周） | Word 工具集 + 原生修订（接受/拒绝面板） | Word 可用 |
| M3（2 周） | Excel 数据分析 + 图表工具集 | Excel 可用 |
| M4（1–2 周） | 快捷指令、用量统计、i18n、可访问性、隐私页、上架材料 | 提交 Partner Center |

## 12. 验证（Verification）
- 单元测试（Vitest）：各协议适配器的流式解析与工具调用归一化（用录制的响应夹具）。
- 侧载测试：Mac 桌面版 Word/Excel/PPT + Office on the web（Edge/Chrome），`npm run start` + `office-addin-debugging`。
- 厂商矩阵：7 家 × {拉模型, 测试连接, 一次流式工具调用} 全部通过或明确标注"需中转"。
- 审阅验收：Word 中 AI 修改出现在原生"审阅"窗格且可接受/拒绝，结束后修订模式复原；PPT/Excel "撤销本轮"后内容与快照逐项一致。
- 场景验收：PPT 由大纲生成 ≥10 页并统一格式；Excel 用示例数据集完成描述统计 + 2 种图表。
- 清单校验：`npx office-addin-manifest validate manifest.prod.xml`。

## 13. 待核实清单
- Anthropic 浏览器直连 header 官方文档；各家实际 CORS 行为（M0 探针确定）。
- Word 修订作者能否显示为 "AI"；各平台 WordApi 1.6 支持度（Mac / Web）。
- 统一 JSON 清单对 PPT / Mac 的当前支持度。
- 候选名商标与重名检索。
