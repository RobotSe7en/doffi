# doffi

面向人和宿主智能体协作的通用办公文档 SDK 与可嵌入工作台。采用 TypeScript 和带语义的 HTML 文档格式，提供 Office 导入、保留样式的网页编辑、框选批注、结构化修改提案和可编辑 Office 导出。达智平台是首个宿主，SDK 同时面向其他应用独立发布。

五个 SDK 包使用 `@doffi/*`，命令行为 `doffi`；文档包和协议标识统一使用 `doffi`，详见[本地开发](docs/development.md)。

**许可：AGPL-3.0-only 或另行取得的商业许可。** 版权所有 © 2026 RobotSe7en。公开版本按 [AGPL 第 3 版](LICENSE)使用；商业授权联系及边界见[商业授权说明](COMMERCIAL-LICENSE.md)。第三方内容保留各自许可。

**当前状态：独立 Word 与 PPT 闭环已实现；第三轮已完成[工作簿模型及 XLSX 只读验收](docs/workbook-reading-acceptance.md)，包含样式、表、规则、图表、绘图和两类批注；Excel 编辑计算实施中，原生输出仍属后续阶段。** 三格式模型设计已明确，`core` 提供严格契约、17 份 JSON Schema、Word 语义及文档包；`office` 提供有界 ZIP 和原生 DOCX 导入／导出。`editor` 提供无脚本预览、文字／对象／矩形选择、批注面板与重定位，见[预览与交互](docs/preview-interaction.md)。Playground 可打开 DOCX、添加或管理批注、缩放、撤销重做、下载文档包并重开，包含中文字体替代和缺字诊断。读取验收使用实际 DOCX、LibreOffice 重存文件与固定字体参考图，范围和限制见 [DOCX 读取](docs/docx-import.md)。隔离候选、七类内容操作、批注事务、单调修订和精确锚点映射见[核心事务与历史](docs/document-session.md)。工作台支持多批注请求、固定提案导入、候选差异和整批应用，见[工作台与协议](docs/workbench.md)。连续正文与表格支持富文本及中文输入，见[正文输入](docs/rich-text.md)。`render` 提供独立只读预览和共享页面画布，`node` 提供受管理的 Chromium、PDF/PNG 与字节协议，见[分页与渲染](docs/pagination-rendering.md)。预览直接显示 HTML，分页覆盖旧 Chromium、Firefox 和 WebKit。原生可编辑导出和批注报告见 [DOCX 导出](docs/docx-export.md)。Playground 原生 DOCX 下载、导出报告和独立打包运行验收见 [Word 闭环](docs/word-closure.md)。R1-15 达智智能体验证按用户决定延期，R1-16 保留尚未完成的宿主验收项；第二轮 PPT 已完成 R2-01 计划、[R2-02 Presentation 契约与文档模型](docs/presentation-model.md)和 [R2-03 独立 PPTX 样例与参考环境](examples/pptx/README.md)，支持 PPT HTML 包的新建、校验、查询与保存重开；[R2-04/05 PPTX 读取](docs/pptx-import.md)已实现，包括三类图表及常量数据、备注和经典／现代批注；[R2-06 只读画布与读取验收](docs/presentation-preview.md)已实现；[R2-07 核心事务](docs/presentation-transactions.md)支持 PPT 几何、尺寸、授权资源及按页锚点映射，[R2-08 至 R2-13 编辑闭环](docs/presentation-editing.md)提供第一方对象、文字、表格、备注、主题／母版／布局、基础图表和两类批注编辑，以及 [PPT 工作台](docs/presentation-workbench.md)。[R2-14 至 R2-16 原生导出与独立 PPT 闭环](docs/pptx-export.md)已实现，包括可编辑 PPTX、图表常量工作簿、两类批注报告、Node PDF/PNG、Playground 和独立 JS/React/Node 打包验收。外部 Office 验收使用固定的 LibreOffice 25.8.1.1；PowerPoint/WPS 未实机验证。尚未发布 npm 包。

## 本地开发

准备 Node.js 24 和 pnpm 10.28.2：

```bash
pnpm install --frozen-lockfile
pnpm dev
```

构建与工程测试使用 `pnpm check`，本地打包使用 `pnpm pack:packages`。浏览器安装、烟测和命令说明见[本地开发](docs/development.md)。

## 已确定的方向

- 独立维护，不依赖 ONLYOFFICE，不将 ibisOffice 源码作为当前实现基础。
- DOCX/PPTX/XLSX 的 OOXML 读写和样式映射自行实现，不使用 docx、PptxGenJS、ExcelJS；ZIP、XML、哈希等通用基础能力可复用成熟库。
- HTML、CSS、资源和格式元数据共同组成可保存的文档源；浏览器 DOM 是其编辑视图。
- Word、PPT、Excel 共用文档身份、事务、批注、修改提案和版本协议，各自使用适合的布局和编辑适配器。
- 人工操作和模型操作均经过同一事务层。模型能够改文字、HTML 区块、样式、图片及布局。
- TypeScript 实现第一版。格式处理保留适配接口，后续可引入 Rust 或 WASM。
- 模型服务、授权、凭据、文件存储和正式版本提交由宿主提供。
- SDK 提供预览、文字划选、矩形框选、批注、差异、应用和撤销界面；宿主复用自己的页面容器与智能体入口。
- 多条批注通过通用文档交互协议统一交给宿主智能体。SDK 不设独立 `ai` 包、模型调用接口或文档内 AI 对话弹窗。
- 第一版以单用户与模型协作为目标；多人实时共同编辑属于后续版本。

## 设计入口

| 文档 | 内容 |
| --- | --- |
| [总体架构](docs/architecture.md) | 产品边界、交付形态、模块职责、依赖方向、状态模型及运行位置 |
| [三格式数据模型](docs/data-model.md) | 唯一持久化来源、公共基座及 Word/PPT/Excel 语义 |
| [文档格式](docs/document-format.md) | HTML 文档包、三种格式的表达和版本规则 |
| [编辑与宿主交互协议](docs/editing-protocol.md) | 事务、锚点、批注请求、修改提案、事件及冲突 |
| [结构契约与 JSON Schema](docs/contracts.md) | 已实现的校验 API、严格字段、类型、静态 Schema 与语义校验边界 |
| [文档包读取与模型](docs/document-model.md) | 只读打开、节点查询、HTML/CSS 支持范围、引用、预算和保存重开 |
| [Presentation 模型](docs/presentation-model.md) | PPT 严格契约、新建文稿、源几何与引用、完整样例和保存重开 |
| [PPT 编辑](docs/presentation-editing.md) | 幻灯片、对象、输入、定义、图表、批注与统一历史 |
| [PPT 工作台](docs/presentation-workbench.md) | 页面／全局请求、只读依赖、候选资源、差异和宿主提案 |
| [PPT 只读画布](docs/presentation-preview.md) | 固定页面、母版／布局投影、形状与图表、字体、参考图及生命周期验收 |
| [核心事务与历史](docs/document-session.md) | 编辑会话、操作、候选、修订、历史、事件及预算 |
| [正文富文本输入](docs/rich-text.md) | 自研输入适配、中文输入法、身份保留、跨段落与表格输入及统一历史 |
| [工作台与协议](docs/workbench.md) | 多批注请求、宿主提交、提案校验、差异应用及 React 包装 |
| [预览与交互](docs/preview-interaction.md) | 普通 JS 挂载、文字与对象选择、矩形坐标、批注面板、布局失效与生命周期 |
| [分页与渲染](docs/pagination-rendering.md) | 纯预览、页面画布、字段、Node PDF/PNG、CLI、浏览器与预算边界 |
| [中文字体](docs/fonts.md) | 常用宋体、仿宋、楷体名称保留、开源字体环境与预览替代 |
| [PPTX 原生导出与闭环](docs/pptx-export.md) | 当前 PPT 源、可编辑部件、两类批注报告、Node、Playground 与独立打包验收 |
| [DOCX 原生导出](docs/docx-export.md) | 当前文档源、样式、字段、批注状态、导出报告及实际重开 |
| [DOCX 读取与只读预览](docs/docx-import.md) | 导入 API、范围、中文字体、样例、参考渲染和限制 |
| [资源与 ZIP](docs/resources-and-zip.md) | 资源摘要、图片容器识别、宿主资源收集、ZIP 预算及保真样例 |
| [整体实施计划](docs/implementation-plan.md) | 五轮交付顺序、具体产物、能力范围和验收门槛 |
| [第一轮实施细则](docs/iteration-1.md) | Word 支持范围、任务依赖、独立闭环及延期的宿主验收 |
| [第二轮实施细则](docs/iteration-2.md) | PPT 已定范围、自研路线、任务依赖及读取／编辑／导出验收门槛 |
| [第三轮实施细则](docs/iteration-3.md) | Excel、自研公式、一万行目标、智能表、绘图、批注与分阶段验收 |
| [独立 XLSX 样例](examples/xlsx/README.md) | 16 个正例、17 个负例、固定 Calc 重存、20 页打印参考与外部差异 |
| [本地开发](docs/development.md) | 环境、启动、构建、类型检查、工程测试及本地打包 |
| [达智集成](docs/dagent-integration.md) | 平台接入、身份隔离、工作区及运行边界 |
| [技术选型](docs/technology-selection.md) | 候选依赖、复用范围及参考工程 |

实现前阅读总体架构、文档格式与交互协议，再按整体计划进入对应迭代细则。后续开发遵循 [AGENTS.md](AGENTS.md)。

## 最小示例

[examples/report](examples/report/) 是一份 `document` 类型的文档包目录，包含中文报告、表格和两条批注。[report.request.json](examples/report.request.json) 表示统一交给宿主的请求，[report.proposal.json](examples/report.proposal.json) 表示宿主智能体针对该请求返回的修改提案。三者共享文档身份和基准修订。报告包已能通过 `openDocument` 读取、查询和保存重开，示例提案中的内容操作已能通过 `prepare/commit` 执行；请求生成和关联提案的 `prepareProposal` 已实现，内容与有效的已处理批注一次提交，旧修订提案明确拒绝。

```text
Office 文件
    ↕ 导入导出适配器
HTML 文档包
    ↕ 统一事务层
浏览器编辑、框选批注、修改提案、差异查看和撤销
    ↕ 文档交互请求与提案
宿主现有智能体、对话与执行链路
```

## 计划中的交付方式

使用一个 TypeScript monorepo，组织 `@doffi/core`、`@doffi/render`、`@doffi/editor`、`@doffi/office`、`@doffi/node` 五个库模块，并提供独立 Playground。编辑器提供普通 JavaScript 挂载、可选 React 包装和底层 API；其他产品主要适配批注提交、提案回传、授权资源和存储。Node 转换进程用于服务端导出和浏览器渲染，具体运行方式由宿主选择。自有内容采用 AGPL-3.0-only 与另行商业授权的模式；正式版本在初次发布时确定，发布前确认 npm scope 归属。

第一轮的三格式模型设计、DOCX 读取验收、事务与历史、富文本、分页、批注、提案、DOCX/PDF 导出与重开已完成独立闭环。R1-15 达智真实智能体验证延期，保留原编号及未通过的宿主验收项。第二轮独立 PPT 已完成。当前按[第三轮细则](docs/iteration-3.md)推进 Excel，依次完成模型与独立样例、完整读取与只读验收、统一事务和自研计算编辑、原生输出与独立闭环；随后为正式宿主存储和发布准备。Playground 使用固定提案文件，不提供独立模型服务。

第一版正式交付覆盖 DOCX、PPTX 和 XLSX 的常见办公内容。WPS 生成的这三种文件走同一条路径；`.wps`、`.et`、`.dps` 通过可选格式转换适配器导入，第一版统一导出标准 OOXML 文件。
