# 本地开发与工程检查

当前完成 R1-02 至 R1-09、R1-05A 模型设计与 R1-10 声明范围的 DOCX 读取验收。`core` 提供严格契约、Word 语义和文档包，`office` 提供 ZIP 与 DOCX 导入，`editor` 提供[预览、选择与批注交互](preview-interaction.md)；`render` 提供共享画布，`node` 提供 PDF/PNG 和字节转换入口。`core` 还提供[编辑会话、事务及历史](document-session.md)、精确文字及结构锚点映射和区域批注重定位；[工作台与协议闭环](workbench.md)支持多批注请求、提案差异和整批应用；[R1-11 正文输入](rich-text.md)也已实现；R1-12 [分页与渲染](pagination-rendering.md)已实现，R1-13 [原生 DOCX 导出](docx-export.md)与 R1-14 [独立 Word 闭环](word-closure.md)也已实现。API 见[DOCX 读取](docx-import.md)，第一轮历史及延期验收见[第一轮细则](iteration-1.md)。当前[第二轮 PPT](iteration-2.md)已实现文档包模型、原生 PPTX 读取、固定画布和 [PPT 编辑协作](presentation-editing.md)，[原生 PPTX、Node 输出与独立闭环](pptx-export.md)也已实现。

## 环境和安装

开发基线为 Node.js 24 和 pnpm 10.28.2，当前验证环境为 Node.js 24.13.0。`.node-version` 和 `packageManager` 记录开发版本，`pnpm-lock.yaml` 固定依赖解析结果。

```bash
pnpm install --frozen-lockfile
pnpm dev
```

`pnpm dev` 先构建库，再同时运行库编译监听与 Vite。开发服务器监听 `0.0.0.0`，可访问终端显示的本机或局域网地址；停止命令时一并结束监听进程。

Playground 通过公开入口打开 DOCX/文档包、选择文字／对象／区域、管理批注、缩放、撤销重做、下载当前会话的文档包及重开。开发宿主将公开合成样例作为静态资源提供；SDK 不自行访问样例目录。它使用 React 生命周期挂载同一 DOM 组件，Office 功能区、字体字段、图标与作用域样式也来自 SDK 的可选 `@doffi/editor/react` 入口；宿主接入文档状态与授权回调。SDK 的 `core`、`office` 和 `node` 不依赖 React，普通 JS 编辑入口也不要求 React。

项目名称为 `doffi`，内部 workspace 使用 `@doffi/core`、`@doffi/editor`、`@doffi/office`、`@doffi/render` 和 `@doffi/node`；Playground 为 `@doffi/playground`，Node CLI 为 `doffi`。版本均为 `0.0.0`，包均设为 `private`；公开许可字段为 `AGPL-3.0-only`，版权署名为 RobotSe7en，也可另行洽谈[商业授权](../COMMERCIAL-LICENSE.md)。npm scope 归属和正式发布仍按发布计划确认。

运行时环境变量统一使用 `DOFFI_*` 前缀，验收使用 `DOFFI_SOURCE_HEAD` 和 `DOFFI_WORKBOOK_*`。Playground 转换路由为 `/__doffi/convert`，请求头为 `X-Doffi-Dev: 1`，浏览器渲染桥为 `window.__doffiRender`，显示设置使用 `doffi:layout`，Word 编辑模式偏好使用 `doffi:editing`，文档列表开关使用 `doffi:library`，标尺开关使用 `doffi:rulers`。

V1 文档包使用 `format: "doffi"`、`urn:doffi:contract:*` Schema 标识及 `doffi/workbook-read/1` Worker 协议标识。本次更名按用户决定不保留旧名称的兼容入口。独立原生样例、实际字体环境路径和历史验收证据保留原始内容。

## 构建与类型检查

| 命令 | 行为 |
| --- | --- |
| `pnpm build:packages` | 按项目引用构建五个库及渲染 IIFE，输出 JS、声明、映射和 15 份 JSON Schema |
| `pnpm typecheck` | 检查库、Playground、工具配置和测试类型 |
| `pnpm build` | 构建库和生产 Playground |
| `pnpm check:boundaries` | 检查允许的依赖方向、公开入口及浏览器模块边界 |
| `pnpm check` | 包边界、一次库构建、其余类型检查、工程测试及生产 Playground 构建 |

库的 TypeScript 项目引用需要声明输出，所以 `typecheck` 会先构建库；Playground 和工具类型检查使用 `noEmit`。生成物进入各包 `dist`，增量元数据进入根 `.cache`，均不提交。PR 和主分支普通推送自动运行 `Quick checks`；整套 CI 在显式发起验收时，以锁定的 Node/pnpm 和依赖构建一次，`CI / check` 下载本次构建产物并执行边界、`typecheck:built` 和 `test:built`。本地 `pnpm check` 默认执行完整模式；触发方式和阶段容量分流见 [CI 执行](ci-performance.md)。

库使用 NodeNext ESM 解析；相对 TypeScript 源码导入写对应的 `.js` 扩展名。跨包使用 workspace 包名和公开 `exports`，不配置指向其他包源码的路径别名。

核心与格式库的类型环境仅包含 ES2022，无 DOM 和自动加载的 Node 类型；编辑器增加 DOM，Node 入口单独启用 Node 类型。工具和测试不参与库发布。

核心通过 Zod 推导并生成普通声明，关闭该包的 `isolatedDeclarations` 限制，仍保持 `strict`、声明输出及 `skipLibCheck: false`。避免为满足隔离声明而手工维护第二份数据类型。静态 Schema 从编译后的同一契约生成；开发监听更新 JS 和声明，修改契约后运行 `pnpm build:packages` 刷新静态 JSON。

## 工程测试

```bash
pnpm test
pnpm test:watch
```

当前 Vitest 覆盖结构契约、JSON Schema、包边界及打包消费。核心用例使用真实报告、请求和提案，并由独立 Ajv 校验器检查导出 Schema 对相同输入的接受/拒绝结果。边界用例检查反向依赖、未声明依赖、内部路径、跨包相对导入、Node 内置模块和不可静态检查的动态导入。

打包用例在 workspace 外、包含空格的临时路径中安装五个 tarball，优先复用工作区 pnpm 依赖缓存；冷缓存时允许获取第三方依赖及其注册表元数据。内部包通过显式 tarball 覆盖固定来源，不从注册表下载内部包。验证真实校验 API、公开 JSON、ES2022 类型、私有路径拒绝，以及不安装 React 的普通 JS 浏览器宿主构建。临时消费工程在测试结束后删除。

测试还覆盖只读文档解析、资源摘要、宿主资源收集、ZIP 原始路径及大小校验、保真样例保存重开；DOCX 用例覆盖实际格式输入、独立生产者版本、样式、字段、批注、修订及保存重开；事务用例另覆盖隔离候选、操作、原子性、历史、冲突和预算，精确 UTF-16/结构映射、跨节点与区域批注，以及 DOCX 原生批注编辑后的历史和文档包重开；正文输入增加身份协调、精确映射证据和中文组合输入；分页与原生 DOCX 导出及独立打包闭环也已有对应验收；PPT 用例覆盖模型、实际 PPTX／独立生产者读取、固定画布、编辑与提案；原生导出、官方 XSD、外部编辑、PDF/PNG、独立打包和源／导出逐页视觉对照均有验收入口。核心用例放入各包 `tests`，浏览器用例放入根 `tests/browser`。

## 浏览器烟测

```bash
pnpm browser:install
pnpm test:browser
```

浏览器安装和测试通过同一包装脚本运行。默认将 Playwright 浏览器保存在仓库的 `.playwright` 缓存，禁用浏览器缓存回收；可用显式的 `PLAYWRIGHT_BROWSERS_PATH` 覆盖路径。

浏览器测试先构建生产 Playground，再启动严格占用本机 4173 端口的临时预览服务，验证 React 挂载、实际 DOCX 读取、文档包下载重开、字体诊断及连续预览截图。服务由 Playwright 管理，测试结束后关闭，不借用已有预览服务。

本机已经安装 Google Chrome 时，可显式选择其渠道：

```bash
DOFFI_BROWSER_CHANNEL=chrome pnpm test:browser
```

未设置渠道时使用对应 Playwright 版本的 Chromium，不进行自动渠道回退。当前截图基线使用锁定的 Chromium 与 Noto 中文字体；补充的仿宋/楷体浏览器用例还需要[固定开源字体环境](fonts.md)。环境和字体摘要见[参考记录](../examples/docx/reference.json)。使用 Chrome 渠道只能作额外验证，不能静默刷新固定基线。独立参考检查使用 `pnpm check:docx-reference`，更新方法见[DOCX 读取](docx-import.md#独立样例与验收)。

仅运行 R1-08 交互验收使用 `node scripts/playwright.mjs test interaction.spec.ts workbench.spec.ts text-editor.spec.ts pagination.spec.ts rendering.spec.ts remote-access.spec.ts word-closure.spec.ts`；它包含普通 JavaScript 宿主及 React Playground。CI 将工程检查与浏览器交互／工作台分开，上传逐用例 JSON 和失败诊断。文档视觉夹具截图时展开滚动容器，继续核对原完整文档基线；交互滚动与缩放另有坐标验收。

第三轮工作簿日常检查在 Chromium、Firefox、WebKit 和固定旧 Chromium 上执行网格、绘图、公共预览、Worker 与 Playground；可用 `node scripts/ci-browser.mjs --browser chromium --suite workbook --mode daily` 运行单片，结果保存在 `.cache/ci-browser/<browser>/workbook/`，上传为 `browser-<browser>-workbook`。一万行用例保留 20 万实际单元格、2 万公式、取消后旧窗口及传输大小报告，改在阶段专项运行。已完成的读取范围和固定资源证据见[读取验收](workbook-reading-acceptance.md)。

浏览器按四引擎及 Word／PPT／工作簿拆分，每片 `workers:1`，保留固定原生页面比较。日常约 5 分钟为优化目标，实际耗时以工作流记录为准。阶段专项独占执行 4 vCPU／8 GiB／无交换验收；首屏 30 秒、选择 p95 100 ms、读取滚轮 p95 120 ms 及内存目标保持，不能用增加日常作业并发代替这些资源条件。

## 本地打包

```bash
pnpm pack:packages
```

五个 tarball 输出到 `.cache/packages`。pnpm 会把内部 `workspace:*` 依赖转换为包版本，`files` 包含构建目录和许可声明；包的公开入口指向构建结果，不暴露 `src` 或 `dist` 内部路径。

各库的 `prepack` 自动从仓库根目录复制 `LICENSE`、`NOTICE` 和 `COMMERCIAL-LICENSE.md`，三份文件随每个 tarball 分发；本地副本由 Git 忽略，根文件为唯一维护来源。`core` 和 `render` 另外保留既有 ECMA-376 许可。打包消费检查核对安装后的许可标识与文件实际内容。发布时还应按所选许可提供与发行物匹配的对应源码及构建材料。

使用 `pack:packages` 避免与 pnpm 自带的 `pack` 命令混淆。此命令只生成本地文件，不发布包或推送仓库。

## 依赖与状态

已安装 Zod、HTML/CSS 解析器、哈希、ZIP、工具及依赖的版本、许可证见[技术选型](technology-selection.md)。XML 已通过内部 saxes 适配器接入，R1-11 使用自研浏览器输入适配，其他格式的编辑适配仍按后续任务实现；Office OOXML 读写和样式映射自行实现，不提前提供占位功能 API。

修改后运行相应的窄范围检查与 `git diff --check`。只有验证所覆盖的能力才可在 README 标为已实现，产品范围和宿主边界以[总体架构](architecture.md)及[交互协议](editing-protocol.md)为准。

R1-12 的浏览器 CI 使用 Ubuntu 24.04、固定 Playwright 引擎、Poppler 和宿主 Noto 字体；另校验官方归档后安装 Chromium 109，并在 macOS 15 使用真实 SafariDriver。Safari 输出实际版本及验收结果，不能用 Linux WebKit 结果冒充 Safari。PDF 开发中间件默认只接受本机连接；远程访问可将 Vite 绑定 `0.0.0.0` 并显式配置确切同源开发地址，见 [Word 闭环](word-closure.md#远程开发访问)。正式转换认证由宿主管理。

PPT 的独立 ZIP/XML 样例与参考环境见 [PPTX 样例](../examples/pptx/README.md)。`pnpm check:pptx-fixtures` 不需要 Office；`pnpm check:pptx-reference` 需要 Impress、Calc、Poppler 和上述字体，会实际打开、重存并重开。环境不同只运行内容检查并明确记录未比较像素，不自动刷新基线。

Office CI 也执行 `pnpm check:xlsx-fixtures` 与 `pnpm check:xlsx-reference`，核对 [XLSX 独立来源](../examples/xlsx/README.md)的可重现字节、官方结构、负例和固定 Calc 重开。源及重存后的结构差异始终对照已记录预期；只有渲染环境完全相同时才比较页数与像素摘要。日常省略一万行重开，专项运行完整来源；结果分别上传为 `office-workbook` 和 `workbook-scale-acceptance`，环境不同会明确标为未比较像素，不更新参考文件。

PPT 导出与独立闭环使用 `pnpm check:pptx-export` 和 `pnpm check:presentation-closure`，需要 `python3-uno`、`python3-lxml`、Impress/Calc、Poppler 和既定字体。普通 JS／React／Node 在工作区外安装 tarball；输出及环境见 `.cache/pptx-export-acceptance/` 和 `.cache/presentation-closure-acceptance/`。完整范围、现代批注限制及 PowerPoint/WPS 未验证项见 [PPTX 导出](pptx-export.md)。

Office CI 固定使用官方 LibreOffice 25.8.1.1 归档，`node scripts/install-acceptance-office.mjs` 校验官方 SHA-256 后仅解包到 `.cache/acceptance-office/`，不替换系统 Office。CI 将其 `bin` 加入当前任务 PATH，并保留版本、来源和摘要；本地 Ubuntu 24.04 x64 在装齐上述原生依赖后可用同一命令，再把该 `bin` 加入 PATH；其他系统可使用已验证的本机版本。归档随附的许可保留在解包目录，工具不会进入 SDK 发布包。

Office CI 另显式运行 `node scripts/install-acceptance-fonts.mjs`，从固定 Debian 归档解出与参考相同的 Noto Sans／Serif CJK Regular/Bold 和 Color Emoji，仅保存到 `.cache/acceptance-fonts/`。归档及逐字体 SHA-256 都要通过，来源、版本与许可见[验收字体清单](../scripts/fixtures/acceptance-fonts.json)。生成的 Fontconfig 配置在当前作业内替换这些字体的解析，系统文件不变；本地使用时显式设置脚本输出的 `FONTCONFIG_FILE`。CI 会通过 `GITHUB_ENV` 传给后续步骤，日志／配置／许可保留在该缓存目录。此为开发验收设置，SDK 不下载字体。

LibreOffice 24.2 在含未使用自定义布局的新建文稿重存中会生成重复原生对象 ID；此结果在隔离 Ubuntu 24.04 中复现，SDK 继续明确拒绝，不放宽重复身份校验。本轮外部打开、实际编辑和重存验收以固定的 25.8.1.1 为准，其他 Office 版本需单独验证。


## CI 运行方式

整体检查、日常容量分流、阶段合并前完整验收、失败追踪和耗时报告见 [CI 执行](ci-performance.md)。
