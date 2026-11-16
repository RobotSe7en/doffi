# DOCX 完整功能执行交接

本文件保留自 2026-11-15 起的阶段交接。**先读 [固定清单](docx-complete-progress.md#当前固定未完成清单)：R01～R14 已完成，剩余 R15～R19 共 5 项。** 后附 R03～R08 记录是对应阶段的补充；不要从历史日志重新推导待办，也不要把 D1／D4／D5／D6／D8 五个大类当成只有五项小工作。

## 分支、服务与工作状态

- 仓库：`/home/danny/work/doffi`；工作分支：`feat/docx-complete`；起点：`main@f568638`。
- 原完整功能提交：`233a7b1`（段落结构粘贴及整组修订）。此前 `c97b814` 完成书签／内部引用迁移，`3b058de` 完成字段／公式／控件等行内语义粘贴。
- 原表格剪贴板 WIP 为 `75121ee`，R01 已在本轮收口。纯格式提交为 `57a2a66`：整理分支改动涉及的 303 个 TypeScript／TSX 文件，302 个产生差异，并核对归一化源语法树及编译后 JavaScript。接手仍须先检查 `git status`、`git log -5 --oneline` 和当前 diff；保留现有改动，不重置或切回 main 覆盖它们。
- 服务已启动：<http://localhost:5174/>，交接检查时返回 HTTP 200。启动命令为 `pnpm --filter @doffi/playground dev --port 5174 --strictPort`。服务仍在运行时直接复用；终端会话消失后检查端口再启动。
- D2、D3、声明范围的 D7 编辑功能已完成。剩余实施以计划里的 R01～R19 为准，不能声称已经全面兼容 Word 或通过 Word 实机像素验收。

## 用户要求与已确定选择

- 继续完成原计划；先完成功能，只做必要的相关测试。本轮不做独立审查、A/B、远程全量 CI，不使用子代理。
- 本轮用户明确授权格式整理和 R01 收口的本地提交及推送。**不创建 PR、不合并、不发布**；用户会主动发起 PR 和审查。当前 R08 按用户本轮要求仅作本地阶段提交，不推送。
- 产品决策主动询问；用户未及时回复时，按已经给出的推荐方案继续并记录。以下决定不要重复询问。
- 未建模的正常原生内容允许完整保留 XML／关系／资源，明确占位；不要求为全部特殊对象实现内部编辑器。安全禁止、结构错误和真实资源预算仍明确拒绝。
- 重叠 run 格式修订保留最初格式、最后作者，已实施。
- 普通合并／拆分采用 Word 常规行为：不记录为修订，界面明确提示，仍可撤销；特殊原生合并修订保留，已实施。
- 跨文档粘贴保留来源格式，重名而不同的样式隔离，定义及依赖均相同则复用，已实施。
- 部分字段／控件／表格选区不自动扩张到选区外；完整选中结构才迁移对应语义。重名书签独立命名并同步复制内容里的引用；未复制目标保持未解析，避免误连目标文档的同名书签，已实施。
- 本轮不推进 Excel 后续阶段，不重开原有 Excel 接受条件；AGPL、第一方 OOXML、单一 HTML 源和统一事务等约束继续遵守仓库 AGENTS.md。

## R01 已完成

整表迁移已接通列网格、合并续接格式、条件表格样式、附属 CSS、制表位／条件标记以及行插入修订。复制全表的来源样式冲突会隔离；部分表格改为选中文字及外观。嵌套表格的独立网格、内表继续输入、成组审阅与撤销已验证；缺失内表元数据和过期候选均整批拒绝。临时 `payload.styles`／`payload.nodes`／目标段落属性的失败诊断已清理，并修正了浏览器续接格式断言的可选值类型检查。

本轮最终源的必要验证：

- 库构建、应用／工具类型、架构边界和 `git diff --check` 通过。
- 来源 CSS、行内／区块粘贴、表格操作、表格修订／样式／续接与 DOCX 导出等 **12 文件 106 项通过**；覆盖严格原生重开、接受／拒绝、撤销和 RevisionConflict。
- Chromium／Firefox／WebKit 的 **16 类 48 项通过**；覆盖来源样式冲突、同文档复用、只读复制、字段／书签／段落区块和图文修订回归。三种整表用例覆盖续接段落的实际底纹／边框、嵌套表格输入、审阅面板拒绝和撤销。
- 将真实浏览器导出保存为标准测试附件后，三种整表用例三引擎 **9 项再次通过**，包含两次严格 DOCX 重开。
- **29 份真实 DOCX 输出、87 个 Word XML 部件**经声明的 W14／MCE 预处理后，通过固定官方 Transitional XSD；包括 20 份普通／追踪粘贴、继续编辑、严格重开和接受／拒绝输出，以及 9 份浏览器当前源输出。生成器直接调用真实 fixture 和 SDK，没有 expect stub；实际断言由前述 Vitest／Playwright 承担。
- 日志：`.cache/docx-gap-audit/r01-formatted-build.log`、`r01-final-*-types.log`、`r01-artifacts-tools-types.log`、`r01-final-tests.log`、`r01-final-browser.log`、`r01-native-browser.log`、`r01-schema.log`。`.cache` 与标准浏览器输出目录均被忽略。

原交接的两个已复现问题已通过本轮回归与 XSD 校验：

1. **粘贴后首行条件粗体变普通。** 原来先移除源样式类，再读取 CSS，使低优先级 `font-weight:inherit` 重置被固化为行内覆盖。新增 `SourceStyleResolver.specifiedOutsideClasses`，按原始胜出声明的所属规则排除被迁移样式，保留真实直接覆盖；格式包改用此方法。本轮普通来源格式、样式复用／冲突与三引擎整表回归均通过。
2. **原生拒绝整表插入后多一个空段落。** 导出将透明 div 内的末尾表格当作整个内容结束，额外生成无修订的段落。`Exporter.blocks` 新增 `hasFollowing`，透传后续源块状态，避免在仍有后续段落时补空段落。普通与嵌套整表的 DOCX 重开／拒绝、相关导出回归和固定官方 XSD 均已通过。

## R02 已完成；下一步优先 R03

- 编号剪贴板只收集选中段落和所选整表续接空段落需要的已建模编号定义、级别与 CSS。保留原生声明的起始／重启、模板、后缀及中西文字体角色；每次粘贴独立分配编号与类身份，目标文档的同名定义保持原样。
- `WordContentSemanticsSchema` 新增可选 `numbering`，行内契约仍不接受编号定义。区块迁移重写全部新插入节点及续接段落的编号引用；没有传输节点身份的片段也不能借用目标已有编号。缺失级别／CSS、重复定义和过期候选均整批拒绝。
- 普通 `ul/ol/li`（含嵌套、起始值、项目重启和后续段落）转换为同一编号段落源，保留内部节点身份。倒序、超过九级和不能表示的标记明确拒绝；未建模原生对象及依赖图迁移仍归 R08。
- 已修复追踪粘贴后段末回车所复制的空修订壳：保留格式，避免新段落产生没有定义的 `ins/del`。真实内容和零宽原生结构保留原归属。新增浏览器验证继续输入、下一项实际序号、编辑后拒绝全部修订和逐步撤销。
- 最终源 **12 文件 291 项**核心／契约／JSON Schema／Office 检查通过；Chromium／Firefox／WebKit **17 用例 51 项**通过，其中新增列表、HTML 列表及续接编号在普通／追踪模式共 18 项。**45 份真实 DOCX、180 个 Word XML 部件**经声明的 W14／MCE 预处理后，通过固定官方 Transitional XSD，包含续接单元格拆分、继续编辑、审阅及浏览器当前源输出。构建、应用／工具类型、架构边界和 diff 检查通过。
- 本轮日志为 `.cache/docx-gap-audit/r02-input-build.log`、`r02-acceptance-tests.log`、`r02-final-browser.log`、`r02-schema.log`、`r02-final-*-types.log` 和 `r02-final-boundaries.log`。只按实际最终通过结果收口，前期失败日志保留供追溯。

接下来优先处理 R03 的跨段接收选区，不要继续拆出新的交付阶段或只追加日志而不更新清单。

## 当前代码入口

| 范围 | 文件 |
| --- | --- |
| 迁移契约 | `packages/core/src/contracts/word-content.ts`、`word-operations.ts`、`packages/core/src/index.ts` |
| 语义附加与 CSS 身份 | `packages/core/src/session/word-content.ts` |
| HTML 列表归一化 | `packages/core/src/session/word-list-content.ts`、`packages/core/src/document/edit.ts` |
| 段落／表格结构粘贴 | `packages/core/src/session/word-block-content.ts`、`word-inline-content.ts` |
| 源拆分、包装和段落合并 | `packages/core/src/document/edit.ts` |
| 原始 CSS 声明归属 | `packages/core/src/document/source-css.ts` |
| 嵌套表格行审阅 | `packages/core/src/session/word-row-revisions.ts`；已补外层表格移除后跳过消失的内层表格 |
| 格式及语义复制 | `packages/render/src/word-clipboard-format.ts`、`word-clipboard-semantics.ts`、`projection.ts` |
| 格式恢复及实际粘贴 | `packages/editor/src/rich-text/clipboard-format.ts`、`packages/editor/src/rich-text.ts` |
| 原生编译 | `packages/office/src/docx-export.ts` |
| 当前新增检查 | `packages/core/tests/word-table-content.test.ts`、`packages/office/tests/docx-table-content.test.ts`、`word-table-content-fixture.ts`；浏览器 `tests/browser/pagination.spec.ts` |
| R02 检查 | `packages/core/tests/word-list-content.test.ts`、`packages/office/tests/docx-list-content.test.ts`、`word-list-content-fixture.ts`、`tests/browser/word-list-clipboard-cases.ts`（由 pagination 注册） |

当前实现边界：

- `pasteWordInlineContent` 与 `pasteWordBlockContent` 支持同故事内准确跨段接收选区；区块内容涵盖段落、整表和列表，见已完成的 R03。
- 已有插入／删除／移动修订的准确覆盖、实际剪切与成组移动见已完成的 R04。原生保留区块的追踪粘贴与统一审阅见已完成的 R08；不能表达的原生修订明确列入导出报告。
- `WordContentSemanticsSchema` 包含编号、表格、tabs、conditionalFormats、extraClasses、对象／注释／批注故事与 annotations／comments，以及所选已索引原生端点。行内接收片段与其拥有的区块故事按既有联合契约区分；传输身份只映射到新插入节点，不能解析到目标已有节点。
- 附属 CSS 分配独立名称；表格 `styleContexts` 的旧绑定移除后由既有表格样式刷新逻辑重建，续接段落／run 的直接格式单独保存。
- R02 已完成已建模编号定义与合法 `ul/ol/li` 列表区块迁移；所有编号引用在粘贴候选内隔离，不复用来源 numberingId 误绑到目标定义。
- 系统剪贴板已覆盖 R05／R06 的绘图、图表、注释引用及附属故事，以及 R07 的批注线程和精确锚点。R08 已补齐所选原生图、旧集合切片、未知属性和间接依赖的迁移；跨原生范围配对、导入保留覆盖及旧集合原位回收已由 R09、R10、R11 收口。

## 后续原生迁移必须保持的边界

- 不用“整份来源文档包”实现剪贴板。共享原生 JSON 集合必须只带所选 fragmentKey；样式、故事、原生 XML 和依赖部件也不得夹带未选正文。
- 字段复制的是当前显示缓存及指令，不能恢复未选中的旧隐藏缓存。部分字段仍复制实际选中文字，不扩大选择。
- 原生对象内部批注已有精确控制符绑定；R07 已将 `NativeWordComments` 按 sourceId 与端点节点共同索引，跨来源和重复粘贴不能退回仅以数字 sourceId 为唯一键。
- 同源 `copyWordStorySemantics` 不能直接套用跨文档数据：它允许回退到已有身份、递归克隆现有故事。跨文档必须先隔离身份与资源。
- 自定义 XML 的 `w:storeItemID` 等依赖不一定体现为片段里的 r:id。需要显式依赖索引；不能为了保留绑定而复制无关根部件。
- JSON 与可安全索引的旧 XML 集合均已完成删除回收。core 只解析通用集合／OPC 信封和不透明 UUID 引用，不解释 OOXML 业务语法；无法安全索引的信封仍保守保留。
- 原生修订编号、关系 ID、部件路径和批注编号各自有边界；保留内容恢复时避免与当前源编译生成的身份冲突。

## 必要检查及执行方式

先完成一批关联代码，再构建一次，构建成功后跑相关检查。测试使用 dist，不能在构建未完成时声称最终源通过。不要并行修改同一生产代码与构建，不重复跑未受影响的整套检查。

```bash
pnpm build:packages
pnpm exec vitest run packages/core/tests/source-css.test.ts packages/core/tests/word-table-content.test.ts packages/office/tests/docx-table-content.test.ts
pnpm exec tsc --project apps/playground/tsconfig.json
pnpm exec tsc --project tsconfig.tools.json
pnpm check:boundaries
git diff --check
```

R01 已按改动补跑来源格式、只读复制、字段与段落区块等用例，覆盖同文档样式复用及冲突隔离。后续只复跑受新改动影响的检查，不启动全量仓库测试／远程 CI。

当前机器的浏览器配置：

```bash
PLAYWRIGHT_BROWSERS_PATH=/home/danny/work/doffi/.playwright pnpm exec playwright test --config .cache/docx-gap-audit/playwright.config.ts tests/browser/pagination.spec.ts --project chromium --project firefox --project webkit --grep 'whole-table semantic clipboard'
```

`.cache` 被忽略。若换工作树或缓存不存在，用根 `playwright.config.ts` 派生临时配置：保留其项目，设置 `baseURL=http://127.0.0.1:5174`、关闭自动 webServer、`retries=0`。不要把失败重试当作通过。

原生校验使用固定的 `scripts/fixtures/ooxml/OfficeOpenXML-XMLSchema-Transitional.zip`，SHA-256 为 `d34187520749998af306faf1b730e568b0ca6d88ad24638a407c0a9bb4ca04fc`。当前机器 `.cache/xlsx-venv/bin/python` 安装了 lxml；系统 Python 没有。既有 `.cache/docx-gap-audit/*-native.mjs` 是收集真实导出字节的辅助生成器，部分使用 expect stub，**不是测试**；先通过实际 Vitest，再校验输出。只按文档声明处理 MCE／W14 等可忽略扩展，不能笼统声称所有现代 OOXML 都通过基础 XSD。

浏览器脚本 DataTransfer 检查不等于操作系统剪贴板实机验收。之前用户上传报 ResourceLimitExceeded 的具体文件未保存在服务器；恢复进展和大量片段门槛已经有复现修复，不能声称验证过该用户原文件。

## 完成与汇报

后续只按 R01～R19 更新具体完成状态，写明功能、必要验证和仍存限制；不再只重复“五个方向”。已有历史检查点可追溯，但未完成项以固定清单为准。若发现破坏本轮已声明保留／编辑边界的缺陷，归入对应条目修复；不自行扩大产品范围或要求新的独立审查流程。

## R03 补充交接（2026-11-15）

跨段接收选区已收口，固定清单更新至 R04～R19。`word-content-range.ts` 负责准确同故事范围、跨表格边界删除及成组标记；两个粘贴操作使用可选 `endNodeId`。`TextRangeTransfer` 按文字编辑序号提供准确迁移，锚点与历史共用。必要检查：105 项事务／锚点／DOCX、12 项三引擎浏览器、12 DOCX／36 XML 固定 XSD、构建及类型／边界通过。用户当前授权继续完成 R04～R06；按项本地提交，不推送、不建 PR、不安排独立审查。

## R04 补充交接（2026-11-15）

`word-content-revisions.ts` 分段删除既有修订；`word-content-move.ts` 提供准确剪切、成对移动范围和锚点迁移。编辑器以实例私有令牌／HTML／修订绑定系统剪切，已有待定插入被取消时保持普通粘贴。原生行修订继承显式移动范围，空表也可成组审阅。9 文件 127 项、三引擎 24 项、16 DOCX／48 XML 固定 XSD 及类型／边界通过。继续 R05、R06，不推送。

## R05 补充交接（2026-11-16）

固定清单 R01～R05 完成，继续 R06。剪贴板 transport 的 `stories` 仅保存传输片段；`attachWordContent` 把它们分配新身份并写入唯一源树。`word-clipboard-stories.ts` 从所选对象遍历所有权，`clipboard-objects.ts` 处理对象面板选择和坐标容器，`clipboard-resources.ts` 暂存选中图表资源。完整原生对象集合的 `cutWordDrawings` 与 R04 移动审阅相通。

图表剪贴板丢弃原始嵌入书的无关单元格，保存格式模板；`docx-chart-patch.ts` 在重建工作簿时同步公式和 externalData。`pruneRemovedWordChartResources` 回收删除后的模板依赖。浮动图形分页使用布局盒，不把旋转后的绘制边界当作文字行。

71 项相关事务／原生、三引擎 12 项和旧浏览器回归 32 项通过；固定 XSD 加所用 WPS／WPG profile 验证 12 DOCX／54 XML。源与 profile 哈希、验收导出和脚本在忽略目录 `.cache/docx-gap-audit/r05-*`。新增对象 fixture 含无关工作簿单元格标记，验证实际剪贴板不带该内容。构建及类型／边界通过。用户在 R05 实施中请求推送，R03／R04 已推至 `origin/feat/docx-complete@ef4a401`；后续提交按当前推送授权完成交付，不建 PR。

## R06 补充交接（2026-11-16）

用户本次要求的 R03～R06 全部完成；固定清单剩余 R07～R19。`WordContentSemantics` 增加 notes／noteReferences，传输故事种类扩展为 textBox／footnote／endnote。`word-clipboard-stories.ts` 只从选中引用／对象遍历所有权；`attachWordContent` 为全部节点和注释／故事分配新身份，检查每个引用与故事的归属，拒绝未选附属内容兜底。

注释正文包含字段、缓存、链接、表格、编号、制表位、图片及对象文字故事。新注释同候选拒绝的清理不再依赖 base 已有引用；注释编号是输入原子边界，旧源中编号节点内的字段仍可按序原生输出。区块粘贴的段落标记修订加入唯一组标识，修复重复粘贴冲突。

156 项相关测试（R06 为 13 项）、三引擎 12 项、浏览器受影响回归 38 项通过。12 DOCX／66 XML 通过固定 Transitional XSD 与所用 WPS／WPG profile；导出、脚本和日志在 `.cache/docx-gap-audit/r06-*`。构建、应用／工具类型、边界与分支 327 个 TS／TSX 的格式检查通过。按用户当前推送授权交付 `feat/docx-complete` 的本地阶段提交；不创建 PR，不启动远程全量 CI。后续如继续，从固定 R07 开始。


## R07 补充交接（2026-11-16）

用户本轮要求的 R07 已完成，固定清单剩余 R08～R19。`word-clipboard-comments.ts` 从实际选区源位置裁剪批注；`clipboard-node-text.ts` 共享源文字计数，`SourceProjection.clipboardSourceRanges` 排除派生内容并保留隐藏修订之前的真实偏移。消息正文加入原有所有权遍历，`word-content-comments.ts` 在候选里分配线程／回复身份，检查正文与回复图，按粘贴检查点映射后续事务；只有完整锚点吻合真实移动证据时复用原线程。

已索引 JSON 原生原子端点通过 `word-clipboard-native.ts` 和 `word-content-native.ts` 传递所选片段及声明关系；资源接收仍走授权字节与新路径。`NativeWordComments` 以数字来源 ID 和实际端点节点共同选择绑定；原生片段仅包含范围一端或只有引用时，补齐／抑制生成控制符的规则分别处理，不丢掉缺失的另一端。字段完整选区按显示替换段映射；旧缓存内部字符没有精确对应时，线程保留为 `needs_reanchor`，不夹带未选旧缓存或猜测偏移。

R08 的通用未知属性／旧集合／间接依赖图仍待完成；未建模原生区块的追踪粘贴明确拒绝，普通原生区块端点复制已验证。不要把本项基本端点迁移当成 R08 完成，也不要把已完成的批注迁移重复列回待办。

必要检查：16 文件 **413 项**事务／契约／锚点／原生检查、三引擎 **30 项**新用例及 **37 项**受影响 Chromium 回归通过。**24 DOCX／102 XML**通过固定 Transitional XSD 与既有 WPS／WPG profile（只按声明处理 W14／MCE）；现代线程／嵌套回复另由原生重开验证。构建、应用／工具类型、边界、格式和 diff 检查通过。日志、脚本与实际浏览器输出位于忽略目录 `.cache/docx-gap-audit/r07-*`。本轮按用户当前要求仅作 R07 本地阶段提交，不推送、不创建 PR、不安排独立审查或远程全量 CI。


## R08 补充交接（2026-11-16）

用户本轮要求的 R08 已完成，固定清单剩余 R09～R19 共 11 项。`word-clipboard-native.ts` 与 `word-clipboard-native-xml.ts` 选择节点／样式残余和必要原生定义，只序列化选中的 JSON／旧 XML 条目与 OPC 依赖。`word-clipboard-native-style.ts` 的原生属性签名参与样式一致性比较。`word-content-native.ts` 检查依赖闭包与所有权，隔离传输身份；资源检查与删除回收同时遍历显式间接依赖和原生编号所有权。

`docx-native-custom-xml.ts` 在 Office 适配器索引数据存储项与属性部件；复制产生新数据存储 GUID，导出重写片段与属性部件，不会误绑定接收方同名项。片段的 `styleReferences`／`numberingReferences` 和 `numberingKind` 定义指向所需样式及原生编号；导出分配独立编号并规范已知级别子元素顺序。绘图、控件及既有原生范围数字身份沿原生作用域隔离。

`nativeInsert` 与普通粘贴共用事务／分组／撤销。`docx-native-revisions.ts` 只对可证明完整归属的段落／文字／表格行生成或恢复整块插入；混合未追踪内容与不同归属不产生猜测逆操作。特殊区块保留内容，并在 `report.unexportedNativeRevisionIds` 报告无法完整输出的原生修订；完整文档包仍保留该审阅。原子内部仍不是普通文字编辑目标。

最终源 21 文件 **403 项**、三引擎 **27 项**新用例及 **43 项**受影响 Chromium 回归通过。**27 DOCX／96 XML** 通过固定官方 Transitional XSD 与其 customXml XSD，仅处理已声明的可忽略扩展和 MCE 标准回退；原始 Choice／Fallback 字节另经复制及多次原生重开验证。构建、应用／工具类型、边界、格式和 diff 检查通过。证据位于 `.cache/docx-gap-audit/r08-*`。只作 R08 本地提交，不推送、不建 PR、不安排独立审查或远程全量 CI。

继续从固定 R09 开始。R11 的旧 XML 集合原位删除回收仍未关闭；R08 的旧集合剪贴板切片不能替代该项。原始 R07 章节的 R08 待办属于历史交接，不再是当前状态。

## R09 已完成

以固定清单和执行记录末尾 R09 为准。原生书签／权限／移动／customXml 范围现在由 Office 索引配对与故事归属；外部边界为零宽源，删除任一端、部分复制和原生回写不会留下孤立边界。特殊跨原生移动继续保留原生内容。6 文件 80 项及三引擎 15 项通过，库构建和工具类型通过；详细日志为 `.cache/docx-gap-audit/r09-*`。继续 R10～R11，仍只做本地阶段提交。四核环境下构建和检查串行执行、测试单 worker，避免同时启动多个 TypeScript 编译。

## R10 已完成

绘图框架和嵌套属性均按实际映射槽检查，`effectExtent`、框架锁定及其他未建模属性保留完整绘图并诊断，严格模式拒绝；已知无效数值仍失败。同类未知属性可批量恢复。当前源 6 文件 82 项、三引擎 6 项以及 12 份 DOCX／36 个 XML 部件的固定 XSD＋声明 WPS 配置通过，构建和工具类型通过。继续 R11；日志为 `.cache/docx-gap-audit/r10-*`。

## R11 已完成

旧 XML 集合现在按条目回收并保留继承命名空间／原始条目字节；UTF-16 旧资源裁剪后使用正确 UTF-8 声明，撤销恢复原始资源。旧引用通过有界通用 XML／OPC 和不透明 UUID 索引恢复，跨对象、独立根、正文图片及完整无键所有权均有区分。core 复用仓库原有的补丁版 saxes；离线 frozen-lockfile 安装已通过，没有引入 Office 转换库。当前源 7 文件 87 项、三引擎 36 项通过；33 份实际浏览器 DOCX／132 个 XML 部件通过固定官方 Transitional XSD。构建、应用／工具类型、边界、格式、Markdown 链接和 diff 检查通过，详见执行记录末尾 R11 与 `.cache/docx-gap-audit/r11-*`。

下一项为 R12／R13 表格布局；固定清单剩 R12～R19 共 8 项。R09／R10／R11 仅作本地阶段提交，不推送。Playground 用单独 Vite 服务维持在 <http://localhost:5174/>，构建／类型与测试继续串行执行，测试单 worker；四核环境避免 `pnpm dev` 同时启动五个独立 TypeScript watcher。


## R12 已完成

行级 `tblPrEx` 的边框／底纹、宽度、缩进／对齐、单元格边距、布局及外观位已接通独立源声明、当前行面板、条件样式、续接／分页显示和 `tblPrExChange` 审阅。修订只恢复例外属性；表格默认边距单独保存。派生行内容框与行间距绘制共用坐标，并验证实际文字命中与截图。必要检查、最终浏览器原生 XSD 见执行记录末尾 R12；日志为 `.cache/docx-gap-audit/r12-*`。

用户当前明确授权按 R12→R15 完成、提交并推送当前 `feat/docx-complete` 分支。继续 R13；不建 PR、不安排独立审查／子代理或远程全量 CI。单独 Vite 服务已重新启动在 <http://localhost:5174/>；四核环境继续串行构建／类型和单 worker 测试。


## R13 已完成

合法空／零宽／动态扩展网格现在通过 `docx-table-grid.ts` 推导为可编辑的正尺寸网格，计入真实合并／省略区和节／栏／嵌套容器。自动布局、首选宽度、拖动、合并拆分与分页均沿唯一源、事务和历史；auto／nil 原生宽度类型保留。面板新增布局算法与自动／磅／百分比宽度选择。相关事务、预算、两次严格重开、三引擎实际拖动及 60 行组合已验证，详见执行记录末尾 R13 与 `.cache/docx-gap-audit/r13-*`。

继续 R14 的合并／省略区域列结构修订，再做 R15；本轮全部完成后提交推送当前分支，不建 PR。


## R14 已完成

横向／纵向合并及首尾省略区域内部、无新增实际单元格的追踪插列已接通事务、原始视图和原生重开。`resolveWordGridInsertion` 只处理有界通用列约束；Office 适配器提供原生跨度／省略历史。拒绝先走同一结构逆操作，再恢复格式；自己的连续待定插入支持部分／全部取消，省略宽度和所属表格分组保持准确。52 项单元／集成检查、三引擎 15 项及 36 DOCX／144 XML 固定 XSD通过，类型／边界通过，证据在 `.cache/docx-gap-audit/r14-*`。

本轮只剩 R15，完成后提交并推送 `feat/docx-complete`。R16～R19 仍是原计划未完成项，不并入本次授权；不创建 PR。
