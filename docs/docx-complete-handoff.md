# DOCX 完整功能执行交接

用户于 2026-11-15 要求停止本 session 的功能实现，写清未完成项，交由另一个 session 继续。**先读 [固定未完成清单 R01～R19](docx-complete-progress.md#当前固定未完成清单)**，不要从历史日志里重新推导待办，也不要把 D1／D4／D5／D6／D8 五个大类当成只有五项小工作。

## 分支、服务与工作状态

- 仓库：`/home/danny/work/doffi`；工作分支：`feat/docx-complete`；起点：`main@f568638`。
- 上一完整功能提交：`233a7b1`（段落结构粘贴及整组修订）。此前 `c97b814` 完成书签／内部引用迁移，`3b058de` 完成字段／公式／控件等行内语义粘贴。
- 这次交接检查点包含尚未收口的表格剪贴板实现及测试。**它是 WIP，不代表 R01 或整个计划完成。** 接手先检查 `git status`、`git log -5 --oneline` 和当前 diff；保留现有改动，不重置或切回 main 覆盖它们。
- 服务已启动：<http://localhost:5174/>，交接检查时返回 HTTP 200。启动命令为 `pnpm --filter @doffi/playground dev --port 5174 --strictPort`。服务仍在运行时直接复用；终端会话消失后检查端口再启动。
- D2、D3、声明范围的 D7 编辑功能已完成。剩余实施以计划里的 R01～R19 为准，不能声称已经全面兼容 Word 或通过 Word 实机像素验收。

## 用户要求与已确定选择

- 继续完成原计划；先完成功能，只做必要的相关测试。本轮不做独立审查、A/B、远程全量 CI，不使用子代理。
- 允许本地阶段提交。**不 push、不创建 PR、不合并、不发布**；用户会主动发起 PR 和审查。
- 产品决策主动询问；用户未及时回复时，按已经给出的推荐方案继续并记录。以下决定不要重复询问。
- 未建模的正常原生内容允许完整保留 XML／关系／资源，明确占位；不要求为全部特殊对象实现内部编辑器。安全禁止、结构错误和真实资源预算仍明确拒绝。
- 重叠 run 格式修订保留最初格式、最后作者，已实施。
- 普通合并／拆分采用 Word 常规行为：不记录为修订，界面明确提示，仍可撤销；特殊原生合并修订保留，已实施。
- 跨文档粘贴保留来源格式，重名而不同的样式隔离，定义及依赖均相同则复用，已实施。
- 部分字段／控件／表格选区不自动扩张到选区外；完整选中结构才迁移对应语义。重名书签独立命名并同步复制内容里的引用；未复制目标保持未解析，避免误连目标文档的同名书签，已实施。
- 本轮不推进 Excel 后续阶段，不重开原有 Excel 接受条件；AGPL、第一方 OOXML、单一 HTML 源和统一事务等约束继续遵守仓库 AGENTS.md。

## 接手首先完成 R01

当前表格迁移已接通列网格、合并续接格式、条件表格样式、附属 CSS、制表位／条件标记以及行插入修订。复制全表的来源样式冲突会隔离；部分表格改为选中文字及外观。尚未完成最终源的必要收口检查，不应直接勾选 R01。

交接前实际复验结果：

- `pnpm build:packages` 通过，日志为 `.cache/docx-gap-audit/word-table-copy-build5.log`。
- `source-css.test.ts`、`word-table-content.test.ts`、`docx-table-content.test.ts` 三文件共 **15 项通过**，日志为 `.cache/docx-gap-audit/handoff-table-tests.log`。
- Chromium 的两个 `whole-table semantic clipboard` 用例通过，涵盖条件样式冲突、合并续接格式、输入及撤销，日志为 `.cache/docx-gap-audit/handoff-table-browser.log`。
- **最后两项修复之后尚未跑**工具／应用类型、受影响旧剪贴板用例、Firefox／WebKit 和表格输出的官方 XSD。较早提交的通过记录不能替代当前源的结果。
- 测试里暂时保留了 `payload.styles`／`payload.nodes`／目标段落属性的失败诊断信息。收口时按需要精简；这些是测试夹具信息，不是产品中的日志功能。

刚修改的两个已复现问题：

1. **粘贴后首行条件粗体变普通。** 原来先移除源样式类，再读取 CSS，使低优先级 `font-weight:inherit` 重置被固化为行内覆盖。新增 `SourceStyleResolver.specifiedOutsideClasses`，按原始胜出声明的所属规则排除被迁移样式，保留真实直接覆盖；格式包改用此方法。此改动也影响普通来源格式复制，须跑相关已有剪贴板回归。
2. **原生拒绝整表插入后多一个空段落。** 导出将透明 div 内的末尾表格当作整个内容结束，额外生成无修订的段落。`Exporter.blocks` 新增 `hasFollowing`，透传后续源块状态，避免在仍有后续段落时补空段落。实际 DOCX 重开／拒绝定向用例现已通过，仍需检查相关导出回归和 XSD。

完成 R01 后优先处理 R02 的编号迁移和 R03 的跨段接收选区，不要继续拆出新的交付阶段或只追加日志而不更新清单。

## 当前代码入口

| 范围 | 文件 |
| --- | --- |
| 迁移契约 | `packages/core/src/contracts/word-content.ts`、`word-operations.ts`、`packages/core/src/index.ts` |
| 语义附加与 CSS 身份 | `packages/core/src/session/word-content.ts` |
| 段落／表格结构粘贴 | `packages/core/src/session/word-block-content.ts`、`word-inline-content.ts` |
| 源拆分、包装和段落合并 | `packages/core/src/document/edit.ts` |
| 原始 CSS 声明归属 | `packages/core/src/document/source-css.ts` |
| 嵌套表格行审阅 | `packages/core/src/session/word-row-revisions.ts`；已补外层表格移除后跳过消失的内层表格 |
| 格式及语义复制 | `packages/render/src/word-clipboard-format.ts`、`word-clipboard-semantics.ts`、`projection.ts` |
| 格式恢复及实际粘贴 | `packages/editor/src/rich-text/clipboard-format.ts`、`packages/editor/src/rich-text.ts` |
| 原生编译 | `packages/office/src/docx-export.ts` |
| 当前新增检查 | `packages/core/tests/word-table-content.test.ts`、`packages/office/tests/docx-table-content.test.ts`、`word-table-content-fixture.ts`；浏览器 `tests/browser/pagination.spec.ts` |

当前实现边界：

- `pasteWordInlineContent` 用于段内内容；`pasteWordBlockContent` 可在同一接收段落中插入多个段落及整表。接收选区跨段仍被明确拒绝，见 R03。
- 区块覆盖已有插入／删除／移动修订仍有限制；修订边界外插入可独立审阅。完整覆盖处理归 R04。
- `WordContentSemanticsSchema` 扩展了表格、tabs、conditionalFormats、extraClasses；行内契约不接受整表／故事。传输身份只映射到新插入节点，不能解析到目标已有节点。
- 附属 CSS 分配独立名称；表格 `styleContexts` 的旧绑定移除后由既有表格样式刷新逻辑重建，续接段落／run 的直接格式单独保存。
- 编号定义、`ul/ol/li` 等列表区块尚未完整迁移，见 R02。不能复用来源 numberingId 造成误绑定。
- 系统剪贴板尚未覆盖绘图／图表／注释引用、批注线程、附属故事、保留原生对象和依赖资源，见 R05～R08。`copyWordClipboardFormat` 仍对 `data-word-native-object` 直接返回。

## 后续原生迁移必须保持的边界

- 不用“整份来源文档包”实现剪贴板。共享原生 JSON 集合必须只带所选 fragmentKey；样式、故事、原生 XML 和依赖部件也不得夹带未选正文。
- 字段复制的是当前显示缓存及指令，不能恢复未选中的旧隐藏缓存。部分字段仍复制实际选中文字，不扩大选择。
- 原生对象内部批注已有精确控制符绑定；跨来源复制时 `NativeWordComments` 的 sourceId 映射需要结合来源／端点节点区分，不能仅以数字 sourceId 为唯一键。
- 同源 `copyWordStorySemantics` 不能直接套用跨文档数据：它允许回退到已有身份、递归克隆现有故事。跨文档必须先隔离身份与资源。
- 自定义 XML 的 `w:storeItemID` 等依赖不一定体现为片段里的 r:id。需要显式依赖索引；不能为了保留绑定而复制无关根部件。
- 新 JSON 集合的资源整理已实现。旧 XML 集合目前仍保守保留，R11 需在不破坏旧包读取及撤销的前提下处理。core 不负责解析 OOXML 业务语法。
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

R01 还要按改动补跑已有来源格式、只读复制、字段与段落区块等用例，覆盖同文档样式复用及冲突隔离。不要因此启动全量仓库测试／远程 CI。

当前机器的浏览器配置：

```bash
PLAYWRIGHT_BROWSERS_PATH=/home/danny/work/doffi/.playwright pnpm exec playwright test --config .cache/docx-gap-audit/playwright.config.ts tests/browser/pagination.spec.ts --project chromium --project firefox --project webkit --grep 'whole-table semantic clipboard'
```

`.cache` 被忽略。若换工作树或缓存不存在，用根 `playwright.config.ts` 派生临时配置：保留其项目，设置 `baseURL=http://127.0.0.1:5174`、关闭自动 webServer、`retries=0`。不要把失败重试当作通过。

原生校验使用固定的 `scripts/fixtures/ooxml/OfficeOpenXML-XMLSchema-Transitional.zip`，SHA-256 为 `d34187520749998af306faf1b730e568b0ca6d88ad24638a407c0a9bb4ca04fc`。当前机器 `.cache/xlsx-venv/bin/python` 安装了 lxml；系统 Python 没有。既有 `.cache/docx-gap-audit/*-native.mjs` 是收集真实导出字节的辅助生成器，部分使用 expect stub，**不是测试**；先通过实际 Vitest，再校验输出。只按文档声明处理 MCE／W14 等可忽略扩展，不能笼统声称所有现代 OOXML 都通过基础 XSD。

浏览器脚本 DataTransfer 检查不等于操作系统剪贴板实机验收。之前用户上传报 ResourceLimitExceeded 的具体文件未保存在服务器；恢复进展和大量片段门槛已经有复现修复，不能声称验证过该用户原文件。

## 完成与汇报

后续只按 R01～R19 更新具体完成状态，写明功能、必要验证和仍存限制；不再只重复“五个方向”。已有历史检查点可追溯，但未完成项以固定清单为准。若发现破坏本轮已声明保留／编辑边界的缺陷，归入对应条目修复；不自行扩大产品范围或要求新的独立审查流程。
