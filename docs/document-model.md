# 文档包读取与模型

三格式模型及来源规则见[数据模型](data-model.md)。本文说明 Word/PPT 内容模型；统一入口新增的工作簿分块快照和窗口查询见 [R3-03](workbook-model.md)。DOCX 读取及只读验收已前移至 R1-06 前，当前状态与细则见[第一轮计划](iteration-1.md)。

R1-04 提供 `document` 类型、格式版本 1 的只读文档模型，R1-05 增加[资源校验、宿主资源收集和 ZIP 适配](resources-and-zip.md)，R2-02 增加 [Presentation 严格契约、新建与文档包模型](presentation-model.md)。输入是宿主已授权的虚拟文件 Map；核心不访问文件系统、网络或模型服务。

```ts
import { openDocument } from "@doffi/core";
import type { DocumentPackage } from "@doffi/core";

declare const files: DocumentPackage; // 宿主提供的 UTF-8 文本和资源字节。
const document = openDocument(files);
const snapshot = document.snapshot();
const summary = document.getNode("n-summary");
const savedFiles = document.serialize();
const reopened = openDocument(savedFiles);
```

`snapshot()` 返回冻结的清单、规范 HTML、有序样式、元数据和批注。`getNode(id)` 返回冻结的节点视图，包含标签、属性、最近的有身份父节点、子节点 ID、文字和完整子树 HTML；查无节点返回 `undefined`。节点视图不暴露 parse5 或 CSS Tree 的对象。

文字按文档顺序拼接原始文本节点，保留空白和 Unicode 组合序列，不做 NFC 归一化；`br` 贡献一个 `\n`，`img` 贡献一个 `U+FFFC`。偏移使用 UTF-16 单元。表格分组等无身份结构不出现在父子 ID 关系中，仍保留在 HTML 中。

加载会复制输入字节。快照及节点视图不能改动模型；`serialize()` 每次返回独立的 Map 和字节副本。修改输入、某次输出或另一份文档不会改变当前模型。序列化不增加修订、不改文档或节点 ID。它规范化 HTML/CSS 和已知 JSON 文件，保留其余包文件和资源字节，不原地覆盖输入。

## 读取规则

- `manifest.json` 决定格式版本、入口及文件角色。未知整数版本返回 `UnsupportedVersion`；当前已支持 `document` 和 `presentation` 文档包，`workbook` 仍返回 `UnsupportedFeature`。Presentation 的几何、引用、内容域及批注规则另见模型说明，原生 PPTX 读取尚未实现。
- 所有文件路径必须为规范 POSIX 相对路径。包内所有文件均不能与另一文件的祖先路径冲突，包括未被清单引用的文件；同一路径不能同时作为文件和目录。清单中的正文、样式、元数据、批注及资源路径不能重复或重叠。
- 正文、样式和 JSON 文件必须存在，文本必须是有效 UTF-8。读取错误返回有界 `DocumentError`，不回显原始内容或解析器错误。
- HTML 采用 HTML5 doctype 和一个 UTF-8 charset 声明。正文仅有一个 `article`、`section` 或 `div` 根，其 `data-office-kind` 为 `document`。解析错误、重复属性和不支持的结构均拒绝，不静默删除内容。
- 正文中除 `br`、`thead`、`tbody`、`tfoot` 外的元素必须有非空且唯一的 `data-node-id`。这些结构也可显式带 ID。普通 HTML `id` 若出现也须唯一。
- 元数据的节 ID 和节根引用分别唯一；每个节根必须指向存在的容器节点。批注 ID 唯一，文档身份一致，基准修订不能晚于包的当前修订。
- `open`、`resolved` 批注引用的节点必须存在，否则返回 `InvalidDocument`；`needs_reanchor` 可保留缺失节点的原始锚点，不自动定位或改写状态。
- HTML 样式链接必须与清单的样式列表及顺序一致。图片必须引用清单声明且已提供字节的包内资源；缺少图片字节返回 `AssetUnavailable`。
- 清单中的全部资源均须提供字节并匹配 SHA-256，包括未被正文引用的声明；PNG、JPEG、WebP 还检查签名与基本容器边界。当前只允许这三种 MIME 用于 `img`，其余附件作为数据保留。检查不等于完整图像解码。

HTML 的资源地址按入口所在目录解析，只做一次 URL 解码；允许安全的包内 `.`、`..` 引用，拒绝越界、协议、绝对地址、反斜杠、控制字符、查询、片段及编码后的路径分隔符。包路径及资源声明都不授予宿主存储权限。

## HTML/CSS 支持范围

以下列出当前读取器实际接受的子集，受[保真目标](document-format.md#保真目标)指导继续扩展。超出支持表的普通样式属于当前能力限制，不代表该样式不安全。当前仍会拒绝未支持的结构和样式，尚未实现任意安全样式的保留，也未完成 Office 视觉保真验收。

Word 原生语义通过可选元数据及 `ins/del` 容器保存；打开时校验内容域、样式图、编号、字段、链接、表格和富文本批注的引用，缺失或循环引用拒绝。原生批注文字范围还须匹配 UTF-16 原文，变更后的精确映射已在 R1-07 实现，见[核心事务与历史](document-session.md)。

正文元素采用[首轮基线](iteration-1.md#html元数据和样式基线)。保留 `id/class/title/lang/dir/style`、`data-*`、`aria-*`，以及图片的 `src/alt/width/height`、单元格的合并及 `scope`、列表的 `start/reversed/type/value`。其他属性需要显式增加支持；处理器、脚本、外部嵌入和外来命名空间均拒绝。样式表通过包内 `link` 引入，当前不接受 HTML `style` 元素。

CSS Tree 解析和检查样式表及行内声明，保存声明顺序、级联源、`!important`、自定义属性与 `var()`。支持元素、类、`data-*` 属性存在/相等、后代、分组和 `:root` 选择器。当前不支持其他伪类、伪元素、ID 选择器、其他组合符、嵌套规则或 at-rule。

支持的属性为：

- 文字：`color`、`font-family/size/weight/style`、`text-decoration` 及其 `line/color/style`、`text-align/indent`、`line-height`、`letter-spacing`、`word-spacing`、`white-space`、`vertical-align`。
- 原生映射扩展：`font-variant`、`text-transform`、`tab-size`、`overflow-wrap`、`table-layout`、`list-style-type/position`、`break-before/after/inside`、`widows/orphans`。
- 尺寸：`width/height` 及其 `min/max`。
- 盒模型：`margin/padding` 及四边属性，`background-color`。
- 表格和边框：`border`、`border-width/style/color`、四边及四边的 `width/style/color`、`border-collapse/spacing`。
- 变量：`--*` 自定义属性，以及 `var()`；计算/颜色函数为 `calc/min/max/clamp/rgb/rgba/hsl/hsla`。

普通声明由 CSS Tree 对照属性语法检查。带变量的值检查可解析性、变量引用及受支持函数，实际替换与有效值留给渲染器；核心不计算 CSS 级联。CSS 中的 URL、未解析内容和未支持函数都会被拒绝，包括自定义属性中隐藏的资源引用。

## 输入预算

`openDocument(files, { limits })` 可覆盖以下默认值；各项须为正的安全整数，未知选项会报错。它们是输入保护默认值，不是已测量的性能承诺。

| 选项 | 默认值 | 范围 |
| --- | --- | --- |
| `maxFiles` | 1024 | 整个包的文件数 |
| `maxPackageBytes` | 32 MiB | 整个包的字节数 |
| `maxFileBytes` | 8 MiB | 单个文件的字节数 |
| `maxNodes` | 100000 | HTML 树节点，包括文字、注释和外壳 |
| `maxDepth` | 128 | HTML/CSS 树深度 |
| `maxCssNodes` | 100000 | 所有样式表和行内样式累计的 AST 节点 |

HTML 和 CSS 在构建树时累计节点数并检查深度，超限立即中止解析并返回 `ResourceLimitExceeded`，不等完整树构建完毕。CSS 预算覆盖所有样式表和行内声明，以及解析器直接补建的节点。同一次 Word 文档读取中，完全相同的行内声明复用规范化结果，每次复用计一个节点。HTML 源标签若被解析器忽略而未生成元素，也会被拒绝；合法表格自动补全 `tbody` 等规范化仍受支持。规范化结果也检查字节预算，避免同一配置下产出无法重开的包。默认值通过冻结的 `DEFAULT_DOCUMENT_READ_LIMITS` 导出。

## 后续边界

本入口不创建编辑会话、不执行操作或提案，不提供历史和界面。R1-06 的 `createDocumentSession(model)` 提供[事务与历史](document-session.md)，复用同一文档包来源。已有文字范围的原文匹配已校验；编辑后的精确映射、失效及显式区域重定位已在 R1-07 实现；不会在读取时猜测新位置或改写批注状态。

资源缺失时先通过[宿主资源收集器](resources-and-zip.md#资源校验与收集)补齐；同步打开不隐式读取宿主存储。ZIP 读写、原始路径/重复项检查及有界解压由 `office` 提供，见[资源与 ZIP](resources-and-zip.md#zip-读取与保存)。只读浏览器样式隔离与基础测量已在 `editor` 实现，见[DOCX 读取](docx-import.md)。完整图像解码及自动分页仍由后续适配器负责。
