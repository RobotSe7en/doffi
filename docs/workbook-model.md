# 工作簿包模型与 XLSX 只读预览

`core` 已实现工作簿严格契约、新建、分块校验、类型化单元格/窗口查询及保存重开，`office` 文档包 ZIP 同样支持工作簿。[公式语法与引用基础](workbook-formulas.md)已实现，声明范围的原生 XLSX 读取、公共只读网格和专用 Worker 已实现；完整计算、编辑和输出仍属后续任务，不能称为 Excel 完整编辑支持。读取矩阵与验收边界见[读取阶段验收](workbook-reading-acceptance.md)。完整范围见[第三轮细则](iteration-3.md)，存储依据见[技术基线](workbook-storage.md)。

## 统一入口与类型分支

按用户选择，三种格式统一使用 `openDocument`。返回模型的 `kind` 是类型判别字段：Word/PPT 使用 `ContentDocumentModel` 和原有 HTML/节点接口，工作簿使用 `WorkbookModel`、`WorkbookSnapshot`、`getCell` 和 `readWindow`。工作簿快照不含 `html`，清单不含 `entrypoint`；查询不先展开全表。现有 Word/PPT 包继续可读，宿主 TypeScript 代码需要在使用格式专属字段前判断类型。

```ts
import { createWorkbook, openDocument } from "@doffi/core";

const files = createWorkbook({ documentId: "host-authorized-id", sheetName: "预算" });
const model = openDocument(files);
if (model.kind === "workbook") {
  const sheetId = model.snapshot().metadata.workbook.activeSheetId;
  const cells = model.readWindow({ sheetId, row: 0, column: 0, rows: 40, columns: 20 });
  // 新建工作簿没有实际单元格，cells 为空；未存储的坐标是隐式空白。
  model.getCell(sheetId, 0, 0);
  model.resolveReference("=A1:B20", { sheetId, row: 0, column: 0 });
  openDocument(model.serialize(), { kind: "workbook" });
}
```

可选 `kind` 参数是预期格式断言，不做转换；不匹配即失败，并帮助 TypeScript 收窄返回类型。`createDocumentSession` 当前明确拒绝工作簿，直到 R3-08 读取门槛之后开放编辑。Playground 已接入工作簿只读预览；编辑器会话和 Node 转换仍拒绝工作簿，不把它当成空白 Word。

`resolveReference` 根据本模型的工作表、名称和智能表，返回有序源区域、引用错误或明确的未解析原因；不展开单元格、不求值，也不使用图表缓存。返回的联合区域保留顺序和重复覆盖，调用者不能直接当作集合去重。具体支持和预算见[引用基础](workbook-formulas.md)。

`getAsset(path, maxBytes?)` 只返回当前清单声明资源的独立字节副本，不序列化单元格块。可选单次字节预算在复制前检查，缺省沿用打开时的单文件限额。路径严格匹配当前模型的资源清单；样式、元数据、内容 HTML、别名路径及网络地址不会通过该接口读取，也不会触发外部访问。修改输入包或查询结果不能改写模型，其他模型中同名路径不共享内容。

内部显示资源适配按文档、修订、路径和摘要核对 PNG/JPEG/WebP 字节，先隔离提供方缓冲区，再异步核对摘要。一次候选窗口内的重复引用共用 Blob URL；成功结果由所属视图释放，失败、取消或 URL 分配异常会释放此前创建的全部 URL，提供方忽略取消也不能发布迟到资源。默认限制 5,000 次引用、256 个不同资源、单项 8 MiB 和累计 64 MiB；旧窗口与候选分别拥有资源，替换前不会撤销旧视图使用的 URL。

Blob URL 创建前检查图片画布和动画帧头：单项最多 16,777,216 像素、累计 67,108,864 像素，每项最多 256 帧。动画按完整画布乘帧数保守计数，APNG 的独立默认图另计一帧；WebP 同时核对容器、子帧和压缩码流声明的尺寸。JPEG 支持基线、扩展顺序和渐进 DCT 帧头，其他编码明确拒绝。这些头部检查不替代浏览器实际解码，也不保证图片内容有效。三个浏览器已解码独立组装的 APNG／动画 WebP 及现有静态 PNG/JPEG/WebP 样例；截断、矛盾尺寸、单项和累计超限均有定向测试。依据 [PNG 第三版](https://www.w3.org/TR/png-3/)、[WebP 容器](https://developers.google.com/speed/webp/docs/riff_container)、[WebP 无损码流](https://developers.google.com/speed/webp/docs/webp_lossless_bitstream_specification)及 [JPEG T.81](https://www.w3.org/Graphics/JPEG/itu-t81.pdf)。像素预算是资源保护边界；绘图组合及实际解码检查见[读取阶段验收](workbook-reading-acceptance.md)，浏览器／Worker 内存另按固定资源方法计量。

只读预览的内部窗口读取器已按文档与修订绑定完整显示帧，跨窗格合并只读取一次内容锚点。外观查询还包含合并四角、可见行列对应的周界及最近可见邻格；邻格属于另一合并时也读取其四角。整张表大小的合并不展开内部坐标，隐藏的百万行整段跳过。可见单元格默认上限仍为 5,000，外观来源另限 25,000 个地址、累计源文本 8 Mi UTF-16 单位；这些是保护限额，不是容量验收结果。整批查询先检查累计单元格和重复地址，再以最多 256 个坐标的小窗口读取；任务间让出执行以接收取消。快速切换窗口、旧响应、超限及已释放实例不会发布部分结果。内部 DOM 投影保持完整单元格尺寸，再按窗格裁剪，RTL 只改变物理位置；合并内容不会在碎片内重新居中，重复片段只有一个辅助技术单元格身份。绘制失败时保留前一个完整窗口，派生 DOM 不写回源包。

`tests/browser/workbook-grid.spec.ts` 在 Chromium、Firefox 和 WebKit 的几何检查覆盖左右方向、四窗格合并、源命中、第 1 万行与 XFD1048576、DOM 单元格预算和释放。这里使用少量稀疏实际单元格，验证坐标与窗口上限，不能替代 20 万实际单元格的完整 XLSX 首屏验收。此适配已由公共 Worker 通道和网格预览复用，数字／日期、本地字体、对象及整窗组合均有定向检查；完整 32 份来源的四引擎矩阵另行记录。

内部单元格组合层现已把当前数值／日期、布尔值、错误值、富文本、填充、周界及对角线装配为一个完整候选窗口。源 CSS 在源区域计算，派生字体与布局在隔离显示区域测量；字体报告、数值溢出和公式缓存状态随候选返回。所有单元格完成后才由所属视图一次替换窗口，源节点、显示节点及累计文字分别受限；失败不改动已发布窗口。网格线遵循工作表显示设置，并位于填充下方。普通边框引起的旋转文字、填充和边框共享物理形状，合并区覆盖角提供的有效边框也参与形状判定。图案／渐变只分配可见碎片，冻结窗格中的内容保持完整合并区的文字布局。

组合检查已接入公开导入的三份独立 XLSX，以及原生旋转样例和四窗格合并／RTL／通用 CSS 隔离用例。布尔值、错误值与数字显示不从 HTML 缓存文字伪造 UTF-16 源位置；富文本保留精确偏移。该组合层由公共只读预览复用；条件格式、普通溢出和视口外旋转来源已接通；自动行高的实际来源索引见下文，行高／原生组合及绘图／图表的完整检查汇总于读取阶段验收；本段只描述窗口组合层。

## 文件与来源

| 内容 | 来源及校验 |
| --- | --- |
| 清单 | `kind: workbook`，`chunks` 数组；每块包含 sheetId、零基准 row/column、路径和 SHA-256 |
| 元数据 | workbook 的日期系统、计算模式、活动表、表次序/名称/可见性、行列度量、冻结、合并、数字格式、作用域名称、超链接和打印定义 |
| 样式 | 清单列出的 CSS；单元格和富文本的 class 引用必须存在；拉丁/东亚字体角色分别声明 |
| 内容 | 每块 HTML 表格；行、列按坐标递增，原点为 256 行及 32 列的整数倍；块不能重叠或引用不存在的表 |
| 批注/资源 | `cells` 锚点保留工作表与精确多区域；正文/回复引用有界独立 HTML story，资源校验摘要；原生备注/线程、基础绘图与图表已接入内部管线，兼容定位核对及只读预览已接通 |

空白远端位置不产生任何分块。合并、冻结和行列尺寸只保存逻辑定义；合并可跨块，不随窗口复制。至少一张表可见，活动表必须可见；表身份、大小写不敏感的表名及数字格式 ID 唯一。尺寸区间必须有序且不重叠，合并重叠通过行扫描和列区间索引检查。

R3-06 接入原生读取时增加了以下可选契约，旧工作簿包无需补字段：

- `names` 保存不透明身份、名称、可选工作表作用域、表达式及明确的引用原点。名称在各自作用域内不区分大小写，不能与单元格地址或保留名称冲突；普通表达式的解析不代表已支持求值。
- 表及行列维度可引用 `numberFormatId`，避免格式化的空行列丢失数字格式。`defaultColumnWidthUnit` 区分原生列宽与不含边距的基础字符宽；缺省沿用原有字符宽语义。`defaultRowsHidden` 表示隐式行的隐藏默认值。`defaultRowHeightCustom` 保留原生手工默认行高；行的 `autoHeight: true` 可显式覆盖固定默认，不能与该行的固定 `heightPt` 同时出现。旧包省略字段仍按自动默认解释。
- 数字格式定义为自定义 `code` 或 `builtinId` 二选一，保留内建日期格式随语言变化的语义，不在导入时固定成某个地区的格式串。`themeCssClass` 与 `tabCssClass` 引用 CSS 中的主题和工作表标签外观；既有 `tabColor` 继续可读，但不能与 `tabCssClass` 同时定义。
- `columnWidthFontCssClass` 单独引用字符列宽的计量字体；原生源按 `cellStyle builtinId=0 → cellStyleXfs → font` 找到 Normal 字体，不以字体记录零或 `cellXfs` 零代替。依据 [Microsoft Open Specifications 的说明](https://learn.microsoft.com/en-us/answers/questions/1110580/where-is-the-normal-style-font-defined-in-an-xlsx)，本地化名称和修改过的 Normal 样式仍通过内建身份识别。原生 Normal 定义有歧义或缺失时不猜测列宽字体；老 HTML 包没有此可选引用时沿用表的默认字体。这个必要的布局引用不替代下述命名单元格样式定义与关联。
- `hyperlinks` 保存单元格区域及 HTTP(S)/邮件目标或内部名称/地址，可保留原生说明文字。读取不访问目标，重叠链接及可执行协议失败。
- `print` 保存区域、重复标题、纸张度量、方向、页边距、适配或百分比缩放、分页及类型化页眉页脚文字/字段。未指定的纸张及页边距保持未指定，未提前猜测打印机设置。该契约可保存重开，实际工作簿分页/PDF 在 R3-18 验收。

原生适配的分部件检查覆盖 32 份独立来源及 Calc 重存文件的名称和打印设置、流式行列/冻结/合并/超链接、共享/内联富文本、公式缓存与共享公式。自动行高是原生布局缓存，固定行高才写入源尺寸。这些模块已接入公开 `importXlsx`，返回经过核心校验的文档包；完整显示的检查与 R3-08 退出条件见读取阶段验收，编辑和输出分阶段交付。

`office.importXlsx(bytes, { documentId, signal?, onProgress?, limits?, workbookLimits?, nativeLimits? })` 是异步入口。身份由宿主提供，结果继续交给统一 `core.openDocument`；不新增 `openWorkbook`。导入前捕获身份、回调、信号身份及独立额度副本，ZIP 在首次让出任务前隔离输入字节。取消信号的 `aborted` 状态保持实时读取；宿主修改原选项或输入数组不会把正在进行的导入改成另一份文档。两个并发导入的批注、资源和输出字节互相隔离。

```ts
import { importXlsx } from "@doffi/office";
import { openDocument } from "@doffi/core";

const files = await importXlsx(authorizedBytes, { documentId: "host-workbook-id" });
const workbook = openDocument(files, { kind: "workbook" });
const cell = workbook.getCell(workbook.snapshot().metadata.workbook.activeSheetId, 0, 0);
```

进度阶段为 `zip`、`metadata`、`worksheet`、`content` 和 `validate`，每条记录冻结；`completed/total` 只在当前阶段及工作表内解释，不是一个跨阶段百分比。`content` 覆盖批注与逐表对象映射。最后一个校验完成通知仍检查取消，取消或宿主回调抛错都会拒绝整次导入，不能得到部分包。原生 ZIP/XML 额度、单元格／公式额度、生成 HTML 包额度分别配置；所有分块、辅助正文、图片、样式、元数据、批注及清单共享最终文件数与字节额度。

公开入口已核对 16 个独立源及对应的 16 个 Calc 重存文件，包含两份一万行／二十万实值单元格／两万公式来源；17 个负例及后期对象映射错误均拒绝整包。测试同时核对表、规则、图表、图片、两类批注、缓存状态、末行查询和保存重开。这是导入 API 检查，完整公共预览矩阵与 4 vCPU／8 GiB 性能分别验证。最终 `openDocument` 校验在调用线程同步执行；Playground 将整个 XLSX／工作簿包的打开和校验放在所属 Worker，主线程只保留元数据和窗口查询。固定资源与完整来源显示结果见读取阶段验收及资源检查记录。

数字格式的内部源码编译器已按照[格式代码规则](https://support.microsoft.com/en-us/excel/review-guidelines-for-customizing-a-number-format)区分分段、引号/转义、条件/颜色、日期/时长字段、科学计数、货币/地区标记及需要测量的填充/留白字符；引号内的分号不会被错误拆段。十进制、科学计数和常用分数分段已能生成文本、按字符宽度留白及填充部件；分组与小数符号由调用方显式提供。显示舍入另行处理，百分比、逗号缩放和科学计数不先乘除浮点值，避免中间溢出；它不改变原值，也不是公式 `ROUND` 的实现。分数使用有界连分数近似，支持假分数、带分数及不约分的固定分母，位数增加不会逐个枚举分母；原生分母常量与引号中的文本数字分别解释。带分数的强制零分子、分组整数等特殊组合仍明确拒绝，不能默默套用其他格式。[固定 Calc 的 130 个格式化结果](evidence/xlsx-number-format-semantics.json)可通过 `python3 scripts/check-xlsx-number-formats.py` 重核，其中 60 个分数结果已与内部显示部件对照；这不代表实际字体宽度与视觉验收。此证据记录了 1900 系统早期日期与 Excel 的差异，SDK 内部日期分支采用用户已确认的 Excel 兼容规则，详见下文。当前公开数字显示入口已接通这些分支，接口和内建格式边界见下文；公式 TEXT 与其他地区的特殊内建格式仍按后续计算／兼容范围处理；公共预览复用当前格式器。

基本分段选择和文本段已接入内部标量格式入口：先用原始值选择正数/负数/零或条件段，再进行显示舍入；提前出现的末尾 `@` 仍是文本段。文本不因外观格式转换为数字，空段保留隐藏含义，纯字面内容不额外加负号。[47 个独立分段/文本结果](evidence/xlsx-format-selection.json)由 `python3 scripts/check-xlsx-format-selection.py` 重核。两个格式参考共 177 项，其中 151 项数值/文本结果已对照；另外 26 项日期或 General 由独立的日历／宽度分支及公共格式入口检查处理，不把原生缓存字符串当成重算结果。未匹配的条件不自动冒充 General，未支持的格式组合明确失败；以上子探针不声称覆盖所有可能的自定义格式。

公开数字格式器对文本值的 `@` 部件现返回 `source: { from, to }`，偏移按原文本的 UTF-16 单位计算。字面前后缀、留白、填充和数值／日期字符串没有该字段，即使恰好与原文相同也不能冒充源文字；重复 `@` 各自指向原文，不改写单元格值。空文本字段保留零长度来源，隐藏的第四分段不产生来源。普通文本已接入组合窗口的格式、颜色、按实际字体测量的留白／末个填充和竖排；前后缀不会污染原文选区偏移。堆叠按完整显示串的字素分段，再映射各原文片段，因而支持派生字符与重复来源混排。

[13 项独立文本格式观察](evidence/xlsx-text-formats.json)由 `python3 scripts/check-xlsx-text-formats.py` 生成原生 XLSX 并使用固定 Calc 重核，普通文本的六项不含宽度填充结果已直接对照公开格式入口。ODS 的文本不包含按列宽绘制的重复字符，不能用它验收填充宽度；浏览器另行核对实测宽度与精确映射。该样例及其原生格式／富文本经公开导入和保存重开保持原值，并通过官方 XSD 与包审计。固定 Calc 对三个带局部颜色的富文本案例忽略文本格式；这与普通文本不同，不代表 Excel/WPS 的行为。用户已明确选择统一应用文本格式：每个 `@` 保留原文的局部字体、字重、斜体和颜色，格式明确声明的颜色覆盖整段显示；前后缀、固定文字及填充字符只属于显示，不能写回原值或获得伪造的源偏移。重复 `@` 分别映射到同一源范围。三个原生富文本组合已接入窗口显示；与固定 Calc 的差异按该产品选择保留，不声称 Excel/WPS 已验证。

日期与时长的内部显示分支已实现 `yy/yyyy`、月份/分钟辨别、星期与月份名称、12/24 小时、累计时长及一至三位小数秒。1900 系统保留零日 `1900-01-00`、虚构的 `1900-02-29` 和 3 月前的兼容星期序列；1904 系统从 `1904-01-01` 开始。日历解释不经过宿主时区，不把虚构日期交给 JavaScript Date 自动归一化。负数、超出 9999 年及舍入后越界返回不可显示日期，交由网格按宽度绘制井号，不修改单元格值。时间显示先舍入到秒或显式小数秒，再省略未请求的单位；日历日期在毫秒精度处理午夜边界。序号换算对恰好半单位附近的一个浮点步长作归一化，包含半秒、半毫秒和跨日回归。此显示规则不是公式日期函数或 ROUND 的实现。

地区解析已固定为“格式声明 → 宿主 → 文档 → zh-CN”。[原生语言 ID 表](https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-lcid/63d3d639-7fd2-4afb-abbe-0d5b5551eef8)的 412 个已分配标记由 `python3 scripts/check-xlsx-locales.py` 核对；保留日历/数字系统的高位参数，不能误作语言 ID。普通拉丁数字的分组和小数符号显式读取所选地区，印度分组独立于西式三位分组。日期词汇使用显式地区、UTC 和 Gregorian 日历，月份根据是否有日期字段选择上下文形式；零日与虚构闰日的星期单独映射。运行时缺失所选语言、未知 LCID 或尚未实现的日历/数字系统明确失败，不能悄悄使用机器默认语言。`F800/F400` 使用所选地区的长日期／时间模式，保留格式颜色但不把原格式中已被系统模式替代的尾部内容再显示一次；模式只在显示时派生。系统时间模式中日间时段前后的不换行空格统一为普通空格，以消除 Node 与浏览器 ICU 版本差异，源格式中的显式空格不改变。不能用原生 h/hh 表达的 h11/h24 时钟扩展仍明确拒绝。Intl 的地区数据版本仍可能造成词汇差异，完整读取验收已记录运行环境。

[130 项独立日期/地区观察](evidence/xlsx-date-format-semantics.json)由 `python3 scripts/check-xlsx-date-formats.py` 使用固定 Calc 重核，覆盖两种日期系统、七种地区与午夜边界。Calc 的早期 1900 日期、整秒截断、`A/P` 大小写及部分地区词汇与所选显示规则存在差异，记录原始结果而不将其全部作为 SDK 预期；既有参考中 12 项相容的日期/时长结果另有直接对照。日期语法依据 [Office 格式说明](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/0e59abdb-7f4e-48fc-9b89-67832fa11789)，1900 兼容依据 [Microsoft 的闰年说明](https://learn.microsoft.com/en-us/office/troubleshoot/excel/wrongly-assumes-1900-is-leap-year)，午夜边界参考 [Stata 对 Excel 时间的独立观察](https://www.stata.com/support/faqs/data-management/excel-datetime-value-behind/index.html)。这仍不是 Excel/WPS 实机验收；公开显示入口与内部数字网格用例已接通，完整网格检查见读取阶段验收。

General 的内部宽度适配已按 [Excel General 说明](https://support.microsoft.com/en-us/excel/reset-a-number-to-the-general-format)和 [OOXML 数字格式规则](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.numberingformat?view=openxml-3.0.1)实现有界候选与实测选择：普通十进制最多 11 个显示位置，12 位及以上整数使用科学计数，窄列可以减少小数显示精度，不能把小非零值当作零。极端有限数也不先乘浮点数缩放。无可用显示时返回宽度溢出，保留原始标量供宿主查询。此规则与固定 Calc 的部分宽列 General 文本不同，不将 Calc 的原文作为该分支的 Excel 预期。显示部件的留白按实际字符测量，填充字符串整体参与字形测量，并受累计输出限额约束；相邻文字合并测量以保留字距。同一段的符号、字面前后缀与重复 General 一起参与宽度选择；百分比缩放也避免浮点中间溢出。多个填充标记按原生规则只应用最后一个。

`createWorkbookNumberFormatter({ dateSystem, hostLocale?, documentLocale? })` 提供公开的当前值显示入口。`format(value, { code } | { builtinId }, measurement?)` 返回带所选地区和颜色的文字／留白／填充部件，或宽度、无效日期、条件未匹配的溢出状态；不读取导入时的显示缓存，不改变标量和格式源。省略定义使用 General，省略测量时 General 使用完整显示位数。格式源码和地区词汇缓存在实例内分别限 256 和 64 项，不缓存单元格值；`dispose()` 清空并拒绝后续调用。已接入常用数值、Gregorian 日期／时间和简体中文内建格式，未声明的内建 ID、其他地区的纪元／数字系统需后续适配，不以 General 代替。

```ts
const formatter = createWorkbookNumberFormatter({ dateSystem: "1900", hostLocale: "zh-CN" });
const display = formatter.format(60, { builtinId: 14 }); // 1900/2/29 的显示部件
formatter.dispose();
```

[19 项独立 General 分段／系统格式结果](evidence/xlsx-general-sections.json)可通过 `python3 scripts/check-xlsx-general-sections.py` 重核，并与公开入口直接对照。内部数字布局先批量解析当前显示文字的本地字体，再按真实字宽选择 General 精度、生成留白与填充、处理井号溢出，并复用单元格缩小、对齐、旋转和竖排。颜色优先使用格式颜色与源索引调色板。源中旧的显示正文保持原样，数值派生文字没有伪造的源 UTF-16 偏移，应按单元格命中。临时测量与最终显示分别受累计 8 Mi UTF-16 和有界节点限额约束，失败不替换已接受窗口。

三浏览器和固定旧 Chromium 已检查数字、日期、地区、格式颜色、窄列、货币留白、旋转竖排、中文字体及失败清理；虚拟网格同时核对 LTR/RTL 中当前值 `12345.678` 显示为 `12,345.68`，源中的旧缓存正文不变。这仍是局部窗口验证。富文本自定义文本格式（如 `"ID "@`）已按上述产品语义接入网格；完整组合器和最终一万行固定资源读取门槛继续实施。

列宽计量已可在样式与本地字体就绪后解析上述 Normal 引用，测量 0–9 的最大数字宽度并返回字体替代诊断。字体解析只作用于所属表面的临时测量节点，不改源 CSS 或其他已挂载界面；成功或失败都会释放临时节点。三浏览器已用独立 DOM 字宽核对该数值，并覆盖旧包默认字体路径。原生列宽、自动行高和固定 Office 参考另有定向几何／视觉检查，汇总见读取阶段验收。

原生表的默认外观及数字格式同样来自 [Normal 样式](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cellstyle?view=openxml-3.0.1)，不以 `cellXfs[0]` 替代。内部预览按单元格、有效行、列、表默认的顺序选择外观和数字格式引用；两类引用分别回退，不把 CSS 类列表顺序当作级联优先级。已存储但没有显式格式的空白与隐式空白采用同一回退规则，显式格式零则保留覆盖含义。行的原生 `customFormat` 未启用时，不应用该行的 `s`；启用但省略 `s` 时应用原生缺省的格式零。稀疏维度通过二分查找定位，不展开百万行。[54 个独立 Calc 结果](evidence/xlsx-cell-style-semantics.json)覆盖这些区别，包含 Normal 与格式零不同的数值显示和背景色；可用 `python3 scripts/check-xlsx-cell-styles.py` 重核。此项不代表完整条件格式或最终视觉验收。

窗口样式测量已保留源块的 `table/tbody/tr/td` 祖先、坐标属性和富文本 `span/br`，因此基于源结构的 CSS 仍按浏览器级联匹配，不受虚拟网格裁剪容器影响。跨窗格的同一锚点只建立一份测量源；临时节点按窗口累计限额构建，全部校验完成后才连接到测量表面，释放时清空节点及索引。惰性 HTML 解析后的结构和属性白名单会拒绝脚本、事件、资源标签及与查询不符的单元格，失败不留下部分测量树。三浏览器已核对远端块选择器、富文本格式、主题变量和释放；完整单元格由下文的统一窗口组合呈现。

单元格字体读取复用 PPT 的字体角色分段，保留 UTF-16 原文和偏移，东亚、拉丁及复杂文字的字体引用分别解析。组合标记和变体选择符沿用前一个基字的角色，跨富文本 span 也不丢失该关联；`br` 计为一个换行单位。分段前先读取原始祖先上的 CSS，不为测量插入会改变源选择器匹配的 span，字体替代和缺字诊断仍由显示表面负责。字体分段受独立限额约束，不将这种字体读取检查等同于最终排版或缺字验收。

内部富文本显示已把这些字体角色、主题明暗色和源 span 样式投影为独立视图，保留每个文本节点对应的 UTF-16 区间。原生普通/双下划线与删除线独立绘制，局部富文本可关闭原生装饰；CSS 装饰仍保留祖先传播关系。`br` 显示为同偏移的换行，源包不被改写。三浏览器的窗口用例已接入真实富文本与纯色填充，并检查冻结合并、远端坐标和有限节点构建。财务下划线采用下述整格绘制；这些内部检查不代表完整 Office 视觉门槛已通过。

内部文字方向按 [SpreadsheetML readingOrder 规则](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.alignment?view=openxml-3.0.1)区分显式 LTR/RTL 与上下文方向。上下文只检查首个非空白 Unicode 标量，因此数字或 emoji 开头的混排文本不能直接套用整段 HTML `dir=auto`。单字符分类使用浏览器 Unicode 双向表，短缓存只属于当前测量表面，读后清空测量文字、释放时清空缓存。三浏览器覆盖希伯来语、阿拉伯语、Adlam、方向标记、弱字符开头及宿主 RTL 隔离；最终对齐及原生定向对照见后续排版记录和[读取矩阵](workbook-reading-acceptance.md)。

原生样式适配已将 32 份样例的有效单元格字体、主题/索引颜色、纯色填充、边框、对齐及数字格式映射到可校验的 CSS/格式定义。主题引用、字体角色、原字体名及明暗参数保留为来源；色阶应用、精确边框和特殊对齐由 R3-08 渲染器解析。原生命名样式现通过 `cellStyles` 保存定义、名称及关联：定义引用唯一 CSS 外观和数字格式，名称保留内建身份、隐藏／定制标志和层级；单元格／行列已选的格式类绑定到定义，保留各组后续更新标志。关联不再重复保存单元格当前外观，也不根据更新标志改写导入时的显示。核心重开检查身份、类名和数字格式引用，并拒绝同一单元格选择两个关联类。32 份独立／Calc 文件保留这些记录，原样保存重开；样式编辑及原生输出仍在后续阶段，不能称为完整样式闭环已通过。

高级单元格外观的源映射已扩展为 19 种标准填充模式、线性／路径渐变、双向斜线边框，以及 0～180 和 255 的旋转／竖排参数。渐变保留原始色标顺序、重复位置、主题／索引色与明暗参数；斜线即使暂未启用也保留样式和颜色。参数继续以 CSS 为唯一外观来源，格式库与条件格式共用映射。旋转范围依据 [MS-OI29500 §2.1.693](https://officeprotocoldocs-f5hpbjgea6b8gneq.b02.azurefd.net/files/MS-OI29500/%5BMS-OI29500%5D.pdf)，填充结构依据固定的官方 Transitional XSD。

内部填充绘制已覆盖上述 19 种模式和两类渐变。17 种非纯色图案的周期、方向和覆盖率以[原始标准示意图的独立测量](evidence/xlsx-fill-pattern-spec.json)核对，运行 `python3 scripts/check-xlsx-fill-patterns.py` 可从固定 ECMA 归档重核。图案点距采用 0.75 文档点；这不是 Excel/WPS 实机像素验收。路径渐变使用指定内部矩形到四边的等色轮廓，线性渐变使用文档物理角度。三浏览器检查 27 个原生填充来源、颜色、渐变方向、冻结／滚动裁剪连续性，以及超大逻辑单元格只生成可见片段、节点与色标预算和源不变。色标排序仅发生于派生绘制参数，原始顺序和重复色标保持不变。

[42 项独立原生来源与固定 Calc 记录](evidence/xlsx-cell-appearance.json)由 `python3 scripts/check-xlsx-cell-appearance.py` 重核，`--record` 才更新证据；参考 PNG 写入 `.cache/xlsx-cell-appearance-reference/`。来源已通过官方 XSD／部件审计，SDK 内部读取与包重开保留全部参数。固定 Calc 将图案及渐变改写为纯色，该外部损失单列记录，不能用其纯色画面认定 SDK 的图案／渐变正确。单格旋转与边框／填充的联动已接入下述内部绘制；最终整窗组合、编辑与原生输出仍需完成，本项尚不构成高级外观闭环验收。

对齐缩进的原生读取范围为 0–255，251–255 不被截断或误拒绝，负数、小数及 256 明确失败；依据为 [MS-OI29500 v20260818 §2.1.693](https://officeprotocoldocs-f5hpbjgea6b8gneq.b02.azurefd.net/files/MS-OI29500/%5BMS-OI29500%5D.pdf#page=295)。实际缩进使用单独测量的 Normal 字体空格宽度，不随该单元格的字体改变。

内部文字布局已接入 General／左中右、上中下、Normal 空格缩进、换行和单行缩小适应，并在同窗口批量解析本地字体后测量。富文本装饰和 UTF-16 投影保留；测量成功后才交给完整单元格的裁剪容器，失败和释放不残留测量文字，也不修改源。未裁剪的内容高度进入共享自动行高流程；窗口组合复用同一路径，覆盖跨空白格溢出、跨格居中及特殊对齐。

内部文字布局现保留 0～180 的原生角度映射及 255 的逐字符竖排。旋转后的完整包围盒参与对齐、缩小和行高测量，不在每个可见片段重新居中。竖排按完整 Unicode 字素分行，保留跨样式片段的异体字选择符、组合符和代理对；派生换行不占用源偏移，原始换行仍保留其 UTF-16 单位。默认竖排行距复用绘图文字的可缩放字体度量，避免浏览器逐行取整；缺少精确度量时显式返回回退标志。节点预算在字素遍历及拆分过程中检查，失败候选不会进入已接受的显示树。

[13 项固定字体原生旋转记录](evidence/xlsx-rotation-semantics.json)及其独立 XLSX 由 `python3 scripts/check-xlsx-rotation.py` 重核。三浏览器及固定旧 Chromium 核对其中 12 项有效文字的实际像素覆盖边界（灰度阈值 224，每边允许 2 文档点的字体／栅格差异），并另验收完整字素及跨片段的偏移。记录同时保留阈值 64 的深色像素边界；旧 Chromium 的旋转抗锯齿边缘不能仅按深色像素确定文字范围，浏览器截图保留 LCD 子像素覆盖，原生参考排除洋红色定位边框。Calc 在最后一个竖排样例中拆开 UTF-16 单位、输出替代字符；该损失记录保留，但不作为 SDK 的文字预期。文字样例使用无边框单元格，边界来自同尺寸未旋转的第一张表；带边框的联动另按下列样例检查，不能把这些局部检查称为完整外观验收。

带普通外框的旋转单元格现使用共用的底边固定轮廓；只有填充或只有斜线边框时保留矩形。侧边按文字物理角度倾斜，上下边仍水平；0°、±90° 和逐字竖排保持矩形，RTL 不反转物理旋转角度。文字的横向位置跟随所在高度的斜移并保留内容边距，竖向仍裁剪在源行范围内。合并边框使用整块高度而非分片高度，虚线相位按真实斜边长度累计。此行为依据 [Microsoft 的旋转表头说明](https://support.microsoft.com/en-au/excel/fit-more-text-in-column-headings-in-excel-for-mac)及[36 项独立固定 Calc 观察](evidence/xlsx-rotated-cells.json)，`python3 scripts/check-xlsx-rotated-cells.py` 可重核；独立 XLSX 另通过官方 XSD／部件审计和包保存重开。

旋转填充先将可见片段反投影到源单元格，再绘制图案／渐变并施加同一仿射映射，超大逻辑单元格不分配完整轮廓大小的表面。普通填充与裁剪片段的实际像素对照覆盖正负角度、纯色、图案及两类渐变，允许最多 3 个颜色通道单位的栅格差异；节点额度包含临时 SVG。原生固色填充角点误差限为 0.12 文档点，文字覆盖边界误差限为 2 文档点；截图统一像素原点，避免长列表滚动改变抗锯齿相位。接近水平时文字可能位于页面之外，参考保留该事实。原生网格尺寸取边框中心线，不能用略大的填充像素覆盖反推行高，否则 1° 的余切会放大误差。

旋转溢出的整窗发现、跨格边框优先级和合并内容已接入统一组合。图案／渐变另有固定字体的派生像素检查；Calc 扁平化后的图案／渐变仍不能充当原源视觉参考。完整范围与生产者差异见[读取验收](workbook-reading-acceptance.md)。

[94 项独立原生 PDF 记录](evidence/xlsx-alignment-semantics.json)可用 `python3 scripts/check-xlsx-alignment.py` 重核，覆盖固定的三种字体、General 类型、显式方向、缩进及横纵对齐。三浏览器对照字宽和物理位置，并检查换行、缩小、字体替代与失败清理；使用关闭字偶距的原生文字度量，WebKit 同时需要显式关闭 OpenType `kern`。这些局部校准不等同于完整页面视觉验收。Calc 的布尔／错误值右对齐与 [OOXML General 居中规则](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.numberingformat?view=openxml-2.8.1)不同，SDK 遵循后者；Calc 的显式 LTR 希伯来文字和 RTL PDF 边框相对文字位置也单独记录，不能当作 Excel 实测。数值负号和错误值标点按类型显示方向保持顺序，普通文本仍遵循显式阅读方向。日期和默认地区语义没有由此推定。

内部颜色显示已按 [SpreadsheetML 的 HLS 亮度规则](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.foregroundcolor?view=openxml-3.0.1)处理正负 tint，区别于 DrawingML 的 RGB tint/shade。浏览器解析源 CSS 中的主题/索引变量及当前文字颜色，保留源引用；原生 ARGB 高字节不作为单元格透明度。无效颜色或明暗参数明确失败。[108 个固定 Calc 颜色观测](evidence/xlsx-color-tint-semantics.json)与连续 HLS 转换的差别最多为一个 8 位通道单位，记录了舍入边界，不声称逐字节颜色一致；运行 `python3 scripts/check-xlsx-color-tints.py` 可复核。三浏览器另核对变量、颜色语法、透明度和无效值处理。

普通共享边框的内部选择器已对照[固定 Calc 的 14 类边框记录及 588 个 PDF 边界颜色结果](evidence/xlsx-border-semantics.json)：线宽等级优先，双线排在中等与粗线之间，同等级取物理下方/右方，因此 RTL 列轴会反转源顺序中的平局选择。两侧均无边框时不产生描边，不套用浏览器 CSS 表格的冲突规则。可运行 `python3 scripts/check-xlsx-borders.py` 重核；字体清单摘要、Office 归档及 PDF 工具环境随记录保存。独立描边层已按 PDF 的点单位线宽、虚线节距和双线间隔核对，三浏览器另检查裁剪后虚线实线区域完全对应、抗锯齿透明度最多一个通道单位的变化及双线留空。描边层只接收已选择的物理边；合并周界及条件格式的属性优先关系已进入完整网格。Calc 将 `slantDashDot` 变为细虚线，这是外部生产者差异，不能把该输出作为斜线保真的证明，该项另按原始标准示例实现；Excel/WPS 仍未实机核对。

斜线层读取合并起始单元格的 up/down 标志，按物理方向连接完整单元格的对角，并与共享周界分开组合。双线沿法向保留间距，冻结／滚动窗格保留原线相位；普通尺寸保持同一组端点，超大逻辑线先裁剪，避免向 SVG 传入巨大坐标。三浏览器核对 9 个原生斜线来源、双线几何与裁剪后的像素覆盖；整表合并只读取一个锚点，输出受统一边数预算约束。`slantDashDot` 使用两条错位点划线保留倾斜端部，对照[ECMA 原图的独立描边测量](evidence/xlsx-border-spec.json)，可用 `python3 scripts/check-xlsx-border-spec.py` 重核。该标准图测量不替代 Excel/WPS 实机验证。

[OOXML Primer §3.2.13](https://download.microsoft.com/download/e/1/4/e14fb96f-83b8-4a2a-84db-7fa8acbe061a/Office%20Open%20XML%20Part%203%20-%20Primer.pdf)将合并内容与各格保存的边框格式区别说明。[80 个合并边框原生 PDF 探针](evidence/xlsx-merged-border-semantics.json)进一步记录了固定 Calc 对不一致周界格式的处理，以及外侧较粗邻格边框对局部边界的影响；运行 `python3 scripts/check-xlsx-merged-borders.py` 可复核。RTL 导出可能重复同一描边，记录保留了原始重复项。内部组合器现按该固定参考从上方起始角读取顶边／起始边、上方末端角读取末端边、下方起始角读取底边，再与外侧邻格逐段比较；源中其他周界格式仍保留，不把绘制选择写回源。80 项边界存在性、外侧覆盖和 RTL 物理位置均有独立预期核对；这些观测不替代 Excel/WPS 实测。冻结和滚动窗格共用稀疏坐标，隐藏行列不展开，整表大小的合并仅需最多四个样式角点；每个窗口统一限制输出边数。分段虚线保留原边起点的相位，三浏览器对照整线裁剪与分段绘制，实线覆盖一致、透明度最多相差一个通道单位。

单元格边框来源解析区分原生逻辑列边与普通 CSS 的物理左右边：原生侧随工作表 RTL 镜像，CSS 侧先转换到逻辑坐标再参与共享边界比较。CSS 实线、虚线、圆点和双线保留浏览器计算后的实际点宽，不量化为原生命名线型；原生线型继续使用固定 PDF 的度量。颜色、主题明暗和透明度统一解析为受校验的 sRGB 描边。原生 start/end 与遗留侧的冲突绘制定义明确失败，不猜测优先级；特殊 CSS 浮雕边仍不属于已实现绘制。来源解析、边框组合和已选边的 SVG 绘制分离，测试窗口在替换前完成全部构建，失败保留原视图。智能表样式已接入属性组合；条件格式按优先级、停止规则及属性声明参与最终组合。

DrawingML 主题的填充、背景填充、线条和效果列表也已映射为共享 CSS 定义，保留占位颜色与对象引用颜色的关系。复用 PPT 的颜色、填充和轮廓原语，Excel 边界额外检查属性、重复选择、颜色变换顺序及累计样式预算。非空对象默认定义、附加配色及未知扩展目前明确拒绝；本项验证覆盖全部 32 份样式来源，尚不代表绘图对象和预览已完成。

`cellXfs` 的 `apply*` 标记控制后续命名样式修改是否覆盖单元格，而不是简单地关闭该格当前格式；实现据 [Microsoft 的 XF 说明](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/68362a4b-5589-4504-b566-e8154dce1de3)保留当前值与更新标记。索引色采用[标准映射](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.indexedcolors?view=openxml-2.8.1)，工作簿明暗参数保留 [HLS 亮度语义](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.foregroundcolor?view=openxml-3.0.1)。

原生一万行分部件链路已执行两遍有界扫描、共享公式处理、40 块 HTML 生成及统一模型重开，核对 200,000 个单元格、20,000 个公式和末行结果。它不依赖技术基线 HTML 生成器；现在同样通过公开导入入口和保存重开，但视觉或最终固定环境验收仍未完成。

只读网格的坐标基础已使用稀疏区间前缀和处理行列度量：隐藏的百万行不展开，远端地址仍保留原始坐标，视口元素受独立限额约束。列宽按照[原生字符宽到像素的换算](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.column?view=openxml-3.0.1)，接收正常样式字体的实测最大数字宽度，再返回文档点单位。冻结区域和滚动主体已可分为四个互不重叠的逻辑窗格，共享可见单元格预算；跨隐藏行列的查询拆为实际可见窗口，不读取巨大包围矩形。合并单元格用区间索引查找，各窗格只裁剪其完整几何，内容始终读取同一左上角源；被隐藏或滚出视口的合并起点也不会丢失。点击映射与查询共用该坐标，重复窗格投影只读取一次内容。三种绘图锚点也使用相同行列度量，旋转后的可见边界参与窗格裁剪；负尺寸明确失败，不通过翻转猜测修复。逻辑滚动将原生滚动条轨道限制在固定长度，仍保留整张表的点坐标和末行定位。坐标测试本身不替代字体测量、虚拟 DOM 及完整预览的[读取验收](workbook-reading-acceptance.md)。

原生记录管线现在贯通异步 ZIP、全包关系审计、样式/共享字符串、逐表双遍扫描、维度压缩和 HTML 分块。32 份独立文件及 Calc 重存文件通过这条内部链路及包含批注、绘图和图表的测试包投影；两份一万行来源均保留 200,000 个单元格和 20,000 个公式。XML 全局节点、单元格、公式与输出字节跨表累计，元数据另设保留节点限制；取消、宿主进度回调异常及输入字节隔离有独立检查。投影结束后另核对所有原生部件：可达但未被工作表、表格、绘图或其他来源消费的部件也会失败。开发分支的公开导入已复用这些严格适配，不提供跳过未知内容的部分导入选项；高级外观的完整组合及读取验收仍是本阶段交付合并的前置条件。

核心文件属性中的标题、语言以及作者、主题、说明、修订、日期等来源信息保存在元数据；`properties` 的作者与应用字段不代表宿主身份或权限，日期字符串保留原始精度及时区。应用中的派生统计列表和包缩略图经过结构/资源检查后丢弃，不作为额外文档源；未知属性内容、签名和非空相对链接基址明确拒绝。原生计算链按[计算链说明](https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/working-with-the-calculation-chain)仅视为旧计算次序提示，逐条复核实际公式地址及继承的工作表身份后丢弃。它不能替代依赖图，也不会使任何导入缓存变成有效计算结果。

`conditionalRules` 已有数值比较、标量公式和双色/三色色阶来源；差异外观及色阶颜色引用 CSS，阈值与优先级保留为元数据。`dataValidations` 保存列表、整数、小数、日期、时间、文本长度和自定义公式，也允许仅有提示的 `none`。规则都保存精确多区域和公式原点；严格检查身份、公式数量、样式引用、优先级及有歧义的重叠验证。已有违规数据不会在读取阶段被修正；规则读取现经第一方标量内核计算当前源依赖，返回条件匹配、色阶插值及验证有效／违规／不可用状态；优先级、停止规则和部分属性叠加已接入预览。导入缓存不冒充当前校验，源中的违规值保留并在报告中列出，选择时显示允许的原生输入／错误提示；工作表禁用提示仍保留可查询的验证状态。各编辑入口的统一执行仍待后续阶段。

规则提供者 `WorkbookRuleSource` 与单元格／资源／图表提供者绑定同一修订；含规则工作表缺少该提供者时预览明确失败。内建模型与 Worker 连接自动提供规则读取，窗口请求覆盖用于可见格及共享边框的来源坐标。端点同时限制坐标、规则、依赖值、步数与文字，接受方校验规则顺序及完整性；失败或迟到结果不替换已接受窗口。规则覆盖的公式显示当前派生值，来源缓存状态仍保留，报告的 `ruleFormulaValues` 单独列出当前／不可用数量。公式派生文字不声明虚假的源 UTF-16 偏移。

原生规则验证依据 [SpreadsheetML 条件格式说明](https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/working-with-conditional-formatting)及[数据验证契约](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.datavalidation?view=openxml-2.8.1)，并单独保存[固定 Calc 的探针结果](evidence/xlsx-rule-semantics.json)：传统 `showDropDown="1"` 隐藏列表；多区域公式原点取最左列中最靠上的实际区域起点，不取列表首项或外包矩形角。后者在跨行列样例中分别为 B5/E9，并经外部 ODS 中的引用重定位核对。这是固定 Calc 的独立观测，不冒充 Excel/WPS 实机结论。运行 `python3 scripts/check-xlsx-rule-semantics.py` 重复核对；只有显式 `--record` 更新记录。

原生智能表保留为 `tables` 定义：表身份和公式名称、区域、表头/汇总行、列身份、计算列、汇总函数/公式/标签、筛选、排序状态和表样式关联。按照 [SpreadsheetML 表定义](https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/working-with-tables)，公式使用的原生 `displayName` 映射为 `name`；不同的对象模型名称另存 `objectName`。工作簿名称与表名不能冲突；列名/身份、区域宽高及表/合并/普通筛选的重叠均校验。显示表头时，列名称必须与实际单元格的常量文字一致；不猜测两者谁覆盖谁。计算列模板使用真实首行原点，不把越过工作表边界的原点截回末行。普通筛选保存值/日期选择、标量比较、前后 N 项/百分比和相对日期/平均值条件；导入不重新应用筛选或排序。原生已过滤行的隐藏状态仍由行维度保留。

内建表样式使用原生名称引用，自定义 `tableStyles` 通过差异 CSS 定义区域和条带，不把表样式铺平成逐格格式。条带采用 Office 声明的 [1–9 范围](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/d91a5d31-0bf4-4801-9139-f6c13a690505)，普通前后 N 项筛选采用 [1–500 范围](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/2961ac37-0f8e-47a8-89f9-981403e619b6)，百分比另限 100。源文件和 Calc 重存文件都通过表格读取与保存重开；计算列执行、筛选重算、表格编辑和准确样式渲染仍属后续验收。

表与表列的原生 `headerRowCellStyle`、`dataCellStyle`、`totalsRowCellStyle` 保留为 `headerCellStyleName`、`dataCellStyleName`、`totalCellStyleName`，按工作簿 `cellStyles.names` 解析共享定义，保留实际引用的别名、大小写与同一份 CSS 外观源。依据 ECMA-376 Part 1 §18.5.1.2/3，名称缺失或未解析时回退到表样式；未解析名称与空名称仍保存重开，不静默删掉来源。差异格式的 `vertical`／`horizontal` 内部边框保留为 CSS 槽，由表区域组合器按位置选择；不展开成每个单元格四边。

窗口组合按规范区域顺序应用内建／自定义表样式、表与列的命名样式及差异格式，再保留源单元格相对 Normal 实际改变的属性。完整 XF 的缺省填充／字体不会覆盖表样式；DXF 则按实际声明组合，显式普通字重或无填充仍可清除下层效果。浏览器在源祖先中解析 CSS 优先级、变量与主题，区域顺序不依赖 CSS 文件或 class 列表顺序。原生填充、颜色和单侧边框的关联参数一起保留，普通 CSS 颜色／填充／物理边框也可参与；RTL 中物理侧与原生逻辑侧分别处理。命名定义的活动组只影响显式区域命名引用，不把单元格绑定的未来传播标志当作当前直接格式掩码。

当前显示假设是隐藏／筛选后按可见行重排行条纹，列条纹继续按源列计数；有界前缀索引处理大段隐藏行，不展开整表或随视口重置条纹。外框和条纹边界落在相应可见边缘，表头、汇总及角落仍按原始表结构识别。组合仅修改临时测量节点，失败清理候选，源包保持不变；每窗组合声明另限 8 Mi UTF-16。独立原生探针和自定义源覆盖 LTR／RTL、直接外观、重复默认记录、命名别名／回退、数字格式、清除、条纹及边框；Calc 未显示表样式的限制仍保留，不把这些 SDK 显示检查称作原生视觉验收或完整读取门槛通过。条件格式将在这个属性边界上继续组合。

每个块是完整的无脚本 HTML 文档，包含一个 table/tbody；以下是其内容示意：

```html
<table data-sheet-id="sheet-1" data-row="0" data-column="0"><tbody>
  <tr data-row="0">
    <td data-column="0" data-value-type="text" class="cell-default">编号 001</td>
    <td data-column="1" data-value-type="number" data-value="12.5" data-format-id="general">12.50</td>
    <td data-column="2" data-value-type="number" data-value="13.5" data-formula="=B1+1" data-cache-state="imported">13.5</td>
  </tr>
</tbody></table>
```

常量类型为 blank/text/number/boolean/error。文本 HTML 是内容源，允许 span 与 br；数字/布尔/错误的原始值在 `data-value`，正文为来源显示缓存。公式表达式是计算源，`data-cache-state` 区分 imported/valid/stale/unavailable；当前阶段只解析并保留声明语法和缓存，不声称已重新求值。新建默认自动计算和 1900 日期系统，也可显式选 1904。

未建模 HTML、未知属性、未知样式或格式、重复坐标、错误摘要、无效 UTF-8、非法标量值及不支持的公式语法均失败。ZIP 打包转发同样的工作簿限额。快照和查询结果冻结，输入及序列化字节隔离，索引不保留全表解析树或第二份单元格 JSON。

## 预算与验证

默认文档字节限制保持原值；`workbookLimits` 单独限制实际单元格 250,000、公式 25,000、一次窗口面积 5,000、模板 25,000 和模板键 4,194,304 个 UTF-16 单位。这些是保护限额，不是容量验收结果。一万行基线为 200,000 实际单元格、20,000 公式。

文档的 `maxNodes`/`maxDepth` 对工作簿每个有界块及元数据适用，完整工作簿另外累计检查实际单元格与公式；源块最大嵌套深度为 16。跨块不能重置整包字节、单元格、公式和模板预算。返回的窗口按行列排序，只包含实际存储的单元格，缺失位置由调用者作为空白显示。

验证覆盖工作簿/原有模型的类型分支、严格 JSON Schema、摘要和跨块身份、中文及变体选择符、类型与公式缓存、稀疏最远坐标、样式引用、合并/维度约束、窗口和累计额度、不可变字节、文档包 ZIP 重开，以及工作区之外的打包 JS/浏览器/Node 类型消费。Word/PPT 原有模型、事务和结构测试继续运行。

```sh
pnpm test packages/core/tests/workbook-model.test.ts packages/office/tests/workbook-package.test.ts
pnpm test tests/package-consumption.test.ts
pnpm probe:workbook-model
```

公开入口的一万行[本地测量](evidence/workbook-model-baseline.json)覆盖 20 万实际单元格和 2 万公式：HTML 包打开约 1.9 秒，通过受限窗口扫描全部单元格约 9.0 秒，快照仅 7,215 字节，进程生涯峰值约 444 MiB。该记录没有宣称原生 XLSX、计算引擎或最终固定环境性能已通过。

## 工作簿候选准备（R3-09 实施中）

候选和历史只复用内建模型私有持有的、已经验证且摘要声明未变的块字节；新字节或新声明重新计算摘要，样式绑定／位置变化重新解析，额度仍按当前模型逐项检查。外部包首次读取始终验证全部块。

工作簿会话提供 `prepareAsync(transaction, options?, execution?)`，同步 `prepare` 与异步入口驱动相同的私有候选步骤，候选仍交由共享生命周期的 `commit`／`discard` 管理。异步入口不自动提交；成功准备不改变当前源、修订或历史。完整编辑阶段的剩余范围见[编辑计算交接](workbook-editing-handoff.md)。

`execution` 可传 `signal`、`yield: () => Promise<void>` 和 `onProgress`。进度报告 `operations`／`cells`／`chunks`／`validation`／`calculation`、`completed` 与 `total`；未知总量用 `total: 0`，不同子阶段的计数不合成为整体百分比。单元格写入及稀疏清除按最多 256 个实际处理格让步，块生成、块校验及来源比较逐块让步，规则扫描和既有计算调度也有检查点。宿主为消息取消提供宏任务让步；core 不创建 Worker 或计时器。事务结构解析、授权资源复制和单个有界块的处理仍在对应检查点之间同步执行。

调用者事务、计算时间／地区上下文、验证确认列表及新增资源在第一次让步前捕获。每个会话最多一项异步准备，正在准备的工作占一个候选名额；取消、失败或释放归还名额且不发布私有源。让步期间若其他操作提交了新源，准备以 `RevisionConflict` 失败；释放或取消以 `OperationCancelled` 失败。进度回调失败不影响文档结果。返回候选后宿主仍需在提交前检查自己的取消状态，取消时丢弃候选；已经提交的事务只能通过共同历史显式撤销。

当前直连小型编辑器的填充及替换已接入异步准备；默认编辑器 Worker 路径、完整结构操作和固定资源大表编辑门槛仍未完成。增量预览基础见下节，不代表完整容量验收。

## Worker 会话通道（R3 编辑阶段实施中）

editor 的 `serveWorkbookSession(port, session, options?)` 绑定宿主已经创建的唯一核心会话，`connectWorkbookSession(port, options?, signal?)` 只连接这条专属授权端口。文档身份与修订用于冲突检查，不能选择另一个会话。宿主创建／终止 Worker、提供原生文件和身份权限；SDK 关闭连接时释放所属候选与视图，保留宿主会话。传输标识为 `doffi/workbook-session/1`。

客户端 `prepare` 接收原始类型事务或输入／TSV 粘贴／复制／填充／替换命令。后者只传文字、范围与选项，在 Worker 中生成操作并经过同一 `prepareAsync`；候选返回只读影响摘要，原始源块、模型、计算图和历史不传到主线程。`commit`／`discard` 接受本连接实际返回的候选，`undo`／`redo` 调用原核心历史。映射采用 JSON 节点对数组，工作簿结构步骤保持原有顺序。`find` 接收当前显示值查找所需的宿主时间和地区；`state()` 与订阅返回修订及历史状态；不构造第二个客户端核心会话。

`readInput(sheetId, at, candidate?, execution?)` 从当前来源或本连接实际候选读取一格可逆输入文字，最多 32,768 个 UTF-16 单位（包含必要的字面量引号），不返回整块 HTML。未知表、错配坐标／修订、外来／失效候选和超额响应明确失败。

`view(candidate?, options?)` 返回调用者负责释放的不可变读取连接，复用已有窗口／资源／图表／规则读取通道。当前源使用会话的增量计算，候选计算保持隔离；自动／手动模式决定缺省是否计算，可用 `calculate` 和宿主 `calculationContext` 显式选择。计算视图在握手中声明当前值能力，规则和图表读取同一完成批次；原始单元格及导入缓存不被改写。新视图可在准备好后更新预览，再释放旧视图。候选被丢弃或失效时，其视图随之关闭。

`options.previous` 可传本连接仍保留的视图，Worker 依据其实际不可变模型和完成计算批次生成单元格变化提示。复制对象、其他连接视图和已释放视图不能代替；提示超出消息额度时省略，预览回到完整测量。更新完成后再释放旧连接，保留视图仍受数量预算限制。

默认限制为 8 个进行中的请求、32 MiB 消息预算、128 层、16 个保留视图和 60 秒请求期限，读取额度另按已有通道设置。准备的取消等待远端清理确认；取消发生在结果送出但尚未接收时，也回收对应候选／端口。提交请求发出后不提供取消；响应超时关闭连接并明确报告结果未知，宿主须检查实际会话后再决定后续操作。已提交但回执超出额度也明确报告实际修订，不能把它当作未修改。

## 编辑后的增量预览（R3 编辑阶段实施中）

core 的 `readWorkbookChangeImpact(previous, current, previousCalculation?, calculation?)` 区分直接源修改和计算结果变化。计算批次须来自实际模型，缺少一侧批次时保守标记公式格，不以导入缓存冒充完成值。此接口仅报告单元格，不授予访问，也不代替元数据、CSS 与辅助内容比较。

直连适配 `createWorkbookPreviewSource(model, calculation?, previousSource?)` 和上述 Worker 会话视图可为预览绑定准确的前后读取提供者；相同文档 ID／修订不能选择旧行高缓存，时间上下文变化也可能在同修订产生不同结果。宿主自定义读取适配器可用 `linkWorkbookPreviewSource(next, previous, impact)` 承诺完整的来源及计算变化；适配器负责正确性，传入未经绑定的 `impact` 字段本身不会启用缓存。弱引用避免保留历次来源链；不支持弱引用的浏览器保留完整路径。

仅在元数据、CSS、辅助内容、视口及字体身份不变时复用脚本隔离 surface。类选择器样式且没有复杂字体层、条件格式或表格的健康测量可保留未变化行高；源修改及远处公式结果所在行重新读取、测量，包括按未隐藏宽度参与行高的隐藏列。字体变化、替代／缺字／不完整审计、行内字体或复杂样式保守走完整测量。新窗口先独立完成，成功后一次发布；失败、取消或迟到响应不移走旧画面。既有滚动预读和准备优化保持。

## Worker 远端输入（R3 编辑阶段实施中）

`createWorkbookInputTransaction` 与 `readWorkbookInputText` 共用直接核心和 Worker 的输入拼写、保守类型转换及精确富文本替换；不变输入保持原类型及源字节。连续 CR／LF 数字实体在共享安全 HTML 适配器正确保留，其他非法控制实体仍失败。

`createWorkbookRemoteInputController(client)` 在主线程只保存草稿、模式、组合状态及原修订选择。`select` 和 `apply` 为异步操作；`apply` 在 Worker 准备候选，再读取其规范输入拼写，最后调用同一核心提交。准备期间继续键入或释放控制器会取消并回收候选；迟到选择不能覆盖已开始的草稿。外部源变化导致修订冲突并保留输入。提交已发出后等待真实回执，此时不能把取消当作撤销；释放控制器不关闭宿主客户端、Worker 或核心会话。

`bindWorkbookAsyncTextInput` 与既有同步绑定共用原生 textarea 的 CR 映射、IME、Enter／Tab／Alt+Enter／Escape／F2 及粘贴判断，异步绑定等提交成功后才通知导航。`applying`／`selecting` 状态供宿主管理输入可用性。以上是默认编辑器的基础，完整编辑器／Playground 界面尚待接入。

Worker 会话的内部让步使用独立消息任务队列，给取消和其他端口请求处理机会，避免连续嵌套计时器的最小延迟。队列不承载文档数据，每个端点独立；释放端点时关闭队列并唤醒等待者执行既有取消／释放守卫。

## 工作簿批注正文

工作簿可选 `manifest.content` 指向同一包内的辅助 HTML，使用 `data-office-kind="workbook-content"` 根节点。它仅包含有界的富文本批注 story，不拼接单元格块，也不创建另一份工作表内容。`snapshot.content` 和 `getNode` 提供正文查询；每个元素有稳定身份，复用无脚本 HTML/CSS 校验。正文和回复各自拥有一个 story；孤立正文、重复归属、未知类名、未来修订和重叠锚点区域明确拒绝。元数据和批注记录共用累计条目限额。

工作表删除后，显式 `needs_reanchor` 的批注可以保留原坐标、正文和回复，等待后续事务重定位；有效批注仍必须指向现有表。Word/PPT 不接受工作簿的 `cells` 锚点，包括需要恢复的项。当前仅实现读取及保存重开；事务映射、完整原生批注验收和交互验收仍在对应任务内。

现代批注的内部读取现在校验工作簿人员、全局 GUID、按表回复图、原生可选时间、正文及解决状态；嵌套回复的父身份保存在格式元数据中。正文仍是辅助 HTML 的唯一可编辑来源。身份、作者显示名和回复关系在打开 HTML 包时重新校验，重复、循环、跨线程父引用及未来修订不接受；导入映射按条目核算重复作者文本后的 JSON 字节，防止共享长名称放大输出。32 份内部记录链路继续通过；固定 Calc 重存的现代线程丢失仍作为外部差异记录。

经典兼容占位备注的核对及缺失定位恢复已在内部管线实现；备注框预览和完整交互仍未完成，不能据此宣称两类原生批注已验收。现代提及现在按同一包中的人员身份、原生 GUID 和 UTF-16 范围保留，详见下文；未知扩展仍明确拒绝。完整读取边界仍需最终复核。格式依据为 Microsoft 的 [CT_ThreadedComment](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/42f9b03d-9662-4204-9783-dbeb324a691c)、[CT_Person](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/6274371e-7c5c-46e3-b661-cbeb4abfe968) 和[兼容占位备注](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/6383f002-c90b-401c-a1d7-66b97b14cb3e)。

经典备注的内部记录也已读取作者、单元格、富文本和兼容线程标记，保留待处理的原生几何属性。线程身份、工作表和位置一致的兼容备注按照[原生占位规则](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/6383f002-c90b-401c-a1d7-66b97b14cb3e)作为缓存消费，不重复显示现代线程；真正的经典备注仍保留。缺失、移位或复制的占位情况保留现代线程的原位置、正文、回复和原解决状态，标记 `needs_reanchor` 并保存具体恢复原因，不自动迁移或复制线程。原生根批注没有位置时，仅在此状态允许空单元格范围，不补造 A1；正常锚点和请求仍要求非空范围。无法匹配的兼容备注保留其完整经典正文与原线程身份，等待恢复，不凭缓存合成现代线程。完全没有旧版备注部件的独立线程继续可读。严格契约、JSON Schema 及保存重开共同核对恢复状态和来源。独立来源的加粗中文与 Calc 重存的普通字重分别按实际来源保存，不互相补造。VML 几何与基础外观现已接入内部管线；[固定 Calc 探针](evidence/xlsx-note-semantics.json)验证其单元格锚点优先于 VML 位置数值，验证脚本为 `python scripts/check-xlsx-note-semantics.py`，仅 `--record` 更新证据。固定 Office 归档和字体清单均检查摘要；此探针不代表 Excel/WPS 或 SDK 视觉验收。

绘图锚点基础契约区分绝对坐标、单端单元格加尺寸和双端单元格锚点。统一使用文档点单位，保留负偏移、超过一行的偏移以及移动/缩放标志；DrawingML EMU/通用长度和 VML 像素在适配边界转换。三类图表、图片、形状及文本框的 12 个独立源/Calc 锚点已通过定向解析检查，零尺寸保留原值；对象与图表映射已接入内部包投影，预览仍在实施。VML 像素解释依据 [Anchor](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.vml.spreadsheet.anchor?view=openxml-3.0.1)；坐标结构和数值界限按仓库固定的官方 Transitional XSD 检查。

工作簿 `objects` 已支持图片、16 个共享预设和文本框的严格源契约。元数据保存工作表、锚点、原生身份和打印/锁定标志，独立 HTML/CSS 保存正文、资源、外观、旋转翻转和裁剪；不重复保存一份由行列推导的位置。对象根必须是辅助内容的独立子节点，正文属于对象内的段落；校验跨表身份、预设调整、裁剪可见区域、资源摘要及累计预算。锁定标志是文件属性，不授予或取消宿主权限。

内部辅助内容索引只解析已校验的 `snapshot.content`，不读取或序列化单元格块。默认限制 8,388,608 个 UTF-16 单位、100,000 个源节点和 64 层嵌套；图片地址在惰性模板文档中移除，连接 DOM 前解析为清单路径，之后由授权资源适配提供 Blob URL。测量投影仅连接当前请求的对象／批注子树，保留源祖先、顺序、身份和精确 UTF-16 正文，另限 5,000 个根及 50,000 个节点。候选构建失败不连接半成品；投影独立释放，索引销毁时释放全部所属投影。三浏览器检查了 8 份独立／Calc 图表、图片形状和两类批注来源，以及空辅助内容、超限和无资源请求；该投影仍不代表对象画面或完整阅读验收已完成。

图片内容层和窗格片段现已分离：裁剪使用完整源比例，旋转／翻转围绕完整对象中心，各冻结窗格仅裁剪同一几何；RTL 改变物理锚点位置，不自动翻转图片内容。绘图可见性支持描边／效果的保守外扩边界。候选图片在连接 DOM 前解码，最多同时请求 8 项；失败或取消清理候选源地址，已连接画面不能传入该候选解码接口。内容层默认最多 5,000 个图片实例、裁剪放大倍数最多 1,000,000；组合层还检查放大后的实际点尺寸，超出测量预算明确失败。原始／Calc 图片的四窗格、旋转、显式翻转、裁剪、取消和失败隔离已通过三浏览器定向检查。Calc 将样例水平裁剪从 0.1 重存为 0.10035、垂直裁剪从 0.05 重存为 0.05039，按实际来源保留；最终原生页面对照仍待完成。

预设路径的 SVG 生成现由 PPT 与工作簿共用，包含基本填充、描边、渐变和箭头的节点预算；多路径图形的箭头使用不同 SVG 身份。工作簿适配保留自己的锚点和源所有权，不构造伪造的 PPT 模型。原始／Calc 的形状颜色及共享渐变、箭头边界已通过三浏览器测试；PPT 的 16 个预设、组变换、连接线、文字自适应、主题颜色和既有截图回归保持通过。零宽／零高线条有描边外扩时继续参与裁剪。复合线、特殊渐变参数、阴影及工作簿文字的最终绘制仍需单独核对，这次共享不等同于完整 DrawingML 外观验收。

共享线性渐变按实际填充区域计算物理方向和完整色标跨度，区分固定角度与随宽高缩放的 `scaled` 标志；文字框自动扩展时重新计算，图表柱条、扇区和图例使用各自区域。非正方形区域的斜向渐变已通过三浏览器像素检查，现有形状／分组／图表回归及 PPT 导出重开检查通过。原生缺省属性保持可继承，最终缺省角度为 0、`scaled=false`，依据 [ECMA 渐变方向说明](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.lineargradientfill.scaled?view=openxml-3.0.1)及 [MS-OI29500 §2.1.1284](https://officeprotocoldocs-f5hpbjgea6b8gneq.b02.azurefd.net/files/MS-OI29500/%5BMS-OI29500%5D.pdf)。旋转／翻转现在按下面的填充模式合成，完整外观矩阵仍待验收。

绘图渐变按 [gradFill 定义](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.gradientfill?view=openxml-3.0.1)保留原生开关和角度。用户已选择 `rotWithShape=true` 时保持原填充随形状一起旋转及原覆盖范围；不按 Calc 重新铺色。false 时，在页坐标的旋转框包围矩形内计算渐变，再将方向反变换到对象局部空间；单翻转会共轭旋转方向，显式反射保留。几何和图表适配都将所属对象方向传给共享填充层，涵盖绘图区、系列、扇区、图例及标记；PPT 自动扩展后也重新计算。所有坐标要求有限，色标节点继续计入同一候选预算，未知旋转标志明确失败。

[32 组独立来源](evidence/xlsx-gradient-rotation.json)记录源／Calc 重存属性和分别重开后的 PDF 色带矢量，包含角度、缩放、单双翻转及省略开关。Calc 将随形状模式改存为固定模式并加上旋转角，删除 `scaled`，部分翻转组合重开会改变颜色方向；另有 90° 变换尺寸变化，锚点、变换尺寸及原生宽高比分别记录。SDK 保留实际来源，不恢复生产者丢失的属性。显示检查分别核对原始旋转覆盖策略、固定模式的方向／原生相同语义矢量及实际颜色采样，不把有差异的项计为保真通过，也不以归一化填充方向代替整个对象的几何验收。Excel/WPS 尚未实机验证。

同一原生继承规则也用于描边：`ln/@w` 缺省时不覆盖主题线宽，完全无继承值时使用 0.75pt；显式零宽保持为零。工作簿主题继承及零宽通过三浏览器检查，共享 PPT 导入、原生导出重开与形状／分组／连接线／图表回归通过。已保存包中的显式线宽保持原值。

形状／文本框已有独立的有界富文本投影：在源祖先仍连接时读取样式，生成不携带源 CSS 类和 DOM 身份的显示节点；共享单元格的字体角色分段，保留跨 span 的组合字符／变体选择符及带样式换行。每个可见文本片段记录原节点身份和对象内 UTF-16 区间，段落之间不额外插入字符；源结构空白占用的偏移仍保留。原始／Calc 两份绘图来源、局部字体诊断、精确节点预算及失败后源不变已通过三浏览器检查。

内部批量排版已接入预设文字矩形、换行、列表、局部字体解析、字号缩小、框体扩展和行距缩减；文字抵消对象翻转，外层继续负责对象旋转。默认限制 50,000 个节点、8 次排版和 1,000,000pt 对象边长；测量临时节点也计入预算，后续对象失败会释放前面候选的附着关系。原生缺省内边距为左右 7.2pt、上下 3.6pt，缺省字号为 18pt；共享源类早于局部声明，显式零边距和字号继续覆盖。依据为 [left inset](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.bodyproperties.leftinset?view=openxml-3.0.1)、[top inset](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.bodyproperties.topinset?view=openxml-3.0.1)及 MS-OI29500 §2.1.1399。Calc 样例将内边距重存为 90000／45000 EMU，按该来源保留。三浏览器检查覆盖自适应、翻转、行距、列表、过大内边距及失败清理；PPT 富文本、自适应和预算回归通过。`anchorCtr=true` 的独立文字块居中已实现，四引擎保持段落自身的对齐；完整页面和定向原生对照由[读取验收](workbook-reading-acceptance.md)汇总。

原生 `horzOverflow=overflow/clip` 与 `vertOverflow=overflow/clip/ellipsis` 保存在对象或图表文字样式的 `--office-text-horizontal-overflow`／`--office-text-vertical-overflow`，缺省均为 overflow。依据 ECMA-376 Part 1 §20.1.10.84，裁切分别服从左右或上下边界；不使用会把另一轴 visible 改为 auto 的 CSS 混合 overflow。私有段落以多边形裁切，完整原文字节点及 UTF-16 投影保留。省略号是无源偏移的派生标记：底部隐藏时放在最后完整可见行尾，顶部隐藏时放在首个完整可见行首；无完整行可容纳时只显示被边界裁切的派生标记。按整段字素分割，包括跨富文本节点的组合字符及变体选择符；浮点字体框与本地行距采用同自动适应的一 CSS 像素取整容差。反射在裁切测量后施加，标记也接受字体缺字诊断和累计节点预算。单体至多 100,000 次几何探测，单个需要省略的段落至多 1,600,000 个 UTF-16 单位；超限失败并释放候选。

[24 组独立源探针](evidence/xlsx-text-overflow.json)覆盖两种横向模式、三种纵向模式、两种换行和顶部／底部锚定，`python3 scripts/check-xlsx-text-overflow.py` 固定归档、字体和摘要重核。Calc 的 PDF 字符提取可能包含裁切字符；该证据单独记录字形 use 数与重存属性，不作为绘图保真通过项。SDK 的实际像素边界、最后完整可见行、源投影、混合字体、RTL、反射、小于一行的尺寸及失败清理分别在浏览器检查。

[27 项独立原生文字探针](evidence/xlsx-drawing-text-semantics.json)记录字号、段落末尾样式、空段落、开头／末尾／连续换行、内边距和默认／级别段落样式；`python3 scripts/check-xlsx-drawing-text.py` 按固定归档、字体及源摘要重核。该 Calc 的无继承字号为 12pt，缺省内边距为 90000／45000 EMU，与上述 Office 缺省值不同。非空段落的末尾字号不改变这些样例的行距，重存时被改为实际文字字号；空段落则保留末尾字号并影响后续文字位置。空行继承相应换行处的字号，末尾空行再由段落末尾样式覆盖；换行字号不撑高前一个有文字的行。Calc 采用 `defPPr`，MS-OI29500 §2.1.1387 则说明 Office 忽略它；该差异仍需在导入和预览门槛中处理。零内边距时字形边界可以略超出图形，证据保留有符号坐标。[段落末尾属性](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.endparagraphrunproperties?view=openxml-3.0.1)的原始状态继续保存在源 CSS 中。

SDK 对其中 22 项明确字号且语义相同的样例执行三浏览器行距及文字边界对照，误差限分别为 0.75pt 和 1.5pt；其余 5 项生产者缺省差异单独保留，不混入保真通过项。默认行距用高精度字体边界测量，避免逐行累积浏览器小字号取整误差；不支持字体边界的引擎使用有界 DOM 测量，并通过 `usedFallbackLineMetrics` 标识回退，不能据此宣称相同原生误差。空行由不参与源投影的视图节点承载，段落默认字号和换行插入字号不额外抬高实际文字行。测试同时检查精确 UTF-16 投影、临时节点预算、源恢复和失败清理；固定旧 Chromium 也通过原生对照及强制回退检查。Excel/WPS 和整页绘图文字验收仍未完成。

字体测量还校验 Canvas 是否真正接受当前字体简写；合法 CSS 的 `all-small-caps` 等值超出该简写语法时，改用上述 DOM 测量，不沿用前一段文字的字体。依据 [Canvas 字体设置规则](https://html.spec.whatwg.org/multipage/canvas.html#dom-context-2d-font)，失败赋值不改变已有字体。三引擎的真实简写拒绝与缺失字体边界两种路径均检查行高、空行、临时预算和失败清理。

`anchorCtr` 按 ECMA-376 Part 1 §21.1.2.1.1 将最小文字块居中，保留块内段落对齐及独立的上／中／下锚定。显示块使用有原矩形上限的固有宽度，自动调整完成后再测量其水平位置，不把文字改为统一居中段落。[24 项独立锚定探针](evidence/xlsx-text-anchoring.json)包含左／中／右及混合段落对齐，固定 Calc 与三种现代浏览器及旧 Chromium 的文字边界误差小于 1.5pt；源摘要、字体及生产者版本由 `python3 scripts/check-xlsx-text-anchoring.py` 重核。原生输入也通过官方 Transitional XSD、关系审计和 SDK 包保存重开。该局部对照不代表整页绘图验收。共享 DrawingML 列表另保留显式零缩进和零左边距，不以缺省值覆盖它们。

内部预览表面为每个候选创建独立、无脚本的 iframe。源 CSS 保留 `:root/html/body` 的字体和变量继承；画布外框的尺寸、边距由 SDK 控制，源节点的 CSS 仍按原祖先解析。生成的网格／绘图视图置于单独的 Shadow DOM，源类型选择器和 `!important` 不会重新匹配这些显示节点。表面缺省限制 CSS 为 4 Mi UTF-16 单位，准备阶段保持隐藏并支持取消，加载期限为 30 秒；上层组合器负责整窗成功后的替换。失败与释放清空源测量、显示节点和样式，保留的表面句柄也不会继续持有这些子树。三浏览器及固定旧 Chromium 检查宿主样式、源强制样式、真实文字／形状布局、并存工作簿、取消及释放，并从实际页面截图核对隔离视图内的 SVG 渐变引用。

隔离视图的选区适配使用明确端点设置选择；读取时优先采用支持 Shadow DOM 的 [`getComposedRanges`](https://www.w3.org/TR/selection-api/#dom-selection-getcomposedranges)，旧 Chromium 沿用其真实 DOM 范围。WebKit 的普通 `getRangeAt()` 会返回外层容器，不能直接用于 UTF-16 投影；适配层只接受两个端点都属于当前视图的单一区间，不猜测外层节点或其他视图的对应文字。反向划选、原生选择文字、精确端点、越界失败后原选择不变均有检查。隔离层已用于下述公共工作簿预览；正文 UTF-16 划选仍是内部适配。

[32 项原生阴影方向探针](evidence/xlsx-drawing-shadow-semantics.json)覆盖零模糊、形状旋转、水平／垂直翻转、两个阴影方向和 `rotWithShape` 两种值；运行 `python3 scripts/check-xlsx-drawing-effects.py` 可按固定 Office 归档和字体重核。该 Calc 版本始终保持阴影的页面方向，重存后写入 `rotWithShape="0"`。[格式定义](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.outershadow?view=openxml-3.0.1)则明确区分是否随形状旋转；因此不能把这一生产者差异作为 SDK 忽略源标志的依据。内部阴影适配将页面方向通过对象逆变换映射到局部过滤器，包含反射；随对象模式保留局部方向。三浏览器中的实际对象变换与原生的 16 项页面方向结果相差小于 0.02pt，另外 16 项检查源标志和对象变换语义，不宣称与该 Calc 差异项一致。

[四组原生模糊轮廓](evidence/xlsx-shadow-blur-semantics.json)使用 0／4／12／24pt 模糊半径、固定归档和字体，并以 144dpi PDF 栅格记录红色阴影的原始边缘像素；运行 `python3 scripts/check-xlsx-shadow-blur.py` 可重核。SDK 内部显示采用半径一半作为 CSS 高斯标准差，`approximateBlur` 明确标识非零模糊近似，保留原生参数而不改写源。原生与浏览器先各自测量锐利边缘的栅格取整位置（均小于 1pt），再对齐比较模糊轮廓；三浏览器实际页面截图的最大通道误差为 21／255，检查上限为 24／255。这不是逐像素等同或 Excel/WPS 的证明。

阴影在连接显示树前检查几何及栅格预算，单对象缺省最多 16,777,216 像素，按实际像素比例和保守外扩范围计量；无阴影不占该预算。调用方需提供包含描边、文字的绘制边界，累加整个候选的像素，并将返回外扩范围接入窗格裁剪。无效单位、非有限值、非法标志和超限明确失败，查询不改源。绘图组合及窗口累计效果预算已接入后述窗口控制器；最终页面与 Excel/WPS 验收仍未完成。

原生图片、预设形状和文本框的内部适配已通过独立第 08 号来源及 Calc 重存的包重开检查：原始图片字节、矩形裁剪、旋转、三类锚点、精确 UTF-16 正文、主题引用、段落/字符样式和换行样式均进入同一来源。对象和文字超链接绑定其 HTML 节点，沿用工作簿的安全外链/内部地址校验。背景填充引用按原生 `fillRef` 的 1001 起始编号解析，依据 [FillReference](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.fillreference?view=openxml-3.0.1)。Calc 把样例文本框改存为普通矩形，读取器保留该来源差异；这些来源已接通开发分支的公开导入；对象显示及原生差异见[读取验收](workbook-reading-acceptance.md)。

图表的源契约与原生引用读取已接入工作表绘图层。系列数值、分类、系列名称和引用标题保留公式；原生缓存单独标为 `imported`，保存稀疏索引、缺值、错误值及整体/单点数字格式，不转成独立可写数据表。没有缓存时不补零或假定结果有效。缓存声明长度与元数据字节跨图表累计计量。六个独立/Calc 图表来源的 30 条引用、标题正文、图例、轴、系列及基础标签外观均通过包保存重开。轴标题同样支持独立富文本或单元格引用；格式、样式和正文所有权在核心再次校验。

原生元素缺省属性按固定官方 XSD 解析，包括百分比间距和标签偏移；可见数据、按点着色和超出最大值的标签开关保留元素是否存在，避免把未指定误当成显式设置。Calc 输出的标签引导线扩展按 [MS-ODRAWXML Datalabels](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-odrawxml/80e97775-804a-4f7e-90a1-8e87e1734d83)及 [showLeaderLines](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.office2013.drawing.chart.showleaderlines?view=openxml-3.0.1)映射；未知、重复和相互矛盾的扩展失败。当前单元格查询驱动的图表显示已接通；完整显示矩阵及视觉边界见[读取验收](workbook-reading-acceptance.md)。无工作表引用的 `strLit`／`numLit` 已映射为独立的 `kind: literal` 来源分支，保留声明长度、稀疏点、空值、错误值及向量／点数字格式；不能同时携带引用表达式、缓存或计算状态。核心重新校验类型、严格递增的点索引和累计声明长度；多个系列的分类来源须一致。保存重开和内部图表读取保留唯一常量来源，不产生伪造单元格或公式。常量与引用共用读取额度，所有来源在读取单元格前核算点数，常量文字按批检查取消；完整图表显示范围与原生差异见[读取验收](workbook-reading-acceptance.md)。

内部图表数据读取器已通过公开模型边界解析分类、系列数值／名称、标题和轴标题引用；名称以所属工作表 A1 为原点，不受当前活动表影响。返回常量或当前求值的类型值、隐藏标志和继承的数字格式身份；公式成功标为 `valid`，不可求值返回 `unavailable` 及类型化原因，不读取其旧缓存补值。来源中的缓存状态不因此改变；多区域顺序、重复坐标、矩形形状及稀疏空白均保留。动态引用、循环名称与无效地址返回明确状态，图表导入缓存继续单独保存在源快照中，不替代这些结果。默认整批限 128 个图表、5,000 条引用、50,000 个坐标和 8 Mi UTF-16 单位；重复出现的值重复计量。公式另限累计 250,000 个依赖单元格读取及 5,000,000 步；同一请求重复引用同坐标复用计算结果，不跨请求共享。求值器在成功、失败或取消时释放，图表字面向量继续只从自身唯一来源读取，不生成伪单元格。完整引用预算在第一次读取单元格前检查，解析每 32 条引用、读取每约 256 个坐标让出任务并检查取消；每次同步窗口查询最多 256 个坐标。14 项检查覆盖取消、修订／文档身份、请求隔离、类型与状态、精确范围和独立／Calc 三类图表。此层不执行数值强制转换、隐藏点筛除或文字格式化，尚不构成图表绘制验收。

无显式填充的图表点／系列现在使用所属文档的六个 accent 主题色。原生 varyColors 的明确 false 会关闭变化；明确 true 仅在原定义只有一个系列时按源点索引变化，隐藏后恰好剩下一个可见系列不改变该判定，依据 [MS-OI29500 §2.1.1571](https://officeprotocoldocs-f5hpbjgea6b8gneq.b02.azurefd.net/files/MS-OI29500/%5BMS-OI29500%5D.pdf)。未指定仍保留现有基本显示缺省：饼图按点，柱／线按系列；元数据不补造该标志。显式系列／点样式继续先于自动色，隐藏点压缩后不重排颜色。饼图及按点变化的单系列使用分类图例，其他图例使用系列名称，键色跟随相应点／系列。

[26 项独立配色探针](evidence/xlsx-chart-colors.json)覆盖默认／替换主题、三类图表、单／多系列、变化开关及显式系列／点色；`python3 scripts/check-xlsx-chart-colors.py` 重核源摘要、固定 Calc 结构、PDF 向量颜色集合和文字计数，不比较字体几何或跨工具像素。16 项柱／饼来源的 SVG 显示颜色集合与独立 PDF 一致，26 项分类／系列文字计数一致；图例键在 SDK 中继续反映源点外观。Calc 的变化折线标记色、显式系列色的变化图例存在差异，不能计作颜色保真通过；正式配色来源的折线重存仅增加逐点外观，原本即可重开；早期探索来源引入的逐点 marker 已在下述批次补齐，不能重复计作正式配色来源的新修复。完整门槛仍未完成。

[独立公式图表来源](../packages/office/tests/fixtures/xlsx-chart-formulas.xlsx)故意携带错误的单元格／图表缓存；[固定 Calc 当前值](evidence/xlsx-chart-formulas.json)使用相同公式但省略缓存的参考输入，避免 Calc 保留错误缓存。25 个公式值覆盖名称、跨表、依赖链、文本、逻辑、空串及错误；公开导入、包重开、传输与图表绑定均保持源缓存不变。此证据是当前值读取检查，完整计算及图表视觉仍有后续门槛。

三类图表的 SVG 绘制已提取为共享模块，PPT 常量适配与工作簿引用读取保持各自职责。工作簿内部绘制适配从源祖先解析样式，生成可放入隔离视图的节点；调用方必须显式提供数字格式化策略和可见标题的派生视图。默认限制 50,000 个节点、50,000 个分类／系列值及累计 8 Mi UTF-16 源／显示文字；文本节点和临时样式探针计入节点峰值，异常时清理源探针和未完成视图。刻度使用有界索引，无法在当前浮点数尺度递增时明确失败，即使坐标轴隐藏也不会持续循环。有限数据先按值域及可见绘制域裁剪，避免无穷坐标；无法在预算内保留渐变覆盖的极端情况明确失败，见下文有限轴探针。三浏览器及固定旧 Chromium 已检查六个独立／Calc 图表的当前常量、基础尺寸与样式隔离、精确节点预算及失败清理，PPT 图表和打印预算回归通过。这里使用受控字体／颜色和数值格式测试策略；后续内部组合层已接入标题、混合类型与隐藏点，原生定向比较和完整工作表合成见[读取矩阵](workbook-reading-acceptance.md)。

[15 项独立图表数据探针](evidence/xlsx-chart-value-semantics.json)由 `python3 scripts/check-xlsx-chart-values.py` 核对，只有 `--record` 更新证据。源图表缓存故意使用 `999` 和过期名称，工作表则包含数字、数字文本、普通文本、布尔、隐式／显式空白、空字符串公式及错误值；另覆盖三种空白显示方式和隐藏行／分类列／数值列。固定 Calc 的重存数值缓存保留数字及布尔转换，省略其余类型；隐藏行可压缩缓存索引，隐藏数值列可产生零长度向量。隐藏分类列的特定案例仍保留并绘制过期缓存，该结果仅作为生产者差异记录，SDK 不仿照此行为。读取器的新增检查验证全部 15 份来源继续返回当前类型值和源坐标，不借用图表旧缓存。混合类型已接入下述当前值组合层，完整图表已接统一窗口组合；内部绘制已按用户选择以绝对值分配饼图扇区，数值标签使用原始带符号值、百分比使用扇区占比，并返回 `negative-pie-magnitude` 诊断及负值数量。比例先按最大绝对值归一化，避免正负抵消或极大数值求和溢出；零／空白不生成扇区。三引擎对照绝对值几何，覆盖混合正负、全负单点、相互抵消、极大／极小数和全零；该检查不等于完整图表页面验收。

富文本图表标题与轴标题已共享形状文字的源样式／字体角色物化代码，按各自正文身份保留精确 UTF-16 投影；只能读取所属图表的直接标题正文。单元格引用标题继续由引用数据与显示格式适配处理。六个独立／Calc 标题的投影和原形状文字的样式、换行、字体行高、自动适应及回退预算共 15 项三浏览器检查通过；此结果不代表图表标题定位与原生排版已经验收。

标题正文现在也复用同一批量排版核心：图表适配提供真实标题正文及其独立的自动适应、换行、字号和行距属性，形状适配提供预设文字区域与反射补偿。所有正文累计节点预算，后续正文失败时释放整批候选；标题保持自身源身份。独立／Calc 标题及新增双轴富文本来源检查本地字体、UTF-16、不同字号缩放、长标题扩展／缩小和失败清理。后续内部组合层已接入标题位置与轴标题旋转；原生页面差异和固定字体定向检查见[读取验收](workbook-reading-acceptance.md)。

经典备注框将单元格锚点、显示/打印和随行列变化标志保存在 `comments[].box`，实体填充、描边、箭头样式、单层阴影、内边距及对齐映射到受校验的 CSS 类。原生 VML 不复制进 HTML。普通 `commentPr` 坐标亦可单独读取。两份来源同时存在时按用户决定选择 VML：`comments[].box` 是唯一显示框，`boxSource` 记录来源；`commentPrBox` 只保留未选中的原生框，不能覆盖显示框。`boxDifferences` 记录锚点、外观、可见／打印、锁定、自动属性及替代文字的差异；外观按实际声明比较，不把生成的 CSS 类名当作原生差异。单独来源也记录 `boxSource`；既有未带来源字段的单框包继续可读。两份来源均经过严格结构与属性校验，备用框不能绕过校验，样式引用和差异字段同样接受包级所有权／唯一性检查，并计入累计元数据及语义额度。未知控件、宏字段、非备注形状、重复定位、孤立形状和矛盾的显示状态明确失败。后者的 Calc 行为另有固定探针记录，不推断 Excel/WPS 相同。两类批注共用累计输出字节核算，源包重开会复核作者、正文、回复和备注框样式引用。32 份内部记录及公开导入检查已包含这些基础映射，仍不代替只读网格验收。


[4 组独立双框／单来源样例](evidence/xlsx-dual-note-boxes.json)覆盖冲突锚点、相同锚点、VML 显示状态及只有 commentPr。源文件和固定 Calc 重存分别通过官方 XSD、关系检查和公开导入／包精确重开；Calc 会删除 commentPr，单独来源时还会生成另一份 VML，因此源保留和生产者重开单独验证，不能当作无损保真。微软实现说明称 Excel 遇到 commentPr 会拒绝打开；VML 优先来自用户选择，并非根据该说明推断出的 Excel 优先级。批注面板列出差异类别并说明当前来源，正文投影仍保持原 UTF-16；冻结的专用 Worker 快照保留两份框和差异列表。

当前窗口按所选原生框显示 visible 的经典备注：与网格共用测量后的行列尺寸，在冻结的各窗格裁切；隐藏备注及现代线程继续由只读面板打开。实体填充／透明度、描边／虚线／端点箭头、单层无模糊阴影、内边距及水平／垂直对齐均由派生视图绘制；分散排版按实际成形行定位，保留跨 span 的 UTF-16 和变体序列。原生导入的备注段落显式使用零外边距，避免浏览器默认段落间距进入原生框。打印、锁定及自动属性继续保留为来源，预览不据此开放编辑。每个窗格使用独立文字节点和源投影，字体在隔离视图内解析；源节点数、显示节点数、正文及重复文字、对象数和有限尺寸分别受累计预算保护。测量探针、失败候选及销毁均释放，分批构建可取消。来源／生产者及受控 LTR／RTL、五种垂直对齐、颜色像素、节点核算和失败保留已纳入现代三引擎与旧 Chromium 检查；整页原生视觉矩阵仍单独验收。


内部图表当前值组合层现将上述类型查询转换为绘图数据：数字保持原值，布尔转换为 1／0，文本和普通空白保留缺点状态；错误值单独标记断点，不能因“空白按零显示”而变成零。内建读取器先通过一方标量内核计算当前单元格公式；不支持、循环或需后续计算的公式及未解析／错误引用返回整个图表不可用，保留明确诊断，始终不读取单元格或图表的旧缓存补值。自定义宿主提供方仍可显式报告 `imported/stale/unavailable`，显示层保留对应诊断；这不等于内建读取器已经重算这些缓存。单点／整体常量格式保留来源，工作表标签使用当前继承格式；多个单元格的标题或系列名引用报告 `non-scalar-label`，不擅自拼接。数字标签及坐标轴的 sourceLinked 格式接入下述布局；留白／重复字段保留结构化部件，最终本地字体解析后测量 SVG 字形。留白用普通空格和显式位移保留占位字形宽度；重复按分配的标签宽度、整串成形和文字额度处理，只有最后一项重复字段生效。固有的引用标题／名称不额外分配填充宽度。

隐藏数据遵循工作表及图表的 `visibleCellsOnly`：所有系列均隐藏的位置从绘图轴移除；仅某系列隐藏的点不画，也不变成零，折线跨过该点；隐藏的分类单元格不显示其文字，但同位置的可见数值仍从当前单元格读取。源点序号保留给单点样式，隐藏整系列也不会改变其他系列的颜色序号。普通空白继续按 gap／zero／span 绘制，错误断点单独处理，依据[微软的空白、错误和隐藏数据说明](https://support.microsoft.com/en-us/excel/display-empty-cells-null-n-a-values-and-hidden-worksheet-data-in-a-chart)；这些规则不等于 Excel/WPS 实机验收，也不仿照固定 Calc 在隐藏分类列时使用过期缓存的异常。

组合层按文档／修订核对响应，并检查图表所属表、重复／缺失来源槽及常量与引用分支。常量响应必须与源快照的唯一向量一致，引用不能替换成常量缓存。默认分别限制输入及补齐后的绘图数据各 50,000 个点、128 个图表、5,000 条来源及累计 8 Mi UTF-16 输入／派生文字；短系列补齐也计入显示点预算。现有两个来源的三类图表浏览器检查已改用此组合层，并新增隐藏点序号／错误断点验证与 PPT 图表回归。15 份独立类型／隐藏数据来源仍不借用故意设置为 999 的图表缓存；完整标签布局与对象网格合成由公共预览及定向浏览器矩阵检查，见读取阶段验收。


内部图表画布已将当前数据、主标题与双轴标题组合为一次候选。富文本标题保留各自正文和 UTF-16 投影；引用标题使用当前格式化标量及原生 txPr 样式，支持字体角色、换行、字号缩放和自动适应，不赋予图表缓存或派生文字伪造的源偏移。引用标题的临时源投影没有规范 nodeId，退出时删除。源与显示文字、生成节点及测量时并存的临时节点分别受累计额度约束；缺少当前引用值、错误所有权和后续标题失败不修改源或已显示画面。

组合器为主标题、轴标题、图例与绘图区分配独立的有界矩形，轴标题按物理方向旋转；顶部／右侧轴及五种图例位置不会重叠这些矩形。字号／自动适应沿用正文测量结果；过长标题返回溢出角色，自动扩展超过预留区域或图表过小则明确失败，不静默改源。三浏览器及旧 Chromium 检查独立原生引用／富文本来源、故意不一致的缓存、空标题、混排和失败清理，并回归 PPT 图表。

图表现在按当前格式和本地字体测量轴标签与图例带，图例可按行／列换行并在带内居中；字体替代解析后，在临时封闭 ShadowRoot 中测量最终 SVG 标签，源中的全局 SVG／text 规则不进入这份视图。未指定分类标签步长时避免重叠，显式步长及左／中／右对齐继续保留。字号比例、字距／词距和多行行距参与显示及边界计算；饼图外部标签分别在两侧排布，显式引导线连接扇区／数据点和文字边界，使用保留的线条外观。节点和临时探针均计入同一候选峰值，失败后释放，不替换当前图表。

`showLabelsOverMaximum` 按 [MS-OI29500 §2.1.1548](https://officeprotocoldocs-f5hpbjgea6b8gneq.b02.azurefd.net/files/MS-OI29500/%5BMS-OI29500%5D.pdf)处理上下轴范围：缺省／false 隐藏范围外的标签，true 将带符号值显示在相应轴边界附近；源中的省略状态仍与显式值区分。用户已选择负柱 `invertIfNegative=true` 时保持负值方向，填充白色并保留原边框；单点 false 覆盖系列 true，图例保留原系列颜色，显式 noFill 继续不绘制填充。General 使用现有测宽算法选择显示形式；显式 `0` 等格式不会被换成另一格式，极长的标签保留全部文字并在画布裁切，返回 `chart-label-overflow` 及数量。轴标签带最多占画布宽度的 30%，防止文字把绘图区完全挤掉；过小画布、不可排布的图例或饼图标签仍按几何预算失败。数字格式留白／重复字段在实际 SVG 字形测量后布局；完整来源矩阵与定向固定字体检查见读取阶段验收。

[12 组独立原生来源及实际重存](evidence/xlsx-chart-labels.json)覆盖标签范围、负值反转、单点外观及饼图引导线，分别记录源／生产者重开的 PDF 文字边界、颜色和实际布尔值。源的 66 个 XML 部件通过官方 XSD／关系审计。生产者共有 69 个 XML 部件，其中 4 个图表的 numFmt 位于逐点 dLbl 之前，原始结构不通过官方 XSD；它们保持实际字节，不计为结构通过，仅在验证副本中调整该顺序以检查其余 XSD 约束及包关系。语义读取与包精确重开另外验证，不把它们混作 XSD 合规。Calc 还会忽略柱图标签范围开关、把负数绘为绝对值并重存反转开关为 false；SDK 保留实际来源，按用户选择显示原始带符号数据，不从初始源恢复生产者已改写的开关。整页像素等同与 Excel/WPS 实机行为没有据此获得验证。


图表数字标签与坐标轴现在从当前来源决定格式。`sourceLinked=false` 使用自身格式，开启关联则读取当前单元格／原生常量格式；属性省略时按[格式定义的 true 缺省值](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.charts.numberingformat?view=openxml-3.0.1)处理。每个显式单点／系列格式都写入其完整关联语义，避免省略开关时误继承父级 false。数字分类轴可单独覆盖数字格式，文本分类保留当前源文字；数值轴取系列顺序中首个实际绘制数值的当前格式，跳过被排除的隐藏点、空白和错误；没有数值时使用 General。百分比标签的关联缺省显示为百分比，显式关闭关联则尊重标签自身格式。格式化使用工作簿传入的日期／地区策略，仍不借用图表引用缓存的格式。

[20 项独立原生格式探针](evidence/xlsx-chart-format-semantics.json)及对应 XLSX 保留源码生成摘要、固定归档与字体。可运行 `python3 scripts/check-xlsx-chart-formats.py` 重核；样例已通过官方 Transitional XSD／OPC 检查、公开导入和文档包重开。覆盖两类轴、值／百分比标签的开关四种状态、首值的不同格式／空白／隐藏及单点缺省覆盖。固定 Calc 的标签与上述关联行为相符，但会把显式坐标轴格式（包括开启或省略关联开关）改成非关联格式；该差异单独记录，不覆盖源设置。SDK 按[微软的源关联说明](https://support.microsoft.com/en-us/office/change-axis-labels-in-a-chart)维持关联；Excel/WPS 仍未实机验证。三浏览器和旧 Chromium 另检查当前显示文字、原始点序号及源不变，PPT 常量图表保留自己的格式策略。字体颜色码、按宽度填充／留白、完整自动标签布局和整页视觉仍未通过读取门槛。

系列及逐点标记现在分别保留可选形状、尺寸与独立 marker 填充／轮廓，点属性省略时继承系列，明确 `none` 只隐藏该标记；源点序号在隐藏压缩后仍参与覆盖。共享绘制遵循 ECMA-376 §21.2.3.27 的九种基础形状：星号包含横竖与对角线，短横／点号为有填充的薄矩形，高为尺寸的 1/5，点号宽为 1/2；加号／叉号／星号保留方形底色及独立线色。图例也显示对应标记。全部派生节点和渐变停止点计入同一预算，失败释放探针而不改源。[16 项独立原生来源及生产者记录](evidence/xlsx-chart-markers.json)分别核对源码、公开导入、包重开和 Calc 重存属性；两种配色主题的 26 个重存图表也已单独重开。固定 Calc 把 star 改为 square、dot 改为 triangle，并物化继承值；这些来源变化单列，不修改规范源去补偿。生产者的标记画面不作为像素保真结论，规范几何在浏览器中另查。共享 PPT 读取及原生写出也保留系列／逐点 marker 外观，避免新增共享字段在已有格式输出时丢失。显式 auto 标记按下一段保留并绘制；隐式标记策略、轴交叉、标签／字体及自动布局见下述独立检查。

分类网格、分类点及刻度已共用一组轴区间：`between` 使用两端半分类留白，`midCat` 使首尾点落在刻度端点，反向分类镜像当前点。单分类仍使用一个有限区间；轴隐藏不删除原有分类／数值网格。网格类现为共享图表外观，PPT 读取、类验证和原生写出也保留它。柱、折线及其标记统一裁切在绘图区内，轴／标题／标签各自留在所属区域；折线先画连线，再画标记。[12 项独立原生来源与 PDF 向量](evidence/xlsx-chart-grid.json)记录正／反向、两种交叉、隐藏轴和单分类；固定 Calc 在柱图中忽略 `midCat` 的显示效果但保留该属性，SDK 保留规范投影，不将此差异计为数据位置保真。分类网格的归一位置另有十项原生对照，溢出柱／线的实际像素在绘图区外不得着色。轴交叉、标签边界、自动布局与极端有限数值另有独立探针及浏览器检查，汇总见读取阶段验收。

内部对象组合层已使用单元格窗口的同一行列度量和冻结窗格，把图片、形状文字及当前数据图表放到网格之上。层级遵循规范 HTML 顺序；填充位于内容下方、轮廓位于内容上方，阴影作用于完整对象。各窗格保留完整旋转／翻转中心并裁剪，拥有独立节点、SVG 标识与 UTF-16 投影；重复片段不会重复进入无障碍树。同一候选共享图片资源，全部解码成功后才交给外层发布，取消、迟到响应、超限及后续失败都释放候选且保持当前画面与源字节不变。

图片保留原生 `blipFill@rotWithShape` 的缺省／true／false。按 [DrawingML 定义](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.spreadsheet.blipfill?view=openxml-3.0.1)，false 时填充不跟随外框旋转；当前显示解释为在页坐标的旋转框包围矩形内拉伸，再在独立旋转的原框内裁剪。实现使用内部反向旋转，并按单双翻转调整方向，原始图片字节及源裁剪不变。派生填充尺寸与裁剪共用有限尺度／最大尺寸检查，额外节点计入完整候选预算。此包围矩形解释已有四象限像素检查，尚无 Excel/WPS 实机确认。[20 组独立来源及 Calc 重存观察](evidence/xlsx-picture-rotation.json)分别保留真实属性；Calc 省略该开关、改写部分旋转／几何，不从生产者反推已丢失的来源属性。

组合层默认限可见对象 128 个、生成／保留测量节点累计 50,000 个、源投影节点 50,000 个、源与派生文字累计 8 Mi UTF-16 单位，以及对象可测量尺寸 1,000,000pt。阴影按实际文本／轮廓边界计算，每对象最多 16 Mi 像素，跨窗格累计最多 64 Mi 像素；文字悬挂到锚点外仍参与可见性判断。当前图表数据过期或不可用时显示明确占位和诊断，不回退到旧缓存。独立／Calc 图表与图片形状来源的左右布局、四窗格、极端裁剪、资源隔离及屏外悬挂文字有浏览器检查；规则、经典备注框和完整来源显示检查见读取阶段验收；一万行固定资源结果另行核对。

旧 `lineChart/marker` 布尔值保存在共享 `nativeShowMarkers` 字段中，仅作为原生来源；显示始终由系列／单点的明确设置决定，不能用它打开或关闭 UI 标记。字段省略与 false、true 不混同，显式 false 不抑制已有标记，true 也不为未指定的系列补造标记。XLSX 与 PPT 使用同一严格读取，重复字段、未知属性、非法布尔值及非空元素正文明确拒绝；XML 属性空白接受标准折叠。PPT 可编辑原生输出保留该值，工作簿的读取端口在冻结快照中携带它，未改变 v1 消息形状或扩大来源授权。

显式 `auto` 已加入共享系列／逐点标记契约，原生源及 PPT 可编辑输出保持该值，预览不把选出的具体形状写回。官方 Transitional 电子 XSD 的 ST_MarkerStyle 包含 auto；[MS-OI29500 §2.1.1588](https://officeprotocoldocs-f5hpbjgea6b8gneq.b02.azurefd.net/files/MS-OI29500/%5BMS-OI29500%5D.pdf)将具体形状交给应用。预览按原始系列顺序循环 circle／square／diamond／triangle／x／plus／star／dash／dot，同系列的自动点和图例一致；隐藏系列及点不会重排其选择。点级具体形状、尺寸、无填充等仍按各自继承规则处理。节点及文字预算、派生所有权和来源不变分别验证。

[8 项独立自动标记来源](evidence/xlsx-chart-auto-markers.json)覆盖系列／点自动、具体点覆盖、部分尺寸／外观、两系列及 none。源文件公开导入和精确包重开通过，源与生产者均经未修改官方 XSD／关系审计。固定 Calc 会删显式 auto 及其尺寸／外观，重存仍写旧 lineChart/marker=true；其结构与 PDF 观察独立记录，SDK 现以 nativeShowMarkers 保留旧布尔值，按用户选择的 Office 语义忽略其显示作用，8 个生产者图表的公开导入／包重开已通过。开关缺失与显式 false／true 分开保留，CT_Boolean 缺少 val 时按 true 处理，属性值按 XML 空白规则解析；不从自动形状的应用选择推导开关语义，不改生产者字节制造通过。完整 R3-08 结果见[读取验收](workbook-reading-acceptance.md)。

图表的派生图例、坐标轴及数据标签复用既有脚本分段，分别保留 ASCII、h-ansi、东亚和复杂文字字体角色；组合符与变体选择符留在原基字片段，不把含汉字的整段文字都切成东亚字体。多行标签保持各行排布，派生片段不制造源 UTF-16 偏移。全局文字类作为私有父层提供缺省字体，局部轴／图例／标签类在子层覆盖，避免全局 CSS 记录较晚时反过来覆盖局部字号。

工作簿在所属源下集中建立临时字体探针，使用既有本地字体解析／诊断，选出的替代字体写入派生 SVG，图例／轴／标签诊断与标题一起进入图表布局报告。探针节点及文本计入同一峰值节点／累计文字额度，成功、限额或异常均释放；原始字体名和源字节不变。PPT 共享脚本分段与默认继承，通过既有已连接画布诊断路径解析字体。受控 SDK 检查覆盖四类字体、缺失字体替代、变体序列、实际物理字号、默认继承、局部覆盖、完整节点核算和失败清理；已有 PPT 图表源／生产者及截图回归保持通过。这不是新增独立原生字体几何或整图自动排版验收，字体／间距／标签另由定向实际字形和完整来源矩阵核对。

坐标轴按已选的 Microsoft Office 语义定位：`axPos` 保留为 `--office-axis-position`，不参与画面或标题带的定位；`crosses` 与 `crossesAt` 是互斥的原生来源。`crossesAt` 按被跨越的轴解释：分类轴上的属性使用垂直数值，数值轴上的属性使用从 1 开始的分类序号。原生有限数值保存为 `--office-axis-crosses-at`，PPT 原生可编辑导出继续保留；非有限值、重复／冲突属性和未消费内容明确失败。Office 的分类序号为整数，非整数来源不推断 Calc 的舍入规则，明确报告超出已选 Office 范围。依据 [MS-OI29500 §2.1.1447](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/c5ba9569-da34-4530-9743-b040df140552)。

当前绘制分类按隐藏数据过滤后的顺序计数，轴外交叉点夹到边界，反向属性同时影响交叉点和高／低端；柱的零值基线独立于轴线的交叉位置。`tickLblPos=high/low/nextTo/none` 分别定位标签，刻度的内外方向独立于标签位置。共享渲染器与工作簿组合器使用相同的交叉与样式继承规则，临时样式节点在成功或失败后释放。[44 组独立来源](evidence/xlsx-chart-axes.json)记录固定 Calc 的归一化 PDF 轴线及重存属性，源文件和生产者文件分别保存重开。Calc 在柱图中保留 `midCat` 却仍按分类间隔绘制，这个差异单列；归一化轴线不等同于完整标签／字体／自动布局或 Excel/WPS 实机验收。

数值轴不再要求两个有限端点之差也必须落在双精度范围内。宽区间的比例、主／次刻度和索引位移使用归一化计算；普通范围仍用原有的直接差值，自动刻度支持次正规数，上界超出容量时收至最大有限数。主刻度至多 1024 步、次刻度至多 4096 步；不能在源量级实际递增或超额时仍明确失败。[14 组独立 XLSX 源与固定 Calc 观察](evidence/xlsx-chart-finite-scales.json)覆盖 ±9e307、自动／显式范围、反向轴和 5e-324 数值，两份文件分别保留数值、轴与包重开。值域线段裁剪保留原斜率，超范围柱体保留可见方向／边框，标记只在完整墨迹不可见时省略；有限原值不会变成 Infinity／NaN 的 SVG。不能保持原覆盖的超大斜向渐变按资源预算失败，不擅自重铺颜色。标签／布局另有测量检查，不代表 Excel/WPS 实机验证。

内部窗口控制器现在为同一不可变来源管理滚动／切表候选。每次候选使用独立窗口读取器，共用已解析辅助内容和按表字体度量；单元格、对象及图片解码全部成功后才同时替换画面与视口尺寸。取消会及时结束等待，即使提供方忽略信号；迟到成功或失败不会改变后来已发布的窗口。组合后的最终节点另限 50,000 个，不把单元格和对象各自额度相加。失败保留前一窗口，替换或销毁释放旧图片地址及显示节点；外层仍拥有 surface、数据提供方和格式化器。三浏览器及旧 Chromium 覆盖切表／调整视口、修订冲突、资源摘要错误／缺失图表响应、两阶段迟到响应、累计额度及重复销毁。开发分支的公共预览已接入此生命周期，专用 Worker 读取通道见后述。

窗口结果已包含批注的源范围投影与待恢复记录。单格批注可沿合并单元格的完整几何跨冻结窗格显示，原始单格锚点保持不变；多区域、重叠批注及解决／关闭状态分别保留，隐藏行列和整表范围只查询边界，不展开单元格。`needs_reanchor` 不生成可靠网格几何，但保留原始内容、回复、原状态和恢复原因；缺失工作表的记录仍可供面板读取。默认累计限制 5,000 条批注、20,000 个源范围及 20,000 个窗格片段。

批注正文另有独立的富文本流投影，沿用源段落／行内样式、SpreadsheetML 下划线和删除线、中文字体角色及精确 UTF-16 正文偏移。段落之间不编造源字符，生成节点没有规范 nodeId；颜色、字体解析和选区只作用于候选。独立原生／Calc 来源覆盖经典批注、现代线程与回复，另检查多段、换行、混排、组合装饰及额度失败。固定 Calc 的现代线程重存样例本身无批注，继续按已记录的生产者差异处理，不从不存在的缓存恢复。

内部只读批注视图把上述正文与作者、时间、状态、回复及父回复关系组成一个可独立替换的候选；定位异常显示恢复原因和原状态。仅当前线程所需的正文子树进入源测量，来源索引按文档／修订绑定，图片形状组合也核对该绑定。候选默认限 1,000 条回复、50,000 个生成／临时节点、50,000 个源节点和 8 Mi UTF-16 正文源／显示与界面文字；任一正文失败不替换当前面板，销毁释放所属节点。宿主的纯文字批注不解释为 HTML。视图、字体、失败隔离及回复关系已有浏览器检查。

内部批注入口组件把同位置批注合并为一个可键盘访问的按钮，向拥有者交付完整批注 ID 列表；每条批注只在首个可见片段提供入口，原始锚点不变。选中批注高亮全部可见区域，左右方向使用相同逻辑几何；冻结／滚动裁剪边缘不画成原始范围边框。超限的高亮更新保留当前选择，关闭和销毁清除显示及事件监听。真实鼠标／键盘检查覆盖两种方向、多区域、重叠批注、失败恢复与源不变。WebKit 需要 iframe 允许父文档安装的 SDK 事件回调；沿用 Word／PPT 的 sandbox 事件机制，同时以 `script-src 'none'` 明确禁止文档脚本和事件属性，三浏览器检查其执行确实被阻止。公共预览已接通批注列表、线程和关闭；可见经典原生备注框已按选定框来源组合，完整读取结果见读取阶段验收。

### 开发分支的公共只读预览

`@doffi/render` 的 `mountDocumentPreview` 按输入的 `kind` 分流。Word／PPT 继续返回原来的分页预览；工作簿返回 `WorkbookPreview`，报告的 `kind` 为 `workbook`、`layout` 为 `grid`，不提前提供工作簿打印或 PDF 支持。

```ts
import { openDocument } from "@doffi/core";
import { mountDocumentPreview } from "@doffi/render";

const model = openDocument(files, { kind: "workbook" });
const preview = mountDocumentPreview(container, model, {
  widthPt: 600,
  heightPt: 360,
  locale: "zh-CN",
  onError(error) { console.error(error.code); },
});
await preview.ready;
await preview.setViewport({ widthPt: 600, heightPt: 360, scrollLeftPt: 0, scrollTopPt: 300 });
const report = await preview.settled();
preview.dispose();
```

视口尺寸使用 pt，默认 600 × 360；行列标题、只读地址／公式栏、工作表导航、滚轮和有界原生滚动条由同一预览管理。滚动条将完整逻辑范围映射到有界轨道，不创建百万行 DOM。隐藏工作表默认不列在导航中；显式 `showSheet(sheetId)` 仍可读取宿主已授权来源中的隐藏表，并标明“隐藏”。点击单元格提供带文档／修订绑定的 `selection()` 与 `onSelectionChange`；合并单元格保留完整范围，显示选框按冻结窗格裁剪。窗口替换会清除当前单元格选择和批注面板。

`openComment(annotationId)`、`closeComment()` 和界面批注入口共用只读面板。重叠入口显示全部线程选择；待恢复记录保留在所属工作表列表中，缺失工作表的待恢复记录也可访问。面板、选区及网格／绘图合用窗口节点额度，字体报告随当前打开的正文更新。标题最多 10,000 个可见标签，导航最多 1,000 张工作表；超限明确失败。所有控件与标题在窗口发布前准备，失败保留上一完整窗口。

直接模型路径仅读取快照、当前窗口、引用和授权资源，不调用整本 `serialize()`。`WorkbookPreviewSource` 也允许宿主提供已验证、授权且不可变的单元格／图片／图表读取接口；专用端口传输见下节。预览不释放宿主提供方或 Worker。`update(input)` 在独立表面准备新来源，`settled()` 等待最新请求；失败保留旧画面，旧请求迟到不会覆盖新来源，销毁中止等待并释放所属图片地址、表面、字体缓存及监听器。

报告区分当前可见公式的 `imported/valid/stale/unavailable` 状态；导入缓存及过期／不可用结果在界面明确提示，不宣称已经重算。图表诊断、字体替代和行高回退继续可查询。规则／智能表完整显示及固定字体定向对照已纳入[读取验收](workbook-reading-acceptance.md)；最终运行时代码的固定资源四引擎结果见[读取资源检查](workbook-reading-budget.md)。Playground 的所属 Worker 接入见下节。

### 宿主 Worker 读取通道

`serveWorkbookReads(port, model, options?)` 把一个宿主授权的专用 `MessagePort` 绑定到一份不可变工作簿；宿主在 Worker 内自行调用 `importXlsx`、`openDocument` 后建立服务。主线程用 `connectWorkbookReads(port, options?, signal?)` 获得 `WorkbookReadConnection`，可直接作为 `mountDocumentPreview` 的输入。SDK 不读取全局 Worker 消息、不创建 Worker，也不通过请求中的文档 ID 选择另一份来源；创建、授权、导入进度、候选切换和终止 Worker 仍属于宿主。

每条连接只传一次元数据快照，随后按需要查询单元格、当前图表引用和图片。工作簿 HTML 块留在 Worker 中；图片只转移 `getAsset` 产生的副本，不分离模型拥有的原始字节。接收快照递归冻结，发送前和接收后分别核算默认 500,000 个属性／数组项及 32 Mi UTF-16 单位。该预算检查不代替核心校验；端口必须来自宿主授权、已校验模型的服务。

两端默认最多 8 个待处理查询、60 秒查询期限；单批最多 25,000 个单元格、8 Mi UTF-16 单位、128 个图表、5,000 条引用和 50,000 个图表坐标，单图片最多 8 MiB。图表公式默认另限 250,000 次依赖单元格读取和 5,000,000 步，两端通过 `maxChartFormulaCells/maxChartFormulaSteps` 核对请求中的 `maxFormulaCells/maxFormulaSteps`。可通过 `WorkbookReadTransportOptions` 设置明确额度；客户端不能提高服务端额度。修订不符、重复请求、未知操作或字段、超限及不可读响应明确失败。普通提供方异常转成有界错误，不返回原始异常文字。取消会在客户端立即结束等待并通知服务端；单元格和图表适配器在已有任务让出点停止工作，迟到响应不会完成新请求。

预览销毁后连接仍可查询。连接或服务的 `dispose()` 关闭专用端口、通知对端、取消待处理查询并释放所属监听与提供方；不会终止宿主 Worker，也不会影响同模型上的另一条连接。宿主取消候选导入时应关闭候选连接并终止其 Worker，保留当前已接受的来源。工作簿编辑、事务和计算命令尚未加入这个只读通道。

[真实 XLSX Worker 记录](evidence/workbook-native-worker.json)使用公开导入、统一打开、公开读取通道和公开预览；独立扫描确认 10,000 × 20 个实际单元格、20,000 个公式、40 块，末行数值和导入缓存可查询，滚动后仍为有界窗口。三种现代浏览器及固定旧 Chromium 检查图片／图表读取、主线程导入期间持续运行、取消候选及释放；实现没有把整本来源序列化给主线程。此本机记录尚未固定 4 vCPU／8 GiB，未测浏览器加 Worker 的峰值内存，不作为完整 R3 性能门槛通过。

### Playground 只读接入

Playground 的本地打开入口接受 `.xlsx` 和工作簿 `.zip`，列表提供真实一万行样例及 Excel 分类。宿主创建一个只服务该来源的 Worker，在其中完成解包、导入、统一打开及校验；工作簿单元格块不传回主线程。普通 Word/PPT 文档包也在该入口中识别，随后交回现有会话路径；不会把工作簿送入 Word 编辑器。宿主直接依赖 render 的公开读取及预览接口，依赖方向仍为应用到 SDK。

界面显示阶段内进度，不把多个阶段累计成虚假的总百分比；取消会关闭候选读取连接并立即终止其 Worker。新工作簿首屏准备成功后才更新当前标题、列表和视图；导入、布局或取消失败保留已接受文档。已接受工作簿的 Worker 随本地列表缓存，切换回来复用同一不可变来源；预览关闭不终止提供方，整个应用卸载时释放缓存及全部所属 Worker。Word/PPT 会话的编辑历史保持隔离。只读阶段隐藏编辑／输出工具，并明确显示仍未完善的表样式、条件格式和排版范围。

浏览器检查覆盖实际一万行打开、窗口缩放、Word/PPT/Excel 切换及 Word 历史、取消实际导入后重试、工作簿包重开、损坏原生文件、注入布局失败保留画面，以及普通 Word 包回归。测试观察真实 Worker 创建和终止，使用公开预览，不替代完整读取或固定 4 vCPU／8 GiB 性能验收。

### 智能表预定义样式数据

内部样式目录已从 [ECMA-376 Part 1，第五版（2016）官方附件](https://www.ecma-international.org/wp-content/uploads/ECMA-376-1_5th_edition_december_2016.zip)的 `OfficeOpenXML-SpreadsheetMLStyles.zip/presetTableStyles.xml` 提取全部 60 种表格样式：Light 21 种、Medium 28 种、Dark 11 种；不引入透视表样式。434 份原始差异格式去重为 180 份只读定义，保留主题色、精确 tint、字体、实色填充、区域外框及内部边框。区域顺序按 §18.8.41 保存，不能用 CSS 文件顺序替代。

[生成器](../scripts/generate-workbook-table-styles.py)固定完整 XML 摘要并检查所有消费字段。该附件的局部 `dxfId` 从 1 起；Light14 的 `dxfs@count` 写成 10 而实际及引用均为 9，内嵌样式名多一个尾空格；Dark9 的内嵌名误写为 `TableStyleDark9 2`。提取仅接受这份固定附件的已记录差异，使用外层元素及 Annex G.1 的规范名称；普通 XLSX 的零起始差异索引、计数及名称校验保持严格。来源摘要、许可与差异见[记录](evidence/xlsx-builtin-table-styles.json)，[Ecma 许可](../packages/render/ECMA-376-TABLE-STYLES-LICENSE.txt)随 render 打包并保留在生成数据注释中。

可运行 `python3 scripts/generate-workbook-table-styles.py /path/to/presetTableStyles.xml --check` 重核。目录的完整名称、典型主题及边框、不可变性和附件差异有单元检查；表样式、直接格式与条件规则已接入上述窗口组合，原生与 SDK 显示结论分别记录，定义目录本身不能替代读取门槛。

内部区域投影按 §18.8.41 的顺序处理十三类区域，并按 §18.18.77 的结构图区分数据条纹、表头和汇总行。条纹大小与缺省空白带独立保留，末带裁到所属区域；列方向使用源坐标，视口和 RTL 不改变区域身份。内边框只作用于区域内部，外边框只作用于区域边界，显式 `none` 与缺省不混同。独立预期覆盖可变条纹、无表头／无汇总行、单列角落重叠及整张工作表尺寸查询；此层不展开全表单元格。隐藏／筛选后的可见行条纹、命名／直接格式和条件属性优先级已接入实际窗口 CSS 组合；原生／SDK 显示检查及差异见读取阶段验收。

[独立原生探针](../scripts/check-xlsx-table-styles.py)包含十七种表开关和直接格式组合，输入通过官方 Transitional XSD 与关系检查。固定 Calc 25.8.1.1 在这些输入上没有显示被引用的 `TableStyleMedium9`，因此不能用其输出证明表样式优先级或隐藏行条纹；[记录](evidence/xlsx-table-style-semantics.json)保留所有单元格观察及这一限制。直接粗体、字体颜色和填充的对照仍有效。可运行 `python3 scripts/check-xlsx-table-styles.py` 重核；Excel/WPS 未实机验证。

内部文字布局增加整段显示值的 `fill` 重复、水平两端／分散对齐及 `justifyLastLine`、垂直行分散。重复项保留各自原始 UTF-16 投影，数字／公式派生值仍不声明源文字偏移；部分原生富文本样式继续生效。分散行根据浏览器已成形的行框拆分私有显示节点，原始换行及跨样式片段保留，自然内容高度与铺开的显示高度分开。节点、行框探测和累计显示文字均受预算约束，失败不替换已接受视图。

[51 项独立排版来源](evidence/xlsx-cell-flow.json)覆盖重复、溢出阻挡、跨格居中、水平／垂直对齐、财务下划线及自动行高，输入通过官方 XSD／关系审计和公开导入重开。当前只有相应已实现部分完成显示检查；整套来源不能解释为完整显示门槛通过。固定 Calc 的重复文字可以比对；其垂直分散仍按普通顶部绘制，SDK 按 ECMA-376 §18.18.88 排列，差异保留。第 50 项相同输入在固定环境产生“合并为一行”或“保留三行”两种 PDF 输出，证据保存两个精确变体，普通检查只接受已记录变体且其余记录逐项相同。

财务下划线按用户选择铺满单元格，与 [Excel 帮助](https://support.microsoft.com/en-us/excel/underline-cell-contents-entire-cells-or-rows-on-a-worksheet)一致；ECMA-376 §18.18.85 与固定 Calc 的字符宽度另记为来源差异。显示层在字体解析后合并同一行的字体角色片段，以裁剪在完整单元格内的 SVG 线段绘制，不复制任何字符或制造源偏移。颜色和单双类型来自实际富文本片段；旋转、缩小、竖排、换行、填充重复和 RTL 复用同一文字几何。节点、行框探测、键文字及缓存均计入累计预算，失败候选清理临时 DOM。

原生下划线的文字盒为墨迹预留字体相关的底部带，单线为 0.125em、双线为 0.35em；该带随文字缩放／旋转并参与自动行高，不改源字号。普通单／双线继续使用浏览器文字装饰，财务线按整格绘制。默认 15pt 仍是自动行高下限，容纳得下的单线不会强制增高；合并格保持既定不驱动 AutoFit 的规则。[20 组独立源与 Calc 观察](evidence/xlsx-accounting-underlines.json)记录原始类型、行信息及固定实际字体，源／生产者分别通过导入和包精确重开。Calc 将两种财务类型分别改写成普通单／双线，SDK 按它实际保存的类型显示，不推回已丢失的财务类型；其行高也不替代 SDK 测得的墨迹空间。Excel/WPS 未实机验证，完整读取门槛继续开放。

只读窗口现在按可见行读取实际存在的分块列，发现视口外普通／旋转文字及跨格居中来源；空白、带样式的空白、公式空字符串、非空值与合并阻挡分别处理。跨格居中只扩展到同样式的显式空白格，不产生合并或改变命中地址。RTL 按实际文字方向选择物理溢出侧，冻结边界保持各自窗格裁剪。倾斜填充和共享边框还按 CSS 中全部可能角度及行／合并高度增加有界来源列，包含通过列样式着色的隐式空格，源文字投影与真实可见 gridcell 分开。

缺省每次窗口至多保留 25000 个来源坐标，行发现累计至多查询 250000 个坐标槽；每个传输批次仍受独立端点额度约束，总源文字累计检查，额外来源／私有 DOM／边框及字形几何另有预算。所有窗格与视口外来源的最终单元格片段共用可见单元格额度，不能通过省略 gridcell 角色绕过计数；报告的 `cellFragmentCount` 包含冻结重复及视口外片段。合并来源只补锚点和必要边界，不展开内部。当前行发现仅是排版来源事实，尚不能据此宣称已完成自动行高或固定资源性能验收。


## 稳定的自动行高索引

公共预览在第一次显示某张表前按声明分块建立有界的行高索引；所有实际来源列参与测量，横向滚动不改变已建立的行前缀和对象锚定。固定行高、手工默认高度和隐藏行保持源尺寸；显式自动行可覆盖手工默认。自动原生 `ht` 作为缓存不写入固定源尺寸。单元格显示与行高共用 `cell-text-layout.ts` 的普通／富文本、数字／文本格式和本地字体测量，结果向上量化到 CSS 像素并遵守 409pt 原生上限；源不改变，也不把公式导入缓存改为当前计算结果。

读取只保留一个批次：每次最多 4096 个查询槽，每 128 个实际来源建立临时样式／文字投影，完成后释放。`limits.autoRows` 缺省累计最多 250 万实际单元格、800 万查询槽和 256 Mi UTF-16 源单位，全部批次共享额度，查询前预检槽数且任务间可取消。简单类选择器的原生 CSS 可以共享已解析的相同外观；任何坐标或祖先选择器均逐格计算。未换行、未旋转的纯文本仅在计算字体、脚本顺序和字形相同的情况下共享高度；非 ASCII 保留完整片段及变体序列，富文本与几何相关布局独立测量。缓存分别限 4096／1024 项及累计键长，属于本预览的派生状态。

[21 项独立行高来源](evidence/xlsx-row-heights.json)覆盖固定默认、自动例外、手工行、隐藏行列、离屏列、跨行／列合并、旋转、竖排、大字、显式换行和第 300 行。`scripts/check-xlsx-row-heights.py` 固定 Calc 归档、字体与来源摘要，分别记录来源和重存行高；输入经官方 XSD／关系检查、公开导入和包重开。Calc 会重算手工默认下的包装行、调整横向合并和隐藏列中的内容；SDK 遵循源手工默认、合并不驱动自动适应；隐藏列现按用户选择以未隐藏宽度参与测量，不把度量或生产者来源差异称为 Excel/WPS 实机一致。

当前索引也覆盖空白单元格和使用区内仅有行／列格式的隐式来源；普通行高以源字体正常行框为下限，缩小、空串和隐藏显示不丢失该下限。字体行框探针是私有派生节点，无源偏移，分批释放；字体高度缓存另限 1024 项、单键 4096 单位。只有空列样式而没有实际使用区时不扩展百万行；只有行样式的明确区间仍参与测量。条件格式／表格单独覆盖隐式空白也已进入同一来源测量，具体有界发现和独立来源见下文；完整组合范围由[读取矩阵](workbook-reading-acceptance.md)记录；四引擎最新结果见执行交接。真实 20 万单元格在本机 Chromium 曾记录首屏约 26.9 秒、单回复最大 530434 个序列化字符；该记录尚未施加固定 4 vCPU／8 GiB，也未测交互 p95 和峰值内存，不能作为完整性能通过。


Playground 宿主目前保留最近一份已接受的工作簿预览，连同宿主已缓存的 Worker 使用；切换 Word/PPT 时隐藏并跳过零宽尺寸更新，返回同一来源不重建行高索引。新工作簿仍经独立候选验证后替换，页面卸载释放预览和 Worker。这里只保留一个视图，不引入跨会话的共享行高缓存；SDK `dispose()` 仍释放其全部视图状态。


自动行高的列度量与显示几何共享同一列宽换算：只在测量时恢复隐藏列的已存宽度，缺省宽度仍取源工作表默认；不修改 hidden、源宽或当前视图列位置。实际格、空白字体及隐式列／条件／表格来源均纳入同一批次，仍受 250 万实际格、800 万发现槽、256 Mi 文字和可取消分批额度约束，手工行、隐藏行和合并排除规则继续生效。最终行前缀与原来的零宽隐藏列共同供网格和绘图锚点使用。

[12 项独立隐藏列来源](evidence/xlsx-hidden-row-heights.json)覆盖三种明确列宽的隐藏／可见配对、未声明宽度、隐式列大字、空白、手工／隐藏行及合并。固定 Calc 对明确同宽的配对给出相同高度；省略宽度时另写入 10.49，和 SDK 的源默认 8.43 不同；合并格的 Calc 调整也单列保留。源与重存文件均通过官方 XSD／关系检查和包重开，不声称浏览器与 Calc 行度量逐像素相同或 Excel/WPS 实机一致。

21 项行高来源新增空白大字、只有行／列格式、空串、缩小字体、短旋转字和隐藏显示；固定 Calc 的相应行高保留在证据中。原生类选择器且没有位置相关的表／条件覆盖时，隐式来源按格式区间取代表空格；复杂选择器逐坐标检查，实际格与合并区域不会伪装成隐式来源。所有查询槽、隐式坐标与逐行区域检查共用自动行高累计额度，未使用工作表区不会因一个列格式产生百万个节点。隐藏列计入行高已由用户选择；源中的列宽仍保留，测量使用未隐藏宽度，显示列继续为零宽。


### 行高测量中的普通来源复用

只含普通类选择器、且没有表／条件格式覆盖的工作表，可在一次行高扫描内复用同一外观／数字格式的私有字体来源原型。每个单元格先独立校验完整属性白名单、字段唯一性、列／类型／值／公式／缓存／格式和文字身份，再查询原型；已有原型不能绕过检查。快速分支只接受双引号、无实体的规范纯文本 TD，实体、CR／NUL、其他合法语法及富文本继续走原有惰性 DOM 解析。两条路径共享同一个来源一致性校验，文字值也必须与文字来源一致。

原型最多 128 个、累计 8 Mi UTF-16 源单位，节点从同批来源额度中扣除，正常完成、失败或取消都释放；不跨预览或模型共享。行高使用当前待测标量构建派生文字，不把原型中首个单元格的地址或 UTF-16 偏移当作其他单元格的源身份。正式窗口的普通／富文本仍从实际单元格投影字符偏移；数字等派生显示按单元格锚定。重复解析减少后，本机主基准首屏记录约 Chromium 21.5 秒、Firefox 28.5 秒、WebKit 20.4 秒；这些仍非固定 4 vCPU／8 GiB、交互与内存验收。

### 有界读取与窗口文字缓存

每个不可变工作簿模型另有最多四块、累计 8 Mi UTF-16 单位的已解码源字符串，以及最多 `min(5000, maxWindowCells)` 格、累计 8 Mi 单位的已解码单元格 LRU。它们都从已验证、私有的原始字节派生，不保留解析树，不展开隐式空白，不影响保存字节。查询范围／数量仍在命中缓存前完整检查。单元格缓存先复制其 HTML 子串，避免某些引擎用一个小切片保留整块父字符串；正文、值、公式和格式字段保持冻结。模型之间不共享缓存，未来新修订必须拥有自己的派生状态。

窗口控制器还有独立的脱离 DOM 的文字 LRU，受条目、节点和源单位三重预算限制。只在普通类选择器、无表／条件覆盖时复用；命中前仍校验每次来源。缓存键包含实际源 HTML、当前值、格式、坐标、几何和视口尺寸，存入／取出均复制显示节点，并重新建立精确 UTF-16 投影和视图所有权。字体集合变化或控制器销毁会清空缓存；缓存与新文字共同接受完整窗口预算。普通文字／富文本使用实际 TD；无规则覆盖的纯数字等派生显示可共享本窗口的纯样式 TD，属性白名单和全部数据身份检查必须先于原型查询。来源树只读样式、不做表格布局，实际显示仍测量真实字体。

来源可用 `allowReadAhead` 授权有界预读；直接模型和专用读取端口启用，普通自定义来源默认不参与。自动行高最多同时持有当前和下一批各不超过 4096 个查询槽的回复，以便 Worker 读取与主线程测量重叠；失败或取消中止所属下一批，迟到失败不会成为未处理拒绝。全部查询仍受累计来源、字符及发现额度约束。

空闲预读仅用于普通类样式、无表／条件格式／数据验证的工作表。`readAheadRows` 默认 8、允许 0～64，0 关闭。首次窗口按相邻行小批预热文字；同表、同尺寸的后续垂直滚动则可按上一次实际位移准备一个下一视口，位移不得超过配置行数。没有对象／批注时，该候选以隐藏且 `aria-hidden` 的网格完成浏览器布局；其片段和节点与当前窗口共用总额度，额度不足就放弃可选准备。候选不会替换画面或拥有源权威，正式请求仍重新读取全部来源，逐项核对身份、修订、视口、坐标、HTML、类型值、格式、公式及字体版本，完全一致才能复用。取消中止未完成的后台工作；失配、释放和指针／键盘交互清理已准备候选，滚轮允许先核对候选。每个控制器最多一个后台任务，忽略取消的来源未结束前不叠加后台读取。原有相邻行预热另限每片段 128 个显示片段、256 个保留源格、4096 个查询槽、2048 个源／派生节点和 32768 个文字单位，超限即停止，不降低前台额度。

纯色透明填充不创建无绘制效果的节点，文字的完整单元格盒子直接位于窗格裁剪中；实际单元格、冻结重复、视口外片段仍全部计数。固定资源方法、原始测量限制及当前未关闭的门槛见[读取资源检查](workbook-reading-budget.md)和[执行交接](workbook-reading-handoff.md)。缓存不是完整读取或性能验收通过的替代。


## 自动行高的隐式外观组合

[12 项 Normal／小字来源](evidence/xlsx-row-font-floor.json)以 20pt Normal、5／8／11pt 显式字和两种默认高度交叉，固定 Calc 保留工作表声明的默认高度下限。SDK 只在实际来源需要更高行框时抬高对应行，不把 Normal 字体高度当作所有行的新默认值。

[9 项条件／表格来源](evidence/xlsx-row-appearance.json)覆盖只有隐式空白的匹配／未匹配／相对条件字体、表自定义字体、数据差异格式、命名单元格样式及优先覆盖。区域发现保留这些完整声明范围，合并重叠列区间，跳过已存储坐标、零宽列和合并，隐藏列以未隐藏宽度参与；所有逐行区域检查计入现有 800 万发现额度；不能靠多条稀疏范围制造未计量循环。只含原生简单类的纯颜色规则不可能改变空白行框，直接排除其全列区域；一般选择器仍逐坐标解析。列字体仍限于实际使用行，新增区域不会把其他列格式扩张到使用区外。

仅有数据验证时，行高也向同一修订的规则提供者请求当前公式值，与实际窗口显示一致；原来的错误缓存保持原字节。来源空白没有伪造值或源偏移。固定 Calc 在其中若干条件字体和表样式探针未产生对应的行高变化；证据明确保留生产者观测，不能据此宣称其视觉无差异或 Excel/WPS 已验证。三引擎与旧 Chromium 的定向行高／当前值检查通过，完整读取及固定资源结果见对应验收记录。


## 现代提及和原始定位恢复

`WorkbookCommentMention` 保留原生 mentionId、已绑定的本包现代 authorId、startIndex 和 length；主评论和回复分别保存自己的提及。依据 [MS-XLSX CT_Mention](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/b03ed619-e307-4d3e-9c67-b68612274128)逐项校验 GUID、人员引用及 unsignedInt，提及计入原生条目、元数据字节和核心对象总额。人员/provider/user 身份均为不透明来源，不授予访问权限；只读面板展示提及的姓名，没有通知或云任务入口。

`resolveWorkbookCommentMention(text, mention)` 将本项目 UTF-16 偏移解析为 from/to，越界、切开代理对或分离变体序列时返回 undefined。导入保留这类原始数字并以 invalid-mention-range 将整个线程标记 needs_reanchor；不会截断、猜测目标或改写正文。核心仅允许有明确恢复原因和原状态的线程保留无效范围。原始公式／单元格区域锚点仍独立于提及的正文范围。

回复记录与主评论落在不同单元格、多个主线程记录同一单元格时，保留所有正文、回复、父关系与 nativeCell 来源，分别记录 mismatched-reply-location／duplicate-thread-location；主评论原锚点不猜测移动，原来的 open/resolved 状态保留为 recoveryStatus。缺失主位置仍使用空范围，不从回复或兼容记录补造位置。无父节点、循环、未知人员和重复身份继续是结构错误，不能用恢复状态掩盖。

[独立提及来源](../packages/office/tests/fixtures/xlsx-comment-mentions.xlsx)和[固定 Calc 记录](evidence/xlsx-comment-mentions.json)包含 6 条现代消息、4 个提及及上述定位问题；官方 Transitional 部件和按 MS-XLSX 扩展的独立检查分别验证。固定 Calc 删除现代线程／人员部件，这项损失不能计为保真通过。SDK 公开导入与包重开保留完整来源，三引擎及旧版面板的精确投影、恢复提示、预算及失败保留另有定向检查。xr:uid 来源保留见下节；可见经典备注框与完整绘图检查见读取阶段验收。


## 备注修订 UID 与兼容身份

经典备注的 `xr:uid` 来自 `http://schemas.microsoft.com/office/spreadsheetml/2014/revision`，按 [MS-XLSX v20261016 §2.5.10、§2.3.7.3](https://officeprotocoldocs-f5hpbjgea6b8gneq.b02.azurefd.net/files/MS-XLSX/%5BMS-XLSX%5D.pdf)保留为 `nativeUid`，原有 guid／annotation.sourceId 独立保留。识别现代兼容占位后，UID 转入该线程的 compatibilityUids；身份与主线程不符时保留实际 UID，并记录 mismatched-placeholder-identity 恢复原因，现代正文和单元格不移动。UID 是不透明来源，不能代替 SDK 节点身份或任何授权。

实现按命名空间 URI 识别属性，兼容不同前缀；仅接受已选修订命名空间的 Ignorable 声明。未知属性或其他兼容命名空间仍明确失败。核心再次校验 GUID，UID 数组逐项计入既有元数据条目和字节预算。独立规范检查先验证这些 UID／声明，再仅从验证投影移除已识别属性，官方 Transitional XSD 保持原样，原始 XML 根不修改。

[独立 UID 来源](../packages/office/tests/fixtures/xlsx-comment-uids.xlsx)同时包含独立经典 guid／uid、匹配的现代兼容 UID 和冲突身份；[固定 Calc 记录](evidence/xlsx-comment-uids.json)显示生产者省略这些 guid／uid，不计为保真通过。SDK 公开导入和包重开完整保留，面板显示明确恢复原因；没有把这些来源身份显示成产品动作或通知入口。


统一预览网格在无合并格时由完整行／列边界直接组成，减少逐格重复短线和接头处的抗锯齿缝。它沿用同一窗格点坐标、裁切、颜色与线宽；隐藏轴不产生空格，冻结与 RTL 的边界相同。有合并格时继续走共享边界解析。该派生优化不改变单元格／文字／对象的源身份和读取结果。
