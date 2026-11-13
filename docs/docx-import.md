# DOCX 读取与只读预览

R1-10 将授权 DOCX 字节转换为同一 HTML 文档包，不创建编辑会话。数据结构见[三格式模型](data-model.md)，任务与范围见[第一轮细则](iteration-1.md)。富文本输入、自动分页和 PDF 渲染已在 R1-11／R1-12 完成；原生导出见 [DOCX 导出](docx-export.md)。

## 公开入口

```ts
import { openDocument } from "@doffi/core";
import { importDocx } from "@doffi/office";
import { mountDocumentPreview } from "@doffi/editor";

const files = importDocx(authorizedBytes, { documentId: "host-supplied-id" });
const document = openDocument(files);
const preview = mountDocumentPreview(container, document);
await preview.ready;
const fonts = preview.fonts();
const savedFiles = document.serialize();
preview.dispose();
```

`DocxImportOptions` 要求宿主提供 `documentId`，继承文档包 ZIP 的 `limits` 和 `maxZipBytes`。默认预算沿用[资源与 ZIP](resources-and-zip.md)：压缩输入 64 MiB，展开包 32 MiB，单文件 8 MiB，最多 1024 项；XML 各部件共享节点预算和深度限制，字段控制符嵌套另按同一深度预算检查。评论展开的范围对象另累计检查 `maxNodes`，在创建范围或扩展摘录前，以 JSON 转义的保守字节上界检查 `maxFileBytes` 与 `maxPackageBytes`，避免重叠评论放大后耗尽内存。

文字体系分段逐段生成，不先收集完整分段数组；在创建文字子串和 HTML 包装节点前，累计检查生成节点数及转义后 HTML 字节预算。正文、页眉、页脚和批注正文分别维护字段、批注范围与书签的解析状态，正文范围可以跨节延续，不能跨内容域闭合。

原生样式通过 CSS 类和继承表达，直接格式只写入对应节点；段落样式中的字符底纹/装饰使用独立的 run 类，避免把文字颜色固化为每段文本的行内覆盖。样式元数据记录原生继承及原始属性归属，便于后续导出。

导入同步执行，不访问文件系统、网络、凭据或模型。先按 OPC 内容类型和关系定位正文，再读取关联部件，生成初始修订 0，最后通过 `openDocument` 的结构、安全和引用校验。失败抛出有界 `DocumentError`，不返回损坏文档、不改原字节。当前 ZIP 适配器的 ZIP64、加密等限制继续适用。

## 兼容导入与原生保留

`importDocx` 默认使用 `unsupportedContent: "preserve"`。尚未映射的普通属性和对象保存为类型化原生残余，相关部件、关系和资源随 HTML 文档包保存；对象在画布上明确占位，周围的已支持内容仍可编辑。当前源导出 DOCX 时恢复这些残余，不依赖第二份原始 DOCX。`importDocxWithReport` 返回同一导入结果及保留项目报告。

需要历史严格行为的宿主可以设置 `unsupportedContent: "reject"`。下表中的普通功能限制在严格模式下拒绝，在兼容模式中尽可能隔离并保留；保留不代表已经支持其内部编辑或完整显示。损坏 XML、非法引用、安全禁止和资源预算仍会拒绝；无法可靠隔离的结构明确报告原因。

兼容片段合并存入一个有界资源，避免每个未支持项目占用一个 ZIP 文件槽。恢复过程检查是否取得实际进展，不使用固定的 1024 次尝试门槛。

## 当前范围

| 类别 | 已实现 | 限制 |
| --- | --- | --- |
| OPC/XML | 内容类型、关系、相对目标、Strict/Transitional 命名空间别名、受支持的兼容分支 | 不执行 DTD、实体或处理指令；不提供全 OOXML XSD 合规性认证 |
| 正文 | 段落、标题、文字片段、空白、换行、明确分页、Unicode | 不支持脚注尾注、文本框、公式、图表及一般内容控件 |
| 样式 | 文档默认、段落/字符样式继承、直接格式、主题颜色及字体、HSL 明暗变换、常用文字和段落属性、自动断字开关及段落禁用标志 | 自动断字取决于浏览器语言词典；标点悬挂等未映射的特殊排版明确失败；主题解析为文档外观，不提供主题编辑命令 |
| 编号 | 抽象定义、实例、多层级、常用数字/字母/罗马数字/项目符号、起始覆盖及标记派生 | 图片项目符号、编号样式链接及完整层级覆盖暂不支持 |
| 表格 | 网格、横纵合并、尺寸、边框、底纹、样式及直接设置的缩进和单元格边距、基本条件样式、表格外观位及条带参数、嵌套内容 | 不支持非零单元格间距、浮动表格、跳过网格位置、行级表属性例外及条件行属性；纵向合并续接单元格只允许经校验的空段落/片段；复杂边框冲突尚无完整兼容性结论 |
| 图片 | 内嵌 PNG/JPEG/WebP、尺寸、替代文字、摘要及不可变字节 | 裁剪、旋转、非矩形图片、浮动环绕、可见轮廓和图片效果明确失败 |
| 节与辅助内容 | 页面大小/边距/装订线、文档网格及 snapToGrid 开关、等宽分栏、分节类型、页码设置、首/奇偶/默认页眉页脚及继承 | 分页预览支持等宽跨页分栏与逐栏表头；不等宽栏及复杂页边框继续原生保留，末页栏平衡和完整字符网格节距仍待完善 |
| 链接与制表位 | 书签、HTTP/HTTPS/mailto 及书签链接、基本左右/居中/小数对齐和引导符；制表位按位置逐层合并，`clear` 删除匹配的继承制表位 | 预览不导航；不自动访问链接或外部图片 |
| 字段 | 简单、复杂、嵌套及跨段落字段、原指令、锁定/脏标记、结果节点与原缓存 | 不执行字段指令；页码/总页数/目录更新在 R1-12；一般表单字段不支持 |
| 批注 | 经典批注、富文本正文、作者/时间、跨节点 UTF-16 范围 | 支持经典批注的原生已解决状态扩展；不支持线程与回复；R1-07 已实现变更映射，R1-08 提供状态与重定位界面；原生富文本正文直接编辑在 R1-11 |
| 修订 | 文字插入/删除、作者/时间、删除内容及标记显示 | 格式、移动和表格结构修订明确失败；不接受/拒绝修订 |

纯校对标记、过期的 `lastRenderedPageBreak` 和样式库 UI 信息不构成文档内容。对于未定义的字符样式引用，按 OOXML 规定不应用该样式，并在节点属性保留原引用，不能将它误判为正文丢失。[Microsoft RunStyle 规则](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.runstyle?view=openxml-3.0.1)

## 中文字体

ASCII、高位西文、东亚及复杂文字字体分别保留在 CSS 自定义属性中。东亚文字、全角标点和扩展区汉字使用对应字体；组合标记不拆散到另一文字体系。中文主题字体按来源语言选择 Hans/Hant，不因主题中日文字体排在前面而误选。

浏览器不自动下载或读取系统字体文件。预览使用本地字宽差异检测字体可用性，缺失时选择可用的 Noto 中文字体并提示替代；源字体名称和原文字不改变。对东亚文字及编号标记，用本地栅格比较检测通用缺字方框，报告缺失的码点。`preview.fonts()` 返回源字体、选择的预览字体、替代标志、发现的缺字以及字形检查是否完整。

字形检查是有界的预览诊断，不是对字体覆盖或所有平台的认证；每次检查至多 4096 个不同的字形/字体组合，超出时明确标为未完整核验。不能把回退字体下的断行当作原字体保真，也不能把罕见字显示成方框当作成功。当前专用字体样例包含 `𠀀`，参考环境确实缺少其字形，验证的是报告缺字并完整保留源字符。

常用宋体、仿宋及楷体名称变体的编辑保留、补充开源字体环境及替代规则见[中文字体](fonts.md)。固定基线使用 Noto Sans CJK SC、Noto Serif CJK SC 和 Liberation Serif。字体文件未随 SDK 分发；参考记录保存实际文件摘要。Noto 字体使用 SIL OFL，Liberation 使用其 OFL 许可；实际使用或分发其他字体由宿主确保授权。

## 预览与生命周期

`mountDocumentPreview` 重新校验传入模型的序列化源，在无脚本的 sandbox iframe 和限制性 CSP 中加载文档样式。包内图片转为当前组件拥有的 Blob URL；辅助标签、编号标记、表格列、字体替代及测量都是派生 DOM，不回写源。

`ready` 在字体准备、派生装饰和初始测量完成后返回；图片加载继续更新容器高度。`dispose()` 幂等，释放观察器、iframe 和全部资源 URL。Playground 切换文件时销毁旧预览，使用加载代次拒绝迟到结果。

R1-08 的 `mountDocumentInteraction` 在同一渲染器上提供文字／对象／区域选择和批注管理，接受宿主拥有的编辑会话，见[预览与交互](preview-interaction.md)。Playground 已接入该入口，下载当前会话而非导入时的初始快照。

R1-10 的读取基线为连续视图，各节分别应用有效纸张宽度和四边页边距；R1-12 已增加跨浏览器的默认分页视图，见[分页与渲染](pagination-rendering.md)。参考渲染用于核对内容、外观和已知差异；浏览器与 LibreOffice 的字形基线、自动行距及表格边界仍可能有小幅差异。本阶段不宣称逐页或逐像素保真，也不以截图回归代替原文件对照。

## 独立样例与验收

[DOCX 样例目录](../examples/docx/) 包含八类功能样例、一个中文字体诊断样例，以及 LibreOffice 实际重存的对应文件。原始样例由独立 Python ZIP/XML 生成器依据公开结构产生，不调用本项目导出器。重存文件验证不同生产者产生的实际 OOXML，文本预期从源文件独立提取。

```bash
pnpm fixtures:docx
pnpm check
pnpm browser:install
pnpm test:browser
pnpm check:docx-reference
```

首次或有意更新参考环境时使用 `pnpm record:docx-reference`，它以 LibreOffice 生成参考页面及重存 DOCX，并记录渲染器、中文字体摘要、输入摘要和参考图像摘要。正常参考检查只更新 `.cache` 并核对记录；环境或输出变化不会静默刷新基线。需要系统的 LibreOffice、Fontconfig 与 Poppler；这些是开发验收工具，不进入 SDK 运行依赖。

自动化覆盖导入内容、样式覆盖、合并、图片字节、字段和嵌套引用、批注范围、修订、危险输入及资源预算；同时检查保存/ZIP/重开、打包消费和固定视口截图。截图是检查过的连续视图回归，不是 DOCX 导出结果。字段缓存按源 XML 验收，参考应用自动更新后的值不能替代原缓存。原生 DOCX 可编辑导出和 LibreOffice 实际重开在 R1-13 验证，Word／WPS 尚未实机验证；完整 Playground 闭环仍在 R1-14。

实现依据：[ECMA-376](https://ecma-international.org/publications-and-standards/standards/ecma-376/)、[Word 条件表格格式优先级](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/f7f85a73-5f60-4f7e-a725-4f1f06198388)、[Word 主题颜色变换](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/8229a077-7fc8-4fba-96cc-c77b6a4fc768)。
