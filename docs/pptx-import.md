# PPTX 读取

R2-04 提供 `importPptx(bytes, { documentId, limits?, maxZipBytes?, mode? })`，将宿主授权的 PPTX 字节映射到同一 HTML/CSS/元数据／资源文档包。它不调用 Office 进程、网络或第三方演示文稿库。R2-05 增加三类图表、备注和两类批注；[R2-06 只读画布与读取验收](presentation-preview.md)已覆盖独立来源与生产者样例，[R2-07 核心事务](presentation-transactions.md)已开放，编辑界面继续按后续阶段推进。

```ts
import { openDocument } from "@doffi/core";
import { importPptx, writeDocumentZip, readDocumentZip } from "@doffi/office";

const files = importPptx(authorizedPptxBytes, { documentId: "host-id" });
const model = openDocument(files);
const snapshot = model.snapshot();
const packageBytes = writeDocumentZip(model.serialize());
const reopened = openDocument(readDocumentZip(packageBytes));
```

输入字节保持不变。文档包不包含原 PPTX 或隐藏的转换回退副本；输出可直接打开、查询、保存重开。ZIP、XML、生成节点与字节、资源和最终源校验均有预算。DTD、实体、危险 XML、外部图片、未知可见对象和失效关系明确失败，不提供默认部分导入。

SDK 默认 `strict`，Playground 打开 PPTX 时显式使用 `compatible`。兼容模式在 `presentation.nativePreview` 中保留主题对象默认定义、未被引用的高级主题样式、对象／页面标签、支持的动画和擦除过渡，以及仅依赖内部主题的讲义母版。动画和过渡以静态方式预览，主题默认定义仅作为原生数据保留。预览报告和 Playground 提示这些限制。含保留定义的文档包可保存重开，原生 PPTX 导出会明确报告限制。引用的高级主题样式、媒体定时、带外部关系的动画和未知可见对象仍会拒绝导入。[讲义母版](https://learn.microsoft.com/en-us/office/open-xml/presentation/working-with-handout-master-slides)

兼容模式为包含完整画幅位图的 EMF 提取 PNG 预览，并把原始 EMF 保存为文档包资源。内置强调色表格样式在单元格已明确提供填充时采用兼容的边框和文字样式，保留原样式标识。[EMF 位图记录](https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-emf/89c0d808-0dea-413f-be40-2e9e51fa36ac)

| 内容 | 当前映射 |
| --- | --- |
| OPC | 按关系和内容类型定位，支持改名／移动部件、根绝对关系、Strict／Transitional 和任意 XML 前缀；不按 `ppt/slides/slide1.xml` 猜路径 |
| 文稿 | 宿主身份、标题、语言、原始 pt 尺寸、幻灯片顺序、隐藏及母版显示标志 |
| 主题和继承 | 颜色、角色字体及补充脚本字体、格式样式列表、颜色映射、母版文字样式、布局与占位符身份；缺失几何继续通过源引用继承 |
| 文字 | UTF-16、emoji／变体序列、符号字体角色、纯色高亮、字符／段落格式、列表、换行、横排和东亚竖排对象、内边距、垂直对齐、三种自动适配、默认制表间隔及文字／对象超链接 |
| 形状 | 声明的 16 种预设、常量调整值、有界直线路径和 `val`／`*/`／`+-` 导引公式、实色／线性渐变填充及轮廓、透明度、虚线和基础外阴影 |
| 组与连接线 | 局部坐标、负坐标、嵌套、旋转、翻转；`line`／`straightConnector1` 及 `bentConnector2` 至 `5` 的原始预设、调整值和对象引用 |
| 图片 | PNG／JPEG／WebP 原始字节、摘要、矩形裁剪、单闭合直线路径蒙版、透明度调节、变换、替代文字和可继承的拉伸图片背景 |
| 表格 | 原生 DrawingML 网格、行高、横纵合并、显式填充／边框／对齐和富文本；合并起点在 HTML 中只有一份内容 |

未支持的表格样式 ID、显式制表位、其他竖排／多栏／变形文字、文字描边／复杂效果、自由曲线及其他自定义路径指令、图案／纹理填充、音视频、嵌入控件、宏及高级演示特性均明确失败。空过渡容器和没有播放内容的时间根节点可读取。其他过渡效果和自动翻页时序仍受导入范围约束。

外观唯一保存在 CSS：主题原色为 `--office-palette-*`，有效颜色角色为 `--office-theme-*`；颜色映射引用原色，避免交换角色时产生变量循环。主要／次要字体保存在 `--office-major-*`／`--office-minor-*`，字符字体分 ascii／h-ansi／east-asia／complex-script。格式样式列表映射为主题 CSS 类；对象引用相应类并保留局部覆盖。颜色变换、渐变停止点、线条和阴影使用 `--office-*` 自定义属性，不保存原生 XML 作为第二份编辑来源。文字样式、列表和字段缓存保持源语义，画面适配在 R2-06 读取这些属性。

位置和组子坐标仍由 `data-office-x/y/width/height` 等源属性保存。原生预设常量为 `data-office-adjust-*`，连接线原始类型为 `data-office-connector-preset`；核心校验检查其类型、数值和 routing 一致性。幻灯片内原生 ID 保存在类型化对象元数据，用于连接引用和后续批注映射。它们只表示身份，不授予访问权限。

`firstSlideNum` 保存为 `presentation.firstSlideNumber`，缺省从 1 开始，可保存 Office 范围内的 0～9999；文档包和原生 PPTX 重开保留。动态页码采用起始值加页面顺序，源字段缓存保持原样。[PowerPoint 页码语义](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.slide.slidenumber)、[Office 属性范围](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/76016a6d-35f9-47cc-8fc3-8d47ba713bbb)。

2026-11-25 的 PowerPoint 16.114.1 实机对照发现：字符属性中位于 `a:cs` 之后的 `a:ea` 不参与东亚字体覆盖，原生重存会移除它。导入沿用该有效继承行为，同时把未生效的声明保存在 `--office-ignored-font-east-asia` 中；文档包重开仍保留源字体名。规范的 `latin`、`ea`、`cs` 顺序继续分别映射字体角色。原生导出写有效字体继承和规范顺序，不把未生效声明重新变成有效覆盖。

验收使用[独立输入与 LibreOffice 重存文件](../examples/pptx/README.md)。基础读取用例对 11 份源及对应 11 份外部生产者文件分别验证文字、对象和引用、输入不变、确定性以及文档包 ZIP 保存重开；另覆盖图片原始字节、主题与占位符、嵌套坐标、连接引用、合并、异常输入及预算。LibreOffice 重存可能改变连接线预设和小数网格尺寸，读取保留重存文件的实际值，不改回原合成值。现代批注在生产者重存中丢失，不使用无批注重存文件代替现代读取验收。

```bash
pnpm build:packages
pnpm exec vitest run packages/office/tests/pptx-import.test.ts
```

此阶段的重开是 HTML 文档包重开，尚未提供 PPTX 导出；原生可编辑导出和外部编辑后的重新导入在 R2-14 至 R2-16 验收。


## R2-05：图表、备注与批注

16 份原始 PPTX 和 16 份 LibreOffice 重存文件分别完成 SDK 导入、输入不变、确定性、HTML 包 ZIP 保存重开验证。原始现代线程保留正文、回复、作者显示信息、原生 GUID、日期、open/resolved/closed 状态及幻灯片／对象／文字锚点；LibreOffice 重存丢失的线程按参考记录计为零，不代替现代线程验收。PowerPoint／WPS 实际应用仍未验证。

图表支持二维簇状柱、简单折线和单系列非负饼图。分类、系列名和数值只进入 `dataSets`；缓存和内嵌 XLSX 的普通行／列范围逐项一致才接受，内嵌表中的单元格公式、外部依赖和不一致缓存明确失败。XLSX 仅作为图表常量表读取，未引用的数据和工作簿排版不进入 SDK，也不是通用 Excel 转换器。共享内嵌表／范围复用一份源数据；独立内嵌表保持独立。无内嵌表的 LibreOffice 缓存常量也可读取，支持其本地数据标记，拒绝可计算表达式及外部范围。PPTX 与内嵌 XLSX 共享 XML 节点预算，并合计展开字节和文件数。

图表标题保存在 HTML，隐藏标题保留内容及 titleVisible 状态；系列、数据点、图例、标签、绘图区、轴及文字样式引用 CSS 类。定义保留图例位置、轴范围与步长、方向、数字格式、空值行为、系列／逐点线条标记、独立标记填充与轮廓、柱间距、重叠及饼图起始角等基础选项。组合图、堆积／平滑／爆炸图、手工绘图区、预设样式 ID、轴标题／对数轴、自定义交叉点和单点标签布局等高级配置明确失败；这些限制不能算作已完成的视觉映射。

备注保存原始 notesSize、备注母版、辅助对象、局部几何和独立正文对象引用 `bodyNodeId`。不会把辅助缩略图、日期或页码混入可编辑备注正文；原有纯富文本备注包仍可读。备注和批注内容域不属于普通幻灯片页面，后续 PDF/PNG 不输出它们的 UI。

经典批注原生位置以 1/8 pt 换成 1 pt 的区域标记，保留实际 x/y；负位置当前不能由公共非负区域表示，明确失败。现代批注保留独立的 callout 位置，锚点通过原生 moniker 定位，不能用位置猜对象。文字范围按 UTF-16 映射，软换行为一个位置，跨段落按原生段落标记转换为源文字范围集合。可选上下文按 [MS-ODRAWXML 1.3.9.1](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-odrawxml/9846e83e-6b3c-4dd3-82fe-244bf1e24e1f)逐代码单元计算 32 位哈希；哈希、长度或目标不一致就失败，不搜索另一个相似范围。最终隐式段落标记之后的范围、混合文字／对象及未实现的 moniker 目前明确失败。

现代原生 `unknownAnchor` 保留为 `needs_reanchor`，原工作流状态单独存于 `recoveryStatus`，避免把“需重定位”和“已解决／关闭”混为一项。`commentAuthors` 只保存格式、原生身份及不授予访问权限的 userId/providerId；显示名字来自批注／回复，不维护第二份名字。引用的作者、回复及显示信息必须一致。Word 仍拒绝回复和 closed 状态，不宣称开放 Word 现代线程。

不支持现代任务、分配和其他未声明的扩展；经典扩展线程未声明读取。现代批注的格式依据 [MS-PPTX](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-pptx/)，锚点依据 [MS-ODRAWXML](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-odrawxml/)。代码和独立样例均自行编写，未复制或翻译其他 Office 编辑器实现。

```bash
pnpm exec vitest run packages/office/tests/pptx-import.test.ts packages/office/tests/pptx-stories-charts.test.ts
```

独立审查补齐标签数字格式与轴选项、隐藏标题、未知批注扩展及省略占位符类型的原生索引解析。显式回车以字符引用在公共 HTML 适配器中稳定保存；普通输入事务仍使用 LF。Word 用例覆盖保留既有回车后的编辑、冲突、撤销重做及重开，其他非法控制字符引用继续拒绝。
