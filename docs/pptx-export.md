# 原生 PPTX 导出与独立闭环

R2-14～R2-16 使用第一方 PresentationML、DrawingML、图表和常量工作簿适配器，从当前 HTML/CSS、元数据及授权资源生成可编辑 PPTX。没有来源 PPTX 回写缓存，不使用富文本、演示、图表编辑器或 Office 转换库。通用 ZIP/XML/CSS 工具沿用既有适配边界；LibreOffice、XSD 和浏览器仅用于开发验收。

```ts
import { openDocument } from "@doffi/core";
import { exportPptx } from "@doffi/office";

const model = openDocument(authorizedPackage);
const { bytes, report } = exportPptx(model, {
  commentFormat: "preserve", // preserve | modern | classic
});
// report.documentId / revision 精确绑定本次源快照。
// bytes 的下载、存储和权限由宿主负责。
```

## 当前源映射

- 幻灯片与对象顺序取自 HTML；元数据数组登记身份和引用。导出生成部件路径、关系及原生数值 ID，内部节点 ID 不成为路径或访问权限。
- 母版、布局、主题、备注母版、占位符及文本默认样式保留为原生定义和关系。未覆盖的占位符几何与文字样式继续继承，局部编辑写为局部覆盖。
- 对象坐标、尺寸、旋转、翻转、组子坐标、预设及常量调整值、图片裁剪、连接对象与连接位点写入原生 DrawingML。图片只读取当前清单中已有的 PNG、JPEG、WebP 字节。
- 主题保留调色板、颜色角色映射、主要／次要字体和补充脚本字体。主题样式的占位颜色与对象引用颜色分别重建，避免把两层颜色变换错误合并为 RGB。
- 含 `satMod` 或超过 100% 的 `lumMod` 的颜色保留有序颜色变换（最多 64 项），包括同一颜色上的 `tint`、`shade` 和亮度调整。预览在解析主题引用后按序计算，原生导出保留顺序；旧文档包的既有颜色属性保持原有解释。
- 富文本、段落、列表、字段缓存、超链接、文字框适配、表格网格和合并、单元格外观及备注正文写为可编辑内容。保留源字体名称、UTF-16 和完整变体序列，不把预览替代字体写回文件。
- 柱、线、饼图的当前数据同时生成缓存、范围和最小内嵌 XLSX 常量表；共享数据引用复用同一工作簿。字符串使用明确的字符串单元格，空值不变成零或公式，不加入外部连接或自动刷新。
- 内嵌工作簿与外层 PPTX 共用文件数、XML 节点和字节预算；所有生成的 XML 再经过有界解析。超预算或无法表达的正文格式直接失败，源文档保持不变。

原生幻灯片尺寸受 OOXML 范围限制，为每轴 72～4032 pt。部分合法 HTML/CSS 不能直接映射为本轮原生格式，例如文本边框、嵌套文本尺寸／内边距、百分比段落边距／内边距、独立词间距、文本底纹、不同 ASCII/HANSI 字体、未支持的字体权重、独立下划线颜色或非整数的项目符号百分比；适配器明确拒绝。组的主题效果引用无法转成所支持的原生组内联效果时也明确失败。已定位的备注保留原几何；普通备注生成确定的原生正文占位框，备注打印版式编辑继续延期。

## 批注与报告

默认 `preserve` 保留每条批注的来源格式；无来源登记时使用文稿的默认格式。现代批注保留富文本、回复、作者、状态，以及能够精确映射的幻灯片／对象／UTF-16 文字锚点。文字上下文包含按规范计算的长度和哈希；对象路径包含组和创建身份。

原生格式无法表达的区域、内部段落节点或丢失锚点，保留完整现代线程并使用未知锚点，报告 `anchor` 损失；没有可归属幻灯片时列入 `unexportedAnnotations`。不猜测新页面或进行语义重定位。缺少原生必需时间戳时使用中性时间，并记录 `metadata` 项；文档包仍保留原始可选信息。

`classic` 兼容导出使用页内点和纯文字，逐项报告精确锚点、线程、状态及富文本损失。回复正文保留在根批注文字中。`annotationLosses` 不代表正文被截断；正文无法表达时整个导出失败。保存文档包可完整保留协作数据。

LibreOffice 会丢失现代批注线程。现代结构及 SDK 精确往返、LibreOffice 实际表现和 PowerPoint/WPS 实机结果分别记录，不能互相替代；当前 PowerPoint/WPS 标记为未验证。

## Node 和 Playground

`createDocumentConverter` 复用预览的第一方固定画布，支持 PPT 文档包的 PDF 和指定页 PNG。字节协议的输入格式增加 `pptx`，需要宿主提供有界 `documentId`；保留已有 `docx` 和 `document-package-zip`。

```ts
const pdf = await converter.renderPdf(model);
const png = await converter.renderPage(model, { pageIndex: 0, scale: 1 });
if (pdf.report.kind === "presentation") {
  console.log(pdf.report.slides); // 稳定 slideId、源根、顺序、隐藏标记、pt 矩形。
}
```

`pages` 保留既有公共画布投影；PPT 的明确语义通过 `kind: "presentation"` 和 `slides` 提供。Node PDF 包含文稿的全部逻辑幻灯片，PNG 的 `pageIndex` 从零开始。隐藏页的预览标签不会进入输出。每项任务拥有独立浏览器／上下文、授权字节、取消和资源预算，完成或失败均关闭任务资源。

Playground 提供新建 PPT、PPTX 与文档包打开、协作样例、第一方编辑、固定提案、原生 PPTX、PDF/PNG、导出报告和包下载重开。PPTX 导出使用短生命周期 Worker；切换文稿、取消、卸载或修订变化会终止任务。宿主检查任务、文档和修订，迟到的输出不会下载。固定提案仅演示宿主协议，不增加模型服务。

## 可复验命令与证据

```bash
pnpm check
pnpm check:pptx-export
pnpm check:presentation-closure
```

外部验收固定 LibreOffice 25.8.1.1 Impress/Calc，另需 Python UNO、lxml、Poppler、Noto Sans CJK SC 和 Liberation Sans；不在 SDK 运行时下载字体或启动 Office。`check:pptx-export` 使用独立样例和生产者版本，保存源、原生文件、SDK 重开包、报告和环境摘要；用未修改的官方 Transitional XSD 校验标准部件及内嵌工作簿，并检查所选 MS-PPTX 现代批注结构。另通过隔离的 UNO 进程实际修改标题、几何、表格、图表数值和备注，重存后用 SDK 读取。

`check:presentation-closure` 在工作区之外安装五个 tarball，验证无 React 的普通 JS、可选 React、原生输入、请求／提案、历史、导出、Node PDF/PNG／子进程协议与 Office 重开；按固定字体和视口比较独立样例每页的源／原生导出预览。源图、输出图、报告、工具和字体摘要及 tarball 留在 `.cache/presentation-closure-acceptance/`。不自动改写参考基线。

固定 Office 工具的校验安装及 24.2 重存重复身份限制见[本地开发](development.md#依赖与状态)。SDK 不修补或忽略外部应用生成的重复对象身份。

标准校验包及许可证位于 [scripts/fixtures/ooxml](../scripts/fixtures/ooxml/README.md)。CI 同时保留 Word 闭环。达智接入仍按用户决定延期，不能以本轮独立 PPT 验收宣称 R1-16 的宿主验收已完成。
