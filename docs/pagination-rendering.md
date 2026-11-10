# 分页、共享画布与转换

R1-12 使用自研测量分页。预览、批注和正文编辑显示相同的 HTML/CSS 页面画布；Node PDF/PNG 使用同一个 `render` 包。没有引入 Tiptap、ProseMirror、Paged.js 或 Office 转换库。浏览器预览直接显示 HTML，不先转换 PDF。

## 单独使用预览

```ts
import { openDocument } from "@doffi/core";
import { mountDocumentPreview } from "@doffi/render";

const model = openDocument(authorizedPackage);
const view = mountDocumentPreview(container, model, { layout: "pages" });
await view.ready;
const report = await view.settled();
// 宿主更换模型时：view.update(nextModel)
// 宿主卸载时：view.dispose()
```

`authorizedPackage` 是宿主已授权的包内路径／字节 Map，`container` 是宿主 DOM 容器。只读入口接受 `DocumentModel`，不创建编辑会话，不提供输入、批注提交或提案应用。`editor` 保留同名兼容导出；达智可先只依赖 `core`、`office` 和 `render`。需要选择／批注时再挂载 `mountDocumentInteraction`，需要输入时使用 `mountDocumentTextEditor` 或完整工作台。工作台的 `editable: false` 只隐藏正文输入入口，仍能应用提案，不能代替纯预览入口。

默认 `layout: "pages"`；显式 `continuous` 使用连续读取视图和字段缓存。Playground 的“页面视图／连续视图”属于宿主显示偏好，不保存到文档源。`ready` 等待首次字体和布局；`settled()` 还等待当前图片解码并再次测量。异步等待期间更新或销毁会使旧的 `settled()` 失效，返回 `InvalidOperation`。

画布处于 `display:none` 等不可测量状态时不执行分页，也不移除已接受的页间间距。隐藏挂载的 `ready` 只表示资源及视图已准备，不保证已有页面坐标；恢复可见后可用 `settled()` 等待有效几何。渲染器同时观察 iframe 与正文，确保宿主显隐也会触发测量。工作台编辑期间暂停隐藏预览的内容重建；`setEditing(false)` 会从当前会话重建可见预览并等待分页完成，避免多页 DOCX 叠在一起。切换与重测不产生源修订或撤销步骤。

## 分页与编辑约束

单一 DOM 保持正文顺序。分页在实际行／表格行边界放置带运行时所有权的空白间距，只切开 DOM Text，不分配新的源段落身份。序列化移除间距、重复故事、表头、字体替代和包装，仍保存原有 HTML 包。输入后的布局更新使用同一模型；中文输入法组合输入期间保留输入法拥有的临时 DOM，在组合结束提交时重新分页。

支持纸张尺寸和边距、连续／新页／奇偶页分节、明确分页、页码格式与重启、标题随下段、孤行控制、表格跨页和连续重复表头。奇偶页空白使用新节纸张尺寸，不显示页眉页脚；此规则与 LibreOffice `IsSkipEmptyPages:false` 的参考结果对照。分节依源 DOM 顺序处理，元数据数组顺序不会重排正文。节外的正文保留位置并沿用前一节设置，首节之前的内容沿用首节设置。

第一页、奇偶页及后续节继承的页眉页脚从源故事生成。可用区域为页边距减页眉／页脚距离；超过区域会裁剪并返回 `ContentOverflow`，不覆盖正文。正文绘制限制在页面正文区，页间空白不进入源文本。纸张边距仍可作为空白矩形批注区域。

当前有界布局要求节根直接位于主故事内且互不重叠；嵌套节、单行／图片超过可打印高度、跨页长行中的非顶端对齐单元格明确返回 `UnsupportedFeature`，可选择连续布局查看源。跨浏览器的字体度量可能不同，分页报告不承诺与 Word/WPS 逐像素一致。字体保留、替代和缺字见[中文字体](fonts.md)。Word／WPS 本机尚未验证；LibreOffice/Poppler 是本轮参考工具。

`DocumentRenderReport` 绑定文档 ID 和修订，包含物理页、打印页码、pt 几何、源节点页映射、字段结果、字体报告及诊断。区域批注仍使用主故事坐标；整个矩形位于一张纸内时增加实际 `pageId`，布局变化仍要求显式重定位。

## 字段和复制

PAGE、NUMPAGES 和基本标题大纲 TOC 是派生显示；TOC 目前支持标题层级及 `\\o` 范围，不生成原生可点击目录链接。锁定字段和未支持的字段／开关组合保留缓存并报告 `UnsupportedField`，不默认为正确的新结果。

```ts
import { createLayoutFieldTransaction } from "@doffi/render";
const transaction = createLayoutFieldTransaction(session, report);
session.commit(session.prepare(transaction));
```

只有显式事务会更新字段缓存；报告必须匹配当前修订，整批更新只占一个撤销步骤。读取和打印不会改写缓存。

复制使用实际选区，排除页面间距、编号标记、重复表头和字体替代。字段按选中的显示子串复制为文字，清除原字段身份；完整保存仍保留原字段及缓存。纯文本复制保留嵌套段落换行和表格单元格制表符。

## 安全和预算

源包始终先经过核心验证。iframe 不含文档源脚本，首位 CSP 使用 `default-src 'none'`，只允许内联样式及授权 Blob 图片，不允许字体或其他网络加载。为使 WebKit 也能执行宿主绑定的交互事件，sandbox 使用 `allow-same-origin allow-scripts`；它本身不再构成脚本隔离，源校验和 CSP 是执行边界。三引擎回归等待实际 CSP violation，确认插入脚本和 HTML `onerror` 均不能执行。SDK 不把源 data 属性当作运行时所有权。

默认 `maxNodes=50000`、`maxPages=256`、`maxLayoutPasses=8`、`maxPrintNodes=250000`，均可由宿主配置正整数。页数、重复内容、测量工作或字段收敛超限均明确失败。布局报告和浏览器 DOM 是派生状态，不提供存储或访问授权。

`createOpaqueId()` 使用 Web Crypto 原生 UUID，或使用 `getRandomValues` 生成 UUID v4，支持 HTTP 开发地址没有 `randomUUID` 的浏览器；不使用弱随机回退。文档、批注和请求 ID 都不授予访问权限。

## Node PDF、PNG 和字节协议

```ts
import { createDocumentConverter } from "@doffi/node";
const converter = createDocumentConverter({ maxConcurrency: 2 });
try {
  const pdf = await converter.renderPdf(model);
  const png = await converter.renderPage(model, { pageIndex: 0, scale: 1 });
  // bytes、mimeType、report、engine 由宿主保存或返回。
} finally {
  await converter.close();
}
```

宿主安装／指定 Chromium，并准备本地字体。SDK 不下载浏览器或字体。每个任务启动独立浏览器和上下文，网络仅允许 Blob／空白页，不继承模型凭据，不缓存文档内容。取消、超时和关闭会终止各任务拥有的浏览器；尚未结束的后台启动／清理继续占用并发槽，`close()` 等待实际清理结束。

PDF 将测量画布按真实纸张裁切，保留可搜索文字；PNG 使用同一页报告和尺寸。默认并发 2、超时 30 秒、输入／输出各 64 MiB、PNG 16M 像素。请求不能提高宿主预算。输出包含 Chromium 实际版本；不同字体、引擎版本的视觉结果分别记录。

`doffi` CLI 从 stdin 读取逐行 JSON V1，从 stdout 输出关联任务结果，支持授权 DOCX／文档包 ZIP 字节。无 URL 或本地文件输入路径：

```json
{"protocolVersion":1,"taskId":"host-task","format":"pdf","input":{"format":"docx","documentId":"host-document","bytesBase64":"UEs..."}}
```

示意的 `bytesBase64` 需替换为真实规范 Base64。DOCX 必须由宿主提供文档 ID；包中已有身份可额外断言匹配。协议严格拒绝未知字段、版本、格式、超限帧和非规范 Base64，错误仍关联 `taskId`；结果字节位于 `output.bytesBase64`。宿主可调用 `executeConversionRequest` 或自行管理 CLI 进程。

`@doffi/node/dev` 的 `runConversionProcess` 提供有界单请求子进程，`createDevelopmentConversionMiddleware` 用于开发：默认检查实际 loopback socket，也允许宿主显式配置确切的 `allowedOrigins`；两种方式均检查 Host、实际监听端口、Origin、自定义请求头及上传期限，中间件不打开端口。Playground 的 PDF 按钮使用它；切换文档、修订变化、取消或卸载会丢弃迟到结果。远程 HTTP 的预览／编辑独立可用；远程开发转换的明确同源配置见 [Word 闭环](word-closure.md#远程开发访问)。正式远程转换服务的认证和部署由宿主提供。

## 验证环境

Playwright 1.64.1 固定 Chromium 156.0.8078.4、Firefox 157.0、WebKit 27.2；额外使用 Chromium 109.0.5414.46 验证没有新列布局能力的内核。旧 Chromium 安装脚本校验官方归档 SHA-256 `1ec3528cecff35bd20184dc7fd07deb794624023607388b0ea8fda696730805d`。分页实现不依赖 Chromium 145 的 `column-height`／`column-wrap`。

Linux WebKit 验证记录为引擎验收，不冒充实际 Safari。CI 另设 macOS 15 的 SafariDriver 验收，结果记录 Safari 实际版本；本机没有 Safari。Word／WPS 尚未验证。固定字体、视口、引擎和图片摘要见[分页参考记录](../examples/render-reference.json)。读取参考使用固定字体，分页、重复表头 PNG、真实 PDF 文字／页数、CLI、取消、资源预算、编辑撤销和三引擎交互均有独立回归。

PPT 的固定画布与 Node PDF/PNG 已接入同一转换器；字节输入支持 `pptx`，报告提供 `kind: "presentation"` 和 `slides`，保留既有公共页面投影。输出包含全部逻辑幻灯片，隐藏页标签不进入 PDF/PNG。具体范围与验收见 [PPTX 导出与闭环](pptx-export.md)。
