# 预览与交互

R1-08 在现有无脚本只读预览上提供文字划选、对象选择、矩形框选和批注面板。`editor` 的普通交互入口只依赖 `core`，普通 JavaScript 宿主通过公开入口挂载；Playground 使用可选 React 包装管理同一组件。请求生成、统一提交和提案界面见已实现的 [R1-09 工作台](workbench.md)，富文本输入见 [R1-11 正文输入](rich-text.md)，R1-12 [共享画布与分页](pagination-rendering.md)也已实现，纯预览可直接使用 `render` 而无需编辑会话。

## 挂载与会话

```ts
import { createDocumentSession, openDocument } from "@doffi/core";
import type { DocumentPackage } from "@doffi/core";
import { mountDocumentInteraction } from "@doffi/editor";

declare const files: DocumentPackage;
declare const container: HTMLElement;
const session = createDocumentSession(openDocument(files));
const view = mountDocumentInteraction(container, session, {
  onSelectionChange(selection) {
    // 临时选择绑定 documentId 和 baseRevision；本回调不推进修订。
    console.log(selection);
  },
  onError(error) { console.error(error.message); },
});
await view.ready;
view.setMode("region");
view.setZoom(1.25);

// 保存唯一的文档包来源，预览 DOM 和高亮不参与序列化。
const saved = session.serialize();
view.dispose();
session.dispose();
```

`mountDocumentPreview(container, model)` 继续提供独立只读入口。两个入口均使用隔离 iframe，文档脚本不执行；交互组件的工具栏和批注面板位于 Shadow DOM，宿主样式不会覆盖文档或面板内部样式。资源 URL 只在当前预览存活期间存在，保存仍使用包内路径。

| 接口 | 行为 |
| --- | --- |
| `element` | 当前组件根元素 |
| `ready` | 当前渲染任务的 Promise；宿主提交内容后，先让会话通知进入微任务，再重新读取此属性等待新预览 |
| `selection()` | 冻结的 `DocumentSelection` 或 `undefined` |
| `fonts()` | 当前预览的字体替代和缺字报告 |
| `setMode("text" / "node" / "region")` | 切换选择模式，清除旧的临时选择 |
| `setZoom(zoom)` | 50% 至 200% 的显示缩放，不推进修订 |
| `openWordNavigation("find" / "outline" / "wordCount")` | Word 只读查找、标题导航和统计；不创建编辑候选，不推进修订或历史 |
| `focusAnnotation(id)` | 滚动并高亮有效批注；失效或不可显示的位置报 `AnchorNotFound` |
| `dispose()` | 可重复调用，解除订阅、监听和观察器，释放 iframe、URL 与 DOM 引用；不销毁宿主传入的会话 |

会话通知在微任务中合并，内容变化重建预览，纯批注变化更新面板。组件不维护第二套可保存的文档或历史。宿主负责结束其拥有的会话；Playground 在文件切换时释放组件，保留当前页面内各文档的隔离会话和历史，React 卸载时统一释放会话。关闭期间尚未完成的渲染不会再连接到新文档。

交互组件和工作台可传入 `appearance: "workspace"`（React 同名属性），使用顶部工具栏、画布与右侧批注卡片；默认外观保留原有内嵌布局。此模式下，文字、对象或矩形选择后在附近显示非模态批注浮窗，点击输入框填写，支持 Ctrl/Command+Enter 提交、Escape 或取消按钮放弃。浮窗不抢夺正在划选的焦点，位置限制在浏览器视口内；切换修订或文档后清除旧选择和草稿。PPT 工作台使用相同外观选项和浮窗，批注仍写入原生经典／现代批注事务。

## 选择与定位

文字选择只参与正文内容域，辅助页眉、页脚和原生批注正文继续显示，不混入正文选择。映射在预览装饰之前记录规范节点的源位置；编号、辅助标签、制表位引导符和字体诊断不进入选区文字。换行占一个 LF，图片占一个 U+FFFC，原生制表位使用源文字长度。文字范围按 UTF-16 记录，拆开代理对或字符与变体选择符的范围明确拒绝。

行内格式下的连续文字优先合并为所在段落的范围；跨段落、单元格或不连续的可见片段生成按文档顺序排列的 `text-set`，隐藏字段缓存不会被当作可见选择。所有范围仍由核心校验原文、顺序及内容域。

对象模式点击稳定节点，Ctrl/Command 加点选可增减对象集合，集合按源文档顺序保存，界面显示对象原文或图片说明。矩形模式支持正向或反向拖动，限制在正文逻辑渲染面内；命中保存最内层的相关节点，空白区域仍生成 `nodeIds: []` 的区域锚点，并显示其属于正文空白区域。

连续视图的 `surfaceId` 是正文容器，矩形原点是该容器的边框左上角，单位为 pt。计算去除 iframe／宿主滚动和组件缩放，不生成虚构的 `pageId`。自动分页及跨页交互进入 R1-12。

临时选择不会推进修订；会话推进修订后清除旧选择，防止将旧偏移用于新内容。批注定位使用源位置重新测量，显示高亮不写入文档。Escape 清除选择；文档和工具栏中的 Ctrl/Command+Z、Shift+Ctrl/Command+Z 使用核心历史，批注输入框保留原生输入历史。

## 批注与布局变化

面板支持新增批注、修改文字说明、删除、处理状态切换和用当前选择重新定位。导入的富文本原生批注显示源正文，保留其富文本内容域；正文直接编辑随 R1-11 实现。所有持久化动作通过 `prepare/commit`，撤销重做使用同一会话历史。过期选择、非法锚点及预算错误保留原文档，并在组件中显示错误。

区域记录 `layoutKey`。该标记由当前渲染实现、正文节点几何和实际选用字体派生；滚动和组件缩放不改变其含义。源内容、样式和资源变化的失效规则由核心负责；字体、容器和视口引起的实际布局变化由预览观察器检测，经批注事务标记 `needs_reanchor`，完整保留原矩形、节点集合和修订。用户重新选择后显式提交新锚点及有效状态。

布局标记是本渲染器的环境标记，不承诺其他渲染器使用同一值，也不将旧矩形自动平移或缩放到新布局。当前仍是连续只读正文，不能以区域批注代替实际编辑或自动分页。

## 验证

`node scripts/playwright.mjs test interaction.spec.ts` 验证独立普通 JavaScript 宿主和 Playground：文字／对象／区域锚点、UTF-16、跨段落及表格、源文字与装饰的区别、缩放坐标、批注历史、保存重开、外部内容提交、样式隔离和资源释放。`pnpm test:browser` 另外运行已固定字体与视口的 DOCX 视觉回归。字体及渲染基线沿用[读取验收](docx-import.md)和[字体记录](fonts.md)，不自动下载字体。
