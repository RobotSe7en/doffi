# 工作台与请求／提案闭环

R1-09 在已有预览、选择和批注交互上增加多批注请求、宿主交接、固定提案导入、候选差异及整批应用。普通 JavaScript 和可选 React 包装使用同一界面。文档源、事务、锚点及历史仍由 `core` 会话管理；模型执行、身份和正式存储由宿主负责。

## 核心 API

`createDocumentSession(model, options)` 返回 `DocumentSession`，包含原有 `DocumentEditingSession` 方法及以下入口：

```ts
const request = session.buildInteractionRequest({
  annotationIds: ["a-summary", "a-heading"],
  supportedOperations: ["replaceText", "setStyle"],
});
const candidate = session.prepareProposal(request, proposal);
// candidate.preview、changes、changedAnnotationIds 供查看；准备不会改变文档。
session.commit(candidate);
// 拒绝时使用 session.discard(candidate)。
```

请求包含当前修订下开放批注的冻结副本。文字所在行内节点提升到完整区块；父子区块重叠时合并，并保持文档顺序。原生富批注包含完整批注正文。空白区域提供所属容器的完整子树。当前采用保守上下文：保留有序的全部样式、解释区块需要的完整元数据以及当前包内可用资源描述，资源字节不进入请求。

默认 `DEFAULT_INTERACTION_LIMITS` 为 4 MiB 请求 JSON、10,000 个具有稳定身份的源节点及 256 个资源描述。`limits: { maxBytes, maxNodes, maxAssets }` 可整体覆盖；各项必须为正安全整数，最终请求还受会话文档包大小预算约束。超限明确失败，不截断子树、不丢弃已选批注。生成请求不创建修订、历史或文档变更事件。

`prepareProposal` 验证版本、文档及请求关联、当前修订、批注副本、完整源区块、样式／元数据／资源和操作子集。操作目标必须位于请求完整子树内；HTML 仍通过原事务解析及安全校验。请求可 JSON 保存后传回；相同文档 ID 不提供访问授权，宿主负责绑定主体和会话。

应用有效内容变化时，提案声明已处理且映射有效的批注与内容在同一修订、同一撤销步骤中提交。未声明批注保持开放；映射失败的目标保留 `needs_reanchor`。空提案或没有净内容变化的提案不自动解决批注、不推进修订，也不清除重做历史。旧请求、旧提案和旧候选均按修订拒绝；查看旧差异不表示可以应用。

## 普通 JavaScript 挂载

```ts
import { mountDocumentWorkbench } from "@doffi/editor";

const view = mountDocumentWorkbench(container, session, {
  onSubmitAnnotations: async (request) => host.submitDocumentRequest(request),
  onError: (error) => host.showError(error),
});
// 宿主取得对应原请求及返回产物后：
view.loadProposal(request, proposal);
// 宿主销毁自己的界面；session 仍由宿主管理。
view.dispose();
```

界面提供开放批注复选、生成请求、下载请求 JSON、统一提交、提案 JSON 文件导入、逐处差异、候选预览、整批应用及拒绝。内容差异显示前后文字，并可展开查看结构和样式。应用后的文档包下载由宿主通过 `session.serialize()` 和 `office` ZIP 入口完成。

公开视图保留 `selection/fonts/setMode/setZoom/focusAnnotation/ready/dispose`，并提供 `request/buildRequest/submitAnnotations/loadProposal/applyProposal/rejectProposal`。宿主可用 `await view.setEditing(true)` 切换正文输入，通过 `executeTextCommand` 执行字号、字重、斜体、下划线、文字颜色或左／中／右对齐；命令复用正文输入的选区、事务和撤销历史，未进入编辑或输入法组合期间明确拒绝。`formatState()` / `onFormatStateChange` 报告当前文字与混合格式；`clipboard("copy" | "cut" | "paste")` 复用原生输入的安全序列化，浏览器不允许读取剪贴板时可使用键盘粘贴。`updateFields()` 显式更新页码与目录。`ready` 等待当前正文及候选预览。卸载释放监听、候选、预览及资源 URL，不销毁外部会话；晚到提交和文件读取结果不重新挂载界面。

工作台一次审阅一个候选；应用或拒绝当前提案后再生成新请求或载入下一份提案。界面在审阅期间禁用对应入口，公开 API 也返回明确的 `InvalidOperation`，保留当前候选。该流程只需一个候选的数量及字节预算，不会叠加未提交候选。

`onEvent` 提供 `selectionChanged`、`annotationsSubmitted` 和 `proposalReady`。提交事件表示宿主成功接受交接，不表示模型执行完成或文件已保存。宿主提交失败返回 `HostSubmissionFailed`，文档与批注不变；未配置提交回调时仍可下载请求。

## React 与 Playground

可选入口 `@doffi/editor/react` 导出 `DocumentWorkbench`，接受 `session`、工作台选项及 `onViewChange`。会话由宿主创建和释放；普通入口不导入 React，非 React 消费者不需要安装 React。回调变化通过最新引用传递，不重建文档视图；会话及上下文配置变化才重新挂载。

同一入口还导出 SDK 自带的 `OfficeRibbon`、`OfficeRibbonProps`、`OfficeActions` 和 `OfficeIcon`。Word、PPT、Excel 共用功能区组件；字体框、菜单、图标、禁用状态和作用域样式随 SDK 提供，不需要复制 Playground 组件或引入它的 CSS。`mode: "edit" | "preview"` 控制界面：预览下默认收起“开始”的编辑面板，禁用编辑标签，仍可打开文件、审阅和视图面板。样式与导航身份按组件实例隔离，支持一个页面挂载多个工作台。普通 JavaScript 的原有 DOM 工作台继续可用；该组件属于可选 React UI，不改变普通入口的依赖。

| `OfficeRibbonProps` 字段 | 责任 |
| --- | --- |
| `kind`、`mode`、`tab`、`onTab` | 文档种类、编辑／预览与当前标签；`kind` 为 `document`、`presentation` 或 `workbook`。 |
| `commands`、`onCommand` | 传入原生视图的命令状态，并执行同一个原生命令入口。 |
| `formatState`、`onTextCommand` | 传入当前格式与类型化文字命令回调；内容和历史仍由 SDK 原生编辑器及核心会话管理。异步工作簿提交完成后才移动字段焦点。 |
| `actions` | 宿主提供已授权的打开、保存和导出等回调及其可用性；组件不自行读取文件、存储或执行模型。 |
| `layout`／`onLayout`、`zoom`／`onZoom`、`rulers`／`onRulers` | 宿主的显示偏好与 SDK 视图控制接线，不创建另一份文档源。 |
| `busy` | 当前模式切换、导入或其他待完成操作的界面状态。 |

正文预览／编辑由工作台的 `await view.setEditing(boolean)` 切换；`editable: false` 隐藏编辑入口。关闭编辑会卸载正文输入及其格式工具栏，保留预览和批注。`await view.settled()` 可在画布可见时等待当前修订的分页；返回预览会自行等待有效分页，隐藏的预览不会随每次输入重复重建。纯读取需求可以直接挂载 `mountDocumentPreview`，不包含编辑工具栏。显示模式和入口可用性不能代替宿主权限控制。

Playground 提供报告协作样例和固定提案，使用公开 React 入口。固定示例提案调整当前请求的关联身份和修订后仍经过完整校验；它只能用于包含相应节点和内容的报告。界面交接给本地演示宿主，不执行模型。测试覆盖双批注请求、提案、应用、撤销重做、文档包下载重开及实际 DOCX 导入语义。

Playground 首页提供三种格式的样例卡片、协作样例和本地文件选择／拖放入口。打开文档后显示 Office 桌面式标题栏、快速访问、SDK 的 `OfficeRibbon`、纸张／幻灯片画布、批注及底部视图／缩放栏。Word 功能区按剪贴板、字体、段落、样式等分组，已接入的命令沿用 SDK 公共入口；未实现或当前模式不可用的命令使用原生禁用状态。导出位于“文件”，请求与提案位于“审阅”；Word 默认打开正文编辑，可切回只读预览；模式偏好存于 `doffi:editing`。预览时默认收起正文编辑面板，并保留读取和批注入口。三种格式的功能区读取各自原生编辑器的命令可用性、选区与文字格式状态，执行同一命令回调；连续切换加粗、斜体、下划线保留选区并正确开关，字体／字号／对齐沿用共同事务历史。Word/PPT 缩放保持可见文档位置和水平居中，不产生修订；工作簿缩放尚未实现，控件禁用。文档列表默认收起，其显示偏好存于 `doffi:library`，提供当前页面导入的文件、搜索与格式筛选。SDK 仍以 `appearance: "workspace"` 提供画布及批注，工作簿编辑器也接受这一外观选项。

Word 默认在文档视口上方和左侧的独立栏显示厘米标尺；“视图 → 标尺”可切换，偏好保存到 `doffi:rulers`。水平刻度以正文左页边距为零点，垂直刻度以正文上页边距为零点，均跟随当前页设置和缩放；连续视图只显示水平标尺。当前标尺用于显示尺寸与页边距，尚未实现拖动调整缩进、制表位或页边距。SDK 可通过 `setRulers(boolean)` 控制显示，派生标尺不进入源或导出。字体／字号字段接受直接输入，保留当前选区并沿用共享历史；缺失字体在字段旁提示替代显示。已核对行为、修复和剩余差异见 [Office 行为核对](playground-office-behavior-audit.md)。

页面左侧选择区支持单击选择实际显示的一行、拖动扩展至多行、双击选择段落及三击／Ctrl 单击选择正文。换行定位使用当前文本几何与 UTF-16 字符边界，不按固定字符数估算。选择本身不提交修订；后续格式或输入沿用现有事务与历史。表格行选择尚未接入此选择区。

返回首页及文档切换保留各自的会话、修订及撤销历史，刷新页面后清除；不提供正式存储。选择文字／对象／区域后，通过“新建批注”打开非模态批注框；选择本身不弹框，填写和取消不推进修订。编辑时批注面板继续可用。Playground 不显示修订调试状态；SDK 默认视图仍保留原有状态信息及选择触发批注行为，可用 `annotationTrigger: "command"` 切换为显式命令。窄屏将文档列表改为横向排列，批注卡片移到画布下方；功能区和 Office 画布在各自容器内滚动。开发服务默认监听 `0.0.0.0`，可通过 `pnpm --filter @doffi/playground dev --port 5174 --strictPort` 启动。

[R1-11 富文本输入](rich-text.md)、[R1-12 分页／PDF 渲染](pagination-rendering.md)及 [R1-13 原生 DOCX 导出](docx-export.md)已实现。第一轮历史与延期的宿主验收见[第一轮细则](iteration-1.md)；[第二轮 PPT](iteration-2.md)已实现格式操作、只读依赖、全局作用域及批注回复，见[PPT 工作台](presentation-workbench.md)。当前进入[第三轮 Excel](iteration-3.md)，工作簿协作尚未实现。

Playground 的 DOCX／PDF 下载、修订守卫、导出报告及独立打包消费验收见 [Word 闭环](word-closure.md)。
