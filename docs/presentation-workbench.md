# PPT 工作台与宿主协议

PPT 复用宿主拥有的 `DocumentSession`、冻结请求、隔离候选和统一事务历史。SDK 提供普通 JS 和 React 入口，不执行模型调用、存储或授权判断。原生导出另见第二轮 R2-14；本页描述 HTML 包与编辑协作。

```ts
import {createDocumentSession, openDocument} from "@doffi/core";
import {mountDocumentWorkbench} from "@doffi/editor";

const session = createDocumentSession(openDocument(authorizedFiles));
const workbench = mountDocumentWorkbench(container, session, {
  format: "presentation",
  onSubmitAnnotations: request => hostSubmit(request),
  onError: error => hostShowError(error),
});
await workbench.ready;
// 宿主保存 session.serialize()；卸载视图不销毁宿主会话。
workbench.dispose();
session.dispose();
```

同一适配器也由 `mountPresentationWorkbench` 导出；`@doffi/editor/react` 提供 `PresentationWorkbench`，接受 `session`、相同选项和 `onViewChange`。普通 JS 子路径不引入 React。原 Word `mountDocumentWorkbench` 默认行为和 `DocumentWorkbench` React 入口保持 Word 类型。

## 冻结范围

核心调用 `session.buildInteractionRequest({kind:"presentation", annotationIds, scope:"slides"})`。页面范围根据所选开放批注确定，提供完整页面对象树、相关备注和请求批注富文本，避免从截取文字猜测几何或关系。引用到的主题、母版、布局、备注母版是只读依赖；其他页面的内容和数据集不进入此范围。样式和资源描述用于解释当前来源，没有资源字节、凭据或 URL 获取权限。

`scope:"global"` 明确冻结整份文稿及可编辑定义。修改主题、母版、布局或文稿尺寸需要此范围。两类请求均受上下文节点、文字、字节和包预算约束，超限明确失败，不静默截断完整页面。共享图表数据或页面链接造成范围外既有页面变化时，页面提案也会拒绝；宿主可以重新生成全局请求。

独立契约名为 `presentation-interaction-request` 与 `presentation-edit-proposal`。它们分别导出类型 `PresentationInteractionRequest`、`PresentationEditProposal` 和相应 JSON Schema 子路径。Word 的 `interaction-request`、`edit-proposal`、七类内容操作名称及 Excel 的未实现边界保持原契约。完整示例见 [PPT 请求](../examples/presentation.request.json)和[对应提案](../examples/presentation.proposal.json)。

## 提案和资源

宿主回传对应请求身份、文稿身份和基准修订的提案。`prepareProposal(request, proposal, prepareOptions?)` 重建并核对冻结上下文，检查操作声明、可写页面及源身份、只读依赖、请求批注和实际间接影响。旧修订、伪造上下文、范围外操作、未知资源均拒绝。新对象、系列和定义的身份由核心分配，不能猜测同批新身份。

新增图片的描述必须由宿主在请求构建时通过 `assets` 声明；授权字节在候选准备时随 `TransactionPrepareOptions.assets` 提供。工作台的 `proposalResources` 同时完成描述声明与候选字节交付。摘要、MIME、路径和预算仍由核心校验。候选的 `model` 是只读的完整下一版本，带候选资源，预览无需污染当前文稿。

```ts
const request = session.buildInteractionRequest({
  kind: "presentation", annotationIds: [annotationId], scope: "slides",
});
const candidate = session.prepareProposal(request, proposal);
// candidate.model 用于只读预览；candidate.changes 与 affectedSlideIds 用于审查。
const receipt = session.commit(candidate); // 内容和处理的批注一起提交，一步撤销。
session.undo();
```

工作台提供多批注勾选、页面／全局范围、请求下载、宿主提交、JSON 提案导入、候选预览、内容与几何差异、图表常量表、定义及受影响页提示。`loadProposal` 准备候选，`applyProposal` 一次应用，`rejectProposal` 释放候选。预览未完成前界面禁用应用；源修订变化或卸载会清理预览和资源，并忽略迟到的宿主回调。视图不修改宿主的保存策略。

## 锚点和验证

人工与提案共用 UTF-16 映射、结构锚点和区域失效规则。测量区域还绑定当前画布、字体和布局环境；环境改变后需重定位，再生成请求。经典批注的原生固定点按页内坐标保留，不作为字体测量区域处理。

核心测试覆盖冻结、严格契约、预算、只读依赖、间接影响、候选槽位释放、资源隔离、原子应用与旧修订拒绝。浏览器用例覆盖普通 JS 与 React、候选图片、差异应用及撤销、人工修改后迟到提案和字体环境变化。CI 还保留 Word 工作台和打包消费回归；外部 PowerPoint/WPS 原生导出重开不由这些编辑测试替代。

宿主自定义功能区可读取 `view.editor.commands()` 或订阅 `onCommandStateChange`，通过 `executeCommand(id)` 调用与 SDK 内置按钮相同的操作。命令可用性依赖当前选区、输入法及异步更新状态；不可用或已卸载的命令明确拒绝。`formatState()` / `onFormatStateChange` 与 `executeTextCommand()` 供宿主显示文字格式并设置开关、字体、字号和对齐，复用原选区及事务历史。`annotationTrigger: "command"` 关闭选择即弹批注的行为，宿主通过 `startAnnotation()` 显式打开批注框。

Playground 通过 `setReviewVisible(boolean)` 显式显示／隐藏请求与提案区，避免仅依赖跨 Shadow DOM 的 CSS 自定义属性动态切换，导致 WebKit 的绘制与控件可见状态不一致。显示状态不改变源、批注选择或历史。
