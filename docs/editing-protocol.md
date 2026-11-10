# 编辑与宿主交互协议

人工操作和宿主智能体返回的修改共用文档操作协议。核心管理事务、修订、批注和撤销；编辑器提供选区、批注和差异界面；宿主负责智能体执行、身份授权和文件保存。

结构契约、推导类型和 JSON Schema 已在 R1-03 实现，公开入口见[结构契约与 JSON Schema](contracts.md)。R1-10 已实现 DOCX 读取及只读预览；R1-06 已实现[会话准备、提交及历史](document-session.md)，工厂现返回完整 `DocumentSession`；R1-09 已实现上下文生成、`prepareProposal` 和[工作台协作 UI](workbench.md)。`protocolVersion: 1` 仍为未发布的开发设计值，首次发布前完成冻结；达智边界使用同一契约生成或校验 Pydantic 模型。

所有协议对象严格拒绝未知字段，包括元数据和嵌套操作。不静默删字段、填默认值或强制转换类型；新增字段通过显式契约演进。结构校验通过不等于 HTML 安全、节点引用、授权或事务语义已经通过。

## 职责与接口

| 责任 | 所属 |
| --- | --- |
| 文档、批注、相关 HTML、样式和支持操作的读取 | `core` |
| 请求与提案关联、操作校验、候选和历史 | `core` |
| 选择、批注填写、统一提交、差异和应用按钮 | `editor` |
| 对话、整体任务要求、提示词、token 预算、模型调用、工具、取消和运行状态 | 宿主智能体 |
| 来源绑定、授权资源、草稿和正式提交 | 宿主 |

SDK 不设置独立模型调用入口或文档内 AI 对话弹窗。统一提交只把数据交给宿主；文档组件不通过模型供应商 SDK 执行请求。

下列签名表达完整协作接口职责。当前工厂返回完整事务及请求／提案接口：

```ts
interface DocumentSession {
  snapshot(): DocumentSnapshot;
  getAnnotations(): readonly Annotation[];
  buildInteractionRequest(input: InteractionRequestInput): DocumentInteractionRequest;
  prepare(transaction: DocumentTransaction): CandidateChange;
  prepareProposal(
    request: DocumentInteractionRequest,
    proposal: EditProposal,
  ): CandidateChange;
  commit(candidate: CandidateChange): TransactionReceipt;
  undo(): TransactionReceipt;
  redo(): TransactionReceipt;
  serialize(): DocumentPackage;
  subscribe(listener: DocumentEventListener): () => void;
  dispose(): void;
}
```

选区由编辑器维护并转换为中性的锚点数据，核心不读取 DOM。导入、导出和宿主保存是独立入口。`commit` 修改会话，不表示文件已保存。`dispose` 释放该会话拥有的订阅、任务和文档引用。

## 选区与批注

锚点由所属批注或请求绑定文档 ID 与基准修订。UI 的临时选择、缩放和高亮不改变修订；批注增删改、状态变化属于持久化编辑，进入事务。

| 锚点 | 表达 |
| --- | --- |
| 文字 `text` | `nodeId`、`from`、`to`、`quote` |
| 跨节点文字 `text-set` | 按文档顺序排列的 `ranges`，每项使用同样的文字范围字段 |
| 对象 `node` | `nodeId` |
| 对象集合 `node-set` | 稳定、有序的 `nodeIds` |
| 区域 `region` | `surfaceId`、pt 矩形 `rect`、命中 `nodeIds`、布局环境 `layoutKey`；真实分页后可附带 `pageId` |
| 单元格范围 `cells` | 稳定工作表 ID 与多个精确矩形片段，随事务映射和修订重绑；工作簿编辑范围见[阶段验收](workbook-editing-acceptance.md) |

文字偏移使用规范节点文本的 UTF-16 索引，与 JavaScript 字符串和 DOM Range 一致。换行占一个 `\n`，内嵌对象占一个 `U+FFFC`；初次导入规范化后不再隐式进行 Unicode 归一化。

第一轮先完成连续内容编辑，再提供派生分页。连续视图使用文档根对应的逻辑渲染面；分页视图使用实际计算的页面及布局标记，不把自动分页写成正文节点。矩形坐标排除界面缩放和滚动偏移。区域可以命中多个节点，也可以位于空白；空白区域附带所属容器和相邻结构，Word 中通过正常文档流插入或受支持的文档布局操作表达修改。

批注包含 `annotationId`、`documentId`、`baseRevision`、`anchor` 和 `status`，正文在纯文本 `instruction` 与富文本 `bodyNodeId` 之间二选一。来源作者显示信息通过严格的 `author.name/initials` 保存，不能作为授权身份。PPT 已定义严格的 `replies`，支持独立回复身份、正文、来源显示和工作流状态；Word 仍拒绝回复。公共状态包含 `open`、`resolved`、`needs_reanchor` 和仅 PPT 线程开放的 `closed`，不能据此宣称 Word 现代线程可用。PPT 的人工线程操作和请求／提案仍按第二轮后续任务实施。

事务产生节点与文字偏移映射。能够明确定位的批注更新到新修订；目标删除、整体替换或区域布局变化后无法确定位置时进入 `needs_reanchor`。原摘录和原区域保留用于人工定位，不猜测语义位置。R1-07 已实现文字与结构操作的精确映射、跨节点整体失效及显式区域重定位；边界插入不纳入选区，局部修改更新摘录，完整覆盖范围时失效。完整规则见[锚点映射与重新定位](document-session.md#锚点映射与重新定位)。

## 文档操作

| 操作 | 作用 | 主要前置条件 |
| --- | --- | --- |
| `replaceText` | 替换节点内文字范围 | 节点、范围、`expectedText` 匹配 |
| `replaceFragment` | 替换指定节点的子内容 | 目标结构、HTML 和引用合法 |
| `reconcileFragment` | 在目标子树保留已有身份并协调内容 | 身份来自基准子树；UTF-16 映射证据与最终文字一致，见[正文输入](rich-text.md) |
| `insertFragment` | 在指定位置插入 HTML 区块 | 位置合法，新节点 ID 由核心分配 |
| `removeNode` | 删除内容或对象节点 | 目标存在且允许删除，不能删除文档根 |
| `setStyle` | 设置目标节点样式 | 属性和值在支持范围内 |
| `replaceAsset` | 替换图片等资源 | 新资源由宿主接收校验且属于当前会话 |
| `setGeometry` | 调整对象几何和层级 | 几何及对象类型匹配，PPT 迭代实现 |
| `setCells` | 设置单元格值或公式 | 工作表、范围、值和公式合法，Excel 迭代实现 |
| `addAnnotation`、`updateAnnotation`、`removeAnnotation` | 管理批注及状态 | 批注身份、锚点、修订和状态合法 |

内容与批注操作均属于类型化事务。`DocumentOperation` 是 `DocumentContentOperation` 与 `AnnotationOperation` 的联合；提案只使用前者。交互请求向智能体声明当前格式和宿主配置允许的内容操作，不把批注管理操作开放为任意模型写入入口。提案应用产生的批注状态更新由核心附加。

`setStyle` 在首轮写回目标节点的行内 CSS 声明，已有样式表继续参与级联。不在元数据中保存另一份颜色、字号等值。

当前结构契约提供七种内容操作和三种批注操作，表中的几何与单元格操作属于后续计划。`insertFragment` 使用 `nodeId/position/html`，位置为同级的 `before/after` 或子内容的 `start/end`。`replaceAsset` 通过 `assetPath` 引用资源，未携带资源字节或新的访问权限。字段约束见[契约 API](contracts.md)。

内容操作目标的节点 ID 必须来自事务基准快照。新插入或复制的节点由核心分配 ID，提案不能猜测并在同一批内引用这些新 ID。首轮不额外引入临时节点引用协议。

## 事务语义

事务包含 `protocolVersion`、`documentId`、`baseRevision` 和操作列表。

1. 校验协议、文档、基准修订、完整操作结构和全部基准节点引用。
2. 在隔离候选中按顺序执行。每一步的范围和 `expectedText` 对照前序操作后的候选状态；目标被前序操作删除时失败。
3. 检查最终结构、样式、资源、格式语义及锚点映射，计算语义差异和受影响节点。
4. 返回绑定当前会话及基准修订的候选。布局测量由编辑器或渲染器派生，不由核心访问 DOM。
5. 提交时再次比较会话和当前修订；有效变化一次增加一个修订并产生一次撤销记录。

`prepare` 和 `prepareProposal` 不改变当前文档、历史和资源引用。任一步失败整批不变；净结果无变化不推进修订、不创建历史、不自动解决批注。候选不能跨会话使用，已消费候选不能重复提交。

撤销重做均产生新的当前修订号，不能复用旧修订。撤销后提交新修改清除重做分支。首轮保存变更前后的状态和必要映射，不可变资源共享字节引用；历史受配置预算约束。

## 统一提交请求

用户在文档中添加多处批注后，选择开放且定位有效的批注，统一提交给宿主。整体任务要求在宿主现有对话或任务入口填写。

`DocumentInteractionRequest` 的设计字段如下：

| 字段 | 含义 |
| --- | --- |
| `protocolVersion` | 精确交互协议版本，设计值为 `1` |
| `requestId` | 该次请求的关联 ID，由 SDK 分配，不构成授权 |
| `documentId`、`baseRevision` | 请求生成时的文档及修订 |
| `annotations` | 本次选中批注的冻结副本，不随会话继续变化 |
| `context.formatVersion`、`context.kind` | 源文档格式版本及类型 |
| `context.blocks` | 完整子树区块列表，每项包含 `nodeId` 与语义 `html` |
| `context.styles` | 有序的相关样式，每项包含包内 `path` 和 `css` |
| `context.metadata` | 解释区块所需的文档及布局元数据 |
| `context.assets` | 当前请求可用的资源描述，不携带任意文件系统路径或凭据 |
| `supportedOperations` | 本次可返回的内容操作名称列表 |

上下文按相关区块提取，父子区块重叠时合并。不得把一个截断的父节点伪装为完整子树，以免区块替换删除未提供的内容。提案内容目标须位于请求提供的完整子树范围内。

SDK 使用字节、节点和资源数量预算限制上下文输出，超限返回明确结果，不截断 HTML 或静默丢弃选中的批注。宿主负责 token 预算、分段规划和提示词，最终仍返回一份具有共同基准的提案。截图可由第一轮 Word 渲染适配器提供，是宿主可选上下文，不是协议请求的必填内容。

提交回调只表示交接。宿主绑定当前主体、请求和任务，决定通过消息、附件或工具传递请求。传递失败不改变文档和批注；仅生成请求也不推进修订。

## 修改提案与应用

宿主返回以下职责的 `EditProposal`：

```ts
interface EditProposal {
  protocolVersion: 1;
  requestId: string;
  proposalId: string;
  documentId: string;
  baseRevision: number;
  annotationIds: string[];
  summary: string;
  operations: DocumentContentOperation[];
}
```

这里的 `operations` 只接受请求声明的内容操作子集。`annotationIds` 表示提案声明已处理的批注，须为请求批注的无重复子集；未处理批注继续开放。摘要和这些声明供用户判断修改是否满足要求，不能代替差异检查。

`prepareProposal(request, proposal)` 检查协议、请求关联、文档、修订、批注和支持操作，再调用同一事务执行边界。模型生成的 HTML 必须经同一解析器、样式和资源校验，不能直接写入 DOM。

候选包含内容差异、受影响节点及拟更新的批注状态。用户接受有效内容变化时，内容与已处理批注状态在同一事务中提交；批注映射失败的目标仍进入 `needs_reanchor`，不被提案声明掩盖。用户拒绝或请求失败时原批注保持。

第一版提供逐处查看、整批应用和整批撤销。按变更组分别应用留待依赖关系明确后扩展。空提案或无净内容变化不自动解决批注，手工解决批注可走独立的批注事务。

当前修订与请求或提案的 `baseRevision` 不同即返回 `RevisionConflict`。用户可以保留旧提案查看，基于当前内容及仍有效批注重新生成请求，不自动语义合并，也不叠加未提交候选。

## 宿主智能体与保存

达智使用既有消息、附件、Agent、工具、Run、Review 和取消机制；其他产品选择自己的运行方式。SDK 不提供模型适配器、供应商调用或智能体循环。

首轮达智将请求作为结构化附件交给智能体，产出提案 JSON，通过宿主选择并读取相应结果后回传组件。可选的后续文档工具采用“读取文档、读取批注、提出修改”语义，由宿主注册与授权；工具产生提案，不直接提交文档或覆盖源文件。

Agent Run 的提案文件遵循宿主既有产物保存规则；它与文档会话提交、正式文件保存是三个结果。宿主保存失败不能回报正式成功，草稿和候选应保留。异步保存修订 N 期间若编辑到 N+1，只能将 N 标记为已保存，N+1 仍未保存。

## 错误、事件与生命周期

边界错误包含稳定 `code`、可读消息和有界详情。首轮至少区分 `UnsupportedFormat`、`UnsupportedVersion`、`UnsupportedFeature`、`InvalidDocument`、`InvalidOperation`、`InvalidProposal`、`AnchorNotFound`、`RevisionConflict`、`AssetUnavailable` 和 `ResourceLimitExceeded`。宿主提交、导出和存储分别使用 `HostSubmissionFailed`、`ExportFailed` 和 `HostCommitFailed`。

核心事件包含 `annotationChanged`、`documentChanged` 和 `error`，结构携带 `protocolVersion/documentId/revision`；编辑器通过工作台 `onEvent` 提供 `selectionChanged`、`annotationsSubmitted` 和 `proposalReady`。导出与宿主保存分别反馈结果及未保存状态，智能体运行状态不由核心推断或持久化。核心会话生产并订阅核心事件，工作台事件不持久化进文档包。

每个会话隔离文档、请求及候选引用。组件创建的会话随卸载释放，外部传入会话由创建者管理。外部智能体任务的取消和晚到结果归宿主处理，回传时仍检查会话、请求和修订。

## 示例与验收

[报告文档包](../examples/report/)包含两条批注；[请求](../examples/report.request.json)冻结这两条批注及上下文；[提案](../examples/report.proposal.json)关联同一请求并修改正文和标题样式。

验收覆盖完整成功、失败不变、无变化、请求错配、越界目标、版本冲突、批注映射、撤销重做和资源生命周期。详细任务与演示见[第一轮实施细则](iteration-1.md)。
