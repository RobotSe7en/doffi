# 工作簿编辑计算阶段验收

本页将 R3-09～16 的已声明范围对应到实现和可复现检查，供完整阶段审查使用。实现核对完成；阶段尚未验收，PR #31 仍为草稿。最终双人独立审查、问题修复复核、最终 HEAD 日常 CI 和 `workbook-scale` 均通过后才能合并。它不包括 R3-17～20 的原生 XLSX、PDF／PNG、Node 输出和独立打包闭环。

唯一延期项是用户于 2026-11-09 明确暂缓的“删除合并区域首行／首列”的 Excel 对照。当前缩小行为及 Calc 差异保留，Excel 兼容性未验证；其他门槛保持。名称求值另有[实际 Excel 网页版观测](evidence/excel-web-name-semantics/README.md)，不得将这份观测推广为 Excel 结构编辑或 WPS 验收。

## 要求与覆盖

以下列出覆盖入口，完整测试由同一日常／阶段任务执行；单独存在 API 或一个局部通过记录不代表整行完成。

| 编号 | 实现范围 | 独立来源及组合检查入口 |
| --- | --- | --- |
| R3-09 | 同一 prepare／commit、不可变块、资源授权、整批校验、精确结构映射、无净变化、撤销、冲突、异步取消和预算 | [会话](../packages/core/tests/workbook-session.test.ts)、[异步准备](../packages/core/tests/workbook-preparation.test.ts)、[混合原生候选](../tests/workbook-editing-combinations-native.test.ts)、[读取快照](../tests/workbook-session-read-snapshot.test.ts)；失败、取消和过期候选不写源或历史 |
| R3-10 | 第一方参数转换／函数求值、两种日期、名称／结构引用、依赖图、循环、增量计算、宿主时间／地区、来源缓存分离和取消 | [公式矩阵](evidence/workbook-calculation-functions.json)、[基础类型矩阵](evidence/xlsx-scalar-semantics.json)、[268 项边界](evidence/workbook-function-boundaries.json)、[筛选当前小计](evidence/workbook-filter-calculation.json)、[Excel 名称](evidence/excel-web-name-semantics/README.md)、[增量／长链](../packages/core/tests/workbook-calculation.test.ts)、[补充容量](../tests/workbook-editing-scale.test.ts) |
| R3-11 | 键盘／IME／公式栏、类型与富文本、TSV／系统 HTML、复制／延迟剪切、选择性粘贴、逻辑填充、查找替换、外观／命名样式 | [系统剪贴板](evidence/workbook-system-clipboard.json)、[HTML 来源](evidence/workbook-html-clipboard.json)、[合并粘贴与填充](evidence/workbook-merged-special.json)、[结构引用填充](evidence/workbook-structured-fill.json)、[名称复制](../tests/workbook-copy-names-native.test.ts)、[原生格式](../tests/workbook-formatting-native.test.ts)、[直连输入](../tests/browser/workbook-input.spec.ts)、[唯一 Worker](../tests/browser/workbook-read-worker.spec.ts) |
| R3-12 | 工作表／行列、维度／隐藏、冻结、合并、剪切、稳定多键排序、普通及表内筛选、各来源引用维护 | [行列](evidence/workbook-axis.json)、[名称](evidence/workbook-axis-names.json)、[智能表](evidence/workbook-axis-tables.json)、[多行合并排序](evidence/workbook-merged-sort.json)、[普通筛选表头删除](evidence/workbook-filter-headers.json)、[局部表剪切](evidence/workbook-partial-table-move.json)、[规则搬移](evidence/workbook-move-definitions.json)；外围排序公式保留地址，剪切精确拆分引用 |
| R3-13 | 表创建／扩展、列身份与配方、汇总、表样式、名称、条件格式及七类验证；所有修改入口共用最终候选验证 | [定义](../packages/core/tests/workbook-definitions.test.ts)、[验证](../packages/core/tests/workbook-validation.test.ts)、[表样式](evidence/workbook-table-style-native.json)、[规则复制](evidence/workbook-copy-rules.json)、[整表复制](evidence/workbook-copy-tables.json)、[混合来源](evidence/workbook-editing-combinations.json)、[人工规则与警告](../tests/browser/workbook-workbench.spec.ts)；停止规则的未知结果拒绝，警告要求明确确认 |
| R3-14 | 三类图表、授权位图、基础形状／文本框、对象文字／外观、裁剪／几何／层次、锚定及测量补偿 | [对象编辑](../tests/workbook-drawing-editing-native.test.ts)、[文字](../tests/workbook-object-text-native.test.ts)、[绘图搬移](evidence/workbook-drawing-transport.json)、[范围复制](evidence/workbook-copy-objects.json)、[图表原点](../tests/workbook-chart-origins.test.ts)、[实际浏览器测量](../tests/browser/workbook-drawing-editing.spec.ts)；固定及混合响应需要来源绑定的布局证明，资源不因节点 ID 获权 |
| R3-15 | 经典备注、现代线程／回复／作者时间／解决状态、多片段映射、排序／剪切跟随、失效／重定位及历史恢复 | [线程原生编辑](../tests/workbook-comment-editing-native.test.ts)、[跨表映射](../tests/workbook-comment-mapping-native.test.ts)、[多片段与恢复](../packages/core/tests/workbook-comment-mapping.test.ts)、[范围复制](evidence/workbook-copy-comments.json)、[两份备注框](evidence/workbook-note-transport.json)、[人工交互](../tests/browser/workbook-comment-editing.spec.ts)；不一致来源保留 needs_reanchor，正文不丢弃 |
| R3-16 | 区域／表／全局请求、冻结当前输入与只读依赖、实际候选差异、显式绘图权限、原子应用、迟到／越界／取消／冲突和宿主所有权 | [上下文](../packages/core/tests/workbook-interaction-context.test.ts)、[范围](../packages/core/tests/workbook-interaction-scope.test.ts)、[提案](../packages/core/tests/workbook-proposals.test.ts)、[完整混合提案](../tests/workbook-editing-combinations-native.test.ts)、[JS／React／唯一 Worker 工作台](../tests/browser/workbook-workbench.spec.ts)；日常浏览器任务明确包含该文件 |

## 跨功能组合

[独立混合 XLSX](../packages/office/tests/fixtures/workbook-editing-combinations.xlsx)使用 Python 标准 ZIP/XML 构造，包含富文本与两个字体角色、原生智能表／计算列／汇总、命名范围、条件格式／停止验证、图片／形状／文本框／图表、单点／双点／绝对锚定、经典备注及有兼容占位的现代线程。源公式缓存故意写成 999，当前预期由独立输入数据计算：初始金额 240、数量 10；B2 改 6 后金额 280、数量 14；按数量大于 2 且单价小于 30 筛选后的可见金额为 120。

同一用户命令修改输入、富文本外观、表筛选、名称、条件规则、形状文字、图表标题和线程回复。检查候选隔离、完整包重开、一个历史步骤及全部源族的恢复；附加停止验证、未授权资源、实际测量取消和冲突检查。整表跨表复制／剪切同时保留富文本、两类批注、单元格所属绘图及公式绑定；全局提案冻结这套来源并显示外部消费者重算。直连与唯一 Worker 在 Chromium／Firefox／WebKit 中用实际第一方字体／布局服务执行混合命令与提案。

核心组合测试显式使用点度量，仅证明核心测量协议与事务组成；浏览器测试使用实际服务。释放编辑器后的独立宿主写入探针在表外空格进行，合成度量只证明宿主所有权，不计字体／原生几何验收。

`check:workbook-editing-combinations` 固定官方 LibreOffice 25.8.1.1 归档摘要，实际修改 B2、重算并保存。该生产者为三份经典兼容记录生成重复 VML 类型／形状标识；[原字节](../packages/office/tests/fixtures/workbook-editing-combinations-calc.xlsx)和[拒绝原因](evidence/workbook-editing-combinations.json)保持。官方 XSD／OPC 不验证完整 VML，不能把其成功当作合法 VML 或无损现代线程重开。原生正例及其他实际编辑复现仍由上表各独立语料覆盖。

## 容量与阶段状态

[固定资源记录](evidence/workbook-editing-budget.json)覆盖 `4ee0710` 的四引擎：一万行／20 万实际格／2 万公式及混合响应绘图，4 CPU、8 GiB、无交换；全部 100 次编辑计入 p95、100 步撤销／重做、真实选择／输入反馈、两类万行取消、克隆包完整重开与释放后宿主写入。各引擎 p95 195.6／494／348／326.9ms，浏览器保守高水位约 1380／1798／1986／1303MiB，均在原门槛内。该记录不是随后公式标记修复后的最终 HEAD 证明。

补充容量用例另验证 1000 个扇出依赖精确重算且复用无关结果，以及一万实际数据行智能表扩展、9999 个真实计算列公式、人工例外、汇总、40 块重开和一次撤销／重做。它们标为 `@scale`，日常 CI 不执行；阶段 `workbook-scale` 必须实际执行。跨表查找／聚合、1200 级依赖链和稀疏远端分别在上述语料与核心检查中覆盖。

最终合并仍须以下证据对应同一最终源：恰好两位完整差异独立审查及问题复核、日常 CI、完整原生参考、容量单元／集成、128 项公开矩阵、固定资源四引擎读取与编辑。历史成功、短样本诊断、基础审查均不能代替。具体运行与修复记录见[交接](workbook-editing-handoff.md)、[审查记录](workbook-editing-review.md)及[CI 执行](ci-performance.md)。
