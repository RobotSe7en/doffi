# 第三轮工作簿读取交接

R3-01～08 已完成并通过 [PR #25](https://github.com/RobotSe7en/doffi/pull/25) 合并，提交 `c1d8dea`。三位子智能体审查、8 项修复复核、最终 HEAD `d4f250e` 的四项 CI 全部成功；该合并提交的树与其最终 HEAD 相同。整体 CI 优化 [PR #26](https://github.com/RobotSe7en/doffi/pull/26) 已合并为 `31cca2f`，最终 21 个作业及容量验收通过；触发频率调整已合并为 `d65c0e9`。2026-11-07 从 main `5451736` 转入 `feat/workbook-editing`，执行 R3-09～16 编辑计算；之后执行 R3-17～20 输出验收，任务编号不变。

用户已选择容量用例移出日常 CI，保留阶段合并前完整验收和一万行目标，具体规则见[CI 执行](ci-performance.md)。

## 授权与约束

用户已授权阶段 PR 的提交、推送、独立复审、修复复核和最终提交全部 CI 成功后的合并；无需再次询问合并许可。不直接提交 main，不发布 npm。2026-11-06 的 PR #25 三位子智能体审查已完成；2026-11-07 用户明确要求 R3-09～16 阶段在合并前使用两位子智能体独立审查，修复并复核通过后由实施者合并。实施不分派给子智能体。产品语义出现新分歧时主动询问；已定选择继续有效，不重复询问。

基础七项选择见 [AGENTS.md](../AGENTS.md)。额外七项为：财务下划线铺满单元格；隐藏列内容按未隐藏宽度参与 AutoFit；冲突经典备注框以 VML 为准且保留两份来源及差异；旧 lineChart/marker 只保留不控制显示；坐标轴按 Office 的跨越轴语义；渐变随形状刚性旋转并保留原覆盖；负柱反转保留符号方向，填白色并保留原边框。

用户只将读取滚轮反馈 p95 调整为 120ms，明确要求保留并完成滚动优化。选择仍为 100ms；首屏 30 秒、真实 20 万单元格／2 万公式／40 块、四逻辑 CPU／8GiB／无交换、DOM 及内存限额和取消检查均不变。历史失败另存，不能删除慢样本或用不受限耗时替代。

Playground 服务监听 `0.0.0.0:5174`，地址 `http://192.168.31.221:5174/`。5173 属于另一个项目，勿操作；保留 5174 服务。

## 已实现的读取范围

- 公开 `importXlsx`、统一 `openDocument`、严格工作簿包、保存重开、专用 Worker 与有界读取通道。
- 单元格类型、富文本及精确 UTF-16；日期／地区及常用数字格式；高级填充、边框、旋转、竖排、命名样式关联。
- 隐藏行列、冻结合并、RTL、跨空白格溢出、跨格居中、实际来源自动行高及隐式字体来源。
- 智能表、基础条件格式及七类验证；当前源求值和明确不可用诊断；打印元数据保留。
- 三类基本图表：当前引用或唯一原生常量，轴、标记、标题、实际字体／数字字段、标签、图例／引导线和极端值裁剪。
- 图片、基础形状／文本框、三种锚点、裁剪／翻转／旋转、渐变覆盖及源投影。
- 经典备注和现代线程、可见备注框及面板、VML 优先、异常定位 needs_reanchor、提及及 uid 保留。
- 公共只读预览、Playground 的打开／取消／切换、资源与字体诊断和释放。

实现细节见[工作簿模型](workbook-model.md)、[公式基础](workbook-formulas.md)及[读取验收](workbook-reading-acceptance.md)。读取验收中的标量内核不等于 R3-10 完整计算；当前事务和计算基础的实施、审查及剩余工作见[编辑计算交接](workbook-editing-handoff.md)，编辑阶段和原生输出仍未验收。

## 滚动优化及独立审查

`bf4d294` 保留有界隐藏预布局候选和 AutoFit 当前／下一批重叠读取；正式窗口仍重新读取并核对身份、修订、值、HTML、格式及字体版本，候选与当前窗口共用 DOM 预算。

用户授权后，`pr25_review_core`、`pr25_review_office`、`pr25_review_render` 分别从 main `2fab5a6` 审查整个阶段差异。共发现 8 项 P2，修复由主实施者完成，原审查代理逐项只读复核关闭。完整记录见[独立审查](workbook-reading-review.md)。

- `04089c3`：累计依赖文本预算、名称作用域、自动行高压缩、局部 XMLNS、分页符缺省、取消清理、viewport 字体 resize 及 Firefox 宿主串滚修复。
- `fd5777a`：单元格与规则读取统一使用逐次可释放的取消等待，避免悬挂来源和共享未结束 Promise 留存旧批次。此提交为最终 SDK 运行时代码；后续交付说明及 Playground 只读提示不改变该运行时。

最终格式检查另清理元数据文件末尾空行及生成表样式许可证注释的尾空白，生成器同步去掉尾空白；60 份规范表样式生成检查通过。注释会进入包摘要，因此 128 矩阵与四引擎资源已在清理后的包上完整重跑、重新汇总。SDK 执行逻辑不变。

网格滚轮最终采用非被动监听并 `preventDefault`，先排除 CTRL 和备注面板；仅 CSS containment／被动监听无法阻止 Firefox 宿主串滚。隐藏预布局和 AutoFit 预读仍保留。不得恢复已撤回的性能实验或重新试探已决门槛。

## 最终本地验证

最后 SDK 代码上的检查均通过：

- `pnpm check`：144 个文件、1745 项测试，类型、边界、所有包及 Playground 构建通过。最终只读提示另做 Playground 构建通过。
- [公开读取矩阵](evidence/workbook-reading-matrix.json)：现代三引擎 96 项，旧 Chromium 32 项；每引擎均含 16 独立源及 16 未修改的实际 Calc 文件，合计 128 份。旧 Chromium 同批追加 7 项回归，共 39 项。
- [固定资源四引擎](workbook-reading-budget.md)：首屏为 14.738／27.773／19.242／21.030 秒；滚轮 p95 为 57.3／27／58／101.9ms。选择、真实计数、DOM、浏览器／Worker 绝对高水位合计 2GiB 和取消均通过。统一使用 cpuset `0,16,32,48`，只声称四个逻辑 CPU。
- 修复前 Office 5 项、core 3 项复现失败；修复后相关单元 37 项通过。现代三引擎字号／滚动／取消 15 项、追加逐次等待／规则取消 9 项通过。
- 两个汇总器已重新构建公共宿主／Worker 并核对当前包摘要，完整记录分别写入矩阵及资源 JSON。旧提交的通过不能替代这些最终记录。

本机执行日志为 `.cache/pr25-reviewed-full-check.log`、`.cache/pr25-final-matrix-{modern,legacy}.log`、`.cache/pr25-final-budget-<browser>.log` 和对应 `*-collect.log`；缓存不是交付来源，正式证据位于 `docs/evidence/`。

主矩阵 32 份来源通过官方 Transitional XSD／关系审计。额外标签探针的 4 份 Calc 图表存在 numFmt／dLbl 子元素顺序错误：保留原文件和精确错误，仅在验证副本检查其余结构，不能称原文件合规。Calc 的线程扁平化、图片／渐变开关丢失及标签／负柱改写分别记录，不从原源偷偷修复。Excel/WPS 未实机验证。

## 交付及下一阶段

每个阶段核对 PR 最新 HEAD 的构建、`check`、`browser-interaction`、`office-reopen`、`safari` 和 `workbook-scale` 均成功，再以该 SHA 守卫 squash merge；不凭旧提交 CI 合并。合并后从最新 main 创建阶段分支，继续 R3-09～16 编辑计算，再做 R3-17～20 输出验收。不直接向 main 提交，也不把当前读取通过记为后续阶段完成。

后续第一项是 R3-09：复用共享会话事务、隔离分块候选、原子验证、历史和精确范围映射；随后完成 R3-10 依赖计算，再开放输入与各类编辑。排序区域外公式仍引用原地址；剪切引用精确跟随移动数据。保持唯一 HTML 包来源、当前源图表／规则和共同历史边界。

## 验证入口

```sh
pnpm build:packages
pnpm exec tsc --project tsconfig.tools.json --noEmit
node scripts/collect-workbook-reading-matrix.mjs --record
node scripts/collect-workbook-reading-budget.mjs --record
pnpm check
```

矩阵运行方法见[读取验收](workbook-reading-acceptance.md)，资源命令见[读取资源检查](workbook-reading-budget.md)。浏览器依次独占运行，不与构建争用，也不在运行期间改写 dist。代码变化后按影响重新验证；普通功能和实际性能不可互相替代。

固定 Office 为官方 LibreOffice 25.8.1.1，归档 SHA-256 `f367c434aee8352e9742e192bedcdaef194ffcbcfeaf760c6300fb4fecbba73c`。本机 `.cache/xlsx-venv/bin/python` 提供 lxml；默认 python3 未装 lxml。正式来源与证据不得被诊断脚本重写。
