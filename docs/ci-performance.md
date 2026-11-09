# CI 执行与阶段验收

本轮优化保留一万行支持目标。用户已选择：真实容量用例移出日常 CI，阶段合并前单独全量验收。其余功能、安全、失败、取消、像素及外部 Office 检查保持；约 5 分钟为日常 CI 关键路径优化目标，实际结果须用 GitHub runner 验证。

[PR #26 最终工作流](https://github.com/RobotSe7en/doffi/actions/runs/42246509253)的 21 个作业已全部成功：从首个构建作业开始到日常汇总完成为 320 秒（5 分 20 秒），容量专项单独通过。该轮包含已授权的容量分流，不能将其当作原全量范围的同负载耗时。

## 原始记录和诊断

[PR #25 最终 CI](https://github.com/RobotSe7en/doffi/actions/runs/42223783754) 的浏览器作业共 2732 秒：Word/PPT 707 秒，现代工作簿 1299 秒，旧 Word/PPT 179 秒，旧工作簿 431 秒。工程检查 190 秒、Office 352 秒、Safari 112 秒。逐步骤记录见[基线](evidence/ci-performance-baseline.json)。

[本机 A/B](evidence/browser-test-overhead.json)在同一浏览器、固定 viewport、独立 context 内各执行相同六次网格渲染，每组四次并轮换顺序。完整首页与完整 trace 平均 1.94 秒；空白宿主与 API trace 0.99 秒。840KB 输入的逐字节数组传输约 3.58 秒，Base64 约 0.041 秒。后者只测测试协议传输；这些诊断不替代 CI 前后对比，也不证明 SDK 导入耗时变化。

复现时先独占启动生产预览端口 4173，再运行：

```sh
PLAYWRIGHT_BROWSERS_PATH="$PWD/.playwright" node scripts/probe-browser-test-overhead.mjs
```

## 日常模式

普通提交／推送和 PR 创建只自动执行 `Quick checks`：差异空白检查、包边界和固定摘要的 actionlint 工作流校验。它不构建 SDK，不启动浏览器或 Office，也不取消正在运行的完整验收。

actionlint 1.7.7 仅用于 Linux CI，采用 MIT 许可；官方 Linux amd64 归档 SHA-256 为 `023070a287cd8cccd71515fedc843f1985bf96c436b7effaecce67290e7e0757`。缓存命中后仍核验归档，再解包工具与许可，不进入 SDK 发布包。

完整 `CI` 仅在显式添加 `run-ci`／`stage-acceptance` 标签或手动触发时运行。草稿转为待审也不自动运行，以免随后请求阶段容量时重复启动并取消完整任务。普通推送、无关标签和合并到 main 不重复执行整套。已存在的触发标签需要移除后重新添加才能再次请求；修改标签以外的推送不会自动重跑。验收后若修改了阶段 PR，须主动对最终 HEAD 再次发起验收，旧提交绿灯不能代替。

```sh
gh pr edit <PR> --add-label run-ci
gh workflow run ci.yml --ref <branch>
```

- SDK 与 Playground 只构建一次，同一 workflow 的后续 job 下载该次不可变产物；类型、边界及测试照常检查。本地 `pnpm check` 也只构建一次。
- 浏览器按四个引擎和 Word／PPT／工作簿分为独立 job，每个仍 `workers:1`。文件集合保留原矩阵；固定资源验收继续独占顺序执行。
- 每片安装实际 UI 引擎；Word／PPT 还安装受管理 PDF/PNG 导出所需的 Chromium，缓存键包含完整工具集合。WebKit／Firefox 工作簿片不安装无关引擎。
- SDK 探针使用同源空白宿主，实际 Playground、远程 HTTP 和工作台流程保留原首页。每例仍新建独立 context/page，不复用模型或 Worker。
- 成功用例不采集完整 trace。首次失败保留错误与截图，随后的一次诊断重跑记录完整 trace；`failOnFlakyTests:true` 使首次失败仍导致 CI 失败。固定资源性能测试明确禁止重试及截图，保留首次失败样本。
- 测试包使用 Base64 跨协议传输，SDK 仍接收独立字节数组，原始来源和容量断言不变。
- Office 按 Word／PPT／工作簿拆分，官方归档缓存按固定 SHA-256 复用；安装器仍核验摘要、重新解包并核对版本。独立打包与实际 Office 编辑／重开保留。
- 每个浏览器分片输出逐用例 JSON 和执行摘要，Office 输出各命令耗时。汇总 job 只有所有分片成功才成功。

`DOFFI_CI_MODE=daily` 只用于日常 CI。六项原生容量单元检查延后；浏览器 `@scale` 用例通过显式过滤延后，原断言保留。工作簿实际打开、切换、失败保留和授权资源仍有小样例覆盖。原生样例的字节、规范及负例审计继续执行；日常 Calc 重开省略一万行项。默认本地检查和参考记录命令保持完整模式，`--record` 永远不产生部分参考基线。

编辑开发分支新增默认 `mountWorkbookEditor` 的 `@scale` 容量场景：100 次原生输入事件／公共确认入口，每次核对同一编辑器当前显示的依赖公式，再通过实际界面按钮执行 100 次撤销／重做／再撤销，并核对原始块和唯一 Worker 会话。此场景的额外 bundle 只在选中该场景时构建；日常检查继续排除它。`check:workbook-performance` 已为编辑增加独立的四引擎固定资源入口，包含混合响应的图片／形状／文本框、真实选择和按键反馈、100 次编辑及完整历史、两类万行取消与新包完整重开；严格汇总当前主线程／Worker 脚本和原生源摘要。带采样器或缩短样本的诊断不能计入验收。`4ee0710` 的本机四引擎编辑已通过，见[记录](evidence/workbook-editing-budget.json)；后续公式及审查修复的最终源仍须重测，读取成功不能代替编辑门槛。

原生公式语料拆为每日小样例与阶段容量两部分；共享公式跟随者、打印名称及图表／规则／表格公式的独立来源检查继续每天执行。

## 阶段合并前

阶段 PR 显式加上 `stage-acceptance` 标签，或手动运行 CI 并开启 `stage_acceptance`。最终 HEAD 必须同时通过全部日常检查和 `workbook-scale`，各分片必须实际成功。标签保留期间的新提交不自动重跑；集中修复完成后重新发起一次，旧 HEAD 的结果不能代替。

专项运行完整原生 XLSX 参考、六项原生及两项编辑补充容量单元检查、四引擎 128 份公开来源／生产者矩阵及其余实际大数据集成用例，然后依次执行四逻辑 CPU／8GiB／无交换资源验收。20 万实际单元格、2 万公式、40 块、首屏、交互、DOM、内存和取消门槛保持。[阶段脚本](../scripts/workbook-scale-acceptance.mjs)核对实际执行、当前包摘要及全部资源条件后才写成功记录。

```sh
node scripts/install-acceptance-fonts.mjs --system-fonts
FONTCONFIG_FILE="$PWD/.cache/acceptance-fonts/fonts.conf" pnpm check:workbook-scale
```

完整 Firefox 工作簿日常集在 `b87098d` 上实际运行约 16.4 分钟（348 通过、3 跳过、6 个夹具失败，均已修复并定向复核），已超过原先整个 job 的 15 分钟；工作簿浏览器 job 的总时限设为 30 分钟，Word／PPT 保持 15 分钟。阶段容量现在串行包含四引擎读取与编辑、全部大表集成及补充用例，总时限设为 60 分钟。它们仅是 CI 调度时限，不改变首屏、计算、编辑／交互 p95、CPU／内存、DOM 或取消的任何 SDK 性能门槛。

前提为 Linux/systemd/cgroup v2、至少四个可用逻辑 CPU、已固定的四个浏览器、字体及 LibreOffice 环境。脚本需要无交互 sudo 创建临时受限服务；不更改持久系统服务。

正式结果位于 `.cache/workbook-scale-acceptance/` 并作为 CI artifact 上传。汇总器的 `--output <file>` 写入显式结果路径；原 `--record` 入口仍用于明确更新已提交证据，普通 CI 不改写参考基线。

## 固定字体的实际浏览器前置检查

`97af0b5` 的远端容量首次运行有 115 项现代集成通过，两项 Firefox 大表编辑各首次／诊断重跑均超过 360 秒，未进入旧引擎及固定资源步骤。固定字体文件位于应用 `.cache` 时，`fc-match` 虽指向正确摘要，Firefox Canvas 与公开矩阵却报告 Noto CJK 替代；完整自动行高因此不能复用，每次编辑重新测量 20 万格。本机用同一 FONTCONFIG 重现，三次诊断的预览各约 15 秒，原日志与 profile 保留，未将诊断算验收。

显式 `--system-fonts` 将相同摘要字节安装到[标准系统字体目录](https://help.gnome.org/system-admin-guide/fonts.html)下由 manifest 摘要命名的独立目录。已有不同字节拒绝覆盖；来源、版本、许可和摘要保持原固定记录。默认无该开关仍使用应用缓存。只有 CI Office／容量 setup 显式调用，SDK 运行时不安装或下载字体，也不更改浏览器安全偏好或允许替代字体复用。

[字体前置脚本](../scripts/check-acceptance-fonts.mjs)在长容量流程前核验仓库规范 manifest、五份实际文件摘要和当前 Fontconfig 选择，再在默认 Firefox context 中测量四个 CJK family／weight，保留不存在字体的负例。正负测量及来源写入阶段 artifact 的 `font-environment.json`。这避免 shell 字体可见却浏览器不可见时继续运行长测试；其成功仍不能代替完整日常／容量。
