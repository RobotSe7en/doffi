# 工作簿输出验收

R3-20 验证原生 XLSX、工作簿打印及 Playground/CLI 在独立安装、外部修改和当前修订重开时的行为。阶段状态以同一最终提交的审查、日常 CI 与容量产物为准。读取和编辑范围分别见[读取验收](workbook-reading-acceptance.md)与[编辑验收](workbook-editing-acceptance.md)。

| 能力 | 入口和证据 | 范围 |
| --- | --- | --- |
| 独立 SDK 消费 | `pnpm check:workbook-closure`，Office 的 `workbook` 作业 | 在仓库外安装五个 tarball。普通 JavaScript 无 React 依赖，React 使用可选入口。两者从独立 XLSX 预览、输入、提交批注提案并审阅、应用、撤销重做，保存当前修订的 XLSX 和文档包。Node 通过公开入口重开两份输出，校验公式、值、备注及修订，生成 PDF/PNG。安装的 `doffi` 命令通过 V1 字节协议处理 XLSX 和文档包。 |
| 原生形式与编辑 | `pnpm check:xlsx-export`，Office 作业和 `workbook-scale` | 日常 30 份独立源／生产者输出，完整模式 32 份。SDK 重开，分别核对官方 Transitional XSD、OPC 与选定的现代批注结构。固定 LibreOffice 25.8.1.1 实际修改值、公式、表范围、图表标题和经典备注，重存后由 SDK 逐项重开核对。 |
| 当前修订的打印与下载 | `workbook-print.spec.ts` 和 `workbook-playground.spec.ts` | 固定字体的两页 A4 横向 PDF/PNG 对照，核对重复表头、页眉页脚与全部数据行。另核对公式缓存与重算、选区、资源限制、当前源、取消、切换、重开和 Worker 释放。 |
| 容量与跨格式回归 | `pnpm check:workbook-scale` 和完整日常 CI | 完整原生输出含两个一万行样例，各 20 万实际单元格、2 万公式。现有 128 项四引擎矩阵、固定资源交互和 Word/PPT、Safari 回归均为阶段门槛。 |

SDK 原生导出保留受支持的名称、智能表、规则、图表、绘图、备注与线程。不能生成的公式缓存、批注锚点降级和不能原生表达的内容以导出报告为准。固定 Calc 在原始语料中的现代线程丢失已有单独记录。打印对照允许已记录的分页差异，比较纸面、页级文字和页眉页脚墨迹边界。[桌面 Excel 抽样](native-office-macos-acceptance.md)通过公式编辑重存，现代线程仍有差异。WPS 尚未实机验证。 [外部参考及限制](../examples/xlsx/README.md#外部参考及限制)

原生筛选保留标准 `row@hidden`。仅被筛选隐藏的行另带一个 SpreadsheetML 行扩展，使 SDK 重开和清除筛选时与手动隐藏区分。外部应用重存可能移除此扩展。凹槽等无法映射的 CSS 边框、独立文字脚本字体及有小数的基础列宽等会在原生报告中说明转换。文档包保存源样式。打印使用图表实际引用工作表的筛选，并从纸面和自动打印范围排除不打印或隐藏的绘图。

本地复现需固定 Chromium、字体、LibreOffice 和 Poppler。Pillow 用于比较 PNG 的页眉页脚。 [CI 环境](ci-performance.md)

```sh
pnpm check
pnpm check:workbook-closure
DOFFI_CI_MODE=full pnpm check:xlsx-export
FONTCONFIG_FILE="$PWD/.cache/acceptance-fonts/fonts.conf" node scripts/playwright.mjs test workbook-print.spec.ts --project=chromium
```

独立打包报告位于 `.cache/workbook-closure-acceptance/`，含源文件和 tarball 摘要、两个宿主的实际输出及原生和打印报告。`.cache/ci-office/workbook/xlsx-export/` 保存运行模式、用例及结构验证结果。打印对照、PDF 和两页 PNG 随浏览器作业上传。正式容量结果位于 `.cache/workbook-scale-acceptance/`。各项 CI 产物按最终源码 SHA 对照。

## 独立审查

两位审查者分别检查了 R3-17～20 的完整差异。复核通过的案例包括筛选与手动隐藏共存、当前字体与样式降级报告、差异数字格式、无效原生格式拒绝、跨表图表筛选，以及不打印对象和自动打印范围。回归覆盖 SDK 重开、官方结构校验和实测 PDF/PNG。

旧格式回归还核对了 Word 兼容导入前后的实际 Writer 重开、OPC 默认命名空间、段落标记以及当前样式的原生编译。Word 格式修复覆盖共享预览和编辑画布。独立来源和历史参考文件保持原样。
