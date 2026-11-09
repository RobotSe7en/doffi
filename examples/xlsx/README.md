# 独立 XLSX 样例与参考

R3-04 使用 [Python ZIP/XML 生成器](../../scripts/generate-xlsx-fixtures.py)创建 16 个正例、17 个负例。样例不依赖 SDK 导出器、第三方电子表格库或计算引擎；通用 DrawingML 样例代码复用本项目的独立 PPTX 生成器。[expected.json](expected.json)保存源摘要、独立单元格/公式数量、选定缓存值、结构和文本预期。这里完成的是测试来源准备，**不代表 SDK 已实现 XLSX 导入或通过 R3-08 读取门槛**。

| 文件 | 主要覆盖 |
| --- | --- |
| 01-values-text | 数字、布尔、空字符串、格式化空白、错误、共享/内联富文本、中文、UTF-16 变体、超长数字文本、外部超链接 |
| 02-styles-dates | 主题、字体、边框、填充、对齐、七种自定义数字格式、1900 序列 0/1/59/60/61 |
| 03-structure | 行列大小与隐藏、冻结、普通合并、跨 256 行分块合并 |
| 04-formulas-names | 共享公式、混合/跨表引用、工作簿/工作表名称、整行整列聚合、字符串/错误处理缓存 |
| 05-native-table | 原生智能表、计算列、当前行结构化引用、汇总行、SUBTOTAL、表样式 |
| 06-rules-validation | 比较/公式条件格式、双色/三色色阶、七种验证、停止/警告提示、来源违规值 |
| 07-three-charts | 二维簇状柱状图、普通折线图、单系列饼图、单元格引用、原生缓存、标题/图例/轴 |
| 08-images-shapes | 内嵌 PNG、裁剪旋转、基础形状/文本框、三种绘图锚定 |
| 09-classic-notes | 经典备注、作者、富文本、VML 定位 |
| 10-modern-threads | 两个现代线程、一条回复、作者时间、解决状态、不同单元格 |
| 11-print-layout | 打印区域、横向 A4、页边距、重复标题行列、手动分页、页眉页脚 |
| 12-multi-sheet-filter | 多表引用、命名范围、隐藏表、普通筛选和隐藏数据行 |
| 13-date-1904 | 1904 日期系统边界 |
| 14-preserved-formulas | 未列入本轮函数矩阵的 FACT、已有/缺失缓存、依赖公式、手动模式 |
| 15-ten-thousand-rows | 1 万行 × 20 列，20 万实际单元格、2 万公式；数据与分块技术基线一致 |
| 16-sparse-far-cell | XFD1048576 远端坐标、三格实际内容、远端引用 |

负例包含缺失部件、重复/越界坐标、非法字符串/样式索引、DTD、数组/外部公式、孤立共享公式、重叠合并、保护、未定义作者/线程环、未支持条件格式/图表、外部图片、截断 ZIP。负例的错误码是后续导入器测试预期；当前完整性检查验证变异实际存在，不冒充已验证 SDK 拒绝行为。资源超限由后续导入测试的显式小额度触发，不在仓库放置巨型恶意包。

## 原生结构检查

[独立检查器](../../scripts/check-xlsx-fixtures.py)重建并逐字节比较源文件，检查内容预期、OPC 关系、样式/共享字符串引用、坐标及作者/线程关系。[原生输出检查器](../../scripts/check-xlsx-output.py)使用已固定 SHA-256 的官方 Transitional XSD；原始电子附件不修改。

官方电子附件将 `sml:t`/`a:t` 定义为简单类型，而 Office 文本使用 `xml:space` 保留空格。检查器单独验证该属性的位置和值，仅在校验投影中移除该属性，文字内容和源文件不变。该处理与 [Microsoft 共享字符串说明](https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/working-with-the-shared-string-table)中的用法一致。现代线程按 [MS-XLSX 结构](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/adb84732-9fc8-48b6-bddc-6b0bcdaad940)检查当前选定字段及引用；VML 只检查 XML 和关系，不宣称通过完整 VML XSD。

[源检查报告](source-schemas.json)包含 75 个标准部件和 2 个现代线程相关部件；[Calc 重存报告](producer-schemas.json)包含 91 个标准部件。两组各 16 个工作簿。XSD 正确不等于 SDK 已能完整表达内容，也不替代读取与编辑验收。

## 外部参考及限制

[reference.json](reference.json)记录官方 LibreOffice 25.8.1.1 归档 SHA-256、实际构建、Poppler/Fontconfig、424 个已安装字体文件摘要、源/重存摘要、20 页 PDF/PNG 摘要及逐文件结构差异。主字体为已安装的 Liberation Sans 和 Noto Sans CJK SC，未下载字体。当前宿主为运行固定归档额外隔离提取了 Ubuntu noble 的 libxml2/libicu74；版本和包摘要也在记录内，没有安装到宿主系统。

全部 16 个源经 Calc 打开、重存及再打开。打印使用样例的有界打印区域；一万行样例只打印前 20 行，完整工作表仍保留 20 万单元格、2 万公式。此处不是 4 vCPU/8 GiB 的最终交互/计算性能验收。

已观察到的外部差异：

- `01`：空字符串 A4 被省略；两个布尔值和两个错误常量被改写成公式，缓存值保留。
- `03`：Calc 为合并区域补充格式化空格，实际单元格从 6 增加到 10。
- `05`：Calc 保留表和单元格公式，但删除计算列定义、汇总标签定义及表样式关联，并将表筛选范围延至汇总行。读取分别保留两份文件实际存在的定义，不能从单元格缓存猜回已丢失的关联。
- `10`：Calc 丢弃两个现代线程、一条回复及解决状态。原始源与独立预期仍保留完整线程。
- `11`：Calc 重复写入同作用域的 `_xlnm.Print_Titles`，两条表达式的行列顺序相反，但都指向第 1 行和 A 列。原生读取只对这种完全等价的重复打印标题作合并；冲突标题及重复普通名称仍拒绝。打印区域的顺序不作重排。

这些变化不能记为无损成功。PDF 不验证批注 UI、未打印区域或隐藏表的交互语义；字形替代与变体显示需要 SDK 自身的字体诊断。Excel/WPS 尚未实机验证。

## 命令

需要 Python 3；XSD 检查另需 `lxml`。外部参考需要 `scripts/install-acceptance-office.mjs` 安装的固定归档，以及已安装的 Poppler、Fontconfig 和上述字体。当前机器使用的旧 ABI 兼容依赖放在 `.cache/acceptance-office-deps/installed/usr/lib/x86_64-linux-gnu`，运行参考时通过 `LD_LIBRARY_PATH` 指向该目录；常规 Ubuntu 验收环境无需该额外配置。

```sh
pnpm fixtures:xlsx
pnpm check:xlsx-fixtures
python3 scripts/check-xlsx-output.py examples/xlsx
python3 scripts/check-xlsx-output.py examples/xlsx/producer
pnpm record:xlsx-reference
pnpm check:xlsx-reference
```

只有 `record:xlsx-reference` 更新参考。普通检查验证已保存摘要并重新打开/重存；始终比较源结构、重存结构及已记录的外部差异，相同环境才比较页数与像素，不同环境明确报告未进行像素等同性比较。Office CI 执行两项检查并上传 `.cache/xlsx-reference/result.json`，不将生产者本身已记录的丢失算作 SDK 无损验收。


## 编辑阶段的自定义表样式来源

新增来源保存在 `packages/office/tests/fixtures/workbook-table-style.xlsx`，由 `scripts/generate-workbook-table-style-fixture.py` 使用 Python 标准 ZIP/XML 从不可变的 `05-native-table.xlsx` 派生，包含两份共享差异格式的自定义定义、四类区域、条纹及默认／表关联；原始十六份样例字节不变。[源摘要](../../docs/evidence/workbook-table-style-source.json)记录基准和新增文件 SHA-256。

[固定 LibreOffice 重存与实际 B2 修改](../../docs/evidence/workbook-table-style-native.json)另外记录官方 25.8.1.1 归档摘要及生产者差异：Calc 移除自定义定义、默认值和样式关联。源及两份输出均通过官方 Transitional XSD／OPC 检查，原始来源和生产者损失分别保留；不将其称为原生 XLSX 导出或 Excel/WPS 实机验收。复现使用 `pnpm check:workbook-table-style-source` 和 `pnpm check:workbook-table-style-native`。


[编辑阶段链接复制](../../docs/evidence/workbook-copy-hyperlinks.json)使用不可变 `01-values-text.xlsx` 和固定 LibreOffice 的普通、重叠、跨表及新建表复制。四份生产者输出保留在 `packages/office/tests/fixtures/workbook-copy-hyperlinks-*-calc.xlsx`；五份源／输出均经官方 XSD／OPC 检查。记录保留显示缓存重算、提示丢失和新建表默认字体页眉不在当前字段读取范围内的拒绝负例，不覆盖原始样例。复现命令为 `pnpm check:workbook-copy-hyperlinks`。

[编辑阶段系统剪贴板接收](../../docs/evidence/workbook-system-clipboard.json)使用同一不可变独立输入，在 SDK 中修改当前源后生成 HTML，由固定 Calc 原生控制器插入并再次修改值验证公式重算。生产者文件 `packages/office/tests/fixtures/workbook-system-clipboard-calc.xlsx` 保留精确字节及摘要，记录小数精度、日期序列、Boolean 公式、字面等号、长编号、合并，以及 CR／LF、制表符和常量错误的生产者差异。官方 XSD／OPC、SDK 再编辑、Worker 当前值、包 ZIP 重开与撤销分别核对。复现命令为 `pnpm check:workbook-system-clipboard`，不称为 Excel/WPS 实机验收。

[名称行列编辑](../../docs/evidence/workbook-axis-names.json)由独立 Python ZIP/XML 生成源和五份固定 Calc 实际插删结果，保存在 `packages/office/tests/fixtures/workbook-axis-names*.xlsx`。相对配方、绝对分量、混合整轴范围、局部／全局 bang、别名及生产者缓存均保留；Calc 的 bang 和插删列后嵌套相对别名差异单独记录。源及原生输出经官方 XSD／OPC，SDK 当前计算、唯一 Worker 候选、再编辑、包 ZIP 重开及撤销分别验证；`pnpm check:workbook-axis-names` 重现原生操作。这不代替 Excel/WPS 结构编辑实机验收。

[合并区域选择性粘贴与数列](../../docs/evidence/workbook-merged-special.json)使用独立 ZIP/XML 来源及二十份固定 Calc 原生输出，文件前缀为 `packages/office/tests/fixtures/workbook-merged-special`。十八次实际控制器粘贴明确覆盖、先重算，两次 `fillAuto` 验证纵向／横向合并数列；保留 Calc 取消目标合并和保存非空覆盖格的差异。全部生产者字节、官方 XSD／OPC、当前值、SDK 再编辑重开及历史分别检查；`pnpm check:workbook-merged-special` 实际复现，不改已有独立样例或宣称 Excel/WPS 实机通过。

[智能表局部剪切](../../docs/evidence/workbook-partial-table-move.json)复用不可变 `workbook-axis-tables.xlsx` 独立源，保留五份实际固定 Calc 剪切和五份再次原生重开输出，前缀为 `packages/office/tests/fixtures/workbook-partial-table-move`。记录列配方和原生筛选差异，以及表外当前行公式初次缓存与重开后错误的差异；SDK 按精确输入映射求当前值，不改生产者字节。复现命令为 `pnpm check:workbook-partial-table-move`，包含官方 XSD／OPC；不称为 Excel/WPS 实机验收。
