# macOS 桌面 Office 实机验收

2026-11-24，macOS 27.2（26B123），Word、PowerPoint、Excel 16.114.1。按用户要求排除 WPS。以下是本机抽样结果，固定 LibreOffice 的独立验收和参考基线继续保留。

## 原生编辑与重开

使用仓库独立 ZIP/XML 样例的副本。SDK 先导入、原生导出并重开 9 份 DOCX、16 份 PPTX 和 16 份 XLSX。从中抽取下列文件交给桌面 Office。Office 实际编辑保存后的文件再用公开 SDK 入口读取。

| 应用与样例 | 实机结果 |
| --- | --- |
| Word，`02-table-merges.docx` | 无修复提示打开，插入文字并原生保存。SDK 严格重开，新增文字和三行合并表格保留。两边 PDF 均为一页 A4；横纵合并正确，行高有可见差异。 |
| Word，`08-comments-revisions.docx` | SDK 导出件在 Word 打开时出现“发现无法读取的内容”及恢复提示。该项未通过；原始输入对照与恢复后差异尚未取得可用结果。 |
| PowerPoint，`01-rich-text.pptx` | 无修复提示打开，通过键盘选中文字对象，追加文字并原生保存。SDK 严格重开并读到新增文字。原生打印及 doffi 幻灯片输出都可读，但纸面尺寸不同，doffi PDF 中列表标记和扩展字存在可见差异。 |
| PowerPoint，`12-modern-comments.pptx` | 原生打开三张幻灯片，评论面板显示线程和回复。此次只检查原生可见性，尚无原生编辑重存线程的结论。 |
| Excel，`04-formulas-names.xlsx` | 原生打开，录入 `D50=123` 和 `E50 = D50*2`。原生保存后，SDK 严格重开，读到公式及 Excel 计算的缓存值 `246`。 |
| Excel，`10-modern-threads.xlsx` | 独立源件和 SDK 导出件的 Excel 评论面板均显示没有批注。修改导出件的单元格并保存后，线程及兼容批注部件被移除，SDK 严格重开得到零条批注。现代线程实机验收未通过。 |
| Excel，`11-print-layout.xlsx` | 原生打开独立源件。在 Excel 打印窗口显式选 A4 横向后，预览显示两页，与 doffi PDF 一致。得到的原生 PDF 使用打印机的 4×6 预设，且为四页。未取得同纸型原生 PDF，页眉页脚、重复标题和像素差异未作同条件对照。 |

Codex 内置浏览器另打开了 Word、PPT 和一万行 Excel 的 Playground 样例。视觉抽样使用同机安装的字体及 Chromium 生成 doffi PDF，使用应用的原生打印 PDF 和 Poppler 页面图比较。字体环境未固定到跨机器像素验收基线。

## 删除合并区域首行或首列

独立源工作表含合并范围 `B2:D4`、左上角文字 `MERGE-TOP-LEFT`，以及区外指向 `B2` 的单格引用和覆盖 `B2:D4` 的范围公式。分别在 Excel 选择区域外的 `A2` 或 `B1`，经编辑菜单的删除对话框选 **整行** 或 **整列**，各执行一次并原生保存。使用区外单元格可避免选择扩张到覆盖整个合并区域。

| 操作 | Excel 原生重存 | doffi 事务及原生重开 |
| --- | --- | --- |
| 删除第 2 行 | 合并缩为 `B2:D3`，原 `B2` 文字被删除，单格引用变成 `#REF!`，范围公式变成 `COUNTA(B2:D3)` | 范围、内容和公式一致 |
| 删除 B 列 | 合并缩为 `B2:C4`，原 `B2` 文字被删除，单格引用变成 `#REF!`，范围公式变成 `COUNTA(B2:C4)` | 范围、内容和公式一致 |

该观测补齐此前延期的首行／首列对照。现有结构编辑行为与 Excel 一致，既有 Calc 差异保留。自动回归检查合并范围、内容、公式、原生重开和撤销。首次重开时，旧 SDK 会拒绝 Excel 保存的附属元数据。本轮已识别试验中出现的来源路径提示、空修订指针、计算引擎版本、工作表默认拼音设置及默认切片器／时间线样式。原生路径和空修订标识不进入文档包。未知扩展、非空修订历史和实际切片器内容仍按原有边界拒绝。参见 [工作簿](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/e29966e2-5baa-4fcf-84c9-082025e1be13)、[样式](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/d8afb021-974d-403a-84de-21af334d3d1d)及 [计算引擎版本](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/c1a60ba0-8af1-4b09-a666-8118f1c39d69)说明。

## 证据

原生输出、来源副本和 PDF 保存在本机 `.cache/native-office-mac-2026-11-24/`。原生文件可能含本地路径及 Office 作者信息，只在本机保留。下表记录与结论直接对应的原生文件 SHA-256。

| 文件 | SHA-256 |
| --- | --- |
| `ui/ms-word.docx` | `515217a8aa7acd4b4546df9618bba59981f7c06209be0b2a940697e2a4c09dbb` |
| `ui/ms-powerpoint.pptx` | `652a220df15d81d8498fe115db8a2ab93858967b658cceafd3bf2733c9f3d303` |
| `ui/ms-excel.xlsx` | `5de1899daa67e498de7c12437de8525a46843601ae2dd070b8353856d8a5a2fd` |
| `smoke/row-first.xlsx` | `29881ab2fb14bb4fcdc061d2f065a50ffef5cb141f6729ff0fdebc829d6a2210` |
| `smoke/column-first.xlsx` | `a3057fcfd4b8f1a326afe4de992003d15b8f56d4a7590994e6b5298e3575bf6b` |

自动验证使用 [Excel 重存回归](../packages/office/tests/xlsx-excel-resave.test.ts)和现有 XLSX、结构编辑测试，共 34 文件 239 项通过。SDK 构建、类型检查、架构边界及差异空白检查通过。
