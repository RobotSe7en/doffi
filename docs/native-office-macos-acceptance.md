# macOS 桌面 Office 实机验收

2026-11-24，macOS 27.2（26B123），Word、PowerPoint、Excel 16.114.1。按用户要求排除 WPS。以下是本机抽样结果，固定 LibreOffice 的独立验收和参考基线继续保留。

## 原生编辑与重开

使用仓库独立 ZIP/XML 样例的副本。SDK 先导入、原生导出并重开 9 份 DOCX、16 份 PPTX 和 16 份 XLSX。从中抽取下列文件交给桌面 Office。Office 实际编辑保存后的文件再用公开 SDK 入口读取。

| 应用与样例 | 实机结果 |
| --- | --- |
| Word，`02-table-merges.docx` | 无修复提示打开，插入文字并原生保存。SDK 严格重开，新增文字和三行合并表格保留。两边 PDF 均为一页 A4；横纵合并正确，行高有可见差异。 |
| Word，`08-comments-revisions.docx` | 11 月 24 日 SDK 导出件出现恢复提示；11 月 26 日已定位并修复，见本页 Word 补充验收。保留首次失败记录。 |
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

## 2026-11-25：11 页 PPTX 静态版式对照

对用户授权的技术汇报源件，在 PowerPoint 16.114.1 实际打开并导出全部 11 页 PNG（1280×721），另存原生副本核对有效属性。使用 Codex 内置浏览器逐页打开同一原文件，修复后再次核对全部页面。原文件没有改写；原生副本、参考 PNG 与对照缩略图仅保存在本机 `.cache/pptx-css-limit/`，未提交用户内容。

| 差异 | 修复与证据 |
| --- | --- |
| 中文字体不同 | 源件 363 处东亚字体声明位于复杂文字字体之后；PowerPoint 重存移除了这些未生效声明。SDK 保留其原名为未生效声明，并按原生有效继承显示。 |
| 多行文字超出文本框 | 非空段落不再由母版较大字号撑高 CSS 行框，实际 run 字号保留。覆盖预览、编辑、撤销和打印。 |
| 页码多 1 | 保留并应用源件的 `firstSlideNum=0`；另覆盖非零起始值及原生导出重开。 |
| 封底标识、组合文字和细线缺失 | 对极小组子坐标作派生坐标调整，绘制零高／零宽线条；不改变源几何。实际指针拖动、缩放与撤销覆盖显示坐标到源坐标的换算。 |
| 自定义方块列表标记不同 | 原生对话框仍保存声明颜色，实际方块使用系统彩色字形；预览采用本机可用的相同回退字体。 |
| 字体诊断误报 | 空白字符不再被判为缺字，预览默认字体不计为文档请求的字体替换。 |

最终内置浏览器逐页检查：11 页均无图片加载失败，无超过 2 px 的文字容器纵向溢出。此项是本机静态版式对照，不把字体栅格差异声明为跨平台像素相等，也不覆盖动画播放或含保留定义的原生 PPTX 导出。自动回归使用公开合成源和固定字体，未收录用户源件。

定向自动验证：357 项核心／Office／字体单元测试、95 项 Chromium 预览与编辑测试通过；SDK 与 Playground 构建、工具类型检查、架构边界和差异空白检查通过。浏览器回归还修复了后台缩略图任务清除输入冲突提示的问题，并等待缩略图完成后验证提示仍然可见。

## 2026-11-25 字体菜单与本地加载

同机 Microsoft Office 16.114.1 抽样记录：Word 字号菜单包含初号至八号和数字字号；PowerPoint 使用 8～96 的演示字号预设；Excel 使用 6～72 的单元格字号预设。Playground 分别采用对应列表，Word 输入数字后显示对应中文字号。字体菜单显示本地配置名称和最近使用项，移除了字体框旁的黄色标记。字体替代与缺字诊断仍由 SDK 报告。

默认开发服务读取根目录 `fonts/local/sources.json`，浏览器按文档引用加载字体；Word、PPT、Excel 的首次测量与本地 PDF／PNG 使用同一配置。字体二进制由 Git 忽略，不进入 SDK。Word 的异步字体完成不再独立替换编辑 DOM；PPT 也通过编辑器重新绑定选区。7 项定向 Chromium 回归覆盖延迟字体、输入法组合、后续输入及撤销、Excel 加载前测量、连续切换和 DOCX 导出重开。原用户 PPT 已在内置浏览器重新打开，默认不显示批注栏、对象列表或内部保留说明。

本次还查看了 PowerPoint“字体／字符间距”和 Excel“设置单元格格式”的原生弹窗。随后通过“访达”前置 Word 应用，取得了其“字体／高级”和“段落”三个选项卡的实际窗口，补齐原先未取得的对照。


## 2026-11-25 SDK Office 外框与功能区

Word、PowerPoint、Excel 16.114.1 的选项卡、Home 命令顺序和字体／字号预设按实际窗口对照。Word 的 1920 px 窗口参考：标题栏 40 px、选项卡 32 px、功能区 74 px，字体组 375 px、段落组 338 px；较窄窗口将样式库收为“样式”按钮。PPT 保留“切换／动画／幻灯片放映／录制”，Excel 保留“页面布局／公式／数据／自动化”。

外框已归入公开 React `OfficeWorkspace`。三种格式的菜单、色板和图标均在 SDK，未依赖 Playground CSS；图标为项目自绘 SVG，粘贴、格式刷、听写、编辑器、项目符号、编号、缩进等使用分层着色。普通 DOM 字体窗口共用 Office 模态框及可编辑下拉列表。Excel 的格式弹窗可从隐藏内部工具栏的完整工作区打开；原先不可见的问题已有回归。

本次针对样式、字体、对话框事务及原生重开进行定向验证，不把未实现的云服务、智能听写、加载项等命令声明为可用；这些入口保持禁用。

自动验证共 55 项 Chromium 用例通过：完整工作区及 Word／PPT 独立闭环 45 项，延迟字体、字体窗口与 Excel 格式事务 10 项；另有 12 项单元测试覆盖直接与 Worker 命令共享的数字格式、原生 XLSX 重开和撤销。独立宿主用例不加载 Playground CSS，检查外框与分组尺寸、32 px 图标、菜单定位、窄窗口样式库和多实例隔离。最终 PPT 外观调整后重跑 5 项功能区、批注、字体、放映及独立宿主用例。

构建、工具类型检查、架构边界和差异空白检查通过。浏览器使用本机 Google Chrome，PDF／PNG 验证使用同一 Chromium 和运行环境附带的 Poppler。原始 11 页用户 PPT 再次在内置浏览器打开，保留原文件；缩略图选中框为 PowerPoint 橙色，画布按工作区剩余高度布局。

## 2026-11-26 Word 批注字体和内置浏览器对照

继续使用本机 Word 16.114.1、macOS 27.2。用户明确授权 AppleScript 控制 Word，只操作仓库合成样例的副本。网页使用 Codex 内置浏览器打开本分支 Playground，同文件上传，不连接云端服务。原生文件和调试副本在忽略目录 `.cache/word-gap-2026-11-26/`；包含 Office 本地元数据，不提交为公开夹具。

| 项目 | 原生观测与修复 |
| --- | --- |
| 批注与文字修订，COMP-01 | 独立 `08-comments-revisions.docx` 正常打开；旧 SDK 导出触发恢复。四类批注扩展部件的 MIME 错写为 `application/vnd.ms-word.*+xml`，更正为 `application/vnd.openxmlformats-officedocument.wordprocessingml.*+xml`。新 SDK 导出无恢复提示，Word 读到一条批注、两条修订；插入 `Native Word verification: ` 并实际原生保存后，SDK 严格重开保留这些内容。 |
| 根批注与回复 | 单独合成线程含 Reviewer A 根消息和 Reviewer B 回复。只有父级元数据而没有回复正文引用时，Word 只显示根消息并在重存中删除回复。为每条回复补点引用，根批注仍独占文字范围。Word 窗格实际显示两条消息；插入 `Native reply roundtrip: ` 并保存后，SDK 严格重开仍有一个线程、一条回复和原文字锚点。没有执行原生解决／重新打开线程，不将其计为通过。 |
| 字体窗口，WORD-03 | 新编辑入口使用已有字符格式事务。合成 DOCX 在 Word 无恢复提示打开，原生字体 API 读到横向缩放 75%、双下划线、下划线 RGB 16 位值 `52428,13107,0`，对应 #CC3300。自动回归另检查预览和编辑的缩放与颜色、无效值不提交、原生导出重开及撤销。 |
| 内置浏览器视觉抽样，WORD-01 | 将独立批注修订样例的字体声明换成两端实际可用的 DengXian，其他内容保留。Word 与网页均使用最终修订视图；按相同纸张显示宽度比较，Word 100%、网页 75%（两者屏幕点换算不同）。1120×728 窗口下纸宽均约 595 px，正文距纸张左侧均约 72 px，两个段落的文字顶端相隔约 26 px；页内首行顶端仍有约 2 px 差异。此为同机视觉抽样，没有跨平台像素相等结论。 |

原样例请求 Noto Sans CJK SC，本机 Word 与网页回退字体不同；两端默认修订显示方式也不同。先统一字体、纸面宽度及修订显示，再比较正文。浏览器预览与编辑切换后的修订视图偏好尚未统一；复杂表格、浮动组合、公式与多页文档没有因这一页样例而获得整体验收。

扩展到同字体合并表格后，内置浏览器显示表格边界比 Word 向左约 4 px，三行高度也有小差异。另造无样式、空默认样式、自定义默认／显式样式、显式边距及更换样式 ID／名称的独立副本，直接查询 Word 的表格边距和行缩进：只有样式 ID 相同不足以确定内置默认，`Normal Table` 名称与普通自定义样式的未声明属性存在差异。仅按是否绑定样式选择 5.4pt 的候选修改无法覆盖这些情况，已撤回，未纳入本轮修复或验收通过项。后续需要统一处理样式有效值、继承／覆盖、缩进原点和兼容模式，再覆盖预览、编辑、原生导出重开；不能用这个样例的坐标补偿代替通用规则。

四类部件 MIME 的依据为 Microsoft Open XML SDK 的原始元数据：[CommentsExtended](https://raw.githubusercontent.com/dotnet/Open-XML-SDK/main/data/parts/WordprocessingCommentsExPart.json)、[CommentsIds](https://raw.githubusercontent.com/dotnet/Open-XML-SDK/main/data/parts/WordprocessingCommentsIdsPart.json)、[CommentsExtensible](https://raw.githubusercontent.com/dotnet/Open-XML-SDK/main/data/parts/WordCommentsExtensiblePart.json)、[People](https://raw.githubusercontent.com/dotnet/Open-XML-SDK/main/data/parts/WordprocessingPeoplePart.json)。新增回归同时断言精确 MIME、根范围与回复引用，保留既有多次 SDK 重开及批注保留检查。

| 实际原生编辑保存文件 | SHA-256 |
| --- | --- |
| `export-fixed.docx` | `0ea9dfd74d6d9491950e0aacebee484f1604f3c089fd90972d9024c338c49b9f` |
| `thread-reference-fixed.docx` | `dfcd8cc24e5fd2b1a0cc52a3ca8e727eaa3d3cab4f9174518c2fd9b37a319e52` |

另外修复 UI-02 的错误路由：绘图选择窗格显示当前故事的对象及组合子项，支持指针多选、键盘定位、属性窗口和删除。父组与子项同时选中时，移动、复制、组合和删除统一使用顶层选项，避免对子对象重复处理，事务可以撤销。只读查找、标题导航和统计已经接通工作区；定位使用当前修订的准确 UTF-16 范围，不触发编辑或新增历史。它们不代表完整 Word 大纲视图及全部选择窗格操作。

定向验证：74 项相关 Office 单元测试、9 项 Chromium 浏览器测试通过，覆盖多个范围／点批注线程、字体导出、组合父子对象去重与撤销、既有拖动缩放、只读查找和导航、外部内容变化后定位刷新、完整工作区入口及独立 SDK 宿主的预览／编辑切换。SDK／Playground 构建、工具类型检查及架构边界通过。没有运行远程全量 CI。
