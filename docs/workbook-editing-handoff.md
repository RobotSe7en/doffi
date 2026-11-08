# 第三轮工作簿编辑计算执行交接

2026-11-07 用户授权开始 R3-09～16，后续明确独立审查安排在提 PR 前：两位子智能体分别独立审查完整改动，实施者修复并交回复核，再提交阶段 PR；已有 #31 保持草稿直至该步骤完成。2026-11-08 用户再次明确要求完成整个 R3，包括 R3-17～20 输出阶段，各阶段在合并前独立审查、修复复核，最终 HEAD 的日常 CI 与 `workbook-scale` 全部成功后由实施者合并。此次授权不包括子智能体实施。原任务范围、产品选择和门槛见 [第三轮细则](iteration-3.md) 与 [AGENTS.md](../AGENTS.md)。

## 基线与当前状态

- main／远端 main 的基线为 `5451736`；当前分支 `feat/workbook-editing`。
- 读取阶段已合并；CI 优化及触发频率调整分别为 `31cca2f`、`d65c0e9`，不再是待办。
- R3-09/10/11/12/13 实施中，R3-09～16 均未完成；公共 Worker 编辑器／Playground 为 `b161098`、跨表复制为 `04ab466`、合并事务为 `5d2eb2f`、默认编辑器容量场景为 `8b64e10`、普通筛选事务为 `3bc5307`、冻结界面及实际坐标查询为 `c917075`，均已推送；R3-13 定义编辑基础随后接入。[阶段草稿 PR #31](https://github.com/RobotSe7en/doffi/pull/31) 未执行阶段最终 CI，不能合并。不能将这些进展计作完整 R3-11、R3-12、R3-13 或 R3-16。
- 两位授权审查者 `workbook_edit_review_one`、`workbook_edit_review_two` 已完成两轮只读基础审查。首轮 8 项、第二轮 6 项原始问题及 3 项追加边界，均由实施者修复并经原审查者复核关闭。整个阶段仍须最终差异复审，不能凭基础审查合并；详见[审查记录](workbook-editing-review.md)。
- 不直接提交 main，不发布 npm；普通推送只执行快速检查。阶段最终验收按 [CI 执行](ci-performance.md) 显式发起。尚有 R3-09～20 共 12 项未完成验收：本阶段 8 项中 5 项实施中、3 项待实施，输出阶段另 4 项。

## 已有改动

- `session/lifecycle.ts` 提取 Word／PPT／工作簿共用候选身份、修订守卫、发布、历史、事件与预算边界；既有内容事务改为适配该边界。
- `createDocumentSession` 按工作簿／内容模型返回专属会话类型；固定 Word／PPT 测试使用已有 `openContentFixture` 保持类型准确。
- 工作簿私有来源注册表保存不可变块字节和索引；候选复用未改块及已验证公式模板。公开序列化仍复制字节，私有来源不公开授权入口。
- `WorkbookSession.prepareAsync` 与同步入口驱动同一候选步骤，捕获输入／授权资源／计算上下文／确认列表；写入、稀疏清除、块生成／校验／比较、规则及计算检查点支持宿主让步、进度和取消。正在准备占候选名额，同会话仅一项异步准备；来源变化以修订冲突失败，取消及释放不发布私有状态。填充和替换的直连界面已接入。事务结构解析、资源捕获与单个有界块处理仍同步；默认 Worker 界面已接入，完整结构命令仍待完成。
- editor 的 `serveWorkbookSession`／`connectWorkbookSession` 绑定宿主唯一核心会话与专属授权端口；主线程只持有状态、影响摘要和有界读取视图。原始事务／复制／填充／替换在 Worker 中准备；当前视图复用会话增量计算，候选视图隔离计算；手动模式保留，显式时间／地区由宿主传入。候选只在所属连接应用，取消等待远端清理确认，迟到候选和视图端口仍回收；提交响应超时明确报告结果未知，超额回执明确已提交修订。关闭端点不释放宿主会话。公共 Worker 编辑界面和 Playground 已复用该通道，仍不是完整 R3-16。
- 工作簿单元格候选只解析修改块；支持常量、公式源、富文本、保留样式的内容清除和完整清除。基础工作簿设置与宿主单元格批注经过同一历史。
- 结构映射原语包含行列插删、精确剪切碎片、排序置换及反向映射；普通区域行列已接入，剪切、排序和原生对象联动仍继续实施。
- 后续 `editWorkbookAxis` 已接通普通区域行列插删：保持块内源正文、缓存和格式，维护跨表公式、绝对限定名称、规则原点／区域、维度、合并范围、冻结、链接、普通筛选键及打印定义。宿主批注精确分段并排除插入空白，完全删除保留正文供重定位；同批逆向的已覆盖场景恢复维度／公式／批注分组，保留重做。格式继承显式为 none／before／after，默认 none；界面默认选择 before，可修改。新增格式格在展开前检查总格预算。直连使用异步准备，Worker 使用原唯一会话；四个按钮按选区确定行列数量，草稿阻止结构操作，完成后回到原生输入。
- 行列基础仍有明确未接通组合：目标工作表含智能表／绘图／原生备注／线程、工作簿含上下文名称、普通筛选表头的部分删除会失败；需在本阶段继续完成，不是范围缩减。现有独立 Calc 删除对照取消合并，SDK 当前缩小范围。2026-11-08 用户已选择按 Microsoft Office 行为实现，以 Microsoft Excel 为兼容目标；不再等待用户选择。具体结果的微软依据／Excel 实测仍缺失，该对照不计为合并删除验收通过，不把 Calc 或当前 SDK 结果冒充 Excel 行为。其余 4 组原生公式／规则／打印来源由 `scripts/check-workbook-axis.py` 生成及实际编辑，官方 LibreOffice 25.8.1.1 的归档 SHA 和构建号精确校验；源及四份实际生产者通过未修改官方 XSD／OPC。证据 `docs/evidence/workbook-axis.json`，运行 `pnpm check:workbook-axis` 可重做实际插删、重算、保存并复核，已加入显式完整 Office 验收。
- 资源暂存抽成公共包级原语，保留宿主授权字节、不可变路径与最终使用检查。
- 停止型验证复用只读规则判断，检查最终修改值、计算结果及阈值／列表／计算名称／相对公式依赖影响的既有目标；不允许空白的规则包含受影响隐含格。无关编辑和格式修改保留导入已有违规值；警告／信息规则必须显式确认，不能绕过停止规则。行列结构及规则编辑的完整影响仍随后续操作继续接入。
- 增加选定函数族、显式宿主时间／地区上下文、区域依赖索引、迭代长链／实际循环与条件读取、增量失效和整批计算。同步与异步计算共用生成器；取消、来源变化和迟到批次不发布部分结果。
- 标量求值器和规则读取会捕获会话的不可变当前模型，旧求值器不会把新单元格值放进旧修订结果。
- `captureWorkbookModel` 无需序列化即可捕获内建会话当前不可变模型；窗口／资源／图表／规则适配器统一使用它，避免旧修订读取混入新会话内容。宿主自定义模型仍须自行满足已有不可变提供者契约。
- 工作表新增、重排、隐藏、属性、改名和删除接入同一候选；同步维护单元格／名称／规则／表公式／图表／内部链接，删除使相关批注需要重新定位并保留正文。富对象内容复用 Word／PPT 的有界 HTML 编辑器；5 份独立原生来源覆盖图表、绘图、经典备注和现代线程的改名／删除／历史恢复。
- SUBTOTAL 使用当前普通区域／表筛选定义：值、通配条件、Top／Bottom 含并列、均值、日期分组及显式宿主时间下的相对日期；手动隐藏行仍由函数编号控制。筛选输入进入依赖发现。完整筛选 UI 与结构操作仍未实施。
- 第一批人工输入原语：`parseWorkbookTextInput` 保守识别公式、数字和布尔，保留前导零／长编号／含糊日期为文本，显式模式允许文本／数字转换；`parseWorkbookTsv` 支持带引号的制表符、换行、双引号及 UTF-16 变体序列，并在生成单一跨块粘贴事务前检查矩形、长度与预算。
- `createWorkbookInputController` 只保存选择和暂存输入，原生 IME 不直接改来源；确认后准备／提交同一核心事务，过期选择拒绝且保留输入。无修改确认不会把数字样文本或空文本重新转型。
- `bindWorkbookTextInput` 将宿主原生 textarea 接到该控制器：Enter／Tab 确认、Alt+Enter 换行、Escape 取消、F2 编辑；导航、矩形清除、TSV 粘贴和撤销重做由宿主回调经过同一事务。过长原生草稿不能误提交控制器中的旧值。跨预览 iframe 返回输入时先激活所属窗口，修复 Firefox 显示聚焦但键盘仍在 iframe 的问题。
- `WorkbookPreview.selectRange` 按来源坐标选择与滚动，复用合并／冻结／RTL 坐标，并拒绝期间更换工作表／来源的迟到选择。`retainSelection` 可在同修订滚动／缩放时保留选择；更新同文档时保留仍可见的当前工作表。
- `WorkbookPreview.moveSelection` 复用实际尺寸索引跳过完整隐藏／零尺寸区间及合并内部，支持 RTL 物理方向、原端点扩选及按活动端滚动；选择不会新增来源节点。连续编辑器导航经排队处理，不丢失重复按键。
- `mountWorkbookEditor` 的小型文档直连路径接入宿主核心会话、公式栏、地址选择、显式文本／数字模式、确认／取消、单步历史、当前计算结果与网格。错误输入切换单元格／工作表时保留草稿和旧选择；外部修改期间保留绑定原修订的输入，取消后才显式读取当前源。跨块 TSV 与矩形清除均为一个核心事务。直连路径供小型嵌入，默认 Worker 入口见下；完整剪贴板、富文本格式和其他 R3-11 边界仍待完成。
- 完成的计算批次可通过 `createWorkbookCalculatedEvaluator` 及 `createWorkbookPreviewSource` 为网格、自动行高、规则和图表提供同一时间／地区上下文与长链结果。私有注册表同时校验实际批次及不可变来源身份，不能用相同文档 ID／修订、复制结果对象或导入缓存替代；释放计算器不会清空已捕获批次。当前值投影不重写单元格源或缓存；Worker 通道的计算视图已为默认编辑器提供同批窗口／规则／图表。
- 工作表操作在每次改名／插入前维护大小写不敏感的唯一名称空间，防止同批中间重名把两个工作表的公式误改到同一表。名称交换须经过唯一临时名，仍为一个原子事务及撤销步；无效批次不改变来源或历史。
- `readWorkbookCellSource` 在固定来源上读取类型化输入，保留中文、CR、IVS 和富文本跨度／字体类，公式缓存不转成常量；源格式身份解析上移至 core，供渲染及命令共用。`copyWorkbookFormula` 公开既有第一方相对／绝对引用复制。
- `createWorkbookFillTransaction` 从固定修订的种子生成一批事务，支持相对公式复制、数列前后延伸、来源格式或保留目标格式、预算、进度及宿主让步／取消；重叠目标不读取先前生成结果，空白填充保留稀疏来源。直连编辑器接入向下／向右／数列填充与取消操作，停止型验证仍在同一候选边界执行。缺少 General 定义时生成有碰撞检查的新定义，与实际需要它的填充格共享一个事务和撤销步；空白无变化时不留下定义。当前填充实现仅覆盖同表未合并区域；合法合并模板及完整复制／超链接等仍需后续补齐，不能据此收缩 R3-11 验收。
- `replaceWorkbookCellText` 按带原文证明的、不重叠的原始 UTF-16 区间修改文本常量，复用核心的代理对／变体序列边界；未修改的字体跨度保留，插入文字继承首个受影响跨度，富文本输出仍须经过同一事务验证。原生单格文字／IME 编辑接入该原语。textarea 的 LF 仅为瞬时显示投影，未编辑的 CR／CRLF 不会被浏览器规范化写回来源。数字样、公式样和空文本使用输入转义以保留原类型，显式类型转换仍可用；转义不写入模型。
- `findWorkbookText` 按固定修订、表／行／列顺序扫描实际存储格，不展开隐含空白区域；支持字面量、大小写、整格、区域／工作表／工作簿范围、源内容／公式／当前值字段，保留实际 UTF-16 位置。显示值使用完成批次与共享的表格／命名样式／直接格式／条件格式数字格式解析，不使用导入公式缓存；缺少批次或不可求值会返回明确的不完整结果。显示查找使用数字格式文本，忽略宽度裁剪与排版填充间距。
- `createWorkbookReplaceTransaction` 生成一个原子事务，默认保留常量类型和富文本跨度；显式模式允许类型转换。单项替换重新核对查询、修订和原文位置，`$&` 等替换内容保持字面量；公式替换仍须为合法公式。预算、取消和修订变化不会提交部分替换。直连编辑器接入查找循环、单项／全部替换、范围保留和取消；核心异步准备、Worker 替换命令和默认界面已接入，完整容量门槛仍待完成。
- 隐藏表的匹配项显示表名和地址，不自动取消隐藏；单项替换仍校验原文和修订。显示值查找的重算结果同步提供给网格，手动计算模式也不会让匹配文字与当前网格使用不同批次；导入缓存、原值和历史不被重算改写。
- `setWorkbookCellFormats` 修改已有类／数字格式身份，保留公式缓存及精确富文本正文；原属性恢复时可恢复原始来源字节。`setWorkbookNumberFormat` 按身份新建／更新共享定义，自定义代码经过第一方格式语法检查。二者均使用同一候选／历史，格式修改保留导入已有验证违规，内容修改仍受停止规则约束。新 CSS 类生成、命名样式传播及格式 UI 继续实施；完整样式命令须按传播标志维护绑定，不能修改共享类而影响未选单元格。
- `setWorkbookHyperlink` 只替换／删除给定矩形的链接，按最多四个精确外侧矩形保留旧链接，不展开空白格；URL 与内部引用沿用安全验证。仅在实际编辑链接的工作表比较按目标分组的范围覆盖，等价碎片恢复原表示，避免同批改回原目标仍新增历史。链接编辑 UI 及剪贴板中的完整链接维护尚待接入。
- `createWorkbookRangeCopyTransaction` 为同／跨表未合并区域提供内容／数值／公式／格式命令，按固定来源读取重叠种子并只生成一个事务。数值模式要求真实完成批次，支持长链与原位公式转常量；公式／数值模式保留目标格式和未修改的富文本跨度，空文本不当成空白。格式模式复用纯格式事务，不读取或覆盖源内容缓存。`targetSheetId` 同时用于公共命令和 Worker `copy`；跨表同坐标仍执行复制，来源计算与目标规则／格式分别取正确所属表。剪贴板生命周期、系统剪贴板、链接／批注复制及完整合并仍须继续实施。
- 候选单元格按最终块一起验证；若中间输入将被覆盖、清除或随工作表删除，则在丢弃前执行原有的有界单格验证。这样保留中间非法输入失败、最终完整包验证和原子失败，同时避免普通批量写入逐格重复构建验证块。富文本输入在 UTF-8 编码前检查孤立代理项和 NUL，避免编码阶段替换原文字。
- 结构契约预检按 Schema 身份缓存是否存在数组路径，不缓存任何文档输入；没有数组的标量子树直接交给原有完整校验。这样避免逐格为无关联合分支构造错误，保留嵌套数组首错停止、未知字段拒绝和错误路径／数量边界。

## 已做检查

- 区域外观／命名样式批次已提交推送为 `8f8d5f4`，PR 草稿说明同步，未触发全量 CI。随后接入 `formatWorkbookText`、公开 `createWorkbookTextFormatTransaction`、Worker `formatText` 和公式栏选中文字的字体按钮。字体子范围检查原文和精确 UTF-16，保留原片段／CRLF／变体字符／单元格关联，不把数字或公式缓存转换成富文本；原生 LF 和转义前缀只在输入转换层映射。局部样式仍走原CSS和共同候选／历史，不增设可写外观副本。字号／角度字段的浏览器未完成数字输入在提交前拒绝。
- 核心字体与既有样式定向 2 文件／11 项通过，后续核心／转换／Worker 3 文件／37 项通过，日志 `.cache/workbook-text-format-unit.log`、`.cache/workbook-text-font-integration.log`；构建与工具类型通过，日志 `.cache/workbook-text-font-integration-{build,types}.log`。初轮工具类型发现测试辅助函数返回整个操作联合导致展开时缺少判别收窄，修正为具体字体操作返回类型，不改运行时。初轮三引擎界面为直连 3 通过／Worker 3 失败，日志 `.cache/workbook-text-font-browser.log`，原因是新增测试查询方法未挂到显式测试桥；补齐后 6 项全部通过，日志 `.cache/workbook-text-font-rechecked-browser.log`，包含实际显示颜色、原生 CRLF／转义前缀、源保留、选择保持、草稿保护、唯一会话及一次撤销。更广的核心／编辑器与完整输入／Worker 三引擎回归运行中，结果另记；这仍不是阶段最终固定资源验收。
- 片段字体最终核心／编辑器及原生定义／排序来源回归 87 文件／1165 项通过，日志 `.cache/workbook-text-font-core-editor-regressions.log`；最终完整输入／Worker 三引擎 126 通过、6 项非 Chromium CDP IME 按声明跳过，日志 `.cache/workbook-text-font-final-browser.log`。工具类型通过，日志 `.cache/workbook-text-font-final-types.log`，`git diff --check` 通过；未启动远端全量 CI。后续表格行列探针临时记录 `.cache/workbook-table-axis-probe.log`：实际 Calc 有七组插删，包含丢失计算列模板、删数量列后将该结构化引用改指向单价、删除表头／汇总行后保留对应标志并提升数据的行为；这些仅为发现记录，尚未做该新探针的完整归档摘要和 XSD 联合验收，不能当作 SDK 或 Excel/WPS 通过。后续实现须保持微软目标和实际生产者差异。

- 排序／剪切提交 `5f4f63a` 已推送，PR #31 的说明已同步，仍为草稿；连接器更新 PR 返回 403，使用已有仓库认证的 `gh pr edit --body-file` 完成，未触发全量 CI。后续新增局部区域外观、共享类及命名样式定义／别名编辑。命名更新按 Microsoft cellXfs 的 false 继承／true 保留局部组规则传播，数字格式维护实际格与稀疏维度，不另存有效外观。局部格式保留富文本源、缓存及共享使用者，通过别名及局部规则保护源层叠和字体直接设置。CSS 编辑器由原 PPT 实现抽为共享工具；原生颜色／旋转／图案／边框等新增字面参数检查。两条编辑器提供折叠的完整单元格格式面板，包含高级图案、双色渐变和斜边框；富文本片段格式与完整定义管理仍未完成。
- 初期命名样式／既有 PPT 编辑 9 文件／116 项通过，扩展后的核心样式／PPT／Worker 11 文件／150 项通过，日志 `.cache/workbook-style-edit-unit.log`、`.cache/workbook-style-edit-regressions.log`；原生颜色及参数拒绝的核心／独立原生来源 2 文件／10 项通过，日志 `.cache/workbook-native-format-guard-unit.log`。源 `02-styles-dates.xlsx` 和未修改实际 Calc 文件分别检查格式保护、原字符／缓存／正文、包 ZIP 重开和精确撤销；这不是在 Calc 运行新样式更新命令。构建／工具类型通过，日志 `.cache/workbook-format-ui-final-{build,types}.log`。首轮界面为 3 通过／3 失败，日志 `.cache/workbook-format-ui-browser.log`；失败是测试写错错误码，另已修正未执行到的当前结果字段，未放宽源不变及历史断言。修正后直连／Worker 三引擎 6 项全部通过，日志 `.cache/workbook-format-ui-rechecked-browser.log`。新增字面参数校验后的完整本地检查与完整三引擎编辑回归待最终记录，不能用前一批 2064 项计作这些改动的最终检查。
- 本批完整 `pnpm check` 193 文件／2072 项、架构边界、类型、SDK／Playground 构建通过，日志 `.cache/workbook-format-complete-check.log`；输入／Worker／工作簿 Playground 完整三引擎为 126 通过、6 项非 Chromium CDP IME 按声明跳过，日志 `.cache/workbook-format-final-browser.log`。最初误将通用 Playground 文件加入该次浏览器命令，发现后中断并改为工作簿 Playground；该中断日志 `.cache/workbook-format-wrong-playground-interrupted-browser.log` 不记作验收通过。随后自查复现同一类名同时用于 TD 和内部 SPAN 时整选择器替换会丢失未改字体的边界，原失败 `.cache/workbook-style-nested-alias-before.log` 保留；别名改为逐元素复合选择器替换，同一复合项的重复类一起映射，其他元素原类保持。修复后核心／PPT／独立原生来源 10 文件／123 项通过，日志 `.cache/workbook-style-nested-alias-rechecked-unit.log`；构建另暴露克隆 AST 的类型收窄遗漏，已补充显式检查，后续构建／类型和新增真实嵌套样式的三引擎复核另记。上述 2072／126 项早于此窄范围修复，不记作阶段最终源验收。
- 上述类型遗漏的失败日志保留 `.cache/workbook-style-alias-before-types{,-build}.log`，在未确认构建成功前误发起的浏览器运行已中断，日志 `.cache/workbook-style-alias-unbuilt-interrupted-browser.log` 不计入复核。最终明确拒绝既有声明外的非默认保护与内部边框；核心／PPT／独立原生来源／Worker 11 文件／153 项、构建与工具类型全部通过，日志 `.cache/workbook-style-alias-published-{unit,build,types}.log`。在此成功构建后重跑三引擎 9 项全部通过，日志 `.cache/workbook-style-alias-published-browser.log`，包含真实嵌套选择器、连续三次格式编辑保留内部斜体、原生富文本改色／渐变／斜边框以及 Worker 的拒绝、格式、当前公式和历史。完整本地检查与随后窄范围修复的结果分开记录；未启动远端全量 CI。

- 定义事务提交 `68cdb0f` 已推送；完整 `pnpm check` 186 文件／2030 项及类型、架构、SDK／Playground 构建通过，日志 `.cache/workbook-definitions-complete-check.log`。随后增加排序、行可见性查询、剪切及编辑器私有剪贴板，尚不是完整 R3-11/12。排序从当前计算构建稳定排列，跳过隐藏／筛除行；外部公式保留地址，内部相对复制。剪切捕获旧源，支持重叠、跨表、精确引用片段、整块合并、链接与需要重定位的跨表批注；部分条件格式／验证规则保留精确目标和相对原点。智能表／筛选／绘图及定位备注的剪切与排序联动仍待完成，普通复制仍拒绝合并。
- 排序基础核心／Worker／独立原生来源定向 5 文件／54 项通过，日志 `.cache/workbook-sort-final-unit.log`；完整三引擎输入和 Worker 回归 105 通过、6 项 CDP IME 按声明跳过，日志 `.cache/workbook-sort-final-browser.log`。实际固定 Calc 六组排序与官方 XSD/OPC 记录及独立复核通过，日志 `.cache/workbook-sort-native.log`、`.cache/workbook-sort-native-recheck.log`；数字两组一致，混合类型／标点／大小写四组差异保留在[证据](evidence/workbook-sort.json)。驱动最初未提供 UNO 类型使排序未执行的错误已由探针发现并修正，原日志 `.cache/workbook-sort-native-initial-noop.log` 不记为通过或生产者差异。共享轴编辑驱动的四组实际 Calc 回归及 XSD/OPC 重新通过，日志 `.cache/workbook-axis-sort-helper-recheck.log`。
- 剪贴板及剪切／排序／定义／原生来源／Worker 定向 7 文件／70 项通过，日志 `.cache/workbook-clipboard-unit.log`；构建、工具类型通过，日志 `.cache/workbook-clipboard-{build,types}.log`。三引擎直连／Worker 私有剪贴板 6 项通过，日志 `.cache/workbook-clipboard-browser.log`，覆盖不提前删除、显式覆盖、引用与撤销、当前公式值、连续复制、过期拒绝、草稿保护及宿主持有会话。后续规则拆分的单独 6 项通过，日志 `.cache/workbook-cut-rules-unit-recheck.log`，类型通过；初轮仅因未变表达式保留原拼写的预期写错而失败，原日志保留 `.cache/workbook-cut-rules-unit.log`。最终完整 `pnpm check` 192 文件／2064 项、类型、架构边界、SDK 与 Playground 构建通过，日志 `.cache/workbook-sort-move-complete-check.log`；最终三引擎排序／剪贴板定向 12 项全部通过，日志 `.cache/workbook-sort-move-final-browser.log`。未启动远端全量 CI；阶段最终双人审查和固定资源验收继续保留。

- 2026-11-08 新增 R3-13 定义事务：名称／表格／表样式／条件格式／数据验证的设置和移除；名称和表格列改名按绑定维护引用，列名交换通过临时身份避免重定向；表移除转为普通格并展开结构化引用。计算列、表头和汇总写入实际源格，同模板的既有例外与未改汇总源保留；展开前检查格数预算并逐批支持取消。无净变化的定义恢复原属性顺序及缺省列表，保留重做。核心 68 文件／1038 项通过，后续汇总源保留修正后的核心／独立原生来源与文档包重开／Worker 定向 3 文件／41 项及包构建、工具类型通过，日志 `.cache/workbook-definitions-{core,integration,build,types}.log`。这不是实际 Excel/Calc 编辑该新命令的对照，也不是完整 R3-13：人工管理界面、命名样式传播、结构联动及容量验收仍须完成。
- 此批最终完整 `pnpm check` 186 文件／2030 项、类型、架构边界、五包和 Playground 构建通过，日志 `.cache/workbook-definitions-complete-check.log`；未启动远端全量 CI，不能代替完整阶段审查及最终源的容量门槛。

- 第一批事务、模型、Word 和 PPT 事务：4 文件／72 项通过。
- 核心全部测试和相关 Office 事务／协作回归：51 文件／935 项通过。
- 独立范围映射：5 项通过。
- 上述通过后增加了映射类型及私有包验证改动，应在后续检查中重新覆盖。
- 包构建通过；固定格式测试的类型分支补齐后 `tsconfig.tools.json` 类型检查通过。
- 未将任何检查记作阶段最终验收；阶段必须在最终源上完成全量检查与容量专项。
- `8136ccd` 上完整 `pnpm check` 全部通过：157 文件／1827 项，类型、边界、五包与 Playground 构建成功，日志 `.cache/workbook-editing-check.log`。覆盖两轮审查修复、验证影响、ISBLANK、不可变读取及新增原生参考；阶段最终检查仍须重新在完整阶段的最终源执行。
- 不可变读取修复后，窗口／图表／资源既有定向 48 项及新增 3 份原生来源会话读取检查通过；新适配器在提交与释放会话后仍返回捕获修订的数据。`build:packages`、工具类型检查、依赖边界及 `git diff --check` 通过。
- 第二轮审查修复新增 11 项回归：工作表身份归一、删除表引用、可见性及计算名称依赖、深度预算与稀疏隐含目标。两人分别复核 4 文件／33 项、5 文件／43 项通过；上述完整检查已覆盖该轮全部修复，PR 仍为草稿。
- 草稿 PR 的 `Quick checks` 在 `85d755f` 成功，运行 `42340514947`；未启动全量 CI。普通提交／推送继续只触发快速检查，不能将其计作阶段验收。
- `8136ccd`、`21526a4` 的远端快速检查均成功；后者运行 `42345431264`，没有启动全量 CI。
- `79a9283` 的远端快速检查成功，运行 `42347498456`，没有启动全量 CI。
- 新增 [81 项独立原生函数来源与 Calc 观察](evidence/workbook-calculation-functions.json)，SDK 对照通过。10 项来源差异按既有逻辑转文本声明、用户日期选择及逐项规范／官方依据保留，不能声称 81 项都与 Calc 相同或 Excel/WPS 实机通过。新增来源的 7 份 XML 部件通过官方 XSD 与关系审计；`check:workbook-calculation` 重新运行固定 Calc 并核对证据，不改基线。
- 新增 [16 项筛选／SUBTOTAL 原生参考](evidence/workbook-filter-calculation.json)，SDK 对照通过。固定 Calc 将一个手动隐藏且符合筛选条件的行在 9／109 中均排除；保留原始观察，按规范区分当前筛选与手动隐藏行，不声称所有值与 Calc 一致。该来源 14 份 XML 部件通过官方 XSD／关系审计；`check:workbook-filter-calculation` 固定生产者复现通过。
- 本地非受限 Node 诊断使用原 10,000×20／200,000 实际格／20,000 公式／40 块：全量约 2.959s；20 次小编辑加增量计算的原报告字段约 362.1ms（该探针以最大项代替 p95，按最近秩应为 358.7ms；原始样本保留），每次计算 2 项、复用 19,998 项，20 步历史保留约 24.74MB。日志和报告位于 `.cache/workbook-calculation-probe.{log,json}`；不是固定资源、浏览器或 100 步历史的验收。
- 更名后的旧 Office 缓存启动脚本仍指向旧目录；已用原固定 SHA-256 归档重建本地工具，使用 `.cache/acceptance-office-deps/installed/usr/lib/x86_64-linux-gnu` 作为进程级 `LD_LIBRARY_PATH`，未变更 SDK 或全局系统安装。
- 输入／计算读取定向单测 5 文件／14 项通过，覆盖 1200 格长链、批次上下文、来源隔离、已释放会话读取、过期输入、跨块粘贴、空文本及一次撤销。新浏览器输入用例已加入日常四引擎文件矩阵；本次本地仅验证 Chromium／Firefox／WebKit，旧 Chromium 待阶段验收。
- 此批完整 `pnpm check` 通过：161 文件／1839 项，类型、边界、五包及 Playground 构建成功，日志 `.cache/workbook-input-check.log`。随后补充的原生草稿状态守卫／图表批次上下文用例经类型、包构建、计算／图表 3 文件／32 项及 3 引擎输入 16 项通过（另 2 项非 Chromium CDP IME 跳过），日志 `.cache/workbook-calculated-recheck.log` 与 `.cache/workbook-native-input-final.log`；对应预览定向 18 项已通过。它们不是完整阶段或容量验收。
- 浏览器定向检查验证原生文字、公式、多行、历史、TSV 注入事件、真实 Chromium IME、并发拒绝、释放及过长草稿恢复；真实系统剪贴板和 Firefox／WebKit 原生 IME 尚不能凭合成事件或 CDP 声称通过。新增 3 引擎长链显示、范围选择、选择保留与已有只读预览回归通过。日志 `.cache/workbook-input-preview-browser.log` 中 Firefox 的过长输入因多发一次原生事件而触发了过严的事件次数断言；按来源不变、错误类型及可恢复行为修正测试后，3 引擎重跑通过，见 `.cache/workbook-oversized-input-browser.log`。
- 本地 WebKit 缓存的 314 个依赖软链接仍指向更名前目录，已改为指向现有依赖的相对链接；没有下载字体、更换浏览器或修改全局系统安装。
- 网格输入接入后完整 `pnpm check` 通过：162 文件／1843 项，类型、边界、包和 Playground 构建成功，日志 `.cache/workbook-grid-input-check.log`。输入／预览的 3 引擎定向 50 项通过，4 项非 Chromium CDP IME 按声明跳过，日志 `.cache/workbook-grid-input-browser.log`；覆盖原生输入、鼠标和工作表切换保护、重复按键、LTR／RTL／合并／隐藏区间、跨块粘贴清除、历史及两个真实 Chromium IME 场景。
- 随后自查复现了中间重名导致公式错误改向，并用失败回归确认：修复前候选把原 `Sheet1!C1`、`Other!C1` 都改为 `Sheet1!C1`；修复后拒绝冲突批次，使用临时名交换能保持原工作表身份并单步撤销。追加修复的事务／计算 3 文件／21 项、包构建及工具类型检查通过，日志 `.cache/workbook-sheet-namespace-recheck.log`；不能将此前完整检查计作该修复的最终阶段检查。
- 填充新增 5 项回归，包含跨块富文本和继承格式、混合公式引用、数列／反向／重叠种子、预算、取消、修订变化、稀疏空白及停止型验证；相关分块／模型／事务／格式／图表 6 文件／71 项通过，日志 `.cache/workbook-fill-regressions.log`。包构建、工具类型及架构边界通过。三引擎输入与预览定向 53 项通过、4 项非 Chromium CDP IME 按声明跳过，包含填充按钮及在进度点触发取消后来源／历史不变，日志 `.cache/workbook-fill-input-preview-browser.log`。这是小型直连路径验证，不是 200,000 格编辑容量或旧 Chromium 验收。
- `abcc53c` 远端快速检查成功，运行 `42349452959`；未启动全量 CI。富文本输入后完整 `pnpm check` 通过：165 文件／1859 项，类型、边界、五包与 Playground 构建成功，日志 `.cache/workbook-rich-text-check.log`。三引擎输入与预览定向 62 项通过、4 项非 Chromium CDP IME 按声明跳过，日志 `.cache/workbook-rich-text-full-browser.log`；新增范围替换、字体跨度、CR／CRLF、IVS、公式样文本、显式类型转换及最大长度转义检查。这些仍不是阶段或容量最终验收。
- 富文本提交 `1c3136e` 已推送，远端快速检查成功，运行 `42351996706`。查找替换及共享数字格式的定向单测 3 文件／21 项通过，新增三引擎界面 6 项通过，日志 `.cache/workbook-find-format-unit.log`、`.cache/workbook-find-browser.log`；期间修复原生 change 事件清空查找结果、替换后刷新丢失原选区，以及显示求值的实例／查询预算不一致。完整 `pnpm check` 166 文件／1867 项、类型、边界与构建通过，日志 `.cache/workbook-find-check.log`。随后补充单项匹配预算和隐藏表交互后，包构建、工具类型、边界与相关单测 4 文件／26 项通过；日志 `.cache/workbook-find-final-{build,types,unit}.log`。三引擎完整输入／网格／预览回归 407 项通过、4 项非 Chromium CDP IME 按声明跳过，包含 96 份三引擎公开原生来源／生产者重开显示及共享尺寸／合并／表格／条件格式回归，日志 `.cache/workbook-find-final-browser.log`。随后手动计算模式下网格使用同批结果的追加修复经包构建、工具类型、查找替换 8 项单测及三引擎查找／隐藏表 9 项复核通过，日志 `.cache/workbook-find-manual-{build,types,unit,browser}.log`。这些不是固定资源四引擎或完整阶段验收。
- `2d910be` 的快速检查发现两份新增文件末尾多空行；工作区差异检查没有覆盖当时的未跟踪新文件。修正提交 `cfaeedc` 的远端快速检查已成功，运行 `42354653086`。提交前必须同时检查暂存差异及整个 main 基线差异，普通推送仍不启动全量 CI。
- 格式／链接及填充 General 的定向事务、验证、输入和查找回归 7 文件／49 项通过，日志 `.cache/workbook-format-link-final.log`；Word／PPT 会话与生成 Schema 3 文件／154 项通过，日志 `.cache/workbook-format-link-shared.log`。3 份原生来源（文字、样式日期、公式名称）的格式修改／保存包重开，加已有原生改名／删除／不可变适配器共 3 文件／11 项通过，日志 `.cache/workbook-format-link-native.log`。包构建、工具类型及架构边界通过；输入、填充、TSV 与富文本的三引擎追加回归 12 项通过，日志 `.cache/workbook-format-link-browser.log`。这些检查不含原生 XLSX 导出；该阶段尚未实施导出。
- `e937f85` 的远端快速检查成功，运行 `42355784146`。选择性复制的新增 6 项单测和相邻输入／填充／查找／格式回归 5 文件／34 项通过；整块验证的覆盖／清除／删除／预算回归相关 6 文件／42 项通过，日志 `.cache/workbook-copy-{unit,regressions,validation-unit}.log`。完整 `pnpm check` 169 文件／1886 项、类型、边界和构建通过，日志 `.cache/workbook-copy-check.log`；追加孤立代理项检查后定向 2 文件／15 项、包构建及工具类型通过，日志 `.cache/workbook-copy-final-{unit,build,types}.log`。原生来源 3 文件／11 项及三引擎输入／填充／富文本 12 项复核通过，日志 `.cache/workbook-copy-{native,browser}.log`。
- Schema 路径优化后，契约／模型／会话／复制 7 文件／422 项通过，新增工作簿值／格式数组首错停止及调用者数据隔离检查；随后完整 `pnpm check` 169 文件／1889 项、类型、边界和构建通过，日志 `.cache/workbook-copy-schema-check.log`。同源生成 Schema 未改变；三引擎输入／TSV／填充／富文本 12 项通过，日志 `.cache/workbook-copy-schema-browser.log`。这些仍不是阶段最终日常 CI 或容量专项。
- 对原有 10,000×20／200,000 实际格、20,000 公式、40 块作本地 Node 重叠复制诊断，向右一列后为 210,000 实际格；一次提交、20,000 公式重算及整批撤销均通过。核心准备／提交从约 41.36s，经整块验证约 16.26s，再经 Schema 路径优化约 6.49s；最后一次生成约 3.57s、重算约 3.35s、撤销约 356ms，进程生命周期峰值 RSS 约 0.84GiB。日志／JSON 依次位于 `.cache/workbook-copy-probe-before-block-validation.*`、`.cache/workbook-copy-probe-after-block-validation.*`、`.cache/workbook-copy-probe.*`。CPU 采样 `.cache/workbook-copy.cpuprofile` 用于定位重复校验开销；不是固定资源、浏览器、取消或历史容量的验收，这是同步基础的诊断；后续异步准备补充取消／进度，证据另列。


- `a0a243e` 的远端快速检查成功，运行 `42359200620`。异步准备新增 14 项测试，连同共享会话／验证／复制／填充等定向回归 7 文件／77 项通过，日志 `.cache/workbook-async-regressions.log`。完整 `pnpm check` 170 文件／1903 项、类型、边界、五包和 Playground 构建通过，日志 `.cache/workbook-async-check.log`；三引擎输入／TSV／查找替换／填充 18 项通过，新增在命令生成完成后、核心处理 256 格时取消的浏览器检查，日志 `.cache/workbook-async-browser.log`。
- 同样的 200,000 实际格／20,000 公式／40 块重叠复制，异步 Node 诊断在写入与第 5 块校验时分别取消，均保持原模型和零步历史；完整准备约 6.20s、同步提交约 435ms、重算约 3.21s、撤销约 346ms，906 次准备检查点，进程峰值 RSS 约 1.01GiB。首次检查点前的输入捕获仍出现约 1.49s 间隔，取消检查的整次调用耗时不是从取消发出算起的延迟。日志／报告 `.cache/workbook-copy-async-probe.{log,json}` 保留真实样本；这不是固定资源、默认 Worker、浏览器或阶段容量验收。


- `5b07a86` 的远端快速检查成功，运行 `42361103875`。Worker 通道完整 `pnpm check` 171 文件／1918 项、类型、边界、五包及 Playground 构建通过，日志 `.cache/workbook-worker-check.log`。随后追加超额已提交回执和过期视图两项检查，定向 3 文件／43 项及工具类型通过，日志 `.cache/workbook-worker-final-{unit,types}.log`。回执超额另有缓存诊断 `.cache/workbook-worker-receipt-probe.log`，确认源实际已撤销且主线程收到修订 2，不能把错误解释为未提交。
- 真实 Worker 三引擎回归 9 项通过，覆盖已有原生图片／图表／规则读取，以及两份原生源的准备、取消、当前值、历史、包重开和断开后宿主继续使用原会话，日志 `.cache/workbook-worker-final-browser.log`。初轮测试宿主误用预览选区 API，修正后复跑；初轮失败日志单独保留 `.cache/workbook-worker-browser-initial.log`。这是 SDK 端口与公开预览集成，尚非默认编辑器／Playground UI。
- Chromium 实际 200,000 格／20,000 公式／40 块的 Worker 编辑集成通过：整表向右一列复制、准备处理 256 格时取消、当前公式结果、复制／种子两次撤销、原始块摘要恢复及包重开均通过，宿主始终只创建一项核心会话。复制准备约 16.16s，782 次命令生成／907 次准备进度，最大控制回复 928,175 字符、窗口回复 778,368 字符；单次报告保存在 `.cache/workbook-worker-capacity-chromium/**/workbook-editing-worker.json`，执行日志 `.cache/workbook-worker-capacity-browser.log`。这里只声称不受限 Chromium 集成，不声称固定资源、反馈 p95、四引擎或 100 步编辑历史门槛。该次测量仍使用完整预览重建；后续增量基础见下。


- `2150842` 的远端快速检查成功，运行 `42363714141`。候选／提交重开时，私有已验证缓冲区且原摘要声明未变的块复用摘要证明；新字节或新声明仍校验 SHA-256，位置／样式变化仍重新解析，字节及累计格数额度独立检查。新增 5 项回归及模型／事务／准备 4 文件／40 项通过，格式原生／端口 2 文件／20 项、原生表编辑及读取快照 2 文件／8 项通过；包构建与工具类型通过，日志 `.cache/workbook-digest-{build,types,unit,native,native-additional}.log`。三引擎原生读取／Worker 编辑 9 项复核通过，日志 `.cache/workbook-digest-browser.log`。本次未将此前完整检查计为摘要修改后的完整阶段检查。
- 相同 200,000 实际格／20,000 公式的 20 次小编辑诊断，按最近秩 `ceil(n×.95)-1` 重新计算双方原始样本，p95 为约 357.93→172.87ms，最大值 403.23→189.22ms；两次都计算 2 项、复用 19,998 项，20 步历史均保留 24,743,329 字节。进程生命周期峰值 RSS 分别约 595.7／707.6MB，不据此宣称内存改善。原探针过于保守的分位数标签、全部样本及单次报告保留在 `.cache/workbook-calculation-{before,after}-digest-reuse.*`；规范化对照为 `.cache/workbook-digest-reuse-comparison.json`。这是不受限 Node 诊断，未代替固定资源浏览器或 100 步验收。

- `48bdb1f` 远端快速检查成功，运行 `42364672639`。后续增量预览基础区分源格与远处计算结果变化，以实际前后提供者／计算批次绑定行高缓存，弱引用避免持有全部历史；可复用测量时只读取变化行，隐藏列仍按未隐藏宽度参与。元数据／CSS／辅助内容／视口／字体变化及复杂或有警告的字体测量保守退回完整路径。独立窗口完成后才发布，失败／取消保留旧画面；默认 UI 仍未切换到 Worker。
- Worker `view` 的 `previous` 只接受本连接仍保留的实际视图；Worker 比较其私有模型／计算，超额提示省略，主线程绑定准确的读取提供者。新增同修订、不同时间上下文及复制／外来视图守卫。定向 4 文件／29 项通过，日志 `.cache/workbook-incremental-unit.log`；三引擎输入／预览／Worker 125 项通过、4 项非 Chromium CDP IME 跳过，日志 `.cache/workbook-incremental-regression.log`。新用例用 6,144 实际格比较增量与完整行高，覆盖隐藏列远处公式变长、无关行不读取、失败保留及忽略取消信号的迟到响应。此处不是固定资源大表验收。
- 复核比较逻辑时补出稳定块路径发生行列／所属表变更的回归：旧窗口被新描述覆盖会漏报原位置。先复现失败后改为各自在原位置读取并分别比较两表；新增行、列、跨表原位及跨表跨行 4 个场景。失败证据 `.cache/workbook-relocated-impact-before.log` 保留；初次完整检查因运行中新增的前两个回归失败，不记为完整通过。修复后完整检查另记，不能引用修复前的完整检查作为最终源验证。
- 增量与块位置修复后的完整 `pnpm check` 174 文件／1940 项、边界、工具／Playground 类型、SDK 与 Playground 构建通过，日志 `.cache/workbook-incremental-final-check.log`。
- 同一构建的 Chromium 增量／原生 Worker／实际容量 3 项复核通过，日志 `.cache/workbook-incremental-capacity.log`；20 万格复制准备约 16.16s，种子准备约 214ms，最大控制／窗口回复约 1,865,821／778,368 字符。复制及取消、两次撤销、40 个原块摘要恢复和包重开均通过；通道影响提示仍有界。此单次复核仍不代表固定资源 p95 或 100 步历史验收。
- 增量预览已提交推送 `94f9997`，远端快速检查 `42370656258` 成功。后续输入命令／远端输入控制器与绑定继续实施：`input`／`paste` 在唯一 Worker 会话生成事务，`readInput` 读取一格可逆拼写；主线程只保存草稿及旧修订，准备中变更／释放回收候选，迟到选择不覆盖草稿，提交发出后等待真实结果。默认编辑器／Playground 界面仍未接入。
- 后续定向 6 文件／53 项通过，日志 `.cache/workbook-remote-binding-unit.log`，包括共享 HTML 的相邻 CR 数字实体和其他非法控制实体、原类型／富文本、远端组合／外部冲突／迟到选择／取消及释放／延迟提交回执。真实 Worker 三引擎输入 7 项通过、2 项非 Chromium CDP IME 跳过，日志 `.cache/workbook-remote-text-browser.log`；宿主始终只创建一项会话，关闭输入绑定后宿主仍可修改。完整检查另记；此处不是默认 GUI、系统剪贴板或固定资源验收。
- 实际 20 万格／2 万公式的 100 次端口输入命令与公开预览、100 次撤销／重做／再撤销通过。原计时器路径和消息任务路径分别记录 `.cache/workbook-ordinary-editing/**/workbook-ordinary-editing.json` 与 `.cache/workbook-ordinary-editing-messages/**/workbook-ordinary-editing.json`；最近秩 p95 约 470.4→247.4ms，最大值约 536.7→283ms，首次画面约 28.10→23.10s，两次历史保留均为 61,116,282 字节，后续读取最远第 23 行，原块与包重开一致。消息队列路径的 Chromium 普通编辑及原生端口 2 项通过，日志 `.cache/workbook-ordinary-editing-messages.log`。两次均为不受限本机诊断，尚非固定资源、整体内存或默认编辑器的验收；随后容量夹具改为调用真正的远端输入控制器／原生绑定并记录源与宿主／Worker bundle 摘要，仍须复跑新夹具。
- 远端输入、共享原生绑定及消息让步队列后的完整 `pnpm check` 177 文件／1962 项、边界、类型、SDK 与 Playground 构建通过，日志 `.cache/workbook-remote-input-final-check.log`；当前三引擎原生输入／直连编辑器／Worker 69 项通过、6 项非 Chromium CDP IME 跳过，日志 `.cache/workbook-remote-input-final-browser.log`。这里不把 `94f9997` 以前的完整检查当作新输入代码的检查。
- 真正远端控制器／原生绑定的 Chromium 100 次编辑及历史、20 万格复制／取消 2 项通过，日志 `.cache/workbook-remote-input-capacity.log`。100 次的端到端 p95 约 246.7ms、最大值约 324ms、首次画面约 23.28s，历史仍保留 61,116,282 字节，撤销／重做／再撤销后输入恢复为 `22`、40 个块摘要恢复；报告 `.cache/workbook-remote-input-capacity/**/workbook-ordinary-editing.json` 记录原生源、Worker 和宿主 bundle SHA-256。此结果仍非固定资源、默认 GUI 或系统剪贴板的完整验收。

- `95d9e43` 的远端快速检查成功，运行 `42373595623`。本批公共 `mountWorkbookEditor` 新增连接重载及 `mountWorkbookRemoteEditor`：公式栏、地址／名称、历史、重算、填充、查找和单项／全部替换使用宿主唯一 Worker 会话。主线程只持草稿和有界视图；显示查找绑定实际视图批次，异步选择核对修订，`settled` 包含输入／命令／导航，释放界面保留宿主连接。提交回执等待中继续键入保留后续草稿且不提前导航；定向单测 3 文件／39 项及三引擎真实延迟回执 3 项通过。
- Playground 已通过该公开入口编辑，缓存保留同一 Worker 会话和历史；草稿及进行中命令阻止文件切换，新候选完成前保留旧画面。新增／缓存／本地打开及“新建 PPT”均须经过草稿检查。隐藏工作簿保留布局环境并跳过新的尺寸更新，避免 Firefox 在 `display:none` iframe 的未完成测量中返回空 CSS；恢复已有当前修订画面不依赖隐藏前的尺寸任务。
- 本批完整 `pnpm check` 177 文件／1966 项、架构边界、类型与构建通过，日志 `.cache/workbook-default-editor-check.log`。随后补充地址异步修订守卫、Playground 隐藏画面修复及浏览器回归，另记包／Playground 构建与类型日志 `.cache/workbook-default-editor-final-{build,types}.log`、`.cache/workbook-hidden-canvas-{build,types}.log`；不能把前述完整检查计作全部后续修复的阶段最终检查。
- 两次三引擎回归均暴露 Firefox 缓存切换问题，原始失败分别保留 `.cache/workbook-worker-editor-regression.log`（86 通过、6 跳过、1 失败）与 `.cache/workbook-default-editor-final-browser.log`（89 通过、6 跳过、1 失败）；后者同时记录布局失败提示及窗口未恢复，不记为通过。独立 Firefox 初轮复测 3 次虽通过，未代替完整回归。仅保留透明画面的中间修复又使内部显式指针事件拦住 Word 提案按钮，已终止该失败本地运行；日志 `.cache/workbook-hidden-canvas-browser.log` 保留。最终将隐藏和候选画面移出交互区域并设 inert，继续保留 CSS 布局环境。
- 修正后 Playground 三引擎 9 项全通过，日志 `.cache/workbook-hidden-canvas-corrected-browser.log`；Chromium 实际一万行打开／Word-PPT 切换和取消导入后重试 2 项通过，日志 `.cache/workbook-default-editor-capacity.log`。后一组是默认公开编辑器的真实打开路径，不是固定资源普通编辑 p95 或阶段最终验收；单次大表用例约 25.5s，不能将整个用例时长当作精确首屏耗时。
- 最终候选界面也采用移出交互区及 inert，避免准备中新画面拦住旧输入；公开重载注释同步。最终构建／工具类型通过，日志 `.cache/workbook-default-gui-published-{build,types}.log`。三引擎 SDK 默认编辑器和 Playground 共 18 项全部通过，日志 `.cache/workbook-default-gui-final-browser.log`，覆盖延迟地址冲突、同一 Worker、当前公式、历史、填充、UTF-16 单项／范围替换、草稿保护及失败恢复。未启动远端全量 CI，阶段最终双人审查与验收仍待全部功能完成。

- `b161098` 已推送，远端快速检查 `42381613553` 成功，未启动远端全量 CI。后续跨表复制定向 4 文件／52 项通过，包含源／目标身份、同坐标、富文本、继承格式、相对及显式引用、名称／结构引用目标上下文、源完成计算与目标停止验证、取消、一次撤销和重做；日志 `.cache/workbook-cross-sheet-regressions.log`。3 份独立原生来源的跨表／跨块复制和包 ZIP 重开，加既有原生格式检查共 2 文件／6 项通过，日志 `.cache/workbook-cross-sheet-native.log`；命名样式定义和原始来源不改写。初次追加验证测试遗漏四个必填提示字段，类型检查先失败，补齐后最终构建、类型和边界全部通过，日志 `.cache/workbook-cross-sheet-final-build.log`、`.cache/workbook-cross-sheet-published-types.log`、`.cache/workbook-cross-sheet-boundaries.log`。
- 真实 Worker 三引擎跨表复制／当前值／取消／历史 3 项通过，每项分别使用两份独立原生来源，日志 `.cache/workbook-cross-sheet-browser.log`。全部仍是一项核心会话，来源和完整历史留在 Worker；复制后预览目标表，撤销两步恢复原始块并包重开。这不是完整剪贴板、原生导出或固定资源阶段验收。

- `04ab466` 的远端快速检查 `42382526674` 成功，未启动远端全量 CI。后续合并事务保留左上角、默认拒绝覆盖内容，显式清除仍经过停止验证、异步取消和历史边界；只扫描相关实际格，不展开稀疏范围。重复／还原的合并集合保持原元数据顺序，部分交叉与表格内合并失败，取消合并可与内部编辑同批执行。直连及 Worker 界面已接入三项命令；完整合并复制模板、结构移动和排序仍待后续完成。
- 合并定向 5 文件／155 项通过，日志 `.cache/workbook-merges-regressions.log`，包括新契约与同源 JSON Schema、跨块／富文本／来源隔离、明确覆盖选择、全表稀疏、无净变化、数据验证、取消及过期候选。初轮定向将计算接口状态写成标量接口字段，修正断言后通过；随后原生检查的可空正文类型先失败，改为保存正文身份与可空节点证明后重新运行完整检查，结果另记。
- 合并及跨表复制后的完整 `pnpm check` 180 文件／1986 项、架构边界、类型、SDK 与 Playground 构建通过，日志 `.cache/workbook-merges-complete-check.log`；新增原生合并还原、经典备注／现代线程正文身份及 ZIP 重开检查均包括在内。三引擎首轮交互为 93 通过／6 项 CDP IME 跳过／3 失败，日志 `.cache/workbook-merges-browser.log`；失败是 Worker 按钮和交互队列重复报告同一次默认合并拒绝，来源与历史未变化。内部按钮改为仅用队列的报告边界，公开方法继续返回可拒绝的 Promise；修复后构建与类型通过，日志 `.cache/workbook-merges-ui-fixed-{build,types}.log`，后续交互复核另记。
- 远端合并拒绝及清除／取消合并／两次撤销的三引擎定向 3 项通过，日志 `.cache/workbook-merges-ui-fixed-browser.log`。随后完整输入／Worker／Playground 三引擎回归 96 通过、6 项非 Chromium CDP IME 按声明跳过，日志 `.cache/workbook-merges-final-browser.log`；未放宽失败断言。这覆盖本批界面错误通知修复，不能替代完整阶段双人审查、最终 HEAD 日常 CI 和固定资源容量门槛。
- 合并事务 `5d2eb2f` 已推送，远端快速检查 `42384970948` 成功。后续新增真实默认编辑器容量探针：同一 Worker 内 200,000 实际格／20,000 公式／40 块，公式栏编辑 C1，屏幕保持 T1；100 次都检查该编辑器自身显示的依赖结果，不再另作验证重算。100 次编辑及实际界面按钮的 100 次撤销／重做／再撤销通过，最终修订 400、输入恢复 `22`、显示结果 `492`、40 个原始块恢复、包重开一致，宿主始终一项核心会话。日志 `.cache/workbook-mounted-capacity.log`，带源／Worker／宿主 bundle 摘要与全部 100 个样本的报告 `.cache/workbook-mounted-capacity/**/workbook-mounted-editing.json`。
- 该次不受限 Chromium 诊断从输入字节解码开始计首屏约 23.323s，单独选择／滚动设置约 128.8ms；编辑到实际公式画面 p95 247.5ms、最大 291.7ms。普通编辑只读取至第 19 行，最大通道回复 530,432 个 UTF-16 单位；100 步历史保留 61,116,282 字节，复用同一 iframe。测试观察只保留当前／准备中的渲染根，不持有全部历史 DOM；场景加入现有 `@scale` 四引擎容量集成，新增 bundle 仅在该场景实际运行时构建，日常检查不增加这项容量负担。工具类型检查通过；SDK 运行时未因该探针改动。
- 默认编辑器容量探针已提交推送为 `8b64e10`。后续 `setWorkbookAutoFilter` 允许设置普通区域定义及以 null 移除，复用现有当前值／小计、验证和历史；不修改源行或手动隐藏／尺寸。定向 6 文件／155 项通过，日志 `.cache/workbook-filter-edit-regressions.log`，包含停止型验证影响、列／区域非法及合并互斥、同源 JSON Schema、唯一 Worker 当前小计与撤销、原生来源移除／恢复与包 ZIP 重开。新增合法非规范化控制 JSON 的无变化事务仍保持原始字节及零历史，实测通过，没有为此修改序列化器。包构建和工具类型通过，日志 `.cache/workbook-filter-edit-final-build.log`、`.cache/workbook-filter-edit-types.log`。这仍是类型化事务基础，人工筛选界面和表内筛选继续实施。
- `8b64e10`、`3bc5307` 的远端快速检查分别为 `42386436498`、`42387742888`，均成功，未启动全量 CI。2026-11-08 继续接入直连／Worker 的冻结选区前行列、取消冻结和两步撤销。公共预览 `sourceOffset` 读取当前已接受画面的实际行列位置及身份，使用已测自动行高、隐藏区间和列宽，不展开空白；界面在提交前拒绝占满当前视口的冻结和过期画面，保留源修订／历史。底层源契约边界不变。
- 冻结界面三引擎 6 项通过，日志 `.cache/workbook-freeze-browser.log`；源坐标查询三引擎 3 项通过，日志 `.cache/workbook-source-offset-browser.log`，覆盖千行隐藏、自动大字体／多行、隐藏列、只返回已发布修订、切表、非法坐标、释放和不修改源。最后冻结、坐标、既有鼠标草稿保护及 Playground 切换共 18 项三引擎定向回归全部通过，日志 `.cache/workbook-freeze-final-browser.log`；构建、最终工具类型及架构边界通过，日志 `.cache/workbook-freeze-guard-build.log`、`.cache/workbook-freeze-published-types.log`、`.cache/workbook-freeze-boundaries.log`。这仍不是阶段最终验收；行列插删／剪切／排序和其他 R3-12～16 工作继续实施。

## 接下来的工作

本批行列基础的定向核心／同源 Schema／原生／Worker 回归 9 文件／197 项通过，日志 `.cache/workbook-axis-regressions.log`；同批逆向的维度／元数据公式曾产生多余历史，原失败保留 `.cache/workbook-axis-noop-before.log`，补充修正与批注分组后单独 15 项通过（`.cache/workbook-axis-annotation-noop.log`）。最终完整 `pnpm check` 184 文件／2017 项、架构边界、类型、SDK 与 Playground 构建通过，日志 `.cache/workbook-axis-complete-check.log`。初轮行列／合并／冻结三引擎定向 18 项通过（`.cache/workbook-axis-ui-browser.log`）；随后直连命令在重新启用输入后再恢复焦点，三引擎完整输入／Worker／Playground 108 项通过、6 项非 Chromium CDP IME 按声明跳过，日志 `.cache/workbook-axis-final-browser.log`。

`pnpm check:workbook-axis` 的四组实际 Calc 插删此前通过，加入 XSD 联合步骤的初轮因系统 pyuno Python 缺少 lxml 失败（`.cache/workbook-axis-native-final-verified.log`），没有把该次计为通过。验收脚本改为独立调用已有 schema 环境，本机复用 `.cache/xlsx-venv`，CI 继续使用自带 lxml 的系统 Python，不安装或修改系统依赖。完整原生命令重跑及官方 XSD／OPC 通过，日志 `.cache/workbook-axis-native-final-rechecked.log`；该环境修复不改 SDK 运行时。独立原生源的所有未删除格、历史缓存、类型、中文变体序列和格式绑定另经 5 项包 ZIP 重开／撤销检查通过，日志 `.cache/workbook-axis-native-retained.log`。未启动远端全量 CI；这些仍不替代完整阶段双人审查及最终 SHA 的日常／固定资源容量验收。

2026-11-08 接入结构操作前增加公式缓存回归，先复现改名重写会把历史数字／文本值清空为 blank（`.cache/workbook-structural-cache-before.log`），再改用保留正文／格式的属性重写：历史缓存标为 stale，原本 unavailable 保持不可用，当前计算仍独立求值。同批改名后改回的初轮检查仍因规范化引号／括号新增修订，失败日志 `.cache/workbook-structural-cache-noop-before.log` 保留；最终只在原语义、其他属性和正文均与基线相同时还原原公式及缓存来源，保留重做且不发事件。最终核心 8 文件／73 项通过，日志 `.cache/workbook-structural-cache-final-core.log`；最终构建、类型和架构边界通过，日志 `.cache/workbook-structural-cache-final-{build,types,boundaries}.log`。原生改名／格式／复制／合并及 Worker 通道／输入 6 文件／49 项通过，日志 `.cache/workbook-structural-cache-final-integration.log`。独立 `04-formulas-names.xlsx` 改名所引用工作表后保留导入的历史值，当前重算、文档包 ZIP 重开及一次撤销均通过；这不是原生 XLSX 输出或完整行列操作验收。

1. 完成 R3-09：验证共享生命周期回归、资源编辑的成功路径与结构映射全面接入；补齐原子失败、历史和冲突边界。工作表元数据及普通区域行列基础已有，继续上下文名称、智能表、绘图及原生备注／线程的行列联动，并完成剪切和排序；合并首行／首列删除已按用户决定以 Microsoft Excel 为准，取得明确微软依据／Excel 实测后修正实现及回归，不能把证据缺口继续列为待用户选择。验收范围见[行列插删](workbook-model.md#行列插删基础r3-12-实施中)。异步准备、专属 Worker 会话及远端输入／IME／TSV 基础已有，默认 Worker 编辑器／Playground 已接入，继续完整结构命令及其界面；候选最终提交仍为同步原子边界。不能将当前直连小型路径或不受限 Node 诊断当作完整大表命令验收。
2. 完成 R3-10 的剩余边界和验收：补足函数参数／类型／错误／1904 系统与 Unicode、数值文本日期边界，以及筛选显示格式与多列／表组合的独立验收；补齐固定资源 2 万公式及普通编辑／100 步历史的完整门槛。现有 `workbook-scale` 的固定资源脚本仍只覆盖读取，新增默认编辑器 `@scale` 是容量集成；还须实现并运行默认编辑器的四引擎固定 CPU／内存／500ms 普通编辑 p95 验收，不能凭本机 247.5ms 或既有读取绿灯替代。现有函数／依赖调度及筛选小计基础不代表整个 R3-10 退出条件通过。
3. 完成 R3-11 的复制剪切、选择性粘贴、填充剩余边界、富文本格式、超链接和完整外观编辑；跨表名称／未限定结构引用还须独立生产者实际复制对照，不能用 SDK 自身的目标上下文断言代替原生复制行为证据。查找替换已有命令和直连界面，仍需阶段容量与 Worker 取消验收。再推进 R3-12～16 的结构、规则表格、绘图、批注、工作台协作，保持一个编辑计算阶段 PR。现有直连网格／公式栏与键盘导航不是完整退出条件，仍须固定资源和容量验收。来源差异与增量失效基础已接入直连预览和 Worker 视图，继续默认 Worker 界面的大表时延验证；复杂字体／样式仍保留完整测量路径，不能凭小型输入回归声称 200,000 格编辑时延已通过。
4. 阶段完成后、正式提 PR 前使用两位子智能体独立审查整个基线差异，逐项修复与复核，再将已有草稿转为正式 PR，在最终 HEAD 发起日常 CI＋容量专项并以 SHA 守卫合并。

Playground 原服务为 5174；5173 属于另一个项目，勿操作。读取的滚动准备优化与全部已定门槛保持。

2026-11-07 用户问及分支 ahead：远端 `feat/workbook-reading` 为 ahead 161／behind 5，来自旧分支与 squash 合并后主线不同的提交身份；其最终 `d4f250e` 与 PR #25 合并提交 `c1d8dea` 文件树完全一致，没有遗漏读取代码。当前 `feat/workbook-editing` 从 main `5451736` 开始，推送基础提交时为 ahead 1／behind 0；不要把旧分支的提交数当作当前阶段新增工作或重复合并依据。

随后已获取并核对全部 18 个非 main 远端分支：17 个历史分支均已纳入 main，其中 13 个分支的最终文件树与已进入 main 的 squash 合并提交相同，另 4 个分支头直接为 main 祖先；唯一未合并的是当前 `feat/workbook-editing`（当时 ahead 2／behind 0）。无需再次合并或删除历史分支。只读审计记录位于本地 `.cache/branch-integration-audit.json`。
