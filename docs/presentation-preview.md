# PPT 只读画布与读取验收

R2-06 为 Presentation 文档包接入 `@doffi/render` 的 `mountDocumentPreview`，不需要编辑会话。浏览器只消费已校验的 HTML/CSS、类型化元数据和授权资源；原 PPTX、第三方演示文稿／图表库和 Office 进程不参与 SDK 运行。

```ts
import { openDocument } from "@doffi/core";
import { importPptx } from "@doffi/office";
import { mountDocumentPreview } from "@doffi/render";

const model = openDocument(importPptx(bytes, { documentId: "host-authorized-id" }));
const view = mountDocumentPreview(container, model);
await view.ready;
const report = await view.settled();
// report.pages 的 pageId/sectionId 对应稳定的 slideId，尺寸与矩形单位为 pt。
// 字体替代、缺字、文字溢出是当前视图的诊断，不改写文档源。
view.dispose();
```

幻灯片始终使用固定页面；continuous 只适用于 Word。默认按照源节点顺序显示所有页，包括有标记的隐藏页。母版、布局、备注母版和批注正文保留各自的源根；母版／布局的可见实例经过投影组合，不能保存为幻灯片的新对象。备注作为逐页可展开面板，批注显示两种格式的正文、作者、时间、状态和回复；这些面板不进入打印画布。

## 映射与边界

| 内容 | 当前映射与验收 |
| --- | --- |
| 页面与继承 | 原始 pt 尺寸；主题、颜色映射、字体、母版／布局占位符及本地覆盖；背景、日期缓存与当前顺序派生页码 |
| 对象与组 | 源几何生成绝对定位；父子坐标、负坐标、旋转、翻转及非均匀缩放；文字保持字形方向和来源字号 |
| 预设形状 | 声明的 16 种预设以及直线／折线连接线；标准调整值、文字区域及连接点；SVG 仅为派生绘制 |
| 外观 | 实色、线性渐变、颜色变换、透明度、轮廓、虚线、箭头和基础外阴影 |
| 文字 | 四种字体角色、UTF-16、变体序列、列表、软换行、超链接外观及有界普通／形状自动适应；不触发外部链接 |
| 图片与表格 | 授权 PNG/JPEG/WebP、矩形裁剪、旋转／翻转；原生网格、合并、边框、填充与文字 |
| 图表 | 二维簇状柱、简单折线、单系列饼图，从唯一常量数据生成 SVG；标题、图例、轴、空值、系列／数据点及基础标签 |
| 备注与批注 | 逐页备注和经典／现代批注的只读正文、富文本回复、状态；不将外部显示作者作为权限 |
| 来源与生命周期 | 图形、字体替代、字段显示、布局和资源 URL 可逆投影；更新失败恢复旧视图；替换、失败挂载及卸载释放私有资源 |

文字 body 属性中的 autoFit、wrap、边距及锚定缺省值保留“未声明”，在实例中解析继承。显式新颜色重设其颜色变换缺省值，避免继承父对象透明度。绑定连接线在所有对象就绪后解析原生连接点；缓存端点已经一致时保留其预设和调整路径，不一致时采用直线或正交路径，源缓存不被静默修改。正交重新布线不承诺与所有 Office 应用采用同一路由算法。

非空段落按实际文字片段的最大有效字号建立行框，避免母版的大字号把小字号正文行距撑高；相对字号先解析再进行视图调整，源样式及撤销历史不变。极小组子坐标先转换为数值稳定的派生坐标系，避免图片裁剪和 SVG 视口在放大前舍入为零；源几何保持原样。零高／零宽的线条使用非零绘制视口，路径位置和线宽不改变。页码和打印标签遵循 `firstSlideNumber`，页面顺序索引仍从 1 开始。

macOS 的自定义 `▪` 项目符号优先使用本机可用的 Apple Color Emoji 字形，以匹配本次 PowerPoint 对照；其他环境回退到可用的声明字体。标记字符和原生颜色声明仍保存在源中。空白字符不以“没有墨迹”判成缺字；未指定的预览默认字体也不产生文档指定字体被替代的误报。

`layout()` 绑定 documentId/revision，返回稳定幻灯片身份、源节点到页面的映射、当前页码显示、字体报告与溢出诊断。来源字段缓存不随只读显示改变。`printDocument()` 只投影每张实际画布，包含动态 CSSOM 继承规则；PDF/PNG 的 Node 分派和 Playground 入口由 [R2-15](pptx-export.md#node-和-playground) 提供。

## 标准数据来源

预设数学数据取自 [ECMA-376 Part 1，第五版（2016）官方附件](https://www.ecma-international.org/wp-content/uploads/ECMA-376-1_5th_edition_december_2016.zip)中的 `OfficeOpenXML-DrawingMLGeometries.zip/presetShapeDefinitions.xml`。仅提取声明预设的公式、路径、文字区域和连接点；求值器、SVG 编译器和图表实现均由本项目编写，未复制其他 Office 编辑器代码。

- 附件 ZIP SHA-256：`a8517838a3a60492d267d1ff92e1ebd831adc04584de79a38a5532a68902abf6`。
- XML SHA-256：`2f7c868d857c1e3c4b5a6068759fe0e07d77ad58377a6618d1b02ba3507b6939`。
- 该附件缺少 upArrow；本项目从 downArrow 数学反射生成。
- [提取器](../scripts/generate-presentation-geometries.py)固定输入摘要，可用 `python3 scripts/generate-presentation-geometries.py /path/to/presetShapeDefinitions.xml --check` 校验生成文件；普通构建不下载附件。
- [Ecma 版权与许可](../packages/core/ECMA-376-GEOMETRY-LICENSE.txt)随 core 打包，生成数据中的 `@license` 也保留在浏览器 bundle 中。

## 验收与已知差异

[独立 PPTX 样例](../examples/pptx/README.md)中的 16 份来源文件和 16 份 LibreOffice 重存文件逐份在 Chromium、Firefox、WebKit 导入、保存重开并显示。检查固定画布、源包不变、投影、打印排除备注／批注、组中文字、连接端点、图表常量、精确 UTF-16 选择、限额、更新失败恢复和 URL 清理。三种浏览器的基础 126 项验收和旧 Chromium 109 的基础 42 项验收通过；另有平台字形栅格不可用、备注展开、备注母版和现代批注主题字体／列表的回归。当前三引擎共 138 项。真实 Safari 的相关烟测进入 CI，其实际结果由 macOS job 记录。

Chromium 固定字体截图覆盖形状、组及三类图表，沿用 `maxDiffPixelRatio=0.005`、`threshold=0.2`。独立原生 PDF 的 PNG 是 1281×721，浏览器画布是 1280×720；[参考检查](../scripts/check-presentation-reference.py)比较共同视口中的 21 处可见形状／图表边界，容差为 3～4 px，不重采样文字，不把跨渲染器像素相等作为门槛。截图和核对输入摘要保存在 `.cache/presentation-preview/`，CI 留存产物。

边界检查使用系统 Python/Pillow，仅为验收工具，不进入 SDK 或打包依赖。本地版本为 Pillow 11.3.0（Debian/Ubuntu 包 11.3.0-1ubuntu1.2），[该版本许可为 MIT-CMU](https://github.com/python-pillow/Pillow/blob/11.3.0/LICENSE)。检查结果记录实际 Python/Pillow 版本及对应许可来源；CI 使用其系统软件包版本。

浏览器和 LibreOffice 的字距、基线及段落度量仍有差异；字体替代、缺字及溢出逐项报告。图表自动布局采用本项目确定性布局，不承诺复刻各应用的标题和图例自动排版。数字格式限 General、普通整数／小数、千分位、百分比及基本正负分段；超出范围明确失败。普通自动适应在布局预算内缩小字形和行距；形状自动适应只调整派生画面尺寸，源几何保持不变。更广泛的样式组合仍需独立样例，不把当前矩阵等同于任意 PPTX 都能打开。

PowerPoint/WPS 未实机验证；LibreOffice 重存丢失现代线程的事实沿用独立参考记录。SDK 保存与源格式结构检查不替代真实现代批注应用验收。[R2-07 核心事务](presentation-transactions.md)、[R2-08～13 编辑协作](presentation-editing.md)及 [R2-14～16 原生导出和独立闭环](pptx-export.md)已实现。渲染报告通过 `kind: "presentation"` 与 `slides` 给出明确幻灯片语义，`pages` 保留现有公共画布投影。

## 推送后审查

首次实现提交推送后，单独对基准差异进行审查，新增用例分别复现了相对行距随本地字号丢失、HTML 文字未应用透明度，以及主题样式引用未合成引用颜色透明度三项问题。修复分别保留相对行距、对文字应用颜色变换，并显式区分直接颜色和主题格式引用。三个回归在 Chromium、Firefox、WebKit 分别复核。真实 Safari 初次验收的字形栅格异常已转为明确的未核验诊断，修复后真实 Safari 通过；备注展开时自动更新页坐标和 iframe 高度，面板保留原生复制行为；备注正文解析备注母版的字体、颜色和背景，现代批注正文解析所属页的主题字体并显示富文本列表。编辑仍保持关闭，阶段合并须以当前 PR 提交的完整 CI 为准。
