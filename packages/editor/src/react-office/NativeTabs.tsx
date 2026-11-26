import type { ReactElement, ReactNode } from "react";
import type { DocumentTextCommand, EditorCommandState } from "../index.js";
import type { OfficeActions } from "./OfficeRibbon.js";
import { Group, Tool } from "./ribbon-controls.js";
import type { Glyph } from "./ribbon-controls.js";
import { OfficeMenu } from "./OfficeMenu.js";
import { OfficeIcon as Icon } from "./icons.js";

type Item = readonly [label: string, icon: Glyph, command?: string];
type Column = Item | { readonly stack: readonly Item[] } | { readonly gallery: readonly string[]; readonly category?: "theme" | "transition" | "pen" };
type Section = readonly [title: string, ...columns: Column[]];
const stack = (...items: Item[]): Column => ({ stack: items });
const gallery = (names: readonly string[], category: "theme" | "transition" | "pen" = "theme"): Column => ({ gallery: names, category });

/** Recorded macOS Office command order. Unsupported commands remain disabled. */
const word: Readonly<Record<string, readonly Section[]>> = {
  插入: [
    ["页面", ["封面", "file"], stack(["空白页", "file"], ["分页符", "pageBreak", "w:pageBreak"])],
    ["表格", ["表格", "table", "w:insertTable"]],
    ["插图", ["图片", "image", "w:insertImage"], ["形状", "shapes", "w:insertShape"], ["图标", "shapes"], ["3D 模型", "shapes"], ["SmartArt", "diagram"], ["图表", "chart", "w:insertChart"], ["屏幕截图", "image"]],
    ["媒体", ["媒体", "slides"]],
    ["链接", ["链接", "link", "w:insertLink"], stack(["书签", "bookmark", "w:bookmarks"], ["交叉引用", "file", "w:crossReference"])],
    ["批注", ["批注", "comment", "newComment"]],
    ["页眉和页脚", ["页眉", "header", "w:pageSettings"], ["页脚", "footer", "w:pageSettings"], ["页码", "numbering", "w:pageNumber"]],
    ["文本", ["文本框", "textBox", "w:insertTextBox"], ["艺术字", "textBox"], ["首字下沉", "textBox"], stack(["域", "file", "w:referenceSettings"], ["日期和时间", "file"], ["对象", "file"])],
    ["符号", ["公式", "textBox", "w:insertEquation"], stack(["高级符号", "textBox"], ["数字", "numbering"])],
  ],
  布局: [
    ["页面设置", ["文本方向", "textDirection", "w:pageSettings"], ["页边距", "margins", "w:pageSettings"], ["方向", "orientation", "w:pageSettings"], ["大小", "file", "w:pageSettings"], ["栏", "columns", "w:pageSettings"], stack(["分隔符", "pageBreak", "w:sectionBreak"], ["行号", "numbering"], ["断字", "file"])],
    ["缩进和间距", stack(["左侧缩进", "outdent", "paragraph"], ["右侧缩进", "indent", "paragraph"]), stack(["段前间距", "lineSpacing", "paragraph"], ["段后间距", "lineSpacing", "paragraph"])],
    ["排列", ["位置", "grid", "w:drawingSettings"], ["自动换行", "lineSpacing", "w:drawingSettings"], stack(["前移一层", "arrow", "w:drawingSettings"], ["后移一层", "arrow", "w:drawingSettings"], ["选择窗格", "panel", "w:selectionPane"]), stack(["对齐", "left", "w:drawingSettings"], ["组合", "grid"], ["旋转", "refresh", "w:drawingSettings"])],
  ],
  引用: [
    ["目录", ["目录", "list", "w:referenceSettings"], stack(["添加文字", "textBox", "w:referenceSettings"], ["更新目录", "refresh", "fields"])],
    ["脚注", ["插入脚注", "file", "w:insertFootnote"], ["插入尾注", "file", "w:insertEndnote"], stack(["下一条脚注", "arrow", "w:referenceSettings"], ["显示备注", "file", "w:referenceSettings"])],
    ["引文与书目", ["插入引文", "book"], stack(["引文", "book"], ["样式", "file"], ["书目", "book"])],
    ["题注", ["插入题注", "image", "w:caption"], stack(["插入表目录", "list"], ["更新表格", "refresh"], ["交叉引用", "file", "w:crossReference"])],
    ["索引", ["标记条目", "bookmark"], stack(["插入索引", "list"], ["更新索引", "refresh"])],
    ["引文目录", ["标记引文", "bookmark"], stack(["插入引文目录", "list"], ["更新表格", "refresh"])],
  ],
  设计: [["文档格式", ["主题", "palette"], gallery(["文档", "独特", "典雅型", "流行", "正式", "手稿", "现代", "新闻纸", "透视", "简单"]), stack(["颜色", "palette"], ["字体", "textBox"], ["段落间距", "lineSpacing", "paragraph"]), ["设为默认值", "check"]], ["页面背景", ["水印", "image", "w:watermark"], ["页面颜色", "fill", "w:pageSettings"], ["页面边框", "table", "borders"]]],
  邮件: [["创建", ["信封", "file"], ["标签", "file"]], ["开始邮件合并", ["开始邮件合并", "file"], ["选择收件人", "book"], stack(["编辑收件人列表", "edit"], ["筛选收件人", "list"])], ["编写和插入域", ["插入合并域", "textBox"], stack(["规则", "list"], ["更新标签", "refresh"])], ["预览结果", ["预览结果", "book"], stack(["第一条记录", "arrow"], ["上一条记录", "arrow"], ["查找收件人", "search"])], ["完成", ["完成并合并", "check"]]],
  审阅: [["校对", ["编辑器", "editor"], stack(["拼写和语法", "proof"], ["同义词库", "book"], ["字数统计", "numbering", "w:wordCount"])], ["语音", ["大声朗读", "book"]], ["辅助功能", ["检查辅助功能", "proof"]], ["语言", ["翻译", "globe"], ["语言", "globe", "w:fontSettings"]], ["批注", ["新建批注", "comment", "newComment"], stack(["删除", "close"], ["解决", "check"], ["上一条批注", "arrow"]), stack(["下一条批注", "arrow"], ["显示批注", "comment", "comments"])], ["修订", ["修订", "edit", "w:review"], ["审阅", "file", "w:review"]], ["更改", ["接受", "check", "w:review"], ["拒绝", "close", "w:review"], stack(["上一条修订", "arrow", "w:review"], ["下一条修订", "arrow", "w:review"])], ["比较", ["比较", "file"]], ["保护", ["保护", "shield"]], ["墨迹", ["隐藏墨迹", "pen"]], ["协作", ["载入固定示例提案", "sparkles", "proposal"]]],
  视图: [["视图", ["页面视图", "file", "view:pages"], ["Web 版式", "globe", "view:continuous"], stack(["大纲", "list", "w:outline"], ["草稿", "file"])], ["沉浸式", ["专注", "file"], ["沉浸式阅读器", "book"]], ["显示", stack(["标尺", "margins"], ["网格线", "grid"], ["导航窗格", "panel", "w:outline"])], ["缩放", ["缩放", "search", "zoom"], ["100%", "search", "zoom100"], stack(["单页", "file", "view:pages"], ["多页", "file", "view:pages"], ["页宽", "file", "view:pages"])], ["窗口", ["新建窗口", "file"], ["全部重排", "grid"], ["拆分", "columns"], ["切换窗口", "file"]], ["宏", ["宏", "file"]]],
};
const ppt: Readonly<Record<string, readonly Section[]>> = {
  插入: [["幻灯片", ["新建幻灯片", "newSlide", "newSlide"]], ["表格", ["表格", "table", "insertTable"]], ["图像", ["图片", "image", "insertImage"], ["屏幕截图", "image"]], ["Cameo", ["Cameo", "image"]], ["插图", ["形状", "shapes", "insertShape"], ["图标", "shapes"], ["3D 模型", "shapes"], ["SmartArt", "diagram"], ["图表", "chart", "insertChart"]], ["链接", ["缩放定位", "slides"], stack(["链接", "link"], ["动作", "arrow"])], ["批注", ["批注", "comment", "newComment"]], ["文本", ["文本框", "textBox", "insertTextBox"], ["页眉和页脚", "header", "definitions"], ["艺术字", "textBox"], stack(["日期和时间", "file"], ["幻灯片编号", "numbering", "definitions"], ["对象", "file"])], ["符号", ["公式", "textBox"], ["符号", "textBox"]], ["媒体", ["视频", "slides"], ["音频", "mic"]]],
  设计: [["主题", gallery(["Office 主题", "Helena", "拱道", "Modulo", "社论", "红利", "障子", "沙丘"])], ["变体", ["变体", "palette", "definitions"], stack(["颜色", "palette", "definitions"], ["字体", "textBox", "definitions"], ["背景样式", "fill", "slideBackground"])], ["自定义", ["版式", "grid", "slideLayout"], ["幻灯片大小", "orientation", "slideSize"]], ["设计建议", ["设计建议", "sparkles"]]],
  切换: [["预览", ["预览", "slides"]], ["切换到此幻灯片", gallery(["无", "平滑", "淡入/淡出", "推入", "擦除", "分割", "显示"], "transition"), ["效果选项", "arrow"]], ["计时", stack(["持续时间", "file"], ["声音", "mic"], ["应用到所有", "slides"]), stack(["单击鼠标时", "cursor"], ["自动换片间隔", "file"])]],
  动画: [["预览", ["预览", "slides"]], ["进入效果", gallery(["出现", "淡入", "飞入", "浮入"], "transition")], ["强调效果", gallery(["脉冲", "放大/缩小", "旋转"], "transition")], ["动画", ["退出效果", "arrow"], ["路径动画", "arrow"], stack(["效果选项", "arrow"], ["动画窗格", "panel"], ["触发器", "cursor"]), ["动画刷", "brush"]], ["计时", stack(["开始", "file"], ["持续时间", "file"])]] ,
  幻灯片放映: [["开始幻灯片放映", ["从头播放", "slides", "slideShowFromStart"], ["从当前幻灯片播放", "slides", "slideShowFromCurrent"], ["演示者视图", "slides"], ["自定义放映", "slides"]], ["设置", ["设置幻灯片放映", "slides"], ["隐藏幻灯片", "slides"], ["排练计时", "file"], ["录制", "mic"], stack(["播放旁白", "mic"], ["使用计时", "file"], ["显示媒体控件", "slides"])], ["字幕", stack(["始终使用字幕", "textBox"], ["字幕设置", "textBox"])]] ,
  录制: [["Cameo", ["Cameo", "image"]], ["录制", ["从头开始", "mic"], ["从当前幻灯片开始", "mic"]], ["录制选项", ["清除录制", "eraser"], ["重置为 Cameo", "image"]]],
  审阅: [["校对", ["拼写", "proof"], ["同义词库", "book"]], ["辅助功能", ["检查辅助功能", "proof"]], ["语言", ["翻译", "globe"], ["语言", "globe", "fontSettings"]], ["更改", stack(["全部标记为已读", "check"], ["显示更改", "edit"])], ["批注", ["新建批注", "comment", "newComment"], stack(["删除", "close"], ["上一条", "arrow"], ["下一条", "arrow"]), ["显示批注", "comment", "comments"]], ["保护", ["始终以只读方式打开", "shield"], ["限制权限", "shield"]], ["墨迹", ["隐藏墨迹", "pen"]], ["协作", ["载入固定示例提案", "sparkles", "proposal"]]],
  视图: [["演示文稿视图", ["普通", "slides", "preview"], ["大纲视图", "list"], ["幻灯片浏览", "grid"], ["备注页", "file", "editNotes"], ["阅读视图", "book", "slideShowFromCurrent"]], ["母版视图", ["幻灯片母版", "slides", "definitions"], ["讲义母版", "file"], ["备注母版", "file", "definitions"]], ["显示", stack(["标尺", "margins"], ["网格线", "grid"], ["参考线", "grid"]), ["备注", "file", "editNotes"]], ["缩放", ["缩放", "search", "zoom"], ["适应窗口大小", "slides", "fitSlide"]], ["宏", ["宏", "file"]]],
};
const excel: Readonly<Record<string, readonly Section[]>> = {
  插入: [["表格", ["数据透视表", "table"], ["建议的数据透视表", "table"], ["表格", "table", "tables"]], ["Forms", ["Forms", "file"]], ["图片数据", ["来自图片", "image"]], ["插图", ["图片", "image", "drawings"], ["形状", "shapes", "drawings"], stack(["图标", "shapes"], ["3D 模型", "shapes"], ["SmartArt", "diagram"]), ["屏幕截图", "image"]], ["控件", ["复选框", "check"]], ["图表", ["建议的图表", "chart", "drawings"], stack(["柱形图", "chart", "drawings"], ["折线图", "chart", "drawings"], ["饼图", "chart", "drawings"]), stack(["层次结构", "chart"], ["统计", "chart", "drawings"], ["X Y (散点图)", "chart", "drawings"]), stack(["瀑布图", "chart"], ["组合图", "chart", "drawings"], ["地图", "globe"])], ["迷你图", ["迷你图", "chart"]], ["筛选器", ["切片器", "grid"], ["日程表", "grid"]], ["链接", ["链接", "link", "hyperlinks"]], ["批注", ["新建批注", "comment", "comments"]], ["文本", ["文本框", "textBox", "drawings"], stack(["页眉和页脚", "header"], ["艺术字", "textBox"], ["对象", "file"])], ["符号", ["公式", "textBox"], ["符号", "textBox"]]],
  页面布局: [["主题", ["主题", "palette"], stack(["颜色", "palette"], ["字体", "textBox"])], ["页面设置", ["页边距", "margins", "printSettings"], ["方向", "orientation", "printSettings"], ["大小", "file", "printSettings"], ["打印区域", "grid", "printSettings"], stack(["分隔符", "pageBreak"], ["背景", "image"], ["打印标题", "header", "printSettings"]), ["页面设置", "file", "printSettings"]], ["调整为合适大小", stack(["宽度", "orientation"], ["高度", "orientation"])], ["工作表选项", stack(["网格线", "grid"], ["标题", "header"])]] ,
  公式: [["函数库", ["插入函数", "textBox", "insertFunction"], ["自动求和", "numbering", "autoSum"], stack(["最近使用", "file"], ["财务", "numbering"], ["逻辑", "numbering"]), stack(["文本", "textBox"], ["日期和时间", "file"], ["查找和引用", "search"]), stack(["数学和三角函数", "numbering"], ["其他函数", "numbering"])], ["定义的名称", ["名称管理器", "bookmark", "names"], stack(["定义名称", "bookmark", "names"], ["在公式中使用", "bookmark", "names"], ["根据所选内容创建", "bookmark", "names"])], ["公式审核", stack(["追踪引用单元格", "arrow"], ["追踪从属单元格", "arrow"], ["删除箭头", "arrow"]), stack(["显示公式", "textBox"], ["错误检查", "proof"], ["监视窗口", "panel"])], ["计算", ["计算选项", "file"], stack(["立即计算", "refresh", "recalculate"], ["计算工作表", "refresh", "recalculate"])]] ,
  数据: [["获取和转换数据", ["获取数据 (Power Query)", "file"], ["来自图片", "image"]], ["查询和连接", ["全部刷新", "refresh"], stack(["查询和连接", "link"], ["属性", "file"], ["工作簿链接", "link"])], ["数据类型", gallery(["股票", "货币", "地理"])], ["排序和筛选", stack(["升序", "sort", "sortAscending"], ["降序", "sort", "sortDescending"]), ["排序", "sort", "sortSettings"], ["筛选", "list", "filters"], stack(["清除", "eraser", "filters"], ["重新应用", "refresh", "filters"], ["高级", "list", "filters"])], ["数据工具", ["分列", "columns"], ["分析数据", "chart"], ["快速填充", "table", "fillSeries"], ["删除重复项", "table"], ["数据验证", "check", "rules"], ["合并计算", "table"]], ["预测", ["模拟分析", "chart"]], ["分级显示", ["组合", "grid", "worksheets"], ["取消组合", "grid", "worksheets"], ["分类汇总", "table"]]],
  审阅: [["校对", ["拼写", "proof"], stack(["同义词库", "book"], ["工作簿统计信息", "numbering"])], ["辅助功能", ["检查辅助功能", "proof"]], ["语言", ["翻译", "globe"]], ["更改", ["显示更改", "edit"]], ["批注", ["新建批注", "comment", "comments"], stack(["删除", "close"], ["上一条批注", "arrow"], ["下一条批注", "arrow"]), ["显示批注", "comment", "comments"]], ["注释", ["注释", "file", "comments"]], ["保护", ["保护工作表", "shield"], ["保护工作簿", "shield"], ["始终以只读方式打开", "shield"]], ["墨迹", ["隐藏墨迹", "pen"]]],
  视图: [["工作表视图", ["默认", "file"], stack(["保留", "save"], ["退出", "close"], ["新建", "plus"])], ["工作簿视图", ["普通", "grid", "preview"], ["分页预览", "file"], ["页面布局", "file"], ["自定义视图", "file"]], ["显示", ["导航", "panel", "worksheets"], stack(["标尺", "margins"], ["网格线", "grid"], ["编辑栏", "textBox"]), stack(["标题", "header"], ["数据类型图标", "grid"], ["焦点单元格", "cursor"])], ["缩放", ["缩放", "search", "zoom"], ["100%", "search", "zoom100"], ["缩放到所选比例", "search"]], ["窗口", ["新建窗口", "file"], ["全部重排", "grid"], ["冻结窗格", "grid", "freezeMenu"], ["拆分", "columns"], stack(["隐藏", "file", "worksheets"], ["取消隐藏", "file", "worksheets"], ["切换窗口", "file"])], ["宏", stack(["查看宏", "file"], ["录制宏", "mic"], ["使用相对引用", "grid"])]] ,
  自动化: [["Office 脚本", ["新脚本", "file"], ["查看脚本", "file"], gallery(["取消隐藏所有行和列", "冻结所选内容", "从所选内容中创建子表", "空行计数"])]] ,
};

export function NativeOfficeTabs({ kind, tab, commands, actions, canFormat, busy, onCommand, onTextCommand, onLayout, onZoom, paragraph, borders, layout, zoom, enableWorkbookZoom }: {
  kind: "document" | "presentation" | "workbook"; tab: string; commands: readonly EditorCommandState[]; actions: OfficeActions; canFormat: boolean; busy: boolean;
  onCommand(id: string): void; onTextCommand(command: DocumentTextCommand): void | Promise<void>; onLayout(value: "pages" | "continuous"): void; onZoom(value: number): void;
  paragraph(): void; borders(): void;
  layout: "pages" | "continuous"; zoom: number; enableWorkbookZoom: boolean;
}): ReactElement {
  const states = new Map(commands.map(command => [command.id, command]));
  const invoke = (key: string): (() => void) | undefined => {
    if (key.startsWith("w:")) return canFormat || kind === "document" && !busy && ["w:find", "w:outline", "w:wordCount"].includes(key) ? () => onTextCommand({ type: "word", action: key.slice(2) } as DocumentTextCommand) : undefined;
    if (key === "view:pages" || key === "view:continuous") return !busy ? () => onLayout(key.slice(5) as "pages" | "continuous") : undefined;
    if (key === "zoom100") return !busy && (kind !== "workbook" || enableWorkbookZoom) ? () => onZoom(1) : undefined;
    if (key === "paragraph") return canFormat ? paragraph : undefined;
    if (key === "borders") return canFormat ? borders : undefined;
    if (states.has(key)) return states.get(key)!.enabled && !busy ? () => onCommand(key) : undefined;
    return actions[key] && !actions[key]!.disabled && !busy ? actions[key]!.run : undefined;
  };
  const drawItem = (item: Item, large = true, key?: number): ReactNode => {
    const [label, icon, command] = item;
    if (command === "freezeMenu") return <OfficeMenu key={key} label={label} icon={<Icon name={icon} size={32} />} large disabled={!states.get("freezePanes")?.enabled}>
      {close => <>{[["freezePanes", "冻结窗格"], ["unfreezePanes", "取消冻结窗格"]].map(([id, label]) => <button key={id} role="menuitem" disabled={!invoke(id!)} onClick={() => { invoke(id!)?.(); close(); }}>{label}</button>)}</>}
    </OfficeMenu>;
    if (command === "zoom") return <OfficeMenu key={key} label={label} icon={<Icon name={icon} size={32} />} large disabled={busy || kind === "workbook" && !enableWorkbookZoom}>
      {close => <>{[.5, .75, 1, 1.25, 1.5, 2].map(value => <button key={value} role="menuitemradio" aria-checked={zoom === value} onClick={() => { onZoom(value); close(); }}>{value * 100}%</button>)}</>}
    </OfficeMenu>;
    return <Tool key={key} label={label} icon={icon} large={large} active={command?.startsWith("view:") ? layout === command.slice(5) : undefined} run={command ? invoke(command) : undefined} />;
  };
  if (tab === "绘图") return <><Group title="工具"><Tool label="绘图" icon="pen" large /><Tool label="橡皮擦" icon="eraser" large />{kind !== "document" && <Tool label="套索选择" icon="cursor" large />}</Group><Group title="笔"><div className="native-pen-gallery">{["#161616", "#d21b23", "#888888", "#ffff00", "#e9cf73"].map((color, index) => <button key={color} type="button" disabled aria-label={["黑色笔", "红色笔", "铅笔", "黄色荧光笔", "荧光笔"][index]} style={{ color }}><Icon name="pen" size={37} /></button>)}</div><Tool label="添加" icon="plus" /></Group>{kind === "presentation" && <Group title="转换"><Tool label="墨迹转换为文本" icon="textBox" large /><Tool label="将墨迹转换为形状" icon="shapes" large /><Tool label="将墨迹转换为数学公式" icon="numbering" large /></Group>}<Group title="触控板"><Tool label="使用触控板绘图" icon="pen" large /></Group></>;
  const sections = (kind === "document" ? word : kind === "presentation" ? ppt : excel)[tab] ?? [];
  return <>{sections.map(([title, ...columns]) => <Group key={title} title={title}>{columns.map((column, index) => {
    if ("stack" in column) return <div key={index} className="ribbon-tool-stack">{column.stack.map((item, index) => drawItem(item, false, index))}</div>;
    if ("gallery" in column) return <div key={index} className={"native-command-gallery " + (column.category ?? "theme")}>{column.gallery.map((label, index) => <button key={label} disabled type="button" title={label} aria-label={label}><span style={{ borderTopColor: ["#4472c4", "#b7472a", "#107c41", "#767171"][index % 4] }}>{column.category === "transition" ? <Icon name="slides" size={26} /> : <><b>Aa</b><i /></>}</span><small>{label}</small></button>)}</div>;
    return drawItem(column as Item, true, index);
  })}</Group>)}{tab === "审阅" && actions.objectSelection && <Group title="选择"><OfficeMenu label="选择" icon={<Icon name="cursor" size={32} />} large disabled={busy}>
    {close => <>{[["textSelection", "文字选择"], ["objectSelection", "对象选择"], ["regionSelection", "矩形框选"]].map(([id, label]) => <button role="menuitem" key={id} disabled={!invoke(id!)} onClick={() => { invoke(id!)?.(); close(); }}>{label}</button>)}</>}
  </OfficeMenu></Group>}</>;
}
