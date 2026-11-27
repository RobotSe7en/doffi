import { useEffect, useState } from "react";
import type { ReactElement, ReactNode } from "react";
import type { DocumentTextCommand, EditorCommandState, TextFormatState, WordPageZoomMode } from "../index.js";
import type { OfficeActions } from "./OfficeRibbon.js";
import { Group, Tool } from "./ribbon-controls.js";
import type { Glyph } from "./ribbon-controls.js";
import { OfficeColorMenu, OfficeMenu } from "./OfficeMenu.js";
import { OfficeIcon as Icon } from "./icons.js";
import {
  WORD_COLOR_SCHEMES,
  WORD_FONT_SCHEMES,
  WORD_STYLE_SETS,
  WORD_THEMES,
} from "../word-theme-presets.js";

type WordShapePreset = Extract<DocumentTextCommand, { type: "word"; action: "quickShape" }>["preset"];
const WORD_SHAPE_GROUPS: readonly {
  readonly heading: string;
  readonly items: readonly (readonly [preset: WordShapePreset, label: string, svgPath: string])[];
}[] = [
  {
    heading: "矩形",
    items: [
      ["rect", "矩形", "M3 5H21V19H3Z"],
      ["roundRect", "圆角矩形", "M6 5H18Q21 5 21 8V16Q21 19 18 19H6Q3 19 3 16V8Q3 5 6 5Z"],
    ],
  },
  {
    heading: "基本形状",
    items: [
      ["ellipse", "椭圆", "M12 4A9 8 0 1 0 12 20A9 8 0 1 0 12 4Z"],
      ["triangle", "等腰三角形", "M12 4L21 19H3Z"],
      ["diamond", "菱形", "M12 3L21 12L12 21L3 12Z"],
      ["parallelogram", "平行四边形", "M6 5H21L18 19H3Z"],
      ["trapezoid", "梯形", "M6 5H18L21 19H3Z"],
      ["pentagon", "五边形", "M12 3L21 10L17.5 20H6.5L3 10Z"],
      ["hexagon", "六边形", "M7 4H17L21 12L17 20H7L3 12Z"],
    ],
  },
  {
    heading: "箭头总汇",
    items: [
      ["rightArrow", "右箭头", "M3 9H13V5L21 12L13 19V15H3Z"],
      ["leftArrow", "左箭头", "M21 9H11V5L3 12L11 19V15H21Z"],
      ["upArrow", "上箭头", "M9 21V11H5L12 3L19 11H15V21Z"],
      ["downArrow", "下箭头", "M9 3V13H5L12 21L19 13H15V3Z"],
      ["leftRightArrow", "左右箭头", "M3 12L8 7V10H16V7L21 12L16 17V14H8V17Z"],
      ["chevron", "燕尾形", "M3 5H16L21 12L16 19H3L8 12Z"],
    ],
  },
  {
    heading: "线条",
    items: [["line", "直线", "M4 20L20 4"]],
  },
];

function parsePtValue(raw: string | undefined): number {
  if (!raw) return 0;
  const trimmed = raw.trim();
  const match = /^(-?\d+(?:\.\d+)?)(pt|px|in|cm|mm)?$/i.exec(trimmed);
  if (!match) return 0;
  const num = Number(match[1]);
  const unit = (match[2] ?? "pt").toLowerCase();
  if (unit === "px") return Math.round(num * 0.75 * 10) / 10;
  if (unit === "in") return Math.round(num * 72 * 10) / 10;
  if (unit === "cm") return Math.round(((num * 72) / 2.54) * 10) / 10;
  if (unit === "mm") return Math.round(((num * 72) / 25.4) * 10) / 10;
  return Math.round(num * 10) / 10;
}

function RibbonNumberField({
  label,
  valuePt,
  disabled,
  min = 0,
  max = 360,
  step = 6,
  onCommit,
}: {
  label: string;
  valuePt: number;
  disabled: boolean;
  min?: number;
  max?: number;
  step?: number;
  onCommit(nextPt: number): void;
}): ReactElement {
  const [draft, setDraft] = useState(`${valuePt} 磅`);
  const [focused, setFocused] = useState(false);
  useEffect(() => {
    if (!focused) setDraft(`${valuePt} 磅`);
  }, [valuePt, focused]);
  const apply = (raw: string) => {
    const parsed = parseFloat(raw.replace(/[^\d.-]/g, ""));
    if (!Number.isFinite(parsed)) {
      setDraft(`${valuePt} 磅`);
      return;
    }
    const clamped = Math.max(min, Math.min(max, Math.round(parsed * 10) / 10));
    setDraft(`${clamped} 磅`);
    if (Math.abs(clamped - valuePt) > 0.01) onCommit(clamped);
  };
  return (
    <label className="ribbon-spin-field">
      <span>{label}</span>
      <span className="ribbon-spin-box">
        <input
          type="text"
          aria-label={label}
          disabled={disabled}
          value={draft}
          onFocus={(e) => {
            setFocused(true);
            e.currentTarget.select();
          }}
          onChange={(e) => setDraft(e.currentTarget.value)}
          onBlur={() => {
            setFocused(false);
            apply(draft);
          }}
          onKeyDown={(e) => {
            if (e.key === "Enter") {
              e.preventDefault();
              apply(draft);
              e.currentTarget.blur();
            } else if (e.key === "Escape") {
              e.preventDefault();
              setDraft(`${valuePt} 磅`);
              e.currentTarget.blur();
            } else if (e.key === "ArrowUp") {
              e.preventDefault();
              onCommit(Math.min(max, Math.round((valuePt + step) * 10) / 10));
            } else if (e.key === "ArrowDown") {
              e.preventDefault();
              onCommit(Math.max(min, Math.round((valuePt - step) * 10) / 10));
            }
          }}
        />
        <span className="ribbon-spin-arrows" aria-hidden="true">
          <button
            type="button"
            tabIndex={-1}
            disabled={disabled}
            onMouseDown={(e) => e.preventDefault()}
            onClick={() => onCommit(Math.min(max, Math.round((valuePt + step) * 10) / 10))}
          >
            ▲
          </button>
          <button
            type="button"
            tabIndex={-1}
            disabled={disabled}
            onMouseDown={(e) => e.preventDefault()}
            onClick={() => onCommit(Math.max(min, Math.round((valuePt - step) * 10) / 10))}
          >
            ▼
          </button>
        </span>
      </span>
    </label>
  );
}

type Item = readonly [label: string, icon: Glyph, command?: string];
type Column =
  | Item
  | { readonly stack: readonly Item[] }
  | { readonly gallery: readonly string[]; readonly category?: "theme" | "transition" | "pen" }
  | { readonly wordIndentSpacing: true };
type Section = readonly [title: string, ...columns: Column[]];
const stack = (...items: Item[]): Column => ({ stack: items });
const gallery = (names: readonly string[], category: "theme" | "transition" | "pen" = "theme"): Column => ({ gallery: names, category });

/** Recorded macOS Office command order. Unsupported commands remain disabled. */
const word: Readonly<Record<string, readonly Section[]>> = {
  插入: [
    ["页面", ["封面", "file", "w:coverPage"], stack(["空白页", "file", "w:insertBlankPage"], ["分页符", "pageBreak", "w:pageBreak"])],
    ["表格", ["表格", "table", "wordTableMenu"]],
    ["插图", ["图片", "image", "w:insertImage"], ["形状", "shapes", "wordShapeMenu"], ["图标", "shapes"], ["3D 模型", "shapes"], ["SmartArt", "diagram"], ["图表", "chart", "w:insertChart"], ["屏幕截图", "image"]],
    ["媒体", ["媒体", "slides"]],
    ["链接", ["链接", "link", "w:insertLink"], stack(["书签", "bookmark", "w:bookmarks"], ["交叉引用", "file", "w:crossReference"])],
    ["批注", ["批注", "comment", "newComment"]],
    ["页眉和页脚", ["页眉", "header", "w:editHeader"], ["页脚", "footer", "w:editFooter"], ["页码", "numbering", "wordPageNumberMenu"]],
    ["文本", ["文本框", "textBox", "w:insertTextBox"], ["文档部件", "file", "w:quickParts"], ["艺术字", "textEffects", "w:wordArt"], ["首字下沉", "textBox", "w:dropCap"], stack(["域", "file", "w:insertField"], ["日期和时间", "file", "w:insertDateTime"], ["文档属性", "file", "w:docProperties"]), ["内容控件", "edit", "wordControlMenu"]],
    ["符号", ["公式", "textBox", "w:insertEquation"], stack(["高级符号", "textBox", "w:symbols"], ["拼音指南", "pinyin", "w:ruby"])],
  ],
  布局: [
    ["页面设置", ["文本方向", "textDirection", "w:pageSettings"], ["页边距", "margins", "wordMarginsMenu"], ["方向", "orientation", "wordOrientationMenu"], ["大小", "file", "wordPageSizeMenu"], ["栏", "columns", "wordColumnsMenu"], stack(["分隔符", "pageBreak", "wordBreaksMenu"], ["行号", "numbering", "wordLineNumbersMenu"], ["断字", "file", "wordHyphenationMenu"])],
    ["缩进和间距", { wordIndentSpacing: true }],
    ["排列", ["位置", "grid", "w:drawingSettings"], ["自动换行", "lineSpacing", "wordWrapMenu"], stack(["前移一层", "arrow", "w:bringForward"], ["后移一层", "arrow", "w:sendBackward"], ["选择窗格", "panel", "w:selectionPane"]), stack(["对齐", "left", "w:drawingSettings"], ["组合", "grid", "w:groupDrawings"], ["旋转", "refresh", "wordRotateMenu"])],
  ],
  引用: [
    ["目录", ["目录", "list", "w:toc"], stack(["添加文字", "textBox", "wordAddTextMenu"], ["更新目录", "refresh", "fields"])],
    ["脚注", ["插入脚注", "file", "w:insertFootnote"], ["插入尾注", "file", "w:insertEndnote"], stack(["下一条脚注", "arrow", "w:nextFootnote"], ["显示备注", "file", "w:showNotes"])],
    ["引文与书目", ["插入引文", "book", "w:insertCitation"], stack(["引文", "book", "w:citations"], ["样式", "file", "wordBibStyleMenu"], ["书目", "book", "w:bibliography"])],
    ["题注", ["插入题注", "image", "w:caption"], stack(["插入表目录", "list", "w:insertTof"], ["更新表格", "refresh", "fields"], ["交叉引用", "file", "w:crossReference"])],
    ["索引", ["标记条目", "bookmark", "w:markIndexEntry"], stack(["插入索引", "list", "w:insertIndex"], ["更新索引", "refresh", "fields"])],
    ["引文目录", ["标记引文", "bookmark", "w:markToaCitation"], stack(["插入引文目录", "list", "w:insertToa"], ["更新表格", "refresh", "fields"])],
  ],
  设计: [["文档格式", ["主题", "palette", "wordThemeMenu"], gallery(WORD_STYLE_SETS.map(item => item.label)), stack(["颜色", "palette", "wordColorSchemeMenu"], ["字体", "textBox", "wordFontSchemeMenu"], ["段落间距", "lineSpacing", "wordParaSpacingMenu"]), stack(["设为默认值", "check", "w:saveDefaultTemplate"], ["文档模板", "file", "w:templateSettings"])], ["页面背景", ["水印", "image", "w:watermark"], ["页面颜色", "fill", "wordPageColorMenu"], ["页面边框", "table", "w:pageBorders"]]],
  邮件: [["创建", ["信封", "file"], ["标签", "file"]], ["开始邮件合并", ["开始邮件合并", "file"], ["选择收件人", "book"], stack(["编辑收件人列表", "edit"], ["筛选收件人", "list"])], ["编写和插入域", ["插入合并域", "textBox"], stack(["规则", "list"], ["更新标签", "refresh"])], ["预览结果", ["预览结果", "book"], stack(["第一条记录", "arrow"], ["上一条记录", "arrow"], ["查找收件人", "search"])], ["完成", ["完成并合并", "check"]]],
  审阅: [["校对", ["编辑器", "editor"], stack(["拼写和语法", "proof"], ["同义词库", "book"], ["字数统计", "numbering", "w:wordCount"])], ["语音", ["大声朗读", "book"]], ["辅助功能", ["检查辅助功能", "proof"]], ["语言", ["翻译", "globe"], ["语言", "globe", "w:fontSettings"]], ["批注", ["新建批注", "comment", "newComment"], stack(["删除", "close", "deleteComment"], ["解决", "check", "resolveComment"], ["上一条批注", "arrow", "prevComment"]), stack(["下一条批注", "arrow", "nextComment"], ["显示批注", "comment", "comments"])], ["修订", ["修订", "edit", "w:toggleTrackChanges"], ["审阅", "file", "wordReviewDisplayMenu"]], ["更改", ["接受", "check", "wordAcceptMenu"], ["拒绝", "close", "wordRejectMenu"], stack(["上一条修订", "arrow", "w:prevRevision"], ["下一条修订", "arrow", "w:nextRevision"])], ["比较", ["比较", "file", "compareDocx"]], ["保护", ["保护", "shield", "w:protection"]], ["墨迹", ["隐藏墨迹", "pen"]], ["协作", ["载入固定示例提案", "sparkles", "proposal"]]],
  视图: [["视图", ["页面视图", "file", "view:pages"], ["Web 版式", "globe", "view:continuous"], stack(["大纲", "list", "w:outlineView"], ["草稿", "file", "view:draft"])], ["沉浸式", ["专注", "file", "focusMode"], ["沉浸式阅读器", "book", "immersiveReader"]], ["显示", stack(["网格线", "grid", "gridlines"], ["导航窗格", "panel", "w:outline"])], ["缩放", ["缩放", "search", "zoom"], ["100%", "search", "zoom100"], stack(["单页", "file", "zoom:singlePage"], ["多页", "file", "zoom:multiplePages"], ["页宽", "file", "zoom:pageWidth"])], ["窗口", ["拆分", "columns", "splitWindow"]]],
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

export function NativeOfficeTabs({ kind, tab, commands, actions, canFormat, busy, format, onCommand, onTextCommand, onLayout, onZoom, paragraph, borders, layout, zoom, enableWorkbookZoom, pageZoomMode, onPageZoomMode, draftMode, onDraftMode, focusMode, onFocusMode, showGridlines, onToggleGridlines, immersiveReader, onToggleImmersiveReader, splitWindow, onToggleSplitWindow }: {
  kind: "document" | "presentation" | "workbook"; tab: string; commands: readonly EditorCommandState[]; actions: OfficeActions; canFormat: boolean; busy: boolean;
  format?: TextFormatState | undefined;
  onCommand(id: string): void; onTextCommand(command: DocumentTextCommand): void | Promise<void>; onLayout(value: "pages" | "continuous"): void; onZoom(value: number): void;
  paragraph(): void; borders(): void;
  layout: "pages" | "continuous"; zoom: number; enableWorkbookZoom: boolean;
  pageZoomMode?: WordPageZoomMode | undefined; onPageZoomMode?: ((mode: WordPageZoomMode) => void) | undefined;
  draftMode?: boolean | undefined; onDraftMode?: ((enabled: boolean) => void) | undefined;
  focusMode?: boolean | undefined; onFocusMode?: ((enabled: boolean) => void) | undefined;
  showGridlines?: boolean | undefined; onToggleGridlines?: (() => void) | undefined;
  immersiveReader?: boolean | undefined; onToggleImmersiveReader?: (() => void) | undefined;
  splitWindow?: boolean | undefined; onToggleSplitWindow?: (() => void) | undefined;
}): ReactElement {
  const [tableHover, setTableHover] = useState<readonly [number, number]>([0, 0]);
  const states = new Map(commands.map(command => [command.id, command]));
  const invoke = (key: string): (() => void) | undefined => {
    if (key.startsWith("w:")) return canFormat || kind === "document" && !busy && ["w:find", "w:goTo", "w:outline", "w:wordCount", "w:protection"].includes(key) ? () => onTextCommand({ type: "word", action: key.slice(2) } as DocumentTextCommand) : undefined;
    if (key === "view:pages" || key === "view:continuous") return !busy ? () => { onDraftMode?.(false); onLayout(key.slice(5) as "pages" | "continuous"); } : undefined;
    if (key === "view:draft") return !busy && onDraftMode ? () => onDraftMode(true) : undefined;
    if (key === "focusMode") return !busy && onFocusMode ? () => onFocusMode(!focusMode) : undefined;
    if (key === "gridlines") return !busy && onToggleGridlines ? onToggleGridlines : undefined;
    if (key === "immersiveReader") return !busy && onToggleImmersiveReader ? onToggleImmersiveReader : undefined;
    if (key === "splitWindow") return !busy && onToggleSplitWindow ? onToggleSplitWindow : undefined;
    if (key === "zoom:singlePage" || key === "zoom:multiplePages" || key === "zoom:pageWidth")
      return !busy && onPageZoomMode ? () => onPageZoomMode(key.slice(5) as WordPageZoomMode) : undefined;
    if (key === "zoom100") return !busy && (kind !== "workbook" || enableWorkbookZoom) ? () => onZoom(1) : undefined;
    if (key === "paragraph") return canFormat ? paragraph : undefined;
    if (key === "borders") return canFormat ? borders : undefined;
    if (states.has(key)) return states.get(key)!.enabled && !busy ? () => onCommand(key) : undefined;
    return actions[key] && !actions[key]!.disabled && !busy ? actions[key]!.run : undefined;
  };
  const drawItem = (item: Item, large = true, key?: number): ReactNode => {
    const [label, icon, command] = item;
    if (command === "wordTableMenu") return <OfficeMenu key={key} label={label} icon={<Icon name={icon} size={32} />} large disabled={!canFormat}>
      {close => <div onPointerLeave={() => setTableHover([0, 0])}>
        <div className="office-table-grid-wrap">
          <div className="office-table-grid-title">{tableHover[0] > 0 ? `插入 ${tableHover[0]} × ${tableHover[1]} 表格` : "插入表格"}</div>
          <div className="office-table-grid" role="grid" aria-label="快速插入表格网格">
            {Array.from({ length: 64 }, (_, idx) => {
              const r = Math.floor(idx / 8) + 1, c = (idx % 8) + 1;
              const active = r <= tableHover[0] && c <= tableHover[1];
              return <button key={idx} type="button" role="menuitem" aria-label={`插入 ${r} 行 ${c} 列表格`} data-active={active}
                onPointerEnter={() => setTableHover([r, c])}
                onClick={() => { onTextCommand({ type: "word", action: "insertTableGrid", rows: r, columns: c }); setTableHover([0, 0]); close(); }} />;
            })}
          </div>
        </div>
        <div className="office-menu-divider" />
        <button role="menuitem" type="button" onClick={() => { onTextCommand({ type: "word", action: "insertTable" }); close(); }}>插入表格…</button>
        <button role="menuitem" type="button" onClick={() => { onTextCommand({ type: "word", action: "setTableTool", mode: format?.wordState?.tableToolMode === "draw" ? "off" : "draw" }); close(); }}>绘制表格</button>
        <button role="menuitem" type="button" onClick={() => { onTextCommand({ type: "word", action: "textToTable" }); close(); }}>将文字转换成表格…</button>
        <button role="menuitem" type="button" onClick={() => { onTextCommand({ type: "word", action: "tableToText" }); close(); }}>表格转换成文本…</button>
      </div>}
    </OfficeMenu>;
    if (command === "wordShapeMenu") return <OfficeMenu key={key} label={label} icon={<Icon name={icon} size={32} />} large disabled={!canFormat}>
      {close => <>
        {WORD_SHAPE_GROUPS.map(group => (
          <div key={group.heading}>
            <div className="office-menu-heading">{group.heading}</div>
            <div className="office-shape-grid">
              {group.items.map(([preset, title, d]) => (
                <button key={preset} role="menuitem" type="button" title={title} aria-label={title}
                  onClick={() => { onTextCommand({ type: "word", action: "quickShape", preset }); close(); }}>
                  <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.7" aria-hidden="true">
                    <path d={d} fill={preset === "line" ? "none" : "#edf3fc"} />
                  </svg>
                </button>
              ))}
            </div>
          </div>
        ))}
        <div className="office-menu-divider" />
        <button role="menuitem" type="button" onClick={() => { onTextCommand({ type: "word", action: "insertShape" }); close(); }}>自定义形状选项…</button>
      </>}
    </OfficeMenu>;
    if (command === "wordPageNumberMenu") return <OfficeMenu key={key} label={label} icon={<Icon name={icon} size={32} />} large disabled={!canFormat}>
      {close => <>
        <button role="menuitem" type="button" onClick={() => { onTextCommand({ type: "word", action: "insertPageNumberAt", location: "header" }); close(); }}>页面顶端（居中）</button>
        <button role="menuitem" type="button" onClick={() => { onTextCommand({ type: "word", action: "insertPageNumberAt", location: "footer" }); close(); }}>页面底端（居中）</button>
        <button role="menuitem" type="button" onClick={() => { onTextCommand({ type: "word", action: "insertPageNumberAt", location: "inline" }); close(); }}>当前位置插入页码</button>
        <div className="office-menu-divider" />
        <button role="menuitem" type="button" onClick={() => { onTextCommand({ type: "word", action: "pageSettings" }); close(); }}>设置页码格式…</button>
      </>}
    </OfficeMenu>;
    if (command === "wordParaSpacingMenu") return <OfficeMenu key={key} label={label} icon={<Icon name={icon} size={17} />} disabled={!canFormat}>
      {close => <div onPointerLeave={() => onTextCommand({ type: "word", action: "livePreview", properties: null })}>
        <div className="office-menu-heading">内置段落间距</div>
        {([
          ["默认", "0pt", "8pt", "1.15", "段后 8 磅 · 1.15 倍"],
          ["无段落间距", "0pt", "0pt", "1", "段后 0 磅 · 单倍"],
          ["紧凑", "0pt", "4pt", "1", "段后 4 磅 · 单倍"],
          ["紧密", "0pt", "6pt", "1.15", "段后 6 磅 · 1.15 倍"],
          ["开放", "0pt", "10pt", "1.15", "段后 10 磅 · 1.15 倍"],
          ["宽松", "0pt", "6pt", "1.5", "段后 6 磅 · 1.5 倍"],
          ["双倍", "0pt", "8pt", "2", "段后 8 磅 · 双倍"],
        ] as const).map(([name, top, bottom, lineHeight, desc]) => (
          <button key={name} role="menuitem" type="button"
            onPointerEnter={() => onTextCommand({ type: "word", action: "livePreview", scope: "theme", properties: { "line-height": lineHeight } })}
            onClick={() => {
              onTextCommand({ type: "word", action: "livePreview", properties: null });
              onTextCommand({ type: "word", action: "paragraph", properties: { "margin-top": top, "margin-bottom": bottom, "line-height": lineHeight, "--word-line-rule": "auto" } });
              close();
            }}>{name}<span className="office-menu-sub">{desc}</span></button>
        ))}
        <div className="office-menu-divider" />
        <button role="menuitem" type="button" onClick={() => { onTextCommand({ type: "word", action: "livePreview", properties: null }); paragraph(); close(); }}>自定义段落间距…</button>
      </div>}
    </OfficeMenu>;
    if (command === "wordMarginsMenu") return <OfficeMenu key={key} label={label} icon={<Icon name={icon} size={32} />} large disabled={!canFormat}>
      {close => <>
        {([
          ["常规", 72, 90, 72, 90, "上下 2.54cm · 左右 3.18cm"],
          ["窄", 36, 36, 36, 36, "上下左右 1.27cm"],
          ["适中", 72, 54, 72, 54, "上下 2.54cm · 左右 1.91cm"],
          ["宽", 72, 144, 72, 144, "上下 2.54cm · 左右 5.08cm"],
        ] as const).map(([name, top, right, bottom, left, desc]) => (
          <button key={name} role="menuitem" type="button"
            onClick={() => { onTextCommand({ type: "word", action: "setPageMargins", top, right, bottom, left }); close(); }}>{name}<span className="office-menu-sub">{desc}</span></button>
        ))}
        <div className="office-menu-divider" />
        <button role="menuitem" type="button" onClick={() => { onTextCommand({ type: "word", action: "pageSettings" }); close(); }}>自定义边距…</button>
      </>}
    </OfficeMenu>;
    if (command === "wordOrientationMenu") return <OfficeMenu key={key} label={label} icon={<Icon name={icon} size={32} />} large disabled={!canFormat}>
      {close => <>
        <button role="menuitemradio" type="button" aria-checked={(format?.wordState?.pageOrientation ?? "portrait") === "portrait"}
          onClick={() => { onTextCommand({ type: "word", action: "setPageOrientation", orientation: "portrait" }); close(); }}>纵向</button>
        <button role="menuitemradio" type="button" aria-checked={format?.wordState?.pageOrientation === "landscape"}
          onClick={() => { onTextCommand({ type: "word", action: "setPageOrientation", orientation: "landscape" }); close(); }}>横向</button>
      </>}
    </OfficeMenu>;
    if (command === "wordPageSizeMenu") return <OfficeMenu key={key} label={label} icon={<Icon name={icon} size={32} />} large disabled={!canFormat}>
      {close => <>
        {([
          ["A4", 595.3, 841.9, "21 × 29.7 cm"],
          ["A3", 841.9, 1190.55, "29.7 × 42 cm"],
          ["A5", 419.55, 595.3, "14.8 × 21 cm"],
          ["信纸 (Letter)", 612, 792, "21.59 × 27.94 cm"],
          ["法律专用纸 (Legal)", 612, 1008, "21.59 × 35.56 cm"],
          ["B5 (JIS)", 515.9, 728.5, "18.2 × 25.7 cm"],
        ] as const).map(([name, width, height, desc]) => (
          <button key={name} role="menuitem" type="button"
            onClick={() => { onTextCommand({ type: "word", action: "setPageSize", width, height }); close(); }}>{name}<span className="office-menu-sub">{desc}</span></button>
        ))}
        <div className="office-menu-divider" />
        <button role="menuitem" type="button" onClick={() => { onTextCommand({ type: "word", action: "pageSettings" }); close(); }}>其他页面大小…</button>
      </>}
    </OfficeMenu>;
    if (command === "wordColumnsMenu") return <OfficeMenu key={key} label={label} icon={<Icon name={icon} size={32} />} large disabled={!canFormat}>
      {close => <>
        {([
          [1, "一栏"],
          [2, "两栏"],
          [3, "三栏"],
        ] as const).map(([count, text]) => (
          <button key={count} role="menuitemradio" type="button" aria-checked={(format?.wordState?.columnCount ?? 1) === count}
            onClick={() => { onTextCommand({ type: "word", action: "setPageColumns", count }); close(); }}>{text}</button>
        ))}
        <div className="office-menu-divider" />
        <button role="menuitem" type="button" onClick={() => { onTextCommand({ type: "word", action: "pageSettings" }); close(); }}>更多栏…</button>
      </>}
    </OfficeMenu>;
    if (command === "wordBreaksMenu") return <OfficeMenu key={key} label={label} icon={<Icon name={icon} size={17} />} disabled={!canFormat}>
      {close => <>
        <div className="office-menu-heading">分页符</div>
        <button role="menuitem" type="button" onClick={() => { onTextCommand({ type: "word", action: "pageBreak" }); close(); }}>分页符</button>
        <div className="office-menu-heading">分节符</div>
        {([
          ["nextPage", "下一页"],
          ["continuous", "连续"],
          ["evenPage", "偶数页"],
          ["oddPage", "奇数页"],
        ] as const).map(([breakType, text]) => (
          <button key={breakType} role="menuitem" type="button"
            onClick={() => { onTextCommand({ type: "word", action: "insertSectionBreak", breakType }); close(); }}>{text}</button>
        ))}
      </>}
    </OfficeMenu>;
    if (command === "wordWrapMenu") return <OfficeMenu key={key} label={label} icon={<Icon name={icon} size={32} />} large disabled={!canFormat}>
      {close => <>
        <button role="menuitem" type="button" onClick={() => { onTextCommand({ type: "word", action: "setDrawingWrap", mode: "inline" }); close(); }}>嵌入型</button>
        <button role="menuitem" type="button" onClick={() => { onTextCommand({ type: "word", action: "setDrawingWrap", mode: "floating", wrap: "square", behindText: false }); close(); }}>四周型环绕</button>
        <button role="menuitem" type="button" onClick={() => { onTextCommand({ type: "word", action: "setDrawingWrap", mode: "floating", wrap: "tight", behindText: false }); close(); }}>紧密型环绕</button>
        <button role="menuitem" type="button" onClick={() => { onTextCommand({ type: "word", action: "setDrawingWrap", mode: "floating", wrap: "through", behindText: false }); close(); }}>穿越型环绕</button>
        <button role="menuitem" type="button" onClick={() => { onTextCommand({ type: "word", action: "setDrawingWrap", mode: "floating", wrap: "topAndBottom", behindText: false }); close(); }}>上下型环绕</button>
        <button role="menuitem" type="button" onClick={() => { onTextCommand({ type: "word", action: "setDrawingWrap", mode: "floating", wrap: "none", behindText: true }); close(); }}>衬于文字下方</button>
        <button role="menuitem" type="button" onClick={() => { onTextCommand({ type: "word", action: "setDrawingWrap", mode: "floating", wrap: "none", behindText: false }); close(); }}>浮于文字上方</button>
        <div className="office-menu-divider" />
        <button role="menuitem" type="button" onClick={() => { onTextCommand({ type: "word", action: "drawingSettings" }); close(); }}>其他布局选项…</button>
      </>}
    </OfficeMenu>;
    if (command === "wordRotateMenu") return <OfficeMenu key={key} label={label} icon={<Icon name={icon} size={17} />} disabled={!canFormat}>
      {close => <>
        <button role="menuitem" type="button" onClick={() => { onTextCommand({ type: "word", action: "transformDrawing", mode: "rotateRight" }); close(); }}>向右旋转 90°</button>
        <button role="menuitem" type="button" onClick={() => { onTextCommand({ type: "word", action: "transformDrawing", mode: "rotateLeft" }); close(); }}>向左旋转 90°</button>
        <button role="menuitem" type="button" onClick={() => { onTextCommand({ type: "word", action: "transformDrawing", mode: "flipV" }); close(); }}>垂直翻转</button>
        <button role="menuitem" type="button" onClick={() => { onTextCommand({ type: "word", action: "transformDrawing", mode: "flipH" }); close(); }}>水平翻转</button>
        <div className="office-menu-divider" />
        <button role="menuitem" type="button" onClick={() => { onTextCommand({ type: "word", action: "drawingSettings" }); close(); }}>其他旋转选项…</button>
      </>}
    </OfficeMenu>;
    if (command === "wordBibStyleMenu") return <OfficeMenu key={key} label={label} icon={<Icon name={icon} size={17} />} disabled={!canFormat}>
      {close => <>{([
        ["APA", "APA (第 7 版)"],
        ["MLA", "MLA (第 9 版)"],
        ["GB7714", "GB/T 7714-2015"],
        ["Chicago", "Chicago (第 17 版)"],
        ["IEEE", "IEEE"],
      ] as const).map(([style, text]) => (
        <button key={style} role="menuitemradio" type="button" aria-checked={(format?.wordState?.bibliographyStyle ?? "APA") === style}
          onClick={() => { onTextCommand({ type: "word", action: "setBibliographyStyle", style }); close(); }}>{text}</button>
      ))}</>}
    </OfficeMenu>;
    if (command === "wordControlMenu") return <OfficeMenu key={key} label={label} icon={<Icon name={icon} size={32} />} large disabled={!canFormat} split={() => onTextCommand({ type: "word", action: "insertControl" })}>
      {close => <>
        <button role="menuitem" type="button" onClick={() => { onTextCommand({ type: "word", action: "insertControl" }); close(); }}>插入内容控件…</button>
        <button role="menuitem" type="button" onClick={() => { onTextCommand({ type: "word", action: "controlSettings" }); close(); }}>内容控件属性…</button>
      </>}
    </OfficeMenu>;
    if (command === "wordThemeMenu") return <OfficeMenu key={key} label={label} icon={<Icon name={icon} size={32} />} large disabled={!canFormat}>
      {close => <div onPointerLeave={() => onTextCommand({ type: "word", action: "livePreview", properties: null })}>{WORD_THEMES.map(preset => (
        <button key={preset.id} role="menuitemradio" type="button" aria-checked={format?.wordState?.themeName === preset.theme.name}
          onPointerEnter={() => onTextCommand({ type: "word", action: "livePreview", scope: "theme", properties: {}, ...(preset.theme.colors ? { themeColors: preset.theme.colors } : {}) })}
          onClick={() => { onTextCommand({ type: "word", action: "livePreview", properties: null }); onTextCommand({ type: "word", action: "setTheme", themeId: preset.id }); close(); }}>{preset.label}</button>
      ))}</div>}
    </OfficeMenu>;
    if (command === "wordColorSchemeMenu") return <OfficeMenu key={key} label={label} icon={<Icon name={icon} size={17} />} disabled={!canFormat}>
      {close => <div onPointerLeave={() => onTextCommand({ type: "word", action: "livePreview", properties: null })}>{WORD_COLOR_SCHEMES.map(preset => (
        <button key={preset.id} role="menuitem" type="button"
          onPointerEnter={() => onTextCommand({ type: "word", action: "livePreview", scope: "theme", properties: {}, themeColors: preset.colors })}
          onClick={() => { onTextCommand({ type: "word", action: "livePreview", properties: null }); onTextCommand({ type: "word", action: "setColorScheme", schemeId: preset.id }); close(); }}>
          <span style={{ display: "inline-flex", gap: 2, marginRight: 8 }}>{[preset.colors.accent1, preset.colors.accent2, preset.colors.accent4, preset.colors.accent6].map((c, idx) => <i key={idx} style={{ width: 9, height: 9, background: "#" + c, display: "inline-block" }} />)}</span>
          {preset.label}
        </button>
      ))}</div>}
    </OfficeMenu>;
    if (command === "wordFontSchemeMenu") return <OfficeMenu key={key} label={label} icon={<Icon name={icon} size={17} />} disabled={!canFormat}>
      {close => <>{WORD_FONT_SCHEMES.map(preset => (
        <button key={preset.id} role="menuitem" type="button" onClick={() => { onTextCommand({ type: "word", action: "setFontScheme", schemeId: preset.id }); close(); }}>{preset.label}</button>
      ))}</>}
    </OfficeMenu>;
    if (command === "wordPageColorMenu") return <OfficeColorMenu key={key} label={label} value={format?.wordState?.pageBackgroundColor ?? "#ffffff"} fill large disabled={!canFormat} themeColors={format?.themeColors}
      apply={value => onTextCommand({ type: "word", action: "setPageColor", color: value === "#ffffff" || value === "transparent" ? null : value })} />;
    if (command === "wordLineNumbersMenu") return <OfficeMenu key={key} label={label} icon={<Icon name={icon} size={17} />} disabled={!canFormat}>
      {close => <>{([
        ["none", "无"],
        ["continuous", "连续"],
        ["newPage", "每页重新编号"],
        ["newSection", "每节重新编号"],
      ] as const).map(([mode, text]) => (
        <button key={mode} role="menuitemradio" type="button" aria-checked={(format?.wordState?.lineNumbersRestart ?? "none") === mode}
          onClick={() => { onTextCommand({ type: "word", action: "setLineNumbers", mode }); close(); }}>{text}</button>
      ))}
        <button role="menuitem" type="button" onClick={() => { onTextCommand({ type: "word", action: "pageSettings" }); close(); }}>行号选项…</button>
      </>}
    </OfficeMenu>;
    if (command === "wordHyphenationMenu") return <OfficeMenu key={key} label={label} icon={<Icon name={icon} size={17} />} disabled={!canFormat}>
      {close => <>
        <button role="menuitemradio" type="button" aria-checked={!format?.wordState?.automaticHyphenation} onClick={() => { onTextCommand({ type: "word", action: "setHyphenation", enabled: false }); close(); }}>无</button>
        <button role="menuitemradio" type="button" aria-checked={Boolean(format?.wordState?.automaticHyphenation)} onClick={() => { onTextCommand({ type: "word", action: "setHyphenation", enabled: true }); close(); }}>自动</button>
      </>}
    </OfficeMenu>;
    if (command === "wordAddTextMenu") return <OfficeMenu key={key} label={label} icon={<Icon name={icon} size={17} />} disabled={!canFormat}>
      {close => <>
        <button role="menuitemradio" type="button" aria-checked={format?.wordState?.outlineLevel === null} onClick={() => { onTextCommand({ type: "word", action: "setOutlineLevel", level: null }); close(); }}>不显示在目录中</button>
        {[0, 1, 2, 3].map(level => (
          <button key={level} role="menuitemradio" type="button" aria-checked={format?.wordState?.outlineLevel === level}
            onClick={() => { onTextCommand({ type: "word", action: "setOutlineLevel", level }); close(); }}>{level + 1} 级</button>
        ))}
      </>}
    </OfficeMenu>;
    if (command === "wordReviewDisplayMenu") return <OfficeMenu key={key} label={label} icon={<Icon name={icon} size={32} />} large disabled={!canFormat}>
      {close => <>{([
        ["simple", "简单标记"],
        ["all", "所有标记"],
        ["final", "无标记"],
        ["original", "原始状态"],
      ] as const).map(([mode, text]) => (
        <button key={mode} role="menuitemradio" type="button" aria-checked={(format?.wordState?.reviewDisplay ?? "all") === mode}
          onClick={() => { onTextCommand({ type: "word", action: "setReviewDisplay", mode }); close(); }}>{text}</button>
      ))}
        <button role="menuitem" type="button" onClick={() => { onTextCommand({ type: "word", action: "review" }); close(); }}>审阅窗格…</button>
      </>}
    </OfficeMenu>;
    if (command === "wordAcceptMenu") return <OfficeMenu key={key} label={label} icon={<Icon name={icon} size={32} />} large disabled={!canFormat} split={() => onTextCommand({ type: "word", action: "acceptRevision" })}>
      {close => <>
        <button role="menuitem" type="button" onClick={() => { onTextCommand({ type: "word", action: "acceptRevision" }); close(); }}>接受并移至下一处</button>
        <button role="menuitem" type="button" onClick={() => { onTextCommand({ type: "word", action: "acceptAllRevisions" }); close(); }}>接受所有修订</button>
        <button role="menuitem" type="button" onClick={() => { onTextCommand({ type: "word", action: "review" }); close(); }}>审阅窗格…</button>
      </>}
    </OfficeMenu>;
    if (command === "wordRejectMenu") return <OfficeMenu key={key} label={label} icon={<Icon name={icon} size={32} />} large disabled={!canFormat} split={() => onTextCommand({ type: "word", action: "rejectRevision" })}>
      {close => <>
        <button role="menuitem" type="button" onClick={() => { onTextCommand({ type: "word", action: "rejectRevision" }); close(); }}>拒绝并移至下一处</button>
        <button role="menuitem" type="button" onClick={() => { onTextCommand({ type: "word", action: "rejectAllRevisions" }); close(); }}>拒绝所有修订</button>
        <button role="menuitem" type="button" onClick={() => { onTextCommand({ type: "word", action: "review" }); close(); }}>审阅窗格…</button>
      </>}
    </OfficeMenu>;
    if (command === "freezeMenu") return <OfficeMenu key={key} label={label} icon={<Icon name={icon} size={32} />} large disabled={!states.get("freezePanes")?.enabled}>
      {close => <>{[["freezePanes", "冻结窗格"], ["unfreezePanes", "取消冻结窗格"]].map(([id, label]) => <button key={id} role="menuitem" disabled={!invoke(id!)} onClick={() => { invoke(id!)?.(); close(); }}>{label}</button>)}</>}
    </OfficeMenu>;
    if (command === "zoom") return <OfficeMenu key={key} label={label} icon={<Icon name={icon} size={32} />} large disabled={busy || kind === "workbook" && !enableWorkbookZoom}>
      {close => <>{[.5, .75, 1, 1.25, 1.5, 2, 3, 4].map(value => <button key={value} role="menuitemradio" aria-checked={zoom === value} onClick={() => { onZoom(value); close(); }}>{value * 100}%</button>)}</>}
    </OfficeMenu>;
    const active =
      command === "view:pages"
        ? layout === "pages" && !draftMode
        : command === "view:continuous"
          ? layout === "continuous" && !draftMode
          : command === "view:draft"
            ? Boolean(draftMode)
            : command === "focusMode"
              ? Boolean(focusMode)
              : command === "gridlines"
                ? Boolean(showGridlines)
                : command === "immersiveReader"
                  ? Boolean(immersiveReader)
                  : command === "splitWindow"
                    ? Boolean(splitWindow)
                    : command === "w:toggleTrackChanges"
                      ? Boolean(format?.wordState?.trackChanges)
                      : command === "w:protection"
                        ? Boolean(format?.wordState?.protectionEnforced)
                        : command?.startsWith("zoom:")
                          ? pageZoomMode === command.slice(5)
                          : undefined;
    return <Tool key={key} label={label} icon={icon} large={large} active={active} run={command ? invoke(command) : undefined} />;
  };
  if (tab === "绘图") return <><Group title="工具"><Tool label="绘图" icon="pen" large /><Tool label="橡皮擦" icon="eraser" large />{kind !== "document" && <Tool label="套索选择" icon="cursor" large />}</Group><Group title="笔"><div className="native-pen-gallery">{["#161616", "#d21b23", "#888888", "#ffff00", "#e9cf73"].map((color, index) => <button key={color} type="button" disabled aria-label={["黑色笔", "红色笔", "铅笔", "黄色荧光笔", "荧光笔"][index]} style={{ color }}><Icon name="pen" size={37} /></button>)}</div><Tool label="添加" icon="plus" /></Group>{kind === "presentation" && <Group title="转换"><Tool label="墨迹转换为文本" icon="textBox" large /><Tool label="将墨迹转换为形状" icon="shapes" large /><Tool label="将墨迹转换为数学公式" icon="numbering" large /></Group>}<Group title="触控板"><Tool label="使用触控板绘图" icon="pen" large /></Group></>;
  const sections = (kind === "document" ? word : kind === "presentation" ? ppt : excel)[tab] ?? [];
  const paraProps = format?.paragraphProperties ?? {};
  const groupLauncher = (title: string) => {
    if (kind !== "document") return undefined;
    if (tab === "布局" && title === "页面设置")
      return { label: "页面设置对话框", disabled: !canFormat, run: () => onTextCommand({ type: "word", action: "pageSettings" }) };
    if (tab === "布局" && title === "缩进和间距")
      return { label: "段落对话框", disabled: !canFormat, run: paragraph };
    return undefined;
  };
  return <>{sections.map(([title, ...columns]) => <Group key={title} title={title} launcher={groupLauncher(title)}>{columns.map((column, index) => {
    if ("wordIndentSpacing" in column) {
      return (
        <div key={index} style={{ display: "flex", gap: 10 }}>
          <div className="ribbon-spin-stack">
            <RibbonNumberField
              label="左缩进"
              valuePt={parsePtValue(paraProps["margin-left"])}
              disabled={!canFormat}
              min={-144}
              max={360}
              step={6}
              onCommit={(pt) => onTextCommand({ type: "word", action: "paragraph", properties: { "margin-left": `${pt}pt` } })}
            />
            <RibbonNumberField
              label="右缩进"
              valuePt={parsePtValue(paraProps["margin-right"])}
              disabled={!canFormat}
              min={-144}
              max={360}
              step={6}
              onCommit={(pt) => onTextCommand({ type: "word", action: "paragraph", properties: { "margin-right": `${pt}pt` } })}
            />
          </div>
          <div className="ribbon-spin-stack">
            <RibbonNumberField
              label="段前"
              valuePt={parsePtValue(paraProps["margin-top"])}
              disabled={!canFormat}
              min={0}
              max={360}
              step={6}
              onCommit={(pt) => onTextCommand({ type: "word", action: "paragraph", properties: { "margin-top": `${pt}pt` } })}
            />
            <RibbonNumberField
              label="段后"
              valuePt={parsePtValue(paraProps["margin-bottom"])}
              disabled={!canFormat}
              min={0}
              max={360}
              step={6}
              onCommit={(pt) => onTextCommand({ type: "word", action: "paragraph", properties: { "margin-bottom": `${pt}pt` } })}
            />
          </div>
        </div>
      );
    }
    if ("stack" in column) return <div key={index} className="ribbon-tool-stack">{column.stack.map((item, index) => drawItem(item, false, index))}</div>;
    if ("gallery" in column) {
      const styleSetInteractive = kind === "document" && tab === "设计";
      return <div key={index} className={"native-command-gallery " + (column.category ?? "theme")}
        onPointerLeave={styleSetInteractive ? () => onTextCommand({ type: "word", action: "livePreview", properties: null }) : undefined}>{column.gallery.map((label, idx) => {
        const preset = styleSetInteractive ? WORD_STYLE_SETS.find(s => s.label === label) : undefined;
        return <button key={label} disabled={styleSetInteractive ? !canFormat : true} type="button" title={label} aria-label={label}
          style={styleSetInteractive ? { opacity: 1 } : undefined}
          onPointerEnter={styleSetInteractive && preset ? () => onTextCommand({ type: "word", action: "livePreview", scope: "theme", properties: { "line-height": preset.lineHeight, color: preset.headingColor } }) : undefined}
          onMouseDown={styleSetInteractive ? event => event.preventDefault() : undefined}
          onClick={styleSetInteractive ? () => { onTextCommand({ type: "word", action: "livePreview", properties: null }); onTextCommand({ type: "word", action: "setStyleSet", styleSetId: label }); } : undefined}>
          <span style={{ borderTopColor: preset?.accent ?? ["#4472c4", "#b7472a", "#107c41", "#767171"][idx % 4] }}>{column.category === "transition" ? <Icon name="slides" size={26} /> : <><b style={preset ? { color: preset.headingColor, fontStyle: preset.headingItalic ? "italic" : "normal" } : undefined}>Aa</b><i /></>}</span><small>{label}</small>
        </button>;
      })}</div>;
    }
    return drawItem(column as Item, true, index);
  })}</Group>)}{tab === "审阅" && actions.objectSelection && <Group title="选择"><OfficeMenu label="选择" icon={<Icon name="cursor" size={32} />} large disabled={busy}>
    {close => <>{[["textSelection", "文字选择"], ["objectSelection", "对象选择"], ["regionSelection", "矩形框选"]].map(([id, label]) => <button role="menuitem" key={id} disabled={!invoke(id!)} onClick={() => { invoke(id!)?.(); close(); }}>{label}</button>)}</>}
  </OfficeMenu></Group>}</>;
}
