import { useEffect, useId, useState } from "react";
import { OfficeComboBox } from "./OfficeComboBox.js";
import type { ReactNode, ReactElement } from "react";
import type { DocumentTextCommand, EditorCommandState, TextFormatState } from "../index.js";
import { OfficeIcon as Icon } from "./icons.js";
import { officeRibbonStyles } from "./styles.js";

export type OfficeActions = Partial<Record<string, { run(): void; disabled?: boolean; reason?: string }>>;
const tabsByKind = {
  document: ["文件", "开始", "插入", "绘图", "设计", "布局", "引用", "邮件", "审阅", "视图", "帮助"],
  presentation: ["文件", "开始", "插入", "绘图", "设计", "切换", "动画", "幻灯片放映", "审阅", "视图", "帮助"],
  workbook: ["文件", "开始", "插入", "页面布局", "公式", "数据", "审阅", "视图", "帮助"],
};
type Glyph = Parameters<typeof Icon>[0]["name"];
function Group({ title, children, className = "" }: { title: string; children: ReactNode; className?: string }) {
  return <div className={"ribbon-group " + className} role="group" aria-label={title}><div className="ribbon-group-content">{children}</div><span className="ribbon-group-caption">{title}</span></div>;
}
function Tool({ label, icon = "file", glyph, large, run, disabled = !run, active, reason }: {
  label: string; icon?: Glyph; glyph?: ReactNode; large?: boolean; run?: (() => void) | undefined; disabled?: boolean;
  active?: boolean | undefined; reason?: string | undefined;
}) {
  return <button type="button" className={"ribbon-tool" + (large ? " large" : "")} aria-label={label} aria-pressed={active}
    title={label + (disabled ? " · " + (reason ?? (run ? "当前选区或模式不可用" : "尚未实现")) : "")}
    disabled={disabled} onMouseDown={event => { const focused = document.activeElement; if (focused instanceof HTMLInputElement && focused.closest(".office-combo")) focused.blur(); event.preventDefault(); }} onClick={run}>
    {glyph ? <span className="ribbon-glyph" aria-hidden="true">{glyph}</span> : <Icon name={icon} size={large ? 27 : 17} />}<span>{label}</span>
  </button>;
}
function fontName(value: string | undefined): string {
  if (!value) return "";
  const match = /^\s*(?:"((?:\\.|[^"])*)"|'((?:\\.|[^'])*)'|([^,]+))/.exec(value);
  return (match?.[1] ?? match?.[2] ?? match?.[3] ?? "").replaceAll('\\"', '"').trim();
}
function previewFontAvailable(name: string): boolean {
  if (!name || ["serif", "sans-serif", "monospace", "system-ui"].includes(name)) return true;
  const canvas = document.createElement("canvas"), context = canvas.getContext("2d"); if (!context) return true;
  return ["monospace", "serif", "sans-serif"].some(base => {
    context.font = "72px " + base; const baseline = context.measureText("汉字ABCmw0123").width;
    context.font = "72px " + JSON.stringify(name) + "," + base; return Math.abs(context.measureText("汉字ABCmw0123").width - baseline) > .01;
  });
}
export interface OfficeRibbonProps {
  kind: "document" | "presentation" | "workbook"; tab: string; onTab(tab: string): void; mode: "edit" | "preview"; busy: boolean;
  actions: OfficeActions; commands: readonly EditorCommandState[]; onCommand(id: string): void;
  formatState: TextFormatState; onTextCommand(command: DocumentTextCommand): void | Promise<void>;
  rulers: boolean; onRulers(visible: boolean): void;
  layout: "pages" | "continuous"; onLayout(value: "pages" | "continuous"): void; zoom: number; onZoom(value: number): void;
}

/** Optional SDK chrome; hosts supply authorized file/export callbacks. */
export function OfficeRibbon({ kind, tab: requestedTab, onTab, mode, busy, actions, commands, onCommand, formatState: format, onTextCommand, layout, onLayout, zoom, onZoom, rulers, onRulers }: OfficeRibbonProps): ReactElement {
  const id = useId(), panelId = id + '-panel';
  const preview = mode === "preview", allowedPreviewTabs = ["开始", "文件", "审阅", "视图", "帮助"];
  const tab = preview && !allowedPreviewTabs.includes(requestedTab) ? "开始" : requestedTab, collapsed = preview && tab === "开始";
  const tabs = tabsByKind[kind], states = new Map(commands.map(command => [command.id, command]));
  const canFormat = !preview && format.enabled && !busy;
  const chosenFont = fontName(format.fontFamily), [fontAvailable, setFontAvailable] = useState(true);
  useEffect(() => { setFontAvailable(previewFontAvailable(chosenFont)); }, [chosenFont]);
  const command = (id: string, label: string, icon: Glyph, large = false) => {
    const native = states.get(id), host = actions[id], run = native ? () => onCommand(id) : host?.run;
    return <Tool label={label} icon={icon} large={large} run={run} disabled={busy || (native ? !native.enabled : !host || Boolean(host.disabled))} active={native?.active} reason={host?.reason} />;
  };
  const toggle = (label: string, property: "bold" | "italic" | "underline", glyph: ReactNode) => <Tool label={label} glyph={glyph} disabled={!canFormat} active={format[property]} run={() => onTextCommand({ type: "toggle", property })} />;
  const setFormat = (label: string, property: Extract<DocumentTextCommand, { type: "format" }>["property"], value: string, glyph: ReactNode) => <Tool label={label} glyph={glyph} disabled={!canFormat} run={() => onTextCommand({ type: "format", property, value })} />;
  const families = [...new Set([fontName(format.fontFamily), "等线", "宋体", "微软雅黑", "Arial", "Calibri", "Aptos", "Times New Roman", "Noto Sans CJK SC", "Noto Serif CJK SC", "FandolSong", "FandolHei", "FandolKai", "FandolFang", "LXGW WenKai", "Liberation Serif", "Liberation Sans", "Liberation Mono"].filter(Boolean))];
  const sizes = [...new Set([8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72, ...format.fontSize ? [format.fontSize] : []])].sort((a, b) => a - b);
  const fontGroup = <Group title="字体" className="font-group">
    <div className="ribbon-font-selects">
      <OfficeComboBox label="字体" className="font-family-field" disabled={!canFormat} value={fontName(format.fontFamily)} options={families} hint={!fontAvailable ? `本机没有“${chosenFont}”，预览使用替代字体；文档保留所选字体。` : undefined}
        commit={value => { if (!value || value.length > 256) return "请输入有效的字体名称。"; return onTextCommand({ type: "format", property: "font-family", value: JSON.stringify(value) }); }} />
      <OfficeComboBox label="字号" className="font-size-field" disabled={!canFormat} value={format.fontSize?.toString() ?? ""} options={sizes.map(String)}
        commit={value => { const size = Number(value), maximum = kind === "document" ? 1638 : kind === "workbook" ? 409 : 3600, step = kind === "presentation" ? 10 : 2;
          if (!value || !Number.isFinite(size) || size < 1 || size > maximum || Math.abs(size * step - Math.round(size * step)) > .0001) return `字号须为 1～${maximum}，步长 ${1 / step}。`;
          return onTextCommand({ type: "format", property: "font-size", value: size + "pt" }); }} />
      <Tool label="增大字号" glyph={<>A<sup>⌃</sup></>} disabled={!canFormat || !format.fontSize || format.fontSize >= (kind === "document" ? 1638 : kind === "workbook" ? 409 : 3600)} run={() => onTextCommand({ type: "format", property: "font-size", value: (sizes.find(size => size > (format.fontSize ?? 0)) ?? Math.min(kind === "document" ? 1638 : kind === "workbook" ? 409 : 3600, Math.floor((format.fontSize ?? 72) / 10) * 10 + 10)) + "pt" })} />
      <Tool label="减小字号" glyph={<>A<sup>⌄</sup></>} disabled={!canFormat || !format.fontSize || format.fontSize <= 1} run={() => onTextCommand({ type: "format", property: "font-size", value: (format.fontSize! > 72 ? Math.max(72, Math.ceil(format.fontSize! / 10) * 10 - 10) : sizes.filter(size => size < (format.fontSize ?? 0)).at(-1) ?? Math.max(1, format.fontSize! - 1)) + "pt" })} />
      <Tool label="清除所有格式" icon="eraser" />
    </div>
    <div className="ribbon-icon-row">
      {toggle("加粗", "bold", <b>B</b>)}{toggle("斜体", "italic", <i>I</i>)}{toggle("下划线", "underline", <u>U</u>)}
      <Tool label="删除线" glyph={<s>abc</s>} /><Tool label="下标" glyph={<>x<sub>2</sub></>} /><Tool label="上标" glyph={<>x<sup>2</sup></>} />
      <span className="ribbon-divider" /><Tool label="文本突出显示颜色" icon="highlight" />
      {setFormat("深蓝色", "color", "#1f4b99", <span className="font-color">A</span>)}
      {setFormat("常规字重", "font-weight", "400", <span className="normal-weight">A</span>)}
    </div>
  </Group>;
  const paragraph = <Group title={kind === "workbook" ? "对齐方式" : "段落"} className="paragraph-group">
    <div className="ribbon-icon-row">
      {command("bullets", "项目符号", "list")}{command("numbering", "编号", "numbering")}<Tool label="多级列表" icon="numbering" />
      <span className="ribbon-divider" /><Tool label="减少缩进量" icon="indent" /><Tool label="增加缩进量" icon="indent" /><Tool label="显示编辑标记" glyph="¶" />
    </div>
    <div className="ribbon-icon-row">
      {(["left", "center", "right"] as const).map((value, index) => <Tool key={value} label={["左对齐", "居中", "右对齐"][index]!} icon={value} disabled={!canFormat} active={format.alignment === value} run={() => onTextCommand({ type: "align", value })} />)}
      <Tool label="两端对齐" icon="justify" /><Tool label="行和段落间距" icon="lineSpacing" />
      {kind === "workbook" ? command("formatCells", "设置单元格格式", "table") : <Tool label="边框" icon="table" />}
    </div>
  </Group>;
  const clipboard = <Group title="剪贴板" className="clipboard-group">{command("paste", "粘贴", "paste", true)}<div className="ribbon-tool-stack">{command("cut", "剪切", "cut")}{command("copy", "复制", "copy")}<Tool label="格式刷" icon="brush" /></div></Group>;
  const insertObjects = <Group title="插图">{command(kind === "workbook" ? "drawings" : "insertImage", "图片", "image", true)}{command(kind === "workbook" ? "drawings" : "insertShape", "形状", "shapes", true)}<Tool label="图标" icon="grid" large /><Tool label="SmartArt" icon="diagram" large />{command(kind === "workbook" ? "drawings" : "insertChart", "图表", "chart", true)}</Group>;
  const selectionTools = <Group title="选择与批注"><div className="ribbon-tool-stack">{command("textSelection", "文字划选", "textBox")}{command("objectSelection", "对象选择", "cursor")}{command("regionSelection", "矩形框选", "margins")}</div>{command("newComment", "新建批注", "comment", true)}</Group>;
  let content: ReactNode;
  if (tab === "文件") content = <>
    <Group title="打开">{command("open", "打开", "folder", true)}</Group>
    <Group title="保存">{command("package", "另存为文档包", "save", true)}{command("native", kind === "presentation" ? "导出 PPTX" : kind === "workbook" ? "导出 XLSX" : "导出 DOCX", "file", true)}</Group>
    <Group title="导出">{command("pdf", "导出 PDF", "pdf", true)}{command("png", "导出 PNG", "image", true)}</Group><Group title="打印"><Tool label="打印" icon="print" large /></Group>
  </>;
  else if (tab === "开始") content = <>
    {clipboard}
    {kind === "presentation" && <Group title="幻灯片">{command("newSlide", "新建幻灯片", "slides", true)}<div className="ribbon-tool-stack">{command("slideLayout", "版式", "grid")}{command("duplicateSlide", "复制幻灯片", "copy")}{command("deleteSlide", "删除幻灯片", "close")}</div></Group>}
    {fontGroup}{paragraph}
    {kind === "document" && <Group title="样式" className="styles-group"><div className="style-gallery">{["正文", "无间隔", "标题 1", "标题 2"].map((name, index) => <button key={name} disabled title="段落样式选择尚未实现"><span style={{ fontSize: index > 1 ? 19 : 15, color: index > 1 ? "#315c90" : "#283549" }}>AaBbCc</span><small>{name}</small></button>)}</div></Group>}
    {kind === "workbook" && <><Group title="合并与填充"><div className="ribbon-tool-stack">{command("mergeCells", "合并单元格", "table")}{command("unmergeCells", "取消合并", "table")}{command("fillDown", "向下填充", "download")}</div><div className="ribbon-tool-stack">{command("fillRight", "向右填充", "arrow")}{command("fillSeries", "数列填充", "numbering")}{command("formatCells", "数字格式", "numbering")}</div></Group><Group title="样式">{command("rules", "条件格式", "grid", true)}{command("tables", "套用表格格式", "table", true)}{command("tableStyles", "自定义表格样式", "palette", true)}</Group><Group title="单元格"><div className="ribbon-tool-stack">{command("insertRows", "插入行", "plus")}{command("deleteRows", "删除行", "close")}{command("worksheets", "工作表格式", "grid")}</div></Group></>}
    {kind === "presentation" && <Group title="绘图">{command("insertShape", "形状", "shapes", true)}{command("formatObject", "设置对象格式", "palette", true)}<div className="ribbon-tool-stack">{command("duplicateObject", "复制对象", "copy")}{command("deleteObject", "删除对象", "close")}{command("groupObjects", "组合", "grid")}</div></Group>}
    <Group title="编辑"><div className="ribbon-tool-stack">{command("find", "查找", "search")}{command("replace", "替换", "replace")}{command(kind === "presentation" ? "editText" : "objectSelection", kind === "presentation" ? "编辑文字" : "对象选择", "cursor")}{kind === "presentation" && format.editingText && command("finishInput", "完成输入", "check")}</div></Group>
    <Group title="智能协作">{command("proposal", "载入固定示例提案", "sparkles", true)}</Group>
  </>;
  else if (tab === "插入") content = <>
    {kind === "document" && <Group title="页面"><Tool label="封面" icon="file" large /><div className="ribbon-tool-stack"><Tool label="空白页" icon="file" /><Tool label="分页符" icon="pageBreak" /></div></Group>}
    <Group title="表格">{command(kind === "workbook" ? "tables" : "insertTable", "表格", "table", true)}</Group>{insertObjects}
    <Group title="链接">{command("hyperlinks", "链接", "link", true)}<Tool label="书签" icon="bookmark" large /></Group>
    <Group title="文本">{command(kind === "workbook" ? "drawings" : "insertTextBox", "文本框", "textBox", true)}<Tool label="艺术字" glyph="A" large /></Group>
    <Group title="批注">{command("newComment", "新建批注", "comment", true)}</Group>
    {kind === "document" && <><Group title="页眉和页脚"><Tool label="页眉" icon="header" large /><Tool label="页脚" icon="footer" large /><Tool label="页码" icon="numbering" large /></Group><Group title="符号"><Tool label="公式" glyph="π" large /><Tool label="符号" glyph="Ω" large /></Group></>}
  </>;
  else if (tab === "审阅") content = <>
    <Group title="校对"><Tool label="编辑器" icon="proof" large /><div className="ribbon-tool-stack"><Tool label="同义词库" icon="book" /><Tool label="字数统计" icon="numbering" /></div></Group>
    {selectionTools}<Group title="批注">{command("comments", "显示批注", "comment", true)}</Group>
    {kind === "document" && <><Group title="修订"><Tool label="修订" icon="edit" large /><Tool label="显示标记" icon="file" large /></Group><Group title="更改"><Tool label="接受" icon="check" large /><Tool label="拒绝" icon="close" large /></Group></>}
    <Group title="协作提案">{command("proposal", "载入固定示例提案", "sparkles", true)}</Group>
  </>;
  else if (tab === "数据" && kind === "workbook") content = <>
    <Group title="排序和筛选">{command("sortAscending", "升序排序", "numbering", true)}{command("sortDescending", "降序排序", "numbering", true)}{command("filters", "筛选", "list", true)}</Group>
    <Group title="数据工具">{command("rules", "数据验证", "check", true)}{command("tables", "智能表管理", "table", true)}</Group><Group title="工作表">{command("worksheets", "工作表管理", "grid", true)}</Group>
  </>;
  else if (tab === "公式" && kind === "workbook") content = <>
    <Group title="函数库"><Tool label="插入函数" glyph="fx" large /><Tool label="自动求和" glyph="Σ" large /></Group>
    <Group title="定义的名称">{command("names", "名称管理器", "bookmark", true)}</Group><Group title="计算">{command("recalculate", "重新计算", "refresh", true)}</Group>
  </>;
  else if (tab === "视图") content = <>
    <Group title="视图">{command("preview", kind === "presentation" ? "普通视图" : "阅读模式", "book", true)}<Tool label="页面视图" icon="file" large disabled={kind !== "document" || busy} run={() => onLayout("pages")} /><Tool label="连续视图" icon="globe" large disabled={kind !== "document" || busy} run={() => onLayout("continuous")} /></Group>
    <Group title="显示"><div className="ribbon-tool-stack"><label><input type="checkbox" checked={kind === "document" && rulers} disabled={kind !== "document" || busy} onChange={event => onRulers(event.currentTarget.checked)} />标尺</label><label><input type="checkbox" disabled />网格线</label>{command("library", "导航窗格", "panel")}</div></Group>
    {kind === "workbook" ? <Group title="窗口">{command("freezePanes", "冻结窗格", "grid", true)}{command("unfreezePanes", "取消冻结", "grid", true)}</Group> : <Group title="缩放"><Tool label="100%" icon="search" large disabled={busy} run={() => onZoom(1)} /><select aria-label="预览缩放" value={zoom} disabled={busy} onChange={event => onZoom(Number(event.currentTarget.value))}>{[.5, .75, 1, 1.25, 1.5, 2].map(value => <option key={value} value={value}>{value * 100}%</option>)}</select><span className="view-description">{layout === "pages" ? "打印布局" : "连续阅读"}</span></Group>}
    {kind === "presentation" && <Group title="母版视图">{command("definitions", "幻灯片母版", "slides", true)}{command("editNotes", "备注", "edit", true)}</Group>}
  </>;
  else if (tab === "布局" || tab === "页面布局") content = <><Group title="页面设置">{[["页边距", "margins"], ["纸张方向", "orientation"], ["纸张大小", "file"], ["分栏", "columns"], ["分隔符", "pageBreak"]].map(([label, icon]) => <Tool key={label} label={label!} icon={icon as Glyph} large />)}</Group>{paragraph}<Group title="排列">{command("formatObject", "位置和大小", "grid", true)}</Group></>;
  else if (tab === "引用") content = <><Group title="目录"><Tool label="目录" icon="list" large />{command("fields", "更新页码与目录", "refresh", true)}</Group><Group title="脚注"><Tool label="插入脚注" glyph={<>AB<sup>1</sup></>} large /><Tool label="插入尾注" glyph={<>AB<sup>i</sup></>} large /></Group></>;
  else if (tab === "设计") content = <><Group title={kind === "presentation" ? "主题和版式" : "文档格式"}>{command("definitions", "主题", "palette", true)}{kind === "presentation" && command("slideLayout", "幻灯片版式", "slides", true)}<div className="style-gallery">{["简洁", "正式", "现代", "经典"].map(name => <button key={name} disabled><span>Aa</span><small>{name}</small></button>)}</div></Group><Group title="背景">{command("slideBackground", "设置背景格式", "fill", true)}</Group></>;
  else if (tab === "绘图") content = <><Group title="工具"><Tool label="选择" icon="cursor" /><Tool label="橡皮擦" icon="eraser" /></Group><Group title="笔"><div className="pen-gallery">{["#323b4b", "#cc4242", "#286acb", "#a4b52f"].map(color => <button key={color} disabled aria-label="画笔" style={{ color }}><Icon name="pen" size={38} /></button>)}</div></Group></>;
  else content = <Group title={tab}>{(tab === "邮件" ? ["信封", "标签", "开始邮件合并", "选择收件人"] : tab === "帮助" ? ["帮助", "联系支持", "反馈"] : ["预览", "效果选项", "持续时间", "应用到全部"]).map(label => <Tool key={label} label={label} icon="file" large />)}</Group>;
  return <div className="doffi-office-ribbon office-ribbon" data-mode={mode} data-collapsed={collapsed}><style>{officeRibbonStyles}</style><div className="office-tabs" role="tablist" aria-label="Office 功能区">
    {tabs.map(name => <button key={name} id={id + "-tab-" + name} role="tab" disabled={preview && !allowedPreviewTabs.includes(name)} aria-selected={tab === name} aria-controls={panelId} tabIndex={tab === name ? 0 : -1} onClick={() => onTab(name)} onKeyDown={event => {
      const direction = event.key === "ArrowRight" ? 1 : event.key === "ArrowLeft" ? -1 : 0;
      if (direction || event.key === "Home" || event.key === "End") { event.preventDefault(); const availableTabs = preview ? tabs.filter(name => allowedPreviewTabs.includes(name)) : tabs; const next = event.key === "Home" ? availableTabs[0]! : event.key === "End" ? availableTabs.at(-1)! : availableTabs[(availableTabs.indexOf(name) + direction + availableTabs.length) % availableTabs.length]!; onTab(next); [...event.currentTarget.parentElement!.querySelectorAll<HTMLButtonElement>('[role=tab]')].find(button => button.textContent === next)?.focus(); }
    }}>{name}</button>)}<div className="office-ribbon-note">{kind === "document" ? "Word" : kind === "workbook" ? "Excel" : "PowerPoint"}</div>
  </div><div id={panelId} role="tabpanel" aria-labelledby={id + "-tab-" + tab} className="office-ribbon-panel" hidden={collapsed}>{content}</div></div>;
}
