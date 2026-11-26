import { useEffect, useId, useRef, useState } from "react";
import { WordParagraphDialog } from "./WordParagraphDialog.js";
import { OfficeComboBox } from "./OfficeComboBox.js";
import { OfficeMenu, OfficeColorMenu } from "./OfficeMenu.js";
import { OfficeDialog } from "./OfficeDialog.js";
import type { ReactNode, ReactElement } from "react";
import type { DocumentTextCommand, EditorCommandState, TextFormatState } from "../index.js";
import { OfficeIcon as Icon } from "./icons.js";
import { officeRibbonStyles } from "./styles.js";
import { loadWordWebFonts } from "@doffi/render";
import type { WordWebFont } from "@doffi/render";
import { officeSizeLabel, officeSizePresets, wordSizeNames } from "./font-presets.js";

export type OfficeActions = Partial<
  Record<string, { run(): void; disabled?: boolean; reason?: string }>
>;
const tabsByKind = {
  document: [
    "文件",
    "开始",
    "插入",
    "绘图",
    "设计",
    "布局",
    "引用",
    "邮件",
    "审阅",
    "视图",
    "帮助",
  ],
  presentation: [
    "文件",
    "开始",
    "插入",
    "绘图",
    "设计",
    "切换",
    "动画",
    "幻灯片放映",
    "录制",
    "审阅",
    "视图",
    "帮助",
  ],
  workbook: ["文件", "开始", "插入", "绘图", "页面布局", "公式", "数据", "审阅", "视图", "自动化", "帮助"],
};
import { Group, Tool } from "./ribbon-controls.js";
import type { Glyph } from "./ribbon-controls.js";
import { NativeOfficeTabs } from "./NativeTabs.js";
import { officeNumberLabel, officeNumberPresets } from "./number-presets.js";
function fontName(value: string | undefined): string {
  if (!value) return "";
  const match = /^\s*(?:"((?:\\.|[^"])*)"|'((?:\\.|[^'])*)'|([^,]+))/.exec(value);
  return (match?.[1] ?? match?.[2] ?? match?.[3] ?? "").replaceAll('\\"', '"').trim();
}
export interface OfficeRibbonProps {
  /** A standalone ribbon can retain application commands as tabs. */
  showApplicationTabs?: boolean;
  webFonts?: readonly WordWebFont[];
  kind: "document" | "presentation" | "workbook";
  tab: string;
  onTab(tab: string): void;
  mode: "edit" | "preview";
  busy: boolean;
  actions: OfficeActions;
  commands: readonly EditorCommandState[];
  onCommand(id: string): void;
  formatState: TextFormatState;
  onObjectFormat?(properties: Readonly<Record<string, string>>): void;
  onTextCommand(command: DocumentTextCommand): void | Promise<void>;
  layout: "pages" | "continuous";
  onLayout(value: "pages" | "continuous"): void;
  zoom: number;
  onZoom(value: number): void;
  enableWorkbookZoom?: boolean;
}

/** Optional SDK chrome; hosts supply authorized file/export callbacks. */
export function OfficeRibbon({
  showApplicationTabs = false,
  webFonts,
  kind,
  tab: requestedTab,
  onTab,
  mode,
  busy,
  actions,
  commands,
  onCommand,
  formatState: format,
  onTextCommand,
  onObjectFormat,
  layout,
  onLayout,
  zoom,
  onZoom,
  enableWorkbookZoom = false,
}: OfficeRibbonProps): ReactElement {
  const root = useRef<HTMLDivElement>(null), [compact, setCompact] = useState(false);
  useEffect(() => {
    const element = root.current!, observer = new ResizeObserver(() => setCompact(element.clientWidth <= 1400));
    observer.observe(element); return () => observer.disconnect();
  }, []);
  const id = useId(),
    panelId = id + "-panel";
  const [paragraphOpen, setParagraphOpen] = useState(false),
    [showMarks, setShowMarks] = useState(false);
  const [paint, setPaint] = useState<Record<string, string>>();
  const [styleName, setStyleName] = useState("");
  const [customStyles, setCustomStyles] = useState<string[]>([]);
  const [styleCreatorOpen, setStyleCreatorOpen] = useState(false);
  const [borderOpen, setBorderOpen] = useState(false);
  const [recentFonts, setRecentFonts] = useState<string[]>([]);
  const word = kind === "document";
  const styleIds = ["Normal", "NoSpacing", "Heading1", "Heading2", "Heading3", "Heading4", "Title", "Subtitle", "SubtleEmphasis", "Emphasis", "IntenseEmphasis", "Strong", "Quote", "IntenseQuote", "BookTitle", ...customStyles];
  const wordCommand = (value: Extract<DocumentTextCommand, { type: "word" }>) =>
    onTextCommand(value);

  const preview = mode === "preview",
    allowedPreviewTabs = ["开始", "文件", "审阅", "视图", "帮助"];
  const shapeSelected = word && !preview && !!format.drawingKind && ["shape", "textBox", "connector", "group"].includes(format.drawingKind);
  const tab = preview && !allowedPreviewTabs.includes(requestedTab) || requestedTab === "形状格式" && !shapeSelected ? "开始" : requestedTab,
    collapsed = preview && tab === "开始";
  const tabs = [...tabsByKind[kind].filter(name => showApplicationTabs || name !== "文件" && name !== "帮助"), ...(shapeSelected ? ["形状格式"] : [])],
    states = new Map(commands.map((command) => [command.id, command]));
  const canFormat = !preview && format.enabled && !busy;
  const command = (id: string, label: string, icon: Glyph, large = false) => {
    const native = states.get(id),
      host = actions[id],
      run = native ? () => onCommand(id) : host?.run;
    return (
      <Tool
        label={label}
        icon={icon}
        large={large}
        run={run}
        disabled={busy || (native ? !native.enabled : !host || Boolean(host.disabled))}
        active={native?.active}
        reason={host?.reason}
      />
    );
  };
  const toggle = (label: string, property: "bold" | "italic" | "underline", glyph: ReactNode) => (
    <Tool
      label={label}
      glyph={glyph}
      disabled={!canFormat}
      active={format[property]}
      run={() => onTextCommand({ type: "toggle", property })}
    />
  );
  const families = [
    ...new Set(
      [
        fontName(format.fontFamily),
        ...(webFonts?.filter(font => !font.fallbackOnly).flatMap(font => [font.family, ...font.aliases]) ?? []),
        "等线",
        "宋体",
        "微软雅黑",
        "Arial",
        "Calibri",
        "Aptos",
        "Times New Roman",
        "Noto Sans CJK SC",
        "Noto Serif CJK SC",
        "FandolSong",
        "FandolHei",
        "FandolKai",
        "FandolFang",
        "LXGW WenKai",
        "Liberation Serif",
        "Liberation Sans",
        "Liberation Mono",
      ].filter(Boolean),
    ),
  ];
  const sizes = [
    ...new Set([
      ...officeSizePresets[kind],
      ...(format.fontSize ? [format.fontSize] : []),
    ]),
  ].sort((a, b) => a - b);
  const fontGroup = (
    <Group title="字体" className="font-group">
      <div className="ribbon-font-selects">
        <OfficeComboBox
          label="字体"
          className="font-family-field"
          disabled={!canFormat}
          value={fontName(format.fontFamily)}
          options={[...new Set([...recentFonts, ...families])]}
          more={kind !== "workbook" ? { label: "字体…", run: () => kind === "document" ? wordCommand({ type: "word", action: "fontSettings" }) : onCommand("fontSettings") } : undefined}
          section={(name) => recentFonts.includes(name) ? "最近使用的字体" : "所有字体"}
          previewFamily={(name) => {
            const face = webFonts?.find(font => !font.fallbackOnly && [font.family, ...font.aliases].some(alias => alias.toLowerCase() === name.toLowerCase()));
            return face?.family ?? name;
          }}
          preview={(name) => { if (webFonts?.length) void loadWordWebFonts(document, webFonts, new Set([name.toLowerCase()])); }}
          commit={(value) => {
            if (!value || value.length > 256) return "请输入有效的字体名称。";
            setRecentFonts(current => [value, ...current.filter(name => name !== value)].slice(0, 6));
            return onTextCommand({
              type: "format",
              property: "font-family",
              value: JSON.stringify(value),
            });
          }}
        />
        <OfficeComboBox
          label="字号"
          className="font-size-field"
          disabled={!canFormat}
          value={officeSizeLabel(kind, format.fontSize)}
          displayValue={(value) => officeSizeLabel(kind, wordSizeNames[value] ?? Number(value))}
          options={kind === "document" ? [...Object.keys(wordSizeNames), ...officeSizePresets.document.map(String)] : officeSizePresets[kind].map(String)}
          commit={(value) => {
            const size = kind === "document" && wordSizeNames[value] !== undefined ? wordSizeNames[value]! : Number(value),
              maximum = kind === "document" ? 1638 : kind === "workbook" ? 409 : 3600,
              step = kind === "presentation" ? 10 : 2;
            if (
              !value ||
              !Number.isFinite(size) ||
              size < 1 ||
              size > maximum ||
              Math.abs(size * step - Math.round(size * step)) > 0.0001
            )
              return `字号须为 1～${maximum}，步长 ${1 / step}。`;
            return onTextCommand({ type: "format", property: "font-size", value: size + "pt" });
          }}
        />
        <Tool
          label="增大字号"
          glyph={
            <>
              A<sup>⌃</sup>
            </>
          }
          disabled={
            !canFormat ||
            !format.fontSize ||
            format.fontSize >= (kind === "document" ? 1638 : kind === "workbook" ? 409 : 3600)
          }
          run={() =>
            onTextCommand({
              type: "format",
              property: "font-size",
              value:
                (sizes.find((size) => size > (format.fontSize ?? 0)) ??
                  Math.min(
                    kind === "document" ? 1638 : kind === "workbook" ? 409 : 3600,
                    Math.floor((format.fontSize ?? 72) / 10) * 10 + 10,
                  )) + "pt",
            })
          }
        />
        <Tool
          label="减小字号"
          glyph={
            <>
              A<sup>⌄</sup>
            </>
          }
          disabled={!canFormat || !format.fontSize || format.fontSize <= 1}
          run={() =>
            onTextCommand({
              type: "format",
              property: "font-size",
              value:
                (format.fontSize! > 72
                  ? Math.max(72, Math.ceil(format.fontSize! / 10) * 10 - 10)
                  : (sizes.filter((size) => size < (format.fontSize ?? 0)).at(-1) ??
                    Math.max(1, format.fontSize! - 1))) + "pt",
            })
          }
        />
        {kind !== "workbook" && <>
        <OfficeMenu label="更改大小写" icon={<span className="ribbon-glyph">Aa</span>} disabled={!word || !canFormat}>
          {close => <>{[["全部小写", "lowercase"], ["全部大写", "uppercase"], ["每个单词首字母大写", "capitalize"], ["取消大小写转换", "none"]].map(([label, value]) =>
            <button key={value} role="menuitem" type="button" onClick={() => { wordCommand({ type: "word", action: "paintFormat", properties: { "text-transform": value! } }); close(); }}>{label}</button>)}</>}
        </OfficeMenu>
        <Tool
          label="清除所有格式"
          icon="clearFormat"
          disabled={!word || !canFormat}
          run={word ? () => wordCommand({ type: "word", action: "clearFormat" }) : undefined}
        />
        </>}
        {word && <><Tool label="拼音指南" icon="pinyin" /><Tool label="字符边框" icon="fontBorder" /></>}
      </div>
      <div className="ribbon-icon-row">
        {toggle("加粗", "bold", <b>B</b>)}
        {toggle("斜体", "italic", <i>I</i>)}
        <OfficeMenu label="下划线" active={format.underline} icon={<span className="ribbon-glyph"><u>U</u></span>} disabled={!canFormat}
          split={() => onTextCommand({ type: "toggle", property: "underline" })}>
          {close => <>{[["单线", "solid"], ["双线", "double"], ["无", "none"]].map(([name, value]) => <button key={name} role="menuitem" type="button" disabled={!word && value === "double"}
            onClick={() => { if (word) wordCommand({ type: "word", action: "paintFormat", properties: { "text-decoration-line": value === "none" ? "none" : "underline", "text-decoration-style": value === "none" ? "solid" : value! } }); else onTextCommand({ type: "format", property: "text-decoration", value: value === "none" ? "none" : "underline" }); close(); }}>{name}</button>)}
            {word && <button role="menuitem" type="button" onClick={() => { wordCommand({ type: "word", action: "fontSettings" }); close(); }}>其他下划线…</button>}</>}
        </OfficeMenu>
        {kind !== "workbook" && <>
        <Tool
          label="删除线"
          glyph={<s>abc</s>}
          disabled={!canFormat}
          active={format.strikethrough}
          run={() => onTextCommand({
                    type: "format",
                    property: "text-decoration",
                    value:
                      [
                        format.underline ? "underline" : "",
                        !format.strikethrough ? "line-through" : "",
                      ]
                        .filter(Boolean)
                        .join(" ") || "none",
                  })}
        />
        <Tool
          label="下标"
          glyph={
            <>
              x<sub>2</sub>
            </>
          }
          disabled={!canFormat}
          active={format.verticalAlign === "sub"}
          run={() => onTextCommand({
                    type: "format",
                    property: "vertical-align",
                    value: format.verticalAlign === "sub" ? "baseline" : "sub",
                  })}
        />
        <Tool
          label="上标"
          glyph={
            <>
              x<sup>2</sup>
            </>
          }
          disabled={!canFormat}
          active={format.verticalAlign === "super"}
          run={() => onTextCommand({
                    type: "format",
                    property: "vertical-align",
                    value: format.verticalAlign === "super" ? "baseline" : "super",
                  })}
        />
        </>}
        <span className="ribbon-divider" />
        {word && <Tool label="文本效果" icon="textEffects" />}
        {kind !== "workbook" && <OfficeColorMenu label="文本突出显示颜色" value="#ffff00" highlight disabled={!word || !canFormat}
          apply={value => onTextCommand({ type: "format", property: "background-color", value })} />}
        {kind === "workbook" && <><OfficeMenu label="边框" icon={<Icon name="table" size={17} />} disabled={!canFormat}>
          {close => <button role="menuitem" type="button" onClick={() => { onCommand("formatCells"); close(); }}>设置单元格格式…</button>}
        </OfficeMenu><OfficeColorMenu label="填充颜色" value="#ffff00" fill disabled={!canFormat}
          apply={value => onTextCommand({ type: "format", property: "background-color", value })} /></>}
        <OfficeColorMenu label="字体颜色" value="#ff0000" disabled={!canFormat}
          apply={value => onTextCommand({ type: "format", property: "color", value })} />
        {word && <><Tool label="字符底纹" glyph={<span style={{ background: "#ccc", padding: "0 2px" }}>A</span>}
          disabled={!canFormat} run={() => onTextCommand({ type: "format", property: "background-color", value: "#d9d9d9" })} />
          <Tool label="带圈字符" icon="enclose" /></>}
        {kind === "workbook" && <Tool label="显示拼音字段" icon="pinyin" />}
      </div>
    </Group>
  );
  const paragraph = (
    <Group title={kind === "workbook" ? "对齐方式" : "段落"} className="paragraph-group">
      <div className="ribbon-icon-row">
        {kind === "workbook" ? <>
          {command("verticalTop", "顶端对齐", "verticalTop")}
          {command("verticalCenter", "垂直居中", "verticalCenter")}
          {command("verticalBottom", "底端对齐", "verticalBottom")}
          <OfficeMenu label="方向" icon={<Icon name="textDirection" size={17} />} disabled={!canFormat}>
            {close => <button role="menuitem" onClick={() => { onCommand("formatCells"); close(); }}>设置单元格对齐方式…</button>}
          </OfficeMenu>
          {command("wrapText", "自动换行", "lineSpacing")}
        </> : <>
          {(["bullet", "numbering"] as const).map((kind, index) => <OfficeMenu key={kind} label={index ? "编号" : "项目符号"} icon={<Icon name={index ? "numbering" : "list"} size={17} />}
            disabled={!canFormat} split={() => word ? wordCommand({ type: "word", action: "list", kind }) : onCommand(index ? "numbering" : "bullets")}>
            {close => <><div className="office-menu-heading">{index ? "编号库" : "项目符号库"}</div><div className="office-list-gallery">{(index ? ["1. 2. 3.", "a. b. c.", "I. II. III."] : ["•", "○", "▪"]).map((sample, i) => <button key={sample} role="menuitem" type="button"
              disabled={i > 0 && !word} onClick={() => { if (word && i > 0) wordCommand({ type: "word", action: "listSettings" }); else if (word) wordCommand({ type: "word", action: "list", kind }); else onCommand(index ? "numbering" : "bullets"); close(); }}><span>{sample}</span><i>━━━━<br />━━━━<br />━━━━</i></button>)}</div>
              {word && <button role="menuitem" onClick={() => { wordCommand({ type: "word", action: "list", kind: "none" }); close(); }}>无</button>}
              {word && <button role="menuitem" onClick={() => { wordCommand({ type: "word", action: "listSettings" }); close(); }}>定义新的{index ? "编号格式" : "项目符号"}…</button>}</>}
          </OfficeMenu>)}
          {word && <OfficeMenu label="多级列表" icon={<Icon name="numbering" size={17} />} disabled={!canFormat}>
            {close => <><div className="office-menu-heading">列表库</div>{Array.from({ length: 9 }, (_, level) => <button key={level} role="menuitem" onClick={() => { wordCommand({ type: "word", action: "list", kind: "numbering", level }); close(); }}>{level + 1} 级编号</button>)}
              <button role="menuitem" onClick={() => { wordCommand({ type: "word", action: "listSettings" }); close(); }}>定义新的多级列表…</button></>}
          </OfficeMenu>}
          <span className="ribbon-divider" />
          <Tool label="减少缩进量" icon="outdent" disabled={!word || !canFormat} run={word ? () => wordCommand({ type: "word", action: "indent", direction: -1 }) : undefined} />
          <Tool label="增加缩进量" icon="indent" disabled={!word || !canFormat} run={word ? () => wordCommand({ type: "word", action: "indent", direction: 1 }) : undefined} />
          {word ? <><OfficeMenu label="中文版式" icon={<Icon name="textDirection" size={17} />} disabled={!canFormat}>
            {close => <button role="menuitem" onClick={() => { setParagraphOpen(true); close(); }}>段落…</button>}
          </OfficeMenu><Tool label="排序" icon="sort" />
          <Tool label="显示编辑标记" glyph="¶" disabled={!canFormat} active={showMarks} run={() => { setShowMarks(!showMarks); wordCommand({ type: "word", action: "showMarks", visible: !showMarks }); }} /></> : <>
            <OfficeMenu label="行距" icon={<Icon name="lineSpacing" size={17} />} disabled={!states.get("paragraphSettings")?.enabled}>
              {close => <button role="menuitem" onClick={() => { onCommand("paragraphSettings"); close(); }}>行距选项…</button>}
            </OfficeMenu><Tool label="栏" icon="columns" />
          </>}
        </>}
      </div>
      <div className="ribbon-icon-row">
        {(["left", "center", "right"] as const).map((value, index) => <Tool key={value} label={["左对齐", "居中", "右对齐"][index]!} icon={value} disabled={!canFormat} active={format.alignment === value} run={() => onTextCommand({ type: "align", value })} />)}
        {kind !== "workbook" && <><Tool label="两端对齐" icon="justify" disabled={!word || !canFormat} active={format.alignment === "justify"} run={word ? () => onTextCommand({ type: "align", value: "justify" }) : undefined} />
          <Tool label="分散对齐" icon="distributed" disabled={!word || !canFormat} run={word ? () => wordCommand({ type: "word", action: "paragraph", properties: { "text-align": "justify", "--word-distributed": "1" } }) : undefined} /></>}
        <span className="ribbon-divider" />
        {word ? <>
          <OfficeMenu label="行和段落间距" icon={<Icon name="lineSpacing" size={17} />} disabled={!canFormat}>
            {close => <>{[1, 1.15, 1.5, 2, 2.5, 3].map(value => <button key={value} role="menuitem" onClick={() => { wordCommand({ type: "word", action: "paragraph", properties: { "line-height": String(value), "--word-line-rule": "auto" } }); close(); }}>{value.toFixed(1)}</button>)}
              <button role="menuitem" onClick={() => { setParagraphOpen(true); close(); }}>行距选项…</button></>}
          </OfficeMenu>
          <OfficeColorMenu label="底纹" value="#ffffff" fill disabled={!canFormat} apply={value => wordCommand({ type: "word", action: "paragraph", properties: { "background-color": value } })} />
          <OfficeMenu label="边框" icon={<Icon name="table" size={17} />} disabled={!canFormat}>
            {close => <>{["下框线", "上框线", "左框线", "右框线", "无框线", "所有框线", "外侧框线"].map((label, i) => <button key={label} role="menuitem" onClick={() => {
              const sides = i < 4 ? [["bottom", "top", "left", "right"][i]!] : ["top", "right", "bottom", "left"];
              wordCommand({ type: "word", action: "paragraph", properties: Object.fromEntries(sides.map(side => ["border-" + side, i === 4 ? "0pt none #000000" : ".5pt solid #000000"])) }); close();
            }}>{label}</button>)}<button role="menuitem" onClick={() => { setBorderOpen(true); close(); }}>边框和底纹…</button></>}
          </OfficeMenu>
        </> : kind === "workbook" ? <>
          <Tool label="减少缩进量" icon="outdent" /> <Tool label="增加缩进量" icon="indent" />
          <OfficeMenu label="合并后居中" icon={<Icon name="merge" size={17} />} disabled={!states.get("mergeCells")?.enabled} split={() => onCommand("mergeCells")}>
            {close => <>{[["mergeCells", "合并单元格"], ["unmergeCells", "取消合并单元格"]].map(([id, label]) => <button key={id} role="menuitem" disabled={!states.get(id!)?.enabled} onClick={() => { onCommand(id!); close(); }}>{label}</button>)}</>}
          </OfficeMenu>
        </> : <><Tool label="文字方向" icon="textDirection" /><Tool label="对齐文本" icon="verticalCenter" /><Tool label="转换为 SmartArt" icon="diagram" /></>}
      </div>
    </Group>
  );
  const workbookAlignment = <Group title="对齐方式" className="workbook-alignment">
    <div className="workbook-alignment-left"><div className="ribbon-icon-row">
      {command("verticalTop", "顶端对齐", "verticalTop")}{command("verticalCenter", "垂直居中", "verticalCenter")}{command("verticalBottom", "底端对齐", "verticalBottom")}
      <OfficeMenu label="方向" icon={<Icon name="textDirection" size={17} />} disabled={!canFormat}>{close => <button role="menuitem" onClick={() => { onCommand("formatCells"); close(); }}>设置单元格对齐方式…</button>}</OfficeMenu>
    </div><div className="ribbon-icon-row">{(["left", "center", "right"] as const).map((value, index) => <Tool key={value} label={["左对齐", "居中", "右对齐"][index]!} icon={value} disabled={!canFormat} active={format.alignment === value} run={() => onTextCommand({ type: "align", value })} />)}
      {command("decreaseCellIndent", "减少缩进量", "outdent")}{command("increaseCellIndent", "增加缩进量", "indent")}
    </div></div>
    <div className="ribbon-tool-stack">{command("wrapText", "自动换行", "lineSpacing")}
      <OfficeMenu label="合并后居中" icon={<Icon name="merge" size={17} />} disabled={!states.get("mergeCells")?.enabled} split={() => onCommand("mergeCells")}>
        {close => <>{[["mergeCells", "合并单元格"], ["unmergeCells", "取消合并单元格"]].map(([id, label]) => <button key={id} role="menuitem" disabled={!states.get(id!)?.enabled} onClick={() => { onCommand(id!); close(); }}>{label}</button>)}</>}
      </OfficeMenu>
    </div>
  </Group>;
  const clipboard = (
    <Group title="剪贴板" className="clipboard-group">
      {command("paste", "粘贴", "paste", true)}
      <div className="ribbon-tool-stack">
        {command("cut", "剪切", "cut")}
        {command("copy", "复制", "copy")}
        <Tool
          label="格式刷"
          icon="brush"
          disabled={!word || !canFormat}
          active={Boolean(paint)}
          reason="先选择文字，再点击格式刷复制；选择目标文字后再次点击应用"
          run={
            word
              ? () => {
                  if (paint) {
                    wordCommand({ type: "word", action: "paintFormat", properties: paint });
                    setPaint(undefined);
                  } else
                    setPaint({
                      "font-weight": format.bold ? "700" : "400",
                      "font-style": format.italic ? "italic" : "normal",
                      "text-decoration":
                        [
                          format.underline ? "underline" : "",
                          format.strikethrough ? "line-through" : "",
                        ]
                          .filter(Boolean)
                          .join(" ") || "none",
                      ...(format.fontSize ? { "font-size": format.fontSize + "pt" } : {}),
                      ...(format.fontFamily ? { "font-family": format.fontFamily } : {}),
                      ...(format.color ? { color: format.color } : {}),
                    });
                }
              : undefined
          }
        />
      </div>
    </Group>
  );
  let content: ReactNode;
  if (tab === "文件")
    content = (
      <>
        <Group title="打开">{command("open", "打开", "folder", true)}</Group>
        <Group title="保存">
          {command("package", "另存为文档包", "save", true)}
          {command(
            "native",
            kind === "presentation" ? "导出 PPTX" : kind === "workbook" ? "导出 XLSX" : "导出 DOCX",
            "file",
            true,
          )}
        </Group>
        <Group title="导出">
          {command("pdf", "导出 PDF", "pdf", true)}
          {command("png", "导出 PNG", "image", true)}
        </Group>
        <Group title="打印">
          <Tool label="打印" icon="print" large />
        </Group>
      </>
    );
  else if (tab === "开始")
    content = (
      <>
        {clipboard}
        {kind === "presentation" && (
          <Group title="幻灯片">
            {command("newSlide", "新建幻灯片", "newSlide", true)}
            <div className="ribbon-tool-stack">
              {command("slideLayout", "版式", "grid")}
              <Tool label="重置" icon="reset" /><Tool label="节" icon="section" />
            </div>
          </Group>
        )}
        {fontGroup}
        {kind === "workbook" ? workbookAlignment : paragraph}
        {kind === "document" && (
          <Group title="样式" className="styles-group">
            <div className="style-gallery">
              {["正文", "无间隔", "标题 1", "标题 2", "标题 3", "标题 4", "标题", "副标题", "不明显强调", "强调", "明显强调", "要点", "引用", "明显引用", "书籍标题", ...customStyles].map((name, index) => (
                <button key={name} disabled={!canFormat} title={name} aria-label={name} aria-pressed={format.paragraphStyleId === styleIds[index]}
                  onMouseDown={event => event.preventDefault()} onClick={() => wordCommand({ type: "word", action: "style", styleId: styleIds[index]! })}>
                  <span style={{ fontFamily: index < 2 || index > 7 ? "Georgia,serif" : "Arial,sans-serif", fontSize: index === 2 || index === 6 ? 17 : 14, fontWeight: index >= 2 && index <= 6 ? 700 : 400, fontStyle: index >= 7 ? "italic" : "normal", color: index === 4 || index === 5 ? "#5485bd" : "#171717" }}>AaBbCcDdEe</span><small>{name}</small>
                </button>
              ))}
            </div>
            <OfficeMenu label="更多样式" icon={<><Icon name="styles" size={32} /><span className="gallery-chevron">⌄</span></>} disabled={!canFormat} className="style-gallery-menu">
              {close => <><div className="office-menu-heading">样式</div>{["正文", "标题 1", "标题 2", "标题 3", "标题 4", ...customStyles].map((name, index) => <button key={name} role="menuitem" onClick={() => { wordCommand({ type: "word", action: "style", styleId: index === 0 ? "Normal" : index < 5 ? "Heading" + index : name }); close(); }}>{name}</button>)}
                <button role="menuitem" onClick={() => { setStyleCreatorOpen(true); close(); }}>创建样式…</button>
                <button role="menuitem" onClick={() => { wordCommand({ type: "word", action: "styleSettings" }); close(); }}>管理样式…</button></>}
            </OfficeMenu>
          </Group>
        )}
        {kind === "workbook" && (
          <>
            <Group title="数字" className="number-group"><div className="ribbon-tool-stack">
              <OfficeComboBox label="数字格式" className="number-format-field" disabled={!canFormat} value={officeNumberLabel(format)} options={officeNumberPresets.map(([name]) => name)}
                commit={value => { const preset = officeNumberPresets.find(([label]) => label === value); if (!preset) return "请选择数字格式。"; return onCommand(preset[1]); }} more={{ label: "其他数字格式…", run: () => onCommand("formatCells") }} />
              <div className="ribbon-icon-row"><OfficeMenu label="会计数字格式" icon={<span>¥</span>} disabled={!canFormat} split={() => onCommand("numberAccounting")}>{close => <button role="menuitem" onClick={() => { onCommand("numberAccounting"); close(); }}>¥ 中文(中国)</button>}</OfficeMenu>
                <Tool label="百分比样式" glyph="%" disabled={!canFormat} run={() => onCommand("numberPercent")} />
                <Tool label="千位分隔样式" glyph="," disabled={!canFormat} run={() => onCommand("numberComma")} />
                <Tool label="增加小数位数" glyph={<span style={{ fontSize: 11 }}>←.0</span>} /><Tool label="减少小数位数" glyph={<span style={{ fontSize: 11 }}>.0→</span>} />
              </div>
            </div></Group>
            <Group title="样式">
              {command("rules", "条件格式", "conditionalFormat", true)}
              {command("tables", "套用表格格式", "tableFormat", true)}
              <Tool label="单元格样式" icon="cellStyles" large />
            </Group>
            <Group title="单元格"><div className="ribbon-tool-stack">
              {[["插入", "insertCells", [["insertRows", "插入工作表行"], ["insertColumns", "插入工作表列"]]], ["删除", "deleteCells", [["deleteRows", "删除工作表行"], ["deleteColumns", "删除工作表列"]]], ["格式", "table", [["formatCells", "设置单元格格式…"], ["worksheets", "工作表…"]]]].map(([label, icon, choices]) => <OfficeMenu key={label as string} label={label as string} icon={<Icon name={icon as Glyph} size={17} />} disabled={!canFormat}>
                {close => <>{(choices as string[][]).map(([id, name]) => <button key={id} role="menuitem" disabled={!states.get(id!)?.enabled} onClick={() => { onCommand(id!); close(); }}>{name}</button>)}</>}
              </OfficeMenu>)}
            </div></Group>
          </>
        )}
        {kind === "presentation" && <>
          <Group title="绘图"><div className="ribbon-tool-stack">{command("insertImage", "图片", "image")}{command("insertShape", "形状", "shapes")}{command("insertTextBox", "文本框", "textBox")}</div>
            <OfficeMenu label="排列" icon={<Icon name="arrange" size={32} />} large disabled={busy}>{close => <>{[["bringToFront", "置于顶层"], ["sendToBack", "置于底层"], ["groupObjects", "组合"], ["ungroupObjects", "取消组合"], ["duplicateObject", "复制"], ["deleteObject", "删除"]].map(([id, label]) => <button key={id} role="menuitem" disabled={!states.get(id!)?.enabled} onClick={() => { onCommand(id!); close(); }}>{label}</button>)}</>}</OfficeMenu>
            <Tool label="快速样式" icon="quickStyles" large /><div className="ribbon-tool-stack"><OfficeColorMenu label="形状填充" value="#4472c4" fill disabled={!onObjectFormat || !states.get("shapeAppearance")?.enabled} apply={value => onObjectFormat?.({ "--office-fill-kind": "solid", "--office-fill-color": value, "--office-fill-alpha": "1" })} /><OfficeColorMenu label="形状轮廓" value="#44546a" disabled={!onObjectFormat || !states.get("shapeAppearance")?.enabled} apply={value => onObjectFormat?.({ "--office-line-kind": "solid", "--office-line-color": value, "--office-line-alpha": "1" })} /><Tool label="形状效果" icon="shapes" /></div>
          </Group>
          <Group title="加载项"><Tool label="加载项" icon="addins" large /></Group><Group title="设计器"><Tool label="设计器" icon="sparkles" large /></Group>
        </>}
        {word ? <>
          <Group title="样式窗格"><Tool label="样式窗格" icon="styles" large disabled={!canFormat} run={() => wordCommand({ type: "word", action: "styleSettings" })} /></Group>
          <Group title="听写"><Tool label="听写" icon="mic" large /></Group>
          <Group title="加载项"><Tool label="加载项" icon="addins" large /></Group>
          <Group title="编辑器"><Tool label="编辑器" icon="editor" large /></Group>
        </> : kind === "workbook" ? <><Group title="编辑"><div className="ribbon-tool-stack"><Tool label="自动求和" glyph="Σ" /><Tool label="填充" icon="download" /><Tool label="清除" icon="eraser" /></div>
          <OfficeMenu label="排序和筛选" icon={<Icon name="sortFilter" size={32} />} large disabled={busy}>{close => <>{[["sortAscending", "升序"], ["sortDescending", "降序"], ["filters", "筛选…"]].map(([id, name]) => <button key={id} role="menuitem" disabled={!states.get(id!)?.enabled} onClick={() => { onCommand(id!); close(); }}>{name}</button>)}</>}</OfficeMenu><OfficeMenu label="查找和选择" icon={<Icon name="search" size={32} />} large disabled={busy}>{close => <>{[["find", "查找…"], ["replace", "替换…"]].map(([id, label]) => <button key={id} role="menuitem" disabled={!states.get(id!)?.enabled} onClick={() => { onCommand(id!); close(); }}>{label}</button>)}</>}</OfficeMenu>
        </Group><Group title="加载项"><Tool label="加载项" icon="addins" large /></Group></> : null}
      </>
    );
  else content = <Group title="帮助"><Tool label="帮助" icon="book" large /></Group>;
  if (!["开始", "文件", "帮助"].includes(tab)) content = <NativeOfficeTabs kind={kind} tab={tab} commands={commands} actions={actions} canFormat={canFormat} busy={busy}
    onCommand={onCommand} onTextCommand={onTextCommand} onLayout={onLayout} onZoom={onZoom} paragraph={() => setParagraphOpen(true)} borders={() => setBorderOpen(true)} layout={layout} zoom={zoom} enableWorkbookZoom={enableWorkbookZoom} />;
  if (tab === "形状格式" && shapeSelected) content = <>
    <Group title="插入形状"><Tool label="形状" icon="shapes" large disabled={!canFormat} run={() => wordCommand({ type: "word", action: "insertShape" })} /></Group>
    <Group title="文本"><Tool label="编辑文字" icon="textBox" large disabled={!canFormat || format.drawingCount !== 1 || !["shape", "textBox"].includes(format.drawingKind!) || !!format.editingInPlace} run={() => wordCommand({ type: "word", action: "editDrawingText" })} /></Group>
    <Group title="形状样式"><Tool label="填充与线条" icon="shapes" large disabled={!canFormat} run={() => wordCommand({ type: "word", action: "drawingSettings" })} /></Group>
    <Group title="排列"><div className="ribbon-tool-stack"><Tool label="文字环绕" icon="arrange" disabled={!canFormat} run={() => wordCommand({ type: "word", action: "drawingSettings" })} /><Tool label="选择窗格" icon="arrange" disabled={!canFormat} run={() => wordCommand({ type: "word", action: "selectionPane" })} /></div></Group>
    <Group title="大小"><Tool label="大小和位置" icon="shapes" large disabled={!canFormat} run={() => wordCommand({ type: "word", action: "drawingSettings" })} /></Group>
  </>;
  return (
    <div ref={root} data-compact={compact} className="doffi-office-ribbon office-ribbon" data-kind={kind} data-mode={mode} data-collapsed={collapsed}>
      <style>{officeRibbonStyles}</style>
      <div className="office-tabs" role="tablist" aria-label="Office 功能区">
        {tabs.map((name) => (
          <button
            key={name}
            id={id + "-tab-" + name}
            role="tab"
            disabled={preview && !allowedPreviewTabs.includes(name)}
            aria-selected={tab === name}
            aria-controls={panelId}
            tabIndex={tab === name || !tabs.includes(tab) && name === "开始" ? 0 : -1}
            onClick={() => onTab(name)}
            onKeyDown={(event) => {
              const direction = event.key === "ArrowRight" ? 1 : event.key === "ArrowLeft" ? -1 : 0;
              if (direction || event.key === "Home" || event.key === "End") {
                event.preventDefault();
                const availableTabs = preview
                  ? tabs.filter((name) => allowedPreviewTabs.includes(name))
                  : tabs;
                const next =
                  event.key === "Home"
                    ? availableTabs[0]!
                    : event.key === "End"
                      ? availableTabs.at(-1)!
                      : availableTabs[
                          (availableTabs.indexOf(name) + direction + availableTabs.length) %
                            availableTabs.length
                        ]!;
                onTab(next);
                [
                  ...event.currentTarget.parentElement!.querySelectorAll<HTMLButtonElement>(
                    "[role=tab]",
                  ),
                ]
                  .find((button) => button.textContent === next)
                  ?.focus();
              }
            }}
          >
            {name}
          </button>
        ))}
      </div>
      {word && format.storyKind && format.storyKind !== "main" && !format.editingInPlace && (
        <div style={{ padding: "6px 12px", background: "#eef4ff" }}>
          正在编辑
          {format.storyKind === "header"
            ? "页眉"
            : format.storyKind === "footer"
              ? "页脚"
              : "辅助内容"}{" "}
          <button
            onMouseDown={(event) => event.preventDefault()}
            onClick={() => wordCommand({ type: "word", action: "mainStory" })}
          >
            返回正文
          </button>
        </div>
      )}
      {paragraphOpen && (
        <WordParagraphDialog current={format.paragraphProperties ?? {}}
          apply={(properties) => wordCommand({ type: "word", action: "paragraph", properties })}
          close={() => setParagraphOpen(false)}
        />
      )}
      {borderOpen && <WordParagraphDialog variant="borders" apply={properties => wordCommand({ type: "word", action: "paragraph", properties })} close={() => setBorderOpen(false)} />}
      {styleCreatorOpen && <OfficeDialog title="根据格式设置创建新样式" close={() => setStyleCreatorOpen(false)} footer={<><button onClick={() => setStyleCreatorOpen(false)}>取消</button><button className="primary" disabled={!styleName.trim()} onClick={() => {
        const name = styleName.trim(); wordCommand({ type: "word", action: "defineStyle", styleId: name, properties: {
          "font-weight": format.bold ? "700" : "400", "font-style": format.italic ? "italic" : "normal",
          ...(format.fontSize ? { "font-size": format.fontSize + "pt" } : {}), ...(format.fontFamily ? { "font-family": format.fontFamily } : {}), ...(format.color ? { color: format.color } : {}),
        } }); setCustomStyles([...new Set([...customStyles, name])]); setStyleName(""); setStyleCreatorOpen(false);
      }}>确定</button></>}><label>名称<input aria-label="样式名称" value={styleName} onChange={event => setStyleName(event.currentTarget.value)} /></label><div className="dialog-preview">AaBbCcDdEe</div></OfficeDialog>}
      <div
        id={panelId}
        role="tabpanel"
        aria-labelledby={tabs.includes(tab) ? id + "-tab-" + tab : undefined}
        aria-label={tabs.includes(tab) ? undefined : tab}
        className="office-ribbon-panel"
        hidden={collapsed}
      >
        {content}
      </div>
    </div>
  );
}
