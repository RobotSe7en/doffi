import { useEffect, useId, useRef, useState } from "react";
import { WordParagraphDialog } from "./WordParagraphDialog.js";
import { OfficeComboBox } from "./OfficeComboBox.js";
import { OfficeMenu, OfficeColorMenu } from "./OfficeMenu.js";
import { OfficeDialog } from "./OfficeDialog.js";
import type { ReactNode, ReactElement } from "react";
import type { DocumentTextCommand, EditorCommandState, TextFormatState, WordPageZoomMode } from "../index.js";
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
  pageZoomMode?: WordPageZoomMode | undefined;
  onPageZoomMode?(mode: WordPageZoomMode): void;
  draftMode?: boolean;
  onDraftMode?(enabled: boolean): void;
  focusMode?: boolean;
  onFocusMode?(enabled: boolean): void;
  showGridlines?: boolean;
  onToggleGridlines?(): void;
  immersiveReader?: boolean;
  onToggleImmersiveReader?(): void;
  splitWindow?: boolean;
  onToggleSplitWindow?(): void;
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
  pageZoomMode,
  onPageZoomMode,
  draftMode,
  onDraftMode,
  focusMode,
  onFocusMode,
  showGridlines,
  onToggleGridlines,
  immersiveReader,
  onToggleImmersiveReader,
  splitWindow,
  onToggleSplitWindow,
}: OfficeRibbonProps): ReactElement {
  const [localPageZoomMode, setLocalPageZoomMode] = useState<WordPageZoomMode>();
  const effectivePageZoomMode = onPageZoomMode ? pageZoomMode : localPageZoomMode;
  const handlePageZoomMode = (mode: WordPageZoomMode) => {
    if (onPageZoomMode) {
      onPageZoomMode(mode);
      return;
    }
    setLocalPageZoomMode(mode);
    onLayout("pages");
    onZoom(mode === "pageWidth" ? 1.25 : mode === "singlePage" ? 0.65 : 0.45);
  };
  const handleZoom = (value: number) => {
    if (!onPageZoomMode) setLocalPageZoomMode(undefined);
    onZoom(value);
  };
  const handleLayout = (value: "pages" | "continuous") => {
    if (!onPageZoomMode) setLocalPageZoomMode(undefined);
    onLayout(value);
  };
  const root = useRef<HTMLDivElement>(null), [compact, setCompact] = useState(false);
  useEffect(() => {
    const element = root.current!;
    let frame = 0;
    const observer = new ResizeObserver(() => {
      cancelAnimationFrame(frame);
      frame = requestAnimationFrame(() => setCompact(element.clientWidth <= 1400));
    });
    observer.observe(element);
    return () => {
      cancelAnimationFrame(frame);
      observer.disconnect();
    };
  }, []);
  const id = useId(),
    panelId = id + "-panel";
  const [paragraphOpen, setParagraphOpen] = useState(false),
    [showMarks, setShowMarks] = useState(false);
  const [paint, setPaint] = useState<{
    run: Record<string, string>;
    paragraph: Record<string, string>;
    sticky: boolean;
  }>();
  useEffect(() => {
    if (!paint) return;
    const onKey = (event: KeyboardEvent) => {
      if (event.key === "Escape") setPaint(undefined);
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [paint]);
  const [styleName, setStyleName] = useState("");
  const [customStyles, setCustomStyles] = useState<string[]>([]);
  const [styleCreatorOpen, setStyleCreatorOpen] = useState(false);
  const [borderOpen, setBorderOpen] = useState(false);
  const [recentFonts, setRecentFonts] = useState<string[]>([]);
  const [localOsFonts, setLocalOsFonts] = useState<string[]>([]);
  const [keyTips, setKeyTips] = useState(false);
  const word = kind === "document";

  const TAB_KEY_TIPS: Readonly<Record<string, string>> = {
    文件: "F",
    开始: "H",
    插入: "N",
    绘图: "JI",
    设计: "G",
    布局: "P",
    引用: "S",
    邮件: "M",
    审阅: "R",
    视图: "W",
    帮助: "Y",
  };

  useEffect(() => {
    const nav = navigator as Navigator & {
      permissions?: { query(desc: { name: string }): Promise<{ state: string }> };
    };
    const win = window as Window & {
      queryLocalFonts?: () => Promise<readonly { readonly family: string }[]>;
    };
    if (!win.queryLocalFonts) return;
    let active = true;
    void (async () => {
      try {
        const perm = await nav.permissions?.query({ name: "local-fonts" });
        if (perm && perm.state !== "granted") return;
        const list = await win.queryLocalFonts!();
        if (!active) return;
        setLocalOsFonts([...new Set(list.map((f) => f.family).filter(Boolean))].slice(0, 120));
      } catch {
        // Local font access is optional; gracefully fall back to web + built-in fonts.
      }
    })();
    return () => {
      active = false;
    };
  }, []);

  const loadOsFontsOnDemand = () => {
    const win = window as Window & {
      queryLocalFonts?: () => Promise<readonly { readonly family: string }[]>;
    };
    if (!win.queryLocalFonts || localOsFonts.length) return;
    void win
      .queryLocalFonts()
      .then((list) =>
        setLocalOsFonts([...new Set(list.map((f) => f.family).filter(Boolean))].slice(0, 120)),
      )
      .catch(() => {});
  };

  const previewRun = (properties: Readonly<Record<string, string>> | null) => {
    if (!word || !canFormat) return;
    wordCommand({ type: "word", action: "livePreview", scope: "run", properties });
  };
  const styleGalleryItems = (() => {
    const accent = format.themeColors?.[4] ?? "#2f5496";
    const sub = format.themeColors?.[3] ?? "#595959";
    const builtIns: readonly {
      readonly id: string;
      readonly label: string;
      readonly sampleStyle: Record<string, string | number>;
      readonly preview: Readonly<Record<string, string>>;
    }[] = [
      { id: "Normal", label: "正文", sampleStyle: { fontFamily: "等线,Calibri,sans-serif", fontSize: 12, fontWeight: 400, color: "#1f1f1f" }, preview: { "font-size": "10.5pt", "font-weight": "400", "font-style": "normal", color: "#1f1f1f" } },
      { id: "NoSpacing", label: "无间隔", sampleStyle: { fontFamily: "等线,Calibri,sans-serif", fontSize: 12, fontWeight: 400, color: "#1f1f1f" }, preview: { "font-size": "10.5pt", "font-weight": "400", "font-style": "normal", color: "#1f1f1f" } },
      { id: "Heading1", label: "标题 1", sampleStyle: { fontFamily: "等线 Light,Arial,sans-serif", fontSize: 13, fontWeight: 700, letterSpacing: "-0.35px", color: accent }, preview: { "font-size": "16pt", "font-weight": "700", "font-style": "normal", color: accent } },
      { id: "Heading2", label: "标题 2", sampleStyle: { fontFamily: "等线 Light,Arial,sans-serif", fontSize: 12.5, fontWeight: 700, letterSpacing: "-0.3px", color: accent }, preview: { "font-size": "13pt", "font-weight": "700", "font-style": "normal", color: accent } },
      { id: "Heading3", label: "标题 3", sampleStyle: { fontFamily: "等线 Light,Arial,sans-serif", fontSize: 12, fontWeight: 600, letterSpacing: "-0.2px", color: "#1f3763" }, preview: { "font-size": "12pt", "font-weight": "600", "font-style": "normal", color: "#1f3763" } },
      { id: "Heading4", label: "标题 4", sampleStyle: { fontFamily: "等线 Light,Arial,sans-serif", fontSize: 11.5, fontWeight: 600, fontStyle: "italic", letterSpacing: "-0.2px", color: accent }, preview: { "font-size": "11pt", "font-weight": "600", "font-style": "italic", color: accent } },
      { id: "Title", label: "标题", sampleStyle: { fontFamily: "等线 Light,Arial,sans-serif", fontSize: 14, fontWeight: 300, letterSpacing: "-0.4px", color: "#171717" }, preview: { "font-size": "24pt", "font-weight": "300", "font-style": "normal", color: "#171717" } },
      { id: "Subtitle", label: "副标题", sampleStyle: { fontFamily: "等线,Calibri,sans-serif", fontSize: 12, fontWeight: 400, color: sub }, preview: { "font-size": "12pt", "font-weight": "400", "font-style": "normal", color: sub } },
      { id: "SubtleEmphasis", label: "不明显强调", sampleStyle: { fontFamily: "等线,Calibri,sans-serif", fontSize: 11.5, fontStyle: "italic", color: sub }, preview: { "font-style": "italic", color: sub } },
      { id: "Emphasis", label: "强调", sampleStyle: { fontFamily: "等线,Calibri,sans-serif", fontSize: 11.5, fontStyle: "italic", color: "#1f1f1f" }, preview: { "font-style": "italic", color: "#1f1f1f" } },
      { id: "IntenseEmphasis", label: "明显强调", sampleStyle: { fontFamily: "等线,Calibri,sans-serif", fontSize: 11.5, fontStyle: "italic", color: accent }, preview: { "font-style": "italic", color: accent } },
      { id: "Strong", label: "要点", sampleStyle: { fontFamily: "等线,Calibri,sans-serif", fontSize: 11.5, fontWeight: 700, letterSpacing: "-0.2px", color: "#1f1f1f" }, preview: { "font-weight": "700", color: "#1f1f1f" } },
      { id: "Quote", label: "引用", sampleStyle: { fontFamily: "Georgia,serif", fontSize: 11.5, fontStyle: "italic", color: "#404040" }, preview: { "font-style": "italic", color: "#404040" } },
      { id: "IntenseQuote", label: "明显引用", sampleStyle: { fontFamily: "Georgia,serif", fontSize: 11.5, fontStyle: "italic", color: accent, borderBottom: `1px solid ${accent}` }, preview: { "font-style": "italic", color: accent } },
      { id: "BookTitle", label: "书籍标题", sampleStyle: { fontFamily: "等线,Calibri,sans-serif", fontSize: 11, fontWeight: 700, fontVariant: "small-caps", letterSpacing: "0.2px", color: "#1f1f1f" }, preview: { "font-weight": "700", "letter-spacing": "0.5pt", color: "#1f1f1f" } },
    ];
    return [
      ...builtIns,
      ...customStyles.map((name) => ({
        id: name,
        label: name,
        sampleStyle: { fontFamily: "等线,Calibri,sans-serif", fontSize: 12.5, fontWeight: 600, color: accent },
        preview: { "font-weight": "600", color: accent },
      })),
    ];
  })();
  const activeStyleId = format.paragraphStyleId || "Normal";
  const wordCommand = (value: Extract<DocumentTextCommand, { type: "word" }>) =>
    onTextCommand(value);

  const preview = mode === "preview",
    allowedPreviewTabs = ["开始", "文件", "审阅", "视图", "帮助"];
  const shapeSelected = word && !preview && !!format.drawingKind && ["shape", "textBox", "connector", "group"].includes(format.drawingKind);
  const imageSelected = word && !preview && format.drawingKind === "image";
  const chartSelected = word && !preview && format.drawingKind === "chart";
  const tableSelected = word && !preview && Boolean(format.inTable);
  const headerFooterSelected = word && !preview && (format.storyKind === "header" || format.storyKind === "footer");
  const contextualTabs = [
    ...(shapeSelected ? ["形状格式"] : []),
    ...(imageSelected ? ["图片格式"] : []),
    ...(chartSelected ? ["图表设计"] : []),
    ...(tableSelected ? ["表格设计", "表格布局"] : []),
    ...(headerFooterSelected ? ["页眉和页脚"] : []),
  ];
  const contextualOnly = ["形状格式", "图片格式", "图表设计", "表格设计", "表格布局", "页眉和页脚"];
  const tab =
    (preview && !allowedPreviewTabs.includes(requestedTab)) ||
    (contextualOnly.includes(requestedTab) && !contextualTabs.includes(requestedTab))
      ? "开始"
      : requestedTab,
    collapsed = preview && tab === "开始";
  const tabs = [...tabsByKind[kind].filter(name => showApplicationTabs || name !== "文件" && name !== "帮助"), ...contextualTabs],
    states = new Map(commands.map((command) => [command.id, command]));
  useEffect(() => {
    if (!word) return;
    let altAlone = false;
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key === "Alt" && !event.ctrlKey && !event.metaKey && !event.shiftKey) {
        altAlone = true;
        return;
      }
      altAlone = false;
      if (!keyTips) return;
      if (event.key === "Escape") {
        event.preventDefault();
        setKeyTips(false);
        return;
      }
      const upper = event.key.toUpperCase();
      const matchedTab = tabs.find((name) => TAB_KEY_TIPS[name] === upper);
      if (matchedTab) {
        event.preventDefault();
        onTab(matchedTab);
        setKeyTips(false);
      } else {
        setKeyTips(false);
      }
    };
    const onKeyUp = (event: KeyboardEvent) => {
      if (event.key === "Alt" && altAlone) {
        altAlone = false;
        setKeyTips((prev) => !prev);
      }
    };
    window.addEventListener("keydown", onKeyDown);
    window.addEventListener("keyup", onKeyUp);
    return () => {
      window.removeEventListener("keydown", onKeyDown);
      window.removeEventListener("keyup", onKeyUp);
    };
  }, [word, keyTips, tabs, onTab]);
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
  const rawThemeFonts = [
    format.wordState?.themeMajorEastAsia,
    format.wordState?.themeMinorEastAsia,
    format.wordState?.themeMajorLatin,
    format.wordState?.themeMinorLatin,
  ].filter((v): v is string => Boolean(v));
  const themeFontList = [
    ...new Set(rawThemeFonts.length ? rawThemeFonts : word ? ["等线", "Calibri"] : []),
  ];
  const cjkFonts = [
    "等线",
    "宋体",
    "黑体",
    "楷体",
    "仿宋",
    "微软雅黑",
    "PingFang SC",
    "Noto Sans CJK SC",
    "Noto Serif CJK SC",
    "FandolSong",
    "FandolHei",
    "FandolKai",
    "FandolFang",
    "LXGW WenKai",
  ];
  const cjkFontSet = new Set(cjkFonts.map((f) => f.toLowerCase()));
  const families = [
    ...new Set(
      [
        ...themeFontList,
        ...recentFonts,
        fontName(format.fontFamily),
        ...cjkFonts,
        ...(webFonts?.filter(font => !font.fallbackOnly).flatMap(font => [font.family, ...font.aliases]) ?? []),
        "Arial",
        "Calibri",
        "Aptos",
        "Times New Roman",
        "Georgia",
        "Liberation Serif",
        "Liberation Sans",
        "Liberation Mono",
        ...localOsFonts,
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
    <Group
      title="字体"
      className="font-group"
      launcher={
        word
          ? {
              label: "字体设置",
              disabled: !canFormat,
              run: () => {
                loadOsFontsOnDemand();
                wordCommand({ type: "word", action: "fontSettings" });
              },
            }
          : undefined
      }
    >
      <div className="ribbon-font-selects">
        <OfficeComboBox
          label="字体"
          className="font-family-field"
          disabled={!canFormat}
          value={fontName(format.fontFamily)}
          options={families}
          more={kind !== "workbook" ? { label: "字体…", run: () => { loadOsFontsOnDemand(); if (kind === "document") wordCommand({ type: "word", action: "fontSettings" }); else onCommand("fontSettings"); } } : undefined}
          section={(name) =>
            themeFontList.includes(name)
              ? "主题字体"
              : recentFonts.includes(name)
                ? "最近使用的字体"
                : cjkFontSet.has(name.toLowerCase())
                  ? "常用中文字体"
                  : "西文与系统字体"
          }
          previewFamily={(name) => {
            const face = webFonts?.find(font => !font.fallbackOnly && [font.family, ...font.aliases].some(alias => alias.toLowerCase() === name.toLowerCase()));
            return face?.family ?? name;
          }}
          preview={(name) => {
            loadOsFontsOnDemand();
            if (webFonts?.length) void loadWordWebFonts(document, webFonts, new Set([name.toLowerCase()]));
            const quoted = JSON.stringify(name);
            previewRun({
              "font-family": quoted,
              "--word-font-ascii": quoted,
              "--word-font-hAnsi": quoted,
              "--word-font-eastAsia": quoted,
              "--word-font-cs": quoted,
            });
          }}
          onPreviewEnd={() => previewRun(null)}
          commit={(value) => {
            previewRun(null);
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
          preview={(value) => {
            const size = kind === "document" && wordSizeNames[value] !== undefined ? wordSizeNames[value]! : Number(value);
            if (Number.isFinite(size) && size >= 1 && size <= 1638)
              previewRun({ "font-size": size + "pt" });
          }}
          onPreviewEnd={() => previewRun(null)}
          commit={(value) => {
            previewRun(null);
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
          {close => <>{([
            ["句首字母大写", "sentence"],
            ["全部小写", "lower"],
            ["全部大写", "upper"],
            ["每个单词首字母大写", "title"],
            ["切换大小写", "toggle"],
          ] as const).map(([label, mode]) =>
            <button key={mode} role="menuitem" type="button" onClick={() => { wordCommand({ type: "word", action: "changeCase", mode }); close(); }}>{label}</button>)}</>}
        </OfficeMenu>
        <Tool
          label="清除所有格式"
          icon="clearFormat"
          disabled={!word || !canFormat}
          run={word ? () => wordCommand({ type: "word", action: "clearFormat" }) : undefined}
        />
        </>}
        {word && <><Tool label="拼音指南" icon="pinyin" disabled={!canFormat} active={format.wordState?.rubyActive} run={() => wordCommand({ type: "word", action: "ruby" })} /><Tool label="字符边框" icon="fontBorder" disabled={!canFormat} active={format.characterBorder}
          run={() => wordCommand({ type: "word", action: "paintFormat", properties: { "--word-run-border": format.characterBorder ? "none" : "0.5pt solid currentColor" } })} /></>}
      </div>
      <div className="ribbon-icon-row">
        {toggle("加粗", "bold", <b>B</b>)}
        {toggle("斜体", "italic", <i>I</i>)}
        <OfficeMenu label="下划线" active={format.underline} icon={<span className="ribbon-glyph"><u>U</u></span>} disabled={!canFormat}
          split={() => onTextCommand({ type: "toggle", property: "underline" })}>
          {close => <>{([
            ["单线", "single", "solid"],
            ["双线", "double", "double"],
            ["粗线", "thick", "solid"],
            ["点线", "dotted", "dotted"],
            ["虚线", "dash", "dashed"],
            ["点划线", "dotDash", "dashed"],
            ["双点划线", "dotDotDash", "dashed"],
            ["波浪线", "wave", "wavy"],
            ["仅字下划线", "words", "solid"],
            ["无", "none", "solid"],
          ] as const).map(([name, nativeStyle, cssStyle]) => (
            <button key={nativeStyle} role="menuitem" type="button" disabled={!word && nativeStyle !== "single" && nativeStyle !== "none"}
              onClick={() => {
                if (word) {
                  const lines = [
                    nativeStyle === "none" ? "" : "underline",
                    format.strikethrough ? "line-through" : "",
                  ].filter(Boolean).join(" ") || "none";
                  wordCommand({
                    type: "word",
                    action: "paintFormat",
                    properties: {
                      "text-decoration-line": lines,
                      "text-decoration-style": cssStyle,
                      "--word-underline-style": nativeStyle,
                    },
                  });
                } else {
                  onTextCommand({ type: "format", property: "text-decoration", value: nativeStyle === "none" ? "none" : "underline" });
                }
                close();
              }}>{name}</button>
          ))}
            {word && <button role="menuitem" type="button" onClick={() => { wordCommand({ type: "word", action: "fontSettings" }); close(); }}>其他下划线…</button>}</>}
        </OfficeMenu>
        {kind !== "workbook" && <>
        <Tool
          label="删除线"
          glyph={<s>abc</s>}
          disabled={!canFormat}
          active={format.strikethrough}
          run={() => {
            const nextLine =
              [format.underline ? "underline" : "", !format.strikethrough ? "line-through" : ""]
                .filter(Boolean)
                .join(" ") || "none";
            if (word)
              wordCommand({
                type: "word",
                action: "paintFormat",
                properties: {
                  "text-decoration-line": nextLine,
                  ...(!format.strikethrough ? { "--word-dstrike": "0" } : {}),
                },
              });
            else onTextCommand({ type: "format", property: "text-decoration", value: nextLine });
          }}
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
        {word && <OfficeMenu label="文本效果" icon={<Icon name="textEffects" size={17} />} disabled={!canFormat}>
          {close => <div onPointerLeave={() => previewRun(null)}>
            <div className="office-menu-heading">文本效果和版式</div>
            {([
              ["无效果", { "--word-run-native-outline": "0", "--word-run-native-shadow": "0", "--word-run-native-emboss": "0", "--word-run-native-imprint": "0", "--word-text-glow": "none", "--word-text-outline": "none", "--word-text-shadow": "none", "--word-text-reflection": "0", "--word-text-fill-gradient": "none", "--word-text-bevel": "none", "--word-text-3d": "none" }],
              ["空心轮廓", { "--word-run-native-outline": "1" }],
              ["文字阴影", { "--word-run-native-shadow": "1" }],
              ["阳文浮雕", { "--word-run-native-emboss": "1" }],
              ["阴文雕刻", { "--word-run-native-imprint": "1" }],
              ["蓝色发光", { "--word-text-glow": "6pt #4472c4" }],
              ["橙色发光", { "--word-text-glow": "6pt #ed7d31" }],
              ["深蓝轮廓", { "--word-text-outline": "1pt #2f5496" }],
              ["柔和投影", { "--word-text-shadow": "2pt 2pt 45 #44546a" }],
              ["文字倒影", { "--word-text-reflection": "1" }],
              ["蓝色渐变填充", { "--word-text-fill-gradient": "90deg #5b9bd5 0% #1f4e79 100%" }],
              ["橙金渐变·三维棱台", { "--word-text-fill-gradient": "135deg #ffc000 0% #ed7d31 55% #c00000 100%", "--word-text-bevel": "circle 3pt 3pt", "--word-text-3d": "orthographicFront" }],
              ["三维浮雕棱台", { "--word-text-bevel": "relaxedInset 4pt 4pt", "--word-text-3d": "perspectiveFront" }],
            ] as const).map(([label, properties]) => (
              <button key={label} role="menuitem" type="button"
                onPointerEnter={() => previewRun(properties)}
                onClick={() => { previewRun(null); wordCommand({ type: "word", action: "paintFormat", properties }); close(); }}>{label}</button>
            ))}
            <button role="menuitem" type="button" onClick={() => { previewRun(null); wordCommand({ type: "word", action: "wordArt" }); close(); }}>插入艺术字库…</button>
          </div>}
        </OfficeMenu>}
        {kind !== "workbook" && <OfficeColorMenu label="文本突出显示颜色" value="#ffff00" highlight disabled={!word || !canFormat}
          onHover={c => previewRun(c ? { "background-color": c } : null)}
          apply={value => onTextCommand({ type: "format", property: "background-color", value })} />}
        {kind === "workbook" && <><OfficeMenu label="边框" icon={<Icon name="table" size={17} />} disabled={!canFormat}>
          {close => <button role="menuitem" type="button" onClick={() => { onCommand("formatCells"); close(); }}>设置单元格格式…</button>}
        </OfficeMenu><OfficeColorMenu label="填充颜色" value="#ffff00" fill disabled={!canFormat}
          apply={value => onTextCommand({ type: "format", property: "background-color", value })} /></>}
        <OfficeColorMenu label="字体颜色" value="#ff0000" disabled={!canFormat} themeColors={format.themeColors}
          onHover={c => previewRun(c ? { color: c } : null)}
          apply={value => onTextCommand({ type: "format", property: "color", value })} />
        {word && <><Tool label="字符底纹" glyph={<span style={{ background: "#ccc", padding: "0 2px" }}>A</span>}
          disabled={!canFormat} run={() => onTextCommand({ type: "format", property: "background-color", value: "#d9d9d9" })} />
          <Tool label="带圈字符" icon="enclose" disabled={!canFormat} run={() => wordCommand({ type: "word", action: "enclosedChar" })} /></>}
        {kind === "workbook" && <Tool label="显示拼音字段" icon="pinyin" />}
      </div>
    </Group>
  );
  const paragraph = (
    <Group
      title={kind === "workbook" ? "对齐方式" : "段落"}
      className="paragraph-group"
      launcher={
        word
          ? {
              label: "段落设置",
              disabled: !canFormat,
              run: () => setParagraphOpen(true),
            }
          : undefined
      }
    >
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
          {(["bullet", "numbering"] as const).map((listKind, index) => <OfficeMenu key={listKind} label={index ? "编号" : "项目符号"} icon={<Icon name={index ? "numbering" : "list"} size={17} />}
            disabled={!canFormat} split={() => word ? wordCommand({ type: "word", action: "list", kind: listKind }) : onCommand(index ? "numbering" : "bullets")}>
            {close => <><div className="office-menu-heading">{index ? "编号库" : "项目符号库"}</div><div className="office-list-gallery">{(index
              ? ([["1. 2. 3.", "decimal"], ["a. b. c.", "lowerLetter"], ["I. II. III.", "upperRoman"]] as const)
              : ([["•", "disc"], ["○", "circle"], ["▪", "square"]] as const)
            ).map(([sample, preset], i) => <button key={sample} role="menuitem" type="button"
              disabled={i > 0 && !word} onClick={() => { if (word) wordCommand({ type: "word", action: "list", kind: listKind, preset }); else onCommand(index ? "numbering" : "bullets"); close(); }}><span>{sample}</span><i>━━━━<br />━━━━<br />━━━━</i></button>)}</div>
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
            {close => <>
              <button role="menuitem" type="button" onClick={() => { wordCommand({ type: "word", action: "ruby" }); close(); }}>拼音指南…</button>
              <button role="menuitem" type="button" onClick={() => { wordCommand({ type: "word", action: "enclosedChar" }); close(); }}>带圈字符…</button>
              <button role="menuitem" type="button" onClick={() => { wordCommand({ type: "word", action: "chineseLayout", mode: "combine" }); close(); }}>双行合一…</button>
              <button role="menuitem" type="button" onClick={() => { wordCommand({ type: "word", action: "chineseLayout", mode: "vert" }); close(); }}>纵横混排…</button>
              <button role="menuitem" type="button" onClick={() => { wordCommand({ type: "word", action: "chineseLayout", mode: "fit" }); close(); }}>调整宽度…</button>
              <button role="menuitem" type="button" onClick={() => { setParagraphOpen(true); close(); }}>段落…</button>
            </>}
          </OfficeMenu><Tool label="排序" icon="sort" disabled={!canFormat} run={() => wordCommand({ type: "word", action: "sort" })} />
          <Tool label="显示编辑标记" glyph="¶" disabled={!canFormat} active={showMarks} run={() => { setShowMarks(!showMarks); wordCommand({ type: "word", action: "showMarks", visible: !showMarks }); }} /></> : <>
            <OfficeMenu label="行距" icon={<Icon name="lineSpacing" size={17} />} disabled={!states.get("paragraphSettings")?.enabled}>
              {close => <button role="menuitem" onClick={() => { onCommand("paragraphSettings"); close(); }}>行距选项…</button>}
            </OfficeMenu><Tool label="栏" icon="columns" />
          </>}
        </>}
      </div>
      <div className="ribbon-icon-row">
        {(["left", "center", "right"] as const).map((value, index) => <Tool key={value} label={["左对齐", "居中", "右对齐"][index]!} icon={value} disabled={!canFormat} active={format.alignment === value} run={() => onTextCommand({ type: "align", value })} />)}
        {kind !== "workbook" && <><Tool label="两端对齐" icon="justify" disabled={!word || !canFormat} active={format.alignment === "justify" && format.paragraphProperties?.["--word-distributed"] !== "1"} run={word ? () => onTextCommand({ type: "align", value: "justify" }) : undefined} />
          <Tool label="分散对齐" icon="distributed" disabled={!word || !canFormat} active={format.paragraphProperties?.["--word-distributed"] === "1"} run={word ? () => wordCommand({ type: "word", action: "paragraph", properties: { "text-align": "justify", "--word-distributed": "1" } }) : undefined} /></>}
        <span className="ribbon-divider" />
        {word ? <>
          <OfficeMenu label="行和段落间距" icon={<Icon name="lineSpacing" size={17} />} disabled={!canFormat}>
            {close => <>{[1, 1.15, 1.5, 2, 2.5, 3].map(value => <button key={value} role="menuitem" onClick={() => { wordCommand({ type: "word", action: "paragraph", properties: { "line-height": String(value), "--word-line-rule": "auto" } }); close(); }}>{value.toFixed(1)}</button>)}
              <button role="menuitem" onClick={() => { setParagraphOpen(true); close(); }}>行距选项…</button></>}
          </OfficeMenu>
          <OfficeColorMenu label="底纹" value="#ffffff" fill disabled={!canFormat} themeColors={format.themeColors} apply={value => wordCommand({ type: "word", action: "paragraph", properties: { "background-color": value } })} />
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
      {word ? (
        <OfficeMenu
          label="粘贴"
          icon={<Icon name="paste" size={32} />}
          large
          disabled={busy || (!canFormat && (!actions.paste || Boolean(actions.paste.disabled)))}
          split={() => actions.paste?.run()}
        >
          {(close) => (
            <>
              <button role="menuitem" type="button" onClick={() => { actions.paste?.run(); close(); }}>保留源格式</button>
              <button role="menuitem" type="button" disabled={!actions.pasteMerge || actions.pasteMerge.disabled} onClick={() => { actions.pasteMerge?.run(); close(); }}>合并格式</button>
              <button role="menuitem" type="button" disabled={!actions.pastePlain || actions.pastePlain.disabled} onClick={() => { actions.pastePlain?.run(); close(); }}>仅保留文本</button>
              <button role="menuitem" type="button" disabled={!canFormat} onClick={() => { wordCommand({ type: "word", action: "pasteSpecial" }); close(); }}>选择性粘贴…</button>
            </>
          )}
        </OfficeMenu>
      ) : (
        command("paste", "粘贴", "paste", true)
      )}
      <div className="ribbon-tool-stack">
        {command("cut", "剪切", "cut")}
        {command("copy", "复制", "copy")}
        <Tool
          label={paint?.sticky ? "格式刷（已锁定，按 Esc 退出）" : "格式刷"}
          icon="brush"
          disabled={!word || !canFormat}
          active={Boolean(paint)}
          reason="单击复制格式并应用一次；双击锁定连续应用（按 Esc 退出）"
          onDoubleClick={
            word
              ? () => {
                  const charProps = format.characterProperties ?? {};
                  const paraProps = format.paragraphProperties ?? {};
                  const run: Record<string, string> = {
                    "font-weight": format.bold ? "700" : "400",
                    "font-style": format.italic ? "italic" : "normal",
                    "text-decoration-line":
                      [
                        format.underline ? "underline" : "",
                        format.strikethrough ? "line-through" : "",
                      ]
                        .filter(Boolean)
                        .join(" ") || "none",
                    ...(format.fontSize ? { "font-size": format.fontSize + "pt" } : {}),
                    ...(format.fontFamily ? { "font-family": format.fontFamily } : {}),
                    ...(format.color ? { color: format.color } : {}),
                    ...(format.backgroundColor && format.backgroundColor !== "rgba(0, 0, 0, 0)"
                      ? { "background-color": format.backgroundColor }
                      : {}),
                    ...(format.verticalAlign && ["sub", "super", "baseline"].includes(format.verticalAlign)
                      ? { "vertical-align": format.verticalAlign }
                      : {}),
                    ...(charProps["text-decoration-style"] ? { "text-decoration-style": charProps["text-decoration-style"] } : {}),
                    ...(charProps["text-decoration-color"] ? { "text-decoration-color": charProps["text-decoration-color"] } : {}),
                    ...(charProps["--word-underline-style"] ? { "--word-underline-style": charProps["--word-underline-style"] } : {}),
                    ...(charProps["--word-dstrike"] ? { "--word-dstrike": charProps["--word-dstrike"] } : {}),
                    ...(charProps["--word-run-border"] ? { "--word-run-border": charProps["--word-run-border"] } : {}),
                  };
                  const paragraph: Record<string, string> = {};
                  for (const key of ["text-align", "line-height", "--word-line-rule", "margin-left", "margin-right", "text-indent", "margin-top", "margin-bottom"]) {
                    if (paraProps[key]) paragraph[key] = paraProps[key]!;
                  }
                  setPaint({ run, paragraph, sticky: true });
                }
              : undefined
          }
          run={
            word
              ? () => {
                  if (paint) {
                    wordCommand({
                      type: "word",
                      action: "paintFormat",
                      properties: paint.run,
                      ...(Object.keys(paint.paragraph).length ? { paragraphProperties: paint.paragraph } : {}),
                    });
                    if (!paint.sticky) setPaint(undefined);
                  } else {
                    const charProps = format.characterProperties ?? {};
                    const paraProps = format.paragraphProperties ?? {};
                    const run: Record<string, string> = {
                      "font-weight": format.bold ? "700" : "400",
                      "font-style": format.italic ? "italic" : "normal",
                      "text-decoration-line":
                        [
                          format.underline ? "underline" : "",
                          format.strikethrough ? "line-through" : "",
                        ]
                          .filter(Boolean)
                          .join(" ") || "none",
                      ...(format.fontSize ? { "font-size": format.fontSize + "pt" } : {}),
                      ...(format.fontFamily ? { "font-family": format.fontFamily } : {}),
                      ...(format.color ? { color: format.color } : {}),
                      ...(format.backgroundColor && format.backgroundColor !== "rgba(0, 0, 0, 0)"
                        ? { "background-color": format.backgroundColor }
                        : {}),
                      ...(format.verticalAlign && ["sub", "super", "baseline"].includes(format.verticalAlign)
                        ? { "vertical-align": format.verticalAlign }
                        : {}),
                      ...(charProps["text-decoration-style"] ? { "text-decoration-style": charProps["text-decoration-style"] } : {}),
                      ...(charProps["text-decoration-color"] ? { "text-decoration-color": charProps["text-decoration-color"] } : {}),
                      ...(charProps["--word-underline-style"] ? { "--word-underline-style": charProps["--word-underline-style"] } : {}),
                      ...(charProps["--word-dstrike"] ? { "--word-dstrike": charProps["--word-dstrike"] } : {}),
                      ...(charProps["--word-run-border"] ? { "--word-run-border": charProps["--word-run-border"] } : {}),
                    };
                    const paragraph: Record<string, string> = {};
                    for (const key of ["text-align", "line-height", "--word-line-rule", "margin-left", "margin-right", "text-indent", "margin-top", "margin-bottom"]) {
                      if (paraProps[key]) paragraph[key] = paraProps[key]!;
                    }
                    setPaint({ run, paragraph, sticky: false });
                  }
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
          {command("print", "打印", "print", true)}
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
            <div className="style-gallery" onPointerLeave={() => previewRun(null)}>
              {styleGalleryItems.map((item) => (
                <button
                  key={item.id}
                  type="button"
                  disabled={!canFormat}
                  title={item.label}
                  aria-label={item.label}
                  aria-pressed={activeStyleId === item.id}
                  onMouseDown={(event) => event.preventDefault()}
                  onPointerEnter={() => previewRun(item.preview)}
                  onClick={() => {
                    previewRun(null);
                    wordCommand({ type: "word", action: "style", styleId: item.id });
                  }}
                >
                  <span style={item.sampleStyle}>AaBbCcDd</span>
                  <small>{item.label}</small>
                </button>
              ))}
            </div>
            <OfficeMenu label="更多样式" icon={<><Icon name="styles" size={32} /><span className="gallery-chevron" aria-hidden="true">⌄</span></>} disabled={!canFormat} className="style-gallery-menu">
              {close => <div onPointerLeave={() => previewRun(null)}>
                <div className="office-menu-heading">样式</div>
                <div className="office-style-grid">
                  {styleGalleryItems.map((item) => (
                    <button
                      key={item.id}
                      role="menuitemradio"
                      type="button"
                      aria-label={item.label}
                      aria-checked={activeStyleId === item.id}
                      onPointerEnter={() => previewRun(item.preview)}
                      onClick={() => {
                        previewRun(null);
                        wordCommand({ type: "word", action: "style", styleId: item.id });
                        close();
                      }}
                    >
                      <span style={item.sampleStyle}>AaBbCcDd</span>
                      <small>{item.label}</small>
                    </button>
                  ))}
                </div>
                <div className="office-menu-divider" />
                <button role="menuitem" type="button" onClick={() => { previewRun(null); wordCommand({ type: "word", action: "clearFormat" }); close(); }}>清除格式</button>
                <button role="menuitem" type="button" onClick={() => { previewRun(null); setStyleCreatorOpen(true); close(); }}>创建样式…</button>
                <button role="menuitem" type="button" onClick={() => { previewRun(null); wordCommand({ type: "word", action: "styleSettings" }); close(); }}>管理样式…</button>
              </div>}
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
          <Group title="编辑">
            <div className="ribbon-tool-stack">
              <Tool label="查找" icon="search" disabled={busy} run={() => wordCommand({ type: "word", action: "find" })} />
              <Tool label="替换" icon="edit" disabled={!canFormat} run={() => wordCommand({ type: "word", action: "replace" })} />
              <Tool label="定位" icon="arrow" disabled={busy} run={() => wordCommand({ type: "word", action: "goTo" })} />
            </div>
          </Group>
          <Group title="听写"><Tool label="听写" icon="mic" large /></Group>
          <Group title="加载项"><Tool label="加载项" icon="addins" large /></Group>
          <Group title="编辑器"><Tool label="编辑器" icon="editor" large /></Group>
        </> : kind === "workbook" ? <><Group title="编辑"><div className="ribbon-tool-stack"><Tool label="自动求和" glyph="Σ" /><Tool label="填充" icon="download" /><Tool label="清除" icon="eraser" /></div>
          <OfficeMenu label="排序和筛选" icon={<Icon name="sortFilter" size={32} />} large disabled={busy}>{close => <>{[["sortAscending", "升序"], ["sortDescending", "降序"], ["filters", "筛选…"]].map(([id, name]) => <button key={id} role="menuitem" disabled={!states.get(id!)?.enabled} onClick={() => { onCommand(id!); close(); }}>{name}</button>)}</>}</OfficeMenu><OfficeMenu label="查找和选择" icon={<Icon name="search" size={32} />} large disabled={busy}>{close => <>{[["find", "查找…"], ["replace", "替换…"]].map(([id, label]) => <button key={id} role="menuitem" disabled={!states.get(id!)?.enabled} onClick={() => { onCommand(id!); close(); }}>{label}</button>)}</>}</OfficeMenu>
        </Group><Group title="加载项"><Tool label="加载项" icon="addins" large /></Group></> : null}
      </>
    );
  else content = <Group title="帮助"><Tool label="帮助" icon="book" large /></Group>;
  if (!["开始", "文件", "帮助"].includes(tab)) content = <NativeOfficeTabs kind={kind} tab={tab} commands={commands} actions={actions} canFormat={canFormat} busy={busy} format={format}
    onCommand={onCommand} onTextCommand={onTextCommand} onLayout={handleLayout} onZoom={handleZoom} paragraph={() => setParagraphOpen(true)} borders={() => setBorderOpen(true)} layout={layout} zoom={zoom} enableWorkbookZoom={enableWorkbookZoom}
    pageZoomMode={effectivePageZoomMode} onPageZoomMode={word ? handlePageZoomMode : undefined} draftMode={draftMode} onDraftMode={onDraftMode} focusMode={focusMode} onFocusMode={onFocusMode}
    showGridlines={showGridlines} onToggleGridlines={onToggleGridlines} immersiveReader={immersiveReader} onToggleImmersiveReader={onToggleImmersiveReader} splitWindow={splitWindow} onToggleSplitWindow={onToggleSplitWindow} />;
  const renderWrapMenu = (menuLabel: string) => (
    <OfficeMenu label={menuLabel} icon={<Icon name="arrange" size={17} />} disabled={!canFormat}>
      {close => <>
        <button role="menuitem" type="button" onClick={() => { wordCommand({ type: "word", action: "setDrawingWrap", mode: "inline" }); close(); }}>嵌入型</button>
        <button role="menuitem" type="button" onClick={() => { wordCommand({ type: "word", action: "setDrawingWrap", mode: "floating", wrap: "square", behindText: false }); close(); }}>四周型环绕</button>
        <button role="menuitem" type="button" onClick={() => { wordCommand({ type: "word", action: "setDrawingWrap", mode: "floating", wrap: "tight", behindText: false }); close(); }}>紧密型环绕</button>
        <button role="menuitem" type="button" onClick={() => { wordCommand({ type: "word", action: "setDrawingWrap", mode: "floating", wrap: "through", behindText: false }); close(); }}>穿越型环绕</button>
        <button role="menuitem" type="button" onClick={() => { wordCommand({ type: "word", action: "setDrawingWrap", mode: "floating", wrap: "topAndBottom", behindText: false }); close(); }}>上下型环绕</button>
        <button role="menuitem" type="button" onClick={() => { wordCommand({ type: "word", action: "setDrawingWrap", mode: "floating", wrap: "none", behindText: true }); close(); }}>衬于文字下方</button>
        <button role="menuitem" type="button" onClick={() => { wordCommand({ type: "word", action: "setDrawingWrap", mode: "floating", wrap: "none", behindText: false }); close(); }}>浮于文字上方</button>
        <div className="office-menu-divider" />
        <button role="menuitem" type="button" onClick={() => { wordCommand({ type: "word", action: "drawingSettings" }); close(); }}>其他布局选项…</button>
      </>}
    </OfficeMenu>
  );
  const renderRotateMenu = () => (
    <OfficeMenu label="旋转" icon={<Icon name="refresh" size={17} />} disabled={!canFormat}>
      {close => <>
        <button role="menuitem" type="button" onClick={() => { wordCommand({ type: "word", action: "transformDrawing", mode: "rotateRight" }); close(); }}>向右旋转 90°</button>
        <button role="menuitem" type="button" onClick={() => { wordCommand({ type: "word", action: "transformDrawing", mode: "rotateLeft" }); close(); }}>向左旋转 90°</button>
        <button role="menuitem" type="button" onClick={() => { wordCommand({ type: "word", action: "transformDrawing", mode: "flipV" }); close(); }}>垂直翻转</button>
        <button role="menuitem" type="button" onClick={() => { wordCommand({ type: "word", action: "transformDrawing", mode: "flipH" }); close(); }}>水平翻转</button>
        <div className="office-menu-divider" />
        <button role="menuitem" type="button" onClick={() => { wordCommand({ type: "word", action: "drawingSettings" }); close(); }}>其他旋转选项…</button>
      </>}
    </OfficeMenu>
  );
  if (tab === "形状格式" && shapeSelected) content = <>
    <Group title="插入形状"><Tool label="形状" icon="shapes" large disabled={!canFormat} run={() => wordCommand({ type: "word", action: "insertShape" })} /></Group>
    <Group title="文本"><Tool label="编辑文字" icon="textBox" large disabled={!canFormat || format.drawingCount !== 1 || !["shape", "textBox"].includes(format.drawingKind!) || !!format.editingInPlace} run={() => wordCommand({ type: "word", action: "editDrawingText" })} /></Group>
    <Group title="形状样式"><Tool label="填充与线条" icon="shapes" large disabled={!canFormat} run={() => wordCommand({ type: "word", action: "drawingSettings" })} /></Group>
    <Group title="排列">
      <div className="ribbon-tool-stack">
        {renderWrapMenu("文字环绕")}
        <Tool label="前移一层" icon="arrow" disabled={!canFormat} run={() => wordCommand({ type: "word", action: "bringForward" })} />
        <Tool label="后移一层" icon="arrow" disabled={!canFormat} run={() => wordCommand({ type: "word", action: "sendBackward" })} />
      </div>
      <div className="ribbon-tool-stack">
        <Tool label="选择窗格" icon="panel" disabled={!canFormat} run={() => wordCommand({ type: "word", action: "selectionPane" })} />
        <Tool label="组合" icon="grid" disabled={!canFormat} run={() => wordCommand({ type: "word", action: "groupDrawings" })} />
        {renderRotateMenu()}
      </div>
    </Group>
    <Group title="大小"><Tool label="大小和位置" icon="shapes" large disabled={!canFormat} run={() => wordCommand({ type: "word", action: "drawingSettings" })} /></Group>
  </>;
  if (tab === "图片格式" && imageSelected) content = <>
    <Group title="调整"><Tool label="裁剪与属性" icon="image" large disabled={!canFormat} run={() => wordCommand({ type: "word", action: "drawingSettings" })} /></Group>
    <Group title="排列">
      <div className="ribbon-tool-stack">
        <Tool label="位置" icon="grid" disabled={!canFormat} run={() => wordCommand({ type: "word", action: "drawingSettings" })} />
        {renderWrapMenu("自动换行")}
        <Tool label="选择窗格" icon="panel" disabled={!canFormat} run={() => wordCommand({ type: "word", action: "selectionPane" })} />
      </div>
      <div className="ribbon-tool-stack">
        <Tool label="前移一层" icon="arrow" disabled={!canFormat} run={() => wordCommand({ type: "word", action: "bringForward" })} />
        <Tool label="后移一层" icon="arrow" disabled={!canFormat} run={() => wordCommand({ type: "word", action: "sendBackward" })} />
        {renderRotateMenu()}
      </div>
    </Group>
    <Group title="大小"><Tool label="大小和旋转" icon="shapes" large disabled={!canFormat} run={() => wordCommand({ type: "word", action: "drawingSettings" })} /></Group>
  </>;
  if (tab === "图表设计" && chartSelected) content = <>
    <Group title="数据与类型"><Tool label="编辑图表" icon="chart" large disabled={!canFormat} run={() => wordCommand({ type: "word", action: "chartSettings" })} /></Group>
    <Group title="排列">
      <div className="ribbon-tool-stack">
        {renderWrapMenu("文字环绕")}
        <Tool label="前移一层" icon="arrow" disabled={!canFormat} run={() => wordCommand({ type: "word", action: "bringForward" })} />
        <Tool label="后移一层" icon="arrow" disabled={!canFormat} run={() => wordCommand({ type: "word", action: "sendBackward" })} />
      </div>
      <Tool label="选择窗格" icon="panel" large disabled={!canFormat} run={() => wordCommand({ type: "word", action: "selectionPane" })} />
    </Group>
  </>;
  if (tab === "表格设计" && tableSelected) content = <>
    <Group title="表格样式选项">
      <div className="ribbon-tool-stack">
        <label><input type="checkbox" checked={format.wordState?.tableFirstRow ?? true} disabled={!canFormat} onChange={e => wordCommand({ type: "word", action: "setTableStyleFlag", flag: "firstRow", enabled: e.currentTarget.checked })} />标题行</label>
        <label><input type="checkbox" checked={Boolean(format.wordState?.tableLastRow)} disabled={!canFormat} onChange={e => wordCommand({ type: "word", action: "setTableStyleFlag", flag: "lastRow", enabled: e.currentTarget.checked })} />汇总行</label>
        <label><input type="checkbox" checked={format.wordState?.tableBandedRows ?? true} disabled={!canFormat} onChange={e => wordCommand({ type: "word", action: "setTableStyleFlag", flag: "noHBand", enabled: !e.currentTarget.checked })} />镶边行</label>
      </div>
      <div className="ribbon-tool-stack">
        <label><input type="checkbox" checked={Boolean(format.wordState?.tableFirstCol)} disabled={!canFormat} onChange={e => wordCommand({ type: "word", action: "setTableStyleFlag", flag: "firstCol", enabled: e.currentTarget.checked })} />第一列</label>
        <label><input type="checkbox" checked={Boolean(format.wordState?.tableLastCol)} disabled={!canFormat} onChange={e => wordCommand({ type: "word", action: "setTableStyleFlag", flag: "lastCol", enabled: e.currentTarget.checked })} />最后一列</label>
        <label><input type="checkbox" checked={Boolean(format.wordState?.tableBandedCols)} disabled={!canFormat} onChange={e => wordCommand({ type: "word", action: "setTableStyleFlag", flag: "noVBand", enabled: !e.currentTarget.checked })} />镶边列</label>
      </div>
    </Group>
    <Group title="表格样式">
      <div className="office-table-style-gallery">
        {([
          ["grid", "普通网格", "#fff", "#fff", "#999", "#999"],
          ["blueHeader", "蓝色表头", "#2f5496", "#edf2f8", "#8eaadb", "#8eaadb"],
          ["banded", "浅色条带", "#4472c4", "#f2f5fb", "#d0d7e5", "#d0d7e5"],
          ["academic3Line", "三线表", "#fff", "#fff", "#111", "transparent"],
        ] as const).map(([preset, title, hdr, alt, stroke, inner]) => (
          <button key={preset} type="button" disabled={!canFormat} title={title} aria-label={title}
            onMouseDown={e => e.preventDefault()}
            onClick={() => wordCommand({ type: "word", action: "applyTableQuickStyle", preset })}>
            <svg width="46" height="30" viewBox="0 0 46 30" aria-hidden="true">
              <rect x="1" y="1" width="44" height="9" fill={hdr} stroke={stroke} strokeWidth="1" />
              <rect x="1" y="10" width="44" height="9" fill={alt} stroke={inner} strokeWidth="0.7" />
              <rect x="1" y="19" width="44" height="10" fill="#fff" stroke={inner} strokeWidth="0.7" />
              <line x1="1" y1="29" x2="45" y2="29" stroke={stroke} strokeWidth={preset === "academic3Line" ? "1.4" : "1"} />
              <line x1="1" y1="1" x2="45" y2="1" stroke={stroke} strokeWidth={preset === "academic3Line" ? "1.4" : "1"} />
              {inner !== "transparent" && <><line x1="16" y1="1" x2="16" y2="29" stroke={inner} strokeWidth="0.7" /><line x1="31" y1="1" x2="31" y2="29" stroke={inner} strokeWidth="0.7" /></>}
            </svg>
            <small>{title}</small>
          </button>
        ))}
      </div>
      <Tool label="表格属性" icon="table" large disabled={!canFormat} run={() => wordCommand({ type: "word", action: "tableSettings" })} />
    </Group>
    <Group title="边框与底纹">
      <OfficeColorMenu label="单元格底纹" value="#f2f2f2" fill large disabled={!canFormat} themeColors={format.themeColors} apply={value => wordCommand({ type: "word", action: "tableShading", color: value })} />
      <OfficeMenu label="斜线表头" icon={<Icon name="table" size={32} />} large disabled={!canFormat}>
        {close => <>
          <button role="menuitem" type="button" onClick={() => { wordCommand({ type: "word", action: "tableDiagonalBorder", mode: "tl2br" }); close(); }}>下斜框线（左上→右下）</button>
          <button role="menuitem" type="button" onClick={() => { wordCommand({ type: "word", action: "tableDiagonalBorder", mode: "tr2bl" }); close(); }}>上斜框线（右上→左下）</button>
          <button role="menuitem" type="button" onClick={() => { wordCommand({ type: "word", action: "tableDiagonalBorder", mode: "none" }); close(); }}>清除斜框线</button>
        </>}
      </OfficeMenu>
      <Tool label="边框和底纹" icon="table" large disabled={!canFormat} run={() => setBorderOpen(true)} />
    </Group>
  </>;
  if (tab === "表格布局" && tableSelected) content = <>
    <Group title="表">
      <Tool label="属性" icon="table" large disabled={!canFormat} run={() => wordCommand({ type: "word", action: "tableSettings" })} />
      <Tool label="删除表格" icon="close" large disabled={!canFormat} run={() => wordCommand({ type: "word", action: "tableAction", operation: "deleteTable" })} />
    </Group>
    <Group title="绘制">
      <Tool label="绘制表格" icon="pen" large disabled={!canFormat} active={format.wordState?.tableToolMode === "draw"} run={() => wordCommand({ type: "word", action: "setTableTool", mode: format.wordState?.tableToolMode === "draw" ? "off" : "draw" })} />
      <Tool label="擦除" icon="eraser" large disabled={!canFormat} active={format.wordState?.tableToolMode === "erase"} run={() => wordCommand({ type: "word", action: "setTableTool", mode: format.wordState?.tableToolMode === "erase" ? "off" : "erase" })} />
    </Group>
    <Group title="行和列">
      <div className="ribbon-tool-stack">
        <Tool label="在上方插入行" icon="insertCells" disabled={!canFormat} run={() => wordCommand({ type: "word", action: "tableAction", operation: "insertRowAbove" })} />
        <Tool label="在下方插入行" icon="insertCells" disabled={!canFormat} run={() => wordCommand({ type: "word", action: "tableAction", operation: "insertRowBelow" })} />
        <Tool label="删除行" icon="deleteCells" disabled={!canFormat} run={() => wordCommand({ type: "word", action: "tableAction", operation: "deleteRow" })} />
      </div>
      <div className="ribbon-tool-stack">
        <Tool label="在左侧插入列" icon="insertCells" disabled={!canFormat} run={() => wordCommand({ type: "word", action: "tableAction", operation: "insertColLeft" })} />
        <Tool label="在右侧插入列" icon="insertCells" disabled={!canFormat} run={() => wordCommand({ type: "word", action: "tableAction", operation: "insertColRight" })} />
        <Tool label="删除列" icon="deleteCells" disabled={!canFormat} run={() => wordCommand({ type: "word", action: "tableAction", operation: "deleteCol" })} />
      </div>
    </Group>
    <Group title="合并">
      <Tool label="合并/拆分" icon="merge" large disabled={!canFormat} run={() => wordCommand({ type: "word", action: "tableSettings" })} />
      <Tool label="拆分表格" icon="columns" large disabled={!canFormat} run={() => wordCommand({ type: "word", action: "tableAction", operation: "splitTable" })} />
    </Group>
    <Group title="单元格大小">
      <div className="ribbon-tool-stack">
        <Tool label="分布行" icon="lineSpacing" disabled={!canFormat} run={() => wordCommand({ type: "word", action: "tableAction", operation: "distributeRows" })} />
        <Tool label="分布列" icon="columns" disabled={!canFormat} run={() => wordCommand({ type: "word", action: "tableAction", operation: "distributeCols" })} />
        <OfficeMenu label="自动调整" icon={<Icon name="table" size={17} />} disabled={!canFormat}>
          {close => <>
            <button role="menuitem" type="button" onClick={() => { wordCommand({ type: "word", action: "tableAction", operation: "autoFitContents" }); close(); }}>根据内容自动调整表格</button>
            <button role="menuitem" type="button" onClick={() => { wordCommand({ type: "word", action: "tableAction", operation: "autoFitWindow" }); close(); }}>根据窗口自动调整表格</button>
          </>}
        </OfficeMenu>
      </div>
    </Group>
    <Group title="对齐方式">
      <div className="office-align-grid" role="group" aria-label="单元格九宫格对齐">
        {([
          ["left", "top", "靠上左对齐", 3, 4],
          ["center", "top", "靠上居中", 6, 4],
          ["right", "top", "靠上右对齐", 9, 4],
          ["left", "middle", "水平左对齐", 3, 9],
          ["center", "middle", "水平居中", 6, 9],
          ["right", "middle", "水平右对齐", 9, 9],
          ["left", "bottom", "靠下左对齐", 3, 14],
          ["center", "bottom", "靠下居中", 6, 14],
          ["right", "bottom", "靠下右对齐", 9, 14],
        ] as const).map(([horizontal, vertical, title, x, y]) => (
          <button key={title} type="button" disabled={!canFormat} title={title} aria-label={title}
            onMouseDown={e => e.preventDefault()}
            onClick={() => wordCommand({ type: "word", action: "setTableCellAlign", horizontal, vertical })}>
            <svg width="16" height="16" viewBox="0 0 18 18" fill="none" stroke="currentColor" strokeWidth="1.4" aria-hidden="true">
              <line x1={x} y1={y - 1.5} x2={x + 6} y2={y - 1.5} />
              <line x1={horizontal === "right" ? x + 2 : horizontal === "center" ? x + 1 : x} y1={y + 1.5} x2={horizontal === "left" ? x + 4 : horizontal === "center" ? x + 5 : x + 6} y2={y + 1.5} />
            </svg>
          </button>
        ))}
      </div>
      <OfficeMenu label="文字方向" icon={<Icon name="textDirection" size={32} />} large disabled={!canFormat}>
        {close => <>
          <button role="menuitem" type="button" onClick={() => { wordCommand({ type: "word", action: "setTableCellDirection", direction: "lrTb" }); close(); }}>水平（从左到右）</button>
          <button role="menuitem" type="button" onClick={() => { wordCommand({ type: "word", action: "setTableCellDirection", direction: "tbRl" }); close(); }}>垂直（自右向左）</button>
          <button role="menuitem" type="button" onClick={() => { wordCommand({ type: "word", action: "setTableCellDirection", direction: "btLr" }); close(); }}>垂直（自左向右）</button>
        </>}
      </OfficeMenu>
    </Group>
    <Group title="数据">
      <Tool label="排序" icon="sort" large disabled={!canFormat} run={() => wordCommand({ type: "word", action: "sort" })} />
      <Tool label="转换为文本" icon="textBox" large disabled={!canFormat} run={() => wordCommand({ type: "word", action: "tableToText" })} />
      <Tool label="公式" icon="numbering" large disabled={!canFormat} run={() => wordCommand({ type: "word", action: "tableFormula" })} />
    </Group>
  </>;
  if (tab === "页眉和页脚" && headerFooterSelected) content = <>
    <Group title="页眉和页脚">
      <Tool label="页眉" icon="header" large disabled={!canFormat} run={() => wordCommand({ type: "word", action: "editHeader" })} />
      <Tool label="页脚" icon="footer" large disabled={!canFormat} run={() => wordCommand({ type: "word", action: "editFooter" })} />
      <Tool label="页码" icon="numbering" large disabled={!canFormat} run={() => wordCommand({ type: "word", action: "pageNumber" })} />
    </Group>
    <Group title="选项">
      <div className="ribbon-tool-stack">
        <label><input type="checkbox" checked={Boolean(format.wordState?.differentFirstPage)} disabled={!canFormat} onChange={event => wordCommand({ type: "word", action: "headerFooterOptions", differentFirstPage: event.currentTarget.checked })} />首页不同</label>
        <label><input type="checkbox" checked={Boolean(format.wordState?.evenAndOddHeaders)} disabled={!canFormat} onChange={event => wordCommand({ type: "word", action: "headerFooterOptions", evenAndOddHeaders: event.currentTarget.checked })} />奇偶页不同</label>
      </div>
    </Group>
    <Group title="位置">
      <Tool label="页眉页脚设置" icon="margins" large disabled={!canFormat} run={() => wordCommand({ type: "word", action: "pageSettings" })} />
    </Group>
    <Group title="关闭">
      <Tool label="关闭页眉和页脚" icon="close" large disabled={busy} run={() => wordCommand({ type: "word", action: "mainStory" })} />
    </Group>
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
            {keyTips && TAB_KEY_TIPS[name] && (
              <kbd className="office-keytip" style={{ marginLeft: 4, padding: "0 4px", fontSize: 10, lineHeight: "14px", borderRadius: 2, background: "#24292f", color: "#fff", fontFamily: "system-ui" }}>
                {TAB_KEY_TIPS[name]}
              </kbd>
            )}
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
          apply={(properties, outlineLevel) => wordCommand({ type: "word", action: "paragraph", properties, ...(outlineLevel !== undefined ? { outlineLevel } : {}) })}
          onTabs={() => { setParagraphOpen(false); wordCommand({ type: "word", action: "tabSettings" }); }}
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
