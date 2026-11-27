import { useState } from "react";
import type { ReactElement } from "react";
import { OfficeDialog } from "./OfficeDialog.js";

/** Draft values become one source transaction on confirmation. */
export function WordParagraphDialog({ apply, close, variant = "paragraph", current = {}, onTabs }: {
  apply(properties: Record<string, string>, outlineLevel?: number | null): void | Promise<void>; close(): void; variant?: "paragraph" | "borders"; current?: Readonly<Record<string, string>>; onTabs?(): void;
}): ReactElement {
  const [numbers, setNumbers] = useState<Record<string, string>>({}), [choices, setChoices] = useState<Record<string, string>>({});
  const [color, setColor] = useState("#000000"), [fill, setFill] = useState("#ffffff"), [borderWidth, setBorderWidth] = useState("0.5");
  const [error, setError] = useState("");
  const fontSize = parseFloat(current["font-size"] ?? "16") * .75 || 12;
  const initialNumber = (key: string, fallback: string) => {
    const source = current[key]; if (!source || source === "normal") return fallback;
    const points = parseFloat(source) * (source.endsWith("px") ? .75 : 1);
    if (!Number.isFinite(points)) return fallback;
    const value = key === "line-height" ? current["--word-line-rule"] === "exact" || current["--word-line-rule"] === "atLeast" ? points : points / fontSize : points / fontSize;
    return String(Math.round(value * 100) / 100);
  };
  const border = variant === "borders";
  const choose = (label: string, key: string, values: readonly (readonly [string, string])[], initial = "") => <label>{label}<select aria-label={label} value={choices[key] ?? current[key] ?? initial}
    onChange={event => setChoices({ ...choices, [key]: event.currentTarget.value })}>{!initial && <option value=""></option>}{values.map(([value, title]) => <option key={value} value={value}>{title}</option>)}</select></label>;
  const number = (label: string, key: string, initial = "0") => <label>{label}<input aria-label={label} type="number" step="0.1" value={numbers[key] ?? initialNumber(key, initial)}
    onChange={event => setNumbers({ ...numbers, [key]: event.currentTarget.value })} /></label>;
  const flag = (key: string, title: string, yes = "1", no = "0") => <label><input type="checkbox" aria-label={title} checked={(choices[key] ?? current[key] ?? (["widows", "--word-kinsoku", "--word-paragraph-snap-to-grid", "--word-overflow-punctuation", "--word-auto-space-de", "--word-auto-space-dn"].includes(key) ? yes : no)) === yes}
    onChange={event => setChoices({ ...choices, [key]: event.currentTarget.checked ? yes : no })} />{title}</label>;
  async function commit(): Promise<void> {
    setError("");
    try {
    const properties: Record<string, string> = {};
    let outlineLevel: number | null | undefined;
    if (border) {
      if (choices.border) for (const side of ["top", "right", "bottom", "left"]) properties["border-" + side] = choices.border === "none" ? "0pt none " + color : borderWidth + "pt " + choices.border + " " + color;
      if (choices.fill) properties["background-color"] = choices.fill === "none" ? "transparent" : fill;
    } else {
      for (const [key, value] of Object.entries(numbers)) {
        if (!value) continue;
        const n = Number(value); if (!Number.isFinite(n)) return;
        properties[key] = key === "line-height" ? (choices.rule ?? current["--word-line-rule"] ?? "auto") === "auto" ? String(n) : n + "pt" : n * fontSize + "pt";
      }
      if (numbers["line-height"] || choices.rule) properties["--word-line-rule"] = choices.rule || current["--word-line-rule"] || "auto";
      for (const [key, value] of Object.entries(choices)) if (key !== "rule" && key !== "special" && key !== "outlineLevel") properties[key] = value;
      if (choices.outlineLevel !== undefined && choices.outlineLevel !== "") {
        outlineLevel = choices.outlineLevel === "body" ? null : Number(choices.outlineLevel);
      }
      if (choices.special === "hanging" && properties["text-indent"]) properties["text-indent"] = -Math.abs(parseFloat(properties["text-indent"])) + "pt";
      if (choices.special === "none") properties["text-indent"] = "0pt";
      if (choices.widows) properties.orphans = choices.widows;
    }
    if (Object.keys(properties).length || outlineLevel !== undefined) await apply(properties, outlineLevel);
    close();
    } catch (reason) { setError(reason instanceof Error ? reason.message : String(reason)); }
  }
  return <OfficeDialog className="word-paragraph-dialog" title={border ? "边框和底纹" : "段落"} tabs={border ? ["边框", "底纹"] : ["缩进和间距", "换行和分页", "中文版式"]} close={close}
    footer={<>{!border && onTabs && <button type="button" style={{ marginRight: "auto" }} onClick={onTabs}>制表位…</button>}<button type="button" onClick={close}>取消</button><button type="button" className="primary" onClick={() => { void commit(); }}>确定</button></>}>
    {tab => <><style>{`.word-paragraph-dialog{width:580px}.word-paragraph-dialog .paragraph-flags{grid-template-columns:1fr}.word-paragraph-dialog .dialog-content{min-height:485px}.word-paragraph-dialog fieldset{margin-bottom:8px}.word-paragraph-dialog .dialog-stack{gap:8px}`}</style>{error && <p role="alert">{error}</p>}{border ? <>
      <div hidden={tab !== 0} className="dialog-stack">
        <fieldset><legend>设置</legend>{choose("设置", "border", [["none", "无"], ["solid", "方框"], ["double", "双线"], ["dotted", "点线"], ["dashed", "虚线"]])}
          <label>颜色<input type="color" aria-label="段落边框颜色" value={color} onChange={event => setColor(event.currentTarget.value)} /></label>
          <label>宽度<input type="number" aria-label="段落边框宽度（磅）" min="0" step=".25" value={borderWidth} onChange={event => setBorderWidth(event.currentTarget.value)} /></label>
          <div className="dialog-preview" style={{ fontSize: 13, margin: 0, border: choices.border === "none" ? "none" : `${borderWidth}px ${choices.border || "solid"} ${color}` }}>预览<br />文字　文字　文字</div>
        </fieldset>
      </div>
      <div hidden={tab !== 1} className="dialog-stack"><fieldset><legend>填充</legend>{choose("填充", "fill", [["none", "无颜色"], ["set", "设置颜色"]])}
        <label>颜色<input type="color" aria-label="段落底纹颜色" value={fill} onChange={event => { setFill(event.currentTarget.value); setChoices({ ...choices, fill: "set" }); }} /></label></fieldset>
        <div className="dialog-preview" style={{ backgroundColor: choices.fill === "none" ? "white" : fill, fontSize: 13 }}>文字　文字　文字</div>
      </div>
    </> : <>
      <div hidden={tab !== 0} className="dialog-stack">
        <fieldset><legend>常规</legend>{choose("对齐方式:", "text-align", [["left", "左对齐"], ["center", "居中"], ["right", "右对齐"], ["justify", "两端对齐"]])}
          {choose("大纲级别:", "outlineLevel", [["body", "正文文本"], ...Array.from({ length: 9 }, (_, i) => [String(i), `${i + 1} 级`] as const)], current["--doffi-outline-level"] ?? "body")}</fieldset>
        <fieldset><legend>缩进</legend>{number("左侧（字符）", "margin-left")}{number("右侧（字符）", "margin-right")}
          {choose("特殊格式:", "special", [["none", "无"], ["firstLine", "首行缩进"], ["hanging", "悬挂缩进"]])}{number("缩进值（字符）", "text-indent")}</fieldset>
        <fieldset><legend>间距</legend>{number("段前（行）", "margin-top")}{number("段后（行）", "margin-bottom")}
          {choose("行距类型", "rule", [["auto", "多倍行距"], ["exact", "固定值"], ["atLeast", "最小值"]], "auto")}{number(choices.rule === "exact" || choices.rule === "atLeast" ? "设置值（磅）" : "设置值（倍）", "line-height", "1")}</fieldset>
        {flag("--word-contextual-spacing", "相同样式的段落之间不添加间距")}{flag("--word-paragraph-snap-to-grid", "如果定义了文档网格，则对齐到网格")}
      </div>
      <div hidden={tab !== 1} className="dialog-stack"><fieldset className="dialog-options paragraph-flags"><legend>分页</legend>
        {flag("widows", "孤行控制", "2", "1")}{flag("break-after", "与下段同页", "avoid", "auto")}{flag("break-inside", "段中不分页", "avoid", "auto")}{flag("break-before", "段前分页", "page", "auto")}</fieldset></div>
      <div hidden={tab !== 2} className="dialog-stack"><fieldset className="dialog-options paragraph-flags"><legend>换行</legend>
        {flag("--word-kinsoku", "按中文习惯控制首尾字符")}{flag("--word-word-wrap", "允许西文在单词中间换行", "0", "1")}{flag("--word-overflow-punctuation", "允许标点溢出边界")}{flag("--word-top-line-punctuation", "压缩行首标点")}</fieldset>
        <fieldset className="dialog-options paragraph-flags"><legend>字符间距</legend>{flag("--word-auto-space-de", "自动调整中文与西文的间距")}{flag("--word-auto-space-dn", "自动调整中文与数字的间距")}</fieldset></div>
      <div className="paragraph-preview" style={{ marginTop: 20, borderTop: "1px solid #ddd", paddingTop: 10 }}><b>预览</b><div style={{ padding: "12px 25px", background: "white", border: "1px solid #ccc", marginTop: 6 }}>
        {[0, 1, 2].map(i => <div key={i} style={{ height: 4, margin: "5px 0", width: i === 2 ? "76%" : "100%", background: "#b8b8b8" }} />)}</div></div>
    </>}</>}
  </OfficeDialog>;
}
