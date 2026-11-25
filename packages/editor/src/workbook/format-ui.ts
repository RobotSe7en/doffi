import type { TextFormatState } from "../text-format-state.js";
import { createOfficeDialog } from "../office-dialog.js";
import { DocumentError } from "@doffi/core";
import type { DocumentOperation, DocumentSelection, WorkbookSnapshot } from "@doffi/core";

export interface WorkbookFormatUi {
  readonly element: HTMLDetailsElement; readonly tools: HTMLDivElement; readonly fields: readonly (HTMLInputElement | HTMLSelectElement)[];
  operations(snapshot: WorkbookSnapshot, selection: DocumentSelection | undefined): DocumentOperation[];
  fontProperties(): Readonly<Record<string, string>>;
  open(state?: TextFormatState): void; close(): void;
  error(reason: unknown): void;
}
/** Plain form values are transient UI state. The command contains only the
 * chosen appearance changes and is validated/applied by the owning session. */
export function createWorkbookFormatUi(document: Document, availableFonts: readonly string[] = [], readState?: () => TextFormatState): WorkbookFormatUi {
  const element = document.createElement("details"), summary = document.createElement("summary");
  const ui = createOfficeDialog(document, "设置单元格格式"), tools = ui.footer;
  ui.dialog.style.setProperty("--office-accent", "#107c41"); ui.dialog.style.width = "617px";
  summary.textContent = "设置单元格格式"; element.append(summary, ui.dialog);
  const panels = new Map(["数字", "对齐", "字体", "边框", "填充", "保护"].map(title => [title, ui.tab(title)]));
  const fields: (HTMLInputElement | HTMLSelectElement)[] = [];
  const initial = new Map<HTMLInputElement | HTMLSelectElement, string>();
  const valueOf = (field: HTMLInputElement | HTMLSelectElement): string => field.value === initial.get(field) ? "" : field.value;
  const section = (title: string) => { const group = document.createElement("div"); group.className = "dialog-options"; panels.get(title)!.append(group); return group; };
  const font = section("字体"), alignment = section("对齐"), fill = section("填充"), border = section("边框"), number = section("数字");
  const reset = () => { for (const field of fields) field.value = ""; initial.clear(); };
  const close = () => { ui.close(); element.open = false; element.style.removeProperty("display"); summary.hidden = false; reset(); };
  summary.addEventListener("click", event => { event.preventDefault(); open(); });
  ui.dialog.addEventListener("cancel", event => { event.preventDefault(); close(); });
  const cancel = document.createElement("button"); cancel.type = "button"; cancel.textContent = "取消"; cancel.addEventListener("click", close); tools.append(cancel);
  const input = (parent: HTMLElement, title: string, type = "text") => {
    const label = document.createElement("label"), field = document.createElement("input"); field.type = type; field.setAttribute("aria-label", title); field.maxLength = 2048;
    label.append(title, field); parent.append(label); fields.push(field); return field;
  };
  const select = (parent: HTMLElement, title: string, items: readonly (readonly [string, string])[]) => {
    const label = document.createElement("label"), field = document.createElement("select"); field.setAttribute("aria-label", title);
    for (const [value, text] of [["", "保持原设置"], ...items]) { const option = document.createElement("option"); option.value = value!; option.textContent = text!; field.append(option); }
    label.append(title, field); parent.append(label); fields.push(field); return field;
  };
  const latin = input(font, "西文字体"), eastAsian = input(font, "中文字体"), size = input(font, "字号（磅）", "number"), foreground = input(font, "文字颜色", "color");
  size.min = "1"; size.max = "409"; size.step = "0.5"; foreground.placeholder = "#112233";
  const weight = select(font, "字重", [["normal", "常规"], ["italic", "斜体"], ["bold", "加粗"], ["boldItalic", "加粗斜体"]]), italic = select(font, "斜体", [["italic", "斜体"], ["normal", "常规"]]);
  const underline = select(font, "下划线", [["none", "无"], ["single", "单线"], ["double", "双线"], ["singleAccounting", "会计单线"], ["doubleAccounting", "会计双线"]]);
  const strike = select(font, "删除线", [["true", "有"], ["false", "无"]]);
  const horizontal = select(alignment, "水平对齐", [["general", "常规"], ["left", "靠左"], ["center", "居中"], ["right", "靠右"], ["fill", "填充"], ["justify", "两端对齐"], ["centerContinuous", "跨列居中"], ["distributed", "分散对齐"]]);
  const vertical = select(alignment, "垂直对齐", [["top", "靠上"], ["center", "居中"], ["bottom", "靠下"], ["justify", "两端对齐"], ["distributed", "分散对齐"]]);
  const wrap = select(alignment, "自动换行", [["true", "开启"], ["false", "关闭"]]), shrink = select(alignment, "缩小字体填充", [["true", "开启"], ["false", "关闭"]]);
  const rotation = input(alignment, "文字旋转角度", "number"), stacked = select(alignment, "文字方向", [["horizontal", "横排"], ["vertical", "竖排"]]); rotation.min = "-90"; rotation.max = "90"; rotation.step = "1";
  const patterns = ["mediumGray", "darkGray", "lightGray", "darkHorizontal", "darkVertical", "darkDown", "darkUp", "darkGrid", "darkTrellis", "lightHorizontal", "lightVertical", "lightDown", "lightUp", "lightGrid", "lightTrellis", "gray125", "gray0625"];
  const names = ["中灰", "深灰", "浅灰", "深横线", "深竖线", "深向下斜线", "深向上斜线", "深网格", "深斜网格", "浅横线", "浅竖线", "浅向下斜线", "浅向上斜线", "浅网格", "浅斜网格", "稀疏点", "细稀疏点"];
  const pattern = select(fill, "填充方式", [["none", "无填充"], ["solid", "纯色"], ["linear", "线性渐变"], ["path", "路径渐变"], ...patterns.map((value, index) => [value, names[index]!] as const)]);
  const fillColor = input(fill, "填充前景色"), background = input(fill, "填充背景色"), degree = input(fill, "渐变角度", "number"); fillColor.placeholder = "#cceeff"; background.placeholder = "#ffffff"; degree.step = "1";
  const borderTarget = select(border, "边框位置", [["all", "所有边框"], ["left", "左"], ["right", "右"], ["top", "上"], ["bottom", "下"], ["diagonalUp", "上斜线"], ["diagonalDown", "下斜线"], ["diagonalBoth", "双斜线"]]);
  const borderKind = select(border, "边框线型", [["none", "无"], ["thin", "细线"], ["medium", "中线"], ["thick", "粗线"], ["double", "双线"], ["dashed", "虚线"], ["dotted", "点线"], ["hair", "发丝线"], ["mediumDashed", "中虚线"], ["dashDot", "点划线"], ["mediumDashDot", "中点划线"], ["dashDotDot", "双点划线"], ["mediumDashDotDot", "中双点划线"], ["slantDashDot", "斜点划线"]]);
  const borderColor = input(border, "边框颜色"), numberCode = input(number, "数字格式代码"); borderColor.placeholder = "#000000"; numberCode.placeholder = "0.00 或 yyyy-mm-dd";
  const fontPreview = document.createElement("div"); fontPreview.className = "dialog-preview"; fontPreview.textContent = "AaBbCc　字样"; panels.get("字体")!.append(fontPreview);
  latin.parentElement!.firstChild!.textContent = "字体:"; eastAsian.parentElement!.firstChild!.textContent = "中文字体:";
  size.parentElement!.firstChild!.textContent = "字号:"; weight.parentElement!.firstChild!.textContent = "字形:";
  const fontNames = [...new Set([...availableFonts, "等线", "宋体", "微软雅黑", "仿宋", "楷体", "黑体", "Arial", "Aptos", "Calibri", "Cambria", "Times New Roman", "Noto Sans CJK SC", "Noto Serif CJK SC"])];
  const list = (field: HTMLInputElement | HTMLSelectElement, values: readonly string[]) => {
    const holder = document.createElement("div"); holder.className = "dialog-list";
    for (const value of values) { const button = document.createElement("button"); button.type = "button"; button.textContent = value;
      button.addEventListener("click", () => { field.value = value; field.dispatchEvent(new Event("change", { bubbles: true })); }); holder.append(button); }
    field.parentElement!.append(holder);
  };
  font.style.gridTemplateColumns = "2fr 1.1fr .8fr";
  for (const field of [latin, weight, size]) font.append(field.parentElement!);
  for (const field of [eastAsian, foreground, italic, underline, strike]) font.append(field.parentElement!);
  list(latin, fontNames);
  const shapes = document.createElement("div"); shapes.className = "dialog-list";
  for (const [value, name] of [["normal", "常规"], ["italic", "斜体"], ["bold", "加粗"], ["boldItalic", "加粗斜体"]]) { const button = document.createElement("button"); button.type = "button"; button.textContent = name!; button.addEventListener("click", () => { weight.value = value!; weight.dispatchEvent(new Event("change", { bubbles: true })); }); shapes.append(button); }
  weight.parentElement!.append(shapes);
  weight.addEventListener("change", () => { italic.value = ["italic", "boldItalic"].includes(weight.value) ? "italic" : "normal"; });
  italic.parentElement!.hidden = true;
  list(size, [6, 8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72].map(String));
  for (const field of [latin, size, weight, italic]) field.addEventListener("change", () => { fontPreview.style.fontFamily = latin.value; fontPreview.style.fontSize = (size.value || "12") + "pt"; fontPreview.style.fontWeight = ["bold", "boldItalic"].includes(weight.value) ? "bold" : "normal"; fontPreview.style.fontStyle = italic.value; });
  const paintPalette = document.createElement("div"), paletteLabel = document.createElement("span"), paintPreview = document.createElement("div");
  paletteLabel.textContent = "背景色:"; paintPalette.style.cssText = "display:grid;grid-template-columns:repeat(8,25px);gap:4px;margin:10px 0 18px";
  for (const color of ["#ffffff", "#000000", "#e7e6e6", "#44546a", "#4472c4", "#ed7d31", "#a5a5a5", "#ffc000", "#5b9bd5", "#70ad47", "#c00000", "#ff0000", "#ffff00", "#00b050", "#00b0f0", "#7030a0", "#d9e1f2", "#fce4d6", "#ededed", "#fff2cc", "#ddebf7", "#e2efda", "#8ea9db", "#f4b084", "#c9c9c9", "#ffd966", "#9bc2e6", "#a9d18e", "#305496", "#c65911", "#7b7b7b", "#bf9000"]) {
    const button = document.createElement("button"); button.type = "button"; button.setAttribute("aria-label", "背景色 " + color); button.title = color; button.style.cssText = "width:25px;height:22px;border:1px solid #bbb;border-radius:0;background:" + color;
    button.addEventListener("click", () => { pattern.value = "solid"; fillColor.value = color; paintPreview.style.backgroundColor = color; }); paintPalette.append(button);
  }
  paintPreview.className = "dialog-preview"; paintPreview.style.cssText = "height:60px;min-height:60px;margin:20px 0 0;font-size:13px"; paintPreview.textContent = "字样";
  panels.get("填充")!.prepend(paletteLabel, paintPalette); panels.get("填充")!.append(paintPreview);
  const categories = document.createElement("div"), categoryLabel = document.createElement("span"), categoryList = document.createElement("div"), sample = document.createElement("div"), categoryInfo = document.createElement("p");
  categories.style.cssText = "display:grid;grid-template-columns:140px 1fr;gap:20px"; categoryList.className = "dialog-list"; categoryList.style.height = "350px"; categoryList.setAttribute("role", "listbox"); categoryList.setAttribute("aria-label", "分类"); categoryLabel.textContent = "分类:";
  const codes = [["常规", "General"], ["数值", "0.00"], ["货币", "¥#,##0.00"], ["会计专用", '_ ¥* #,##0.00_ ;_ ¥* -#,##0.00_ ;_ ¥* "-"??_ ;_ @_ '], ["日期", "yyyy/m/d"], ["时间", "h:mm:ss"], ["百分比", "0.00%"], ["分数", "# ?/?"], ["科学记数", "0.00E+00"], ["文本", "@"], ["特殊", "000000"], ["自定义", ""]];
  for (const [name, code] of codes) { const button = document.createElement("button"); button.type = "button"; button.setAttribute("role", "option"); button.textContent = name!; button.addEventListener("click", () => {
    for (const node of categoryList.children) node.setAttribute("aria-selected", String(node === button)); numberCode.parentElement!.hidden = name !== "自定义"; numberCode.value = code!; categoryInfo.textContent = name === "常规" ? "常规格式的单元格没有特定的数字格式。" : "";
  }); categoryList.append(button); }
  const left = document.createElement("div"), right = document.createElement("div"); left.append(categoryLabel, categoryList); right.append("示例", sample, number, categoryInfo); sample.className = "dialog-preview"; sample.style.cssText = "font-size:13px;min-height:45px;margin:10px 0 24px;padding:10px"; number.style.display = "block"; numberCode.parentElement!.hidden = true; categories.append(left, right); panels.get("数字")!.append(categories);
  for (const title of ["锁定", "隐藏"]) { const label = document.createElement("label"), checkbox = document.createElement("input"); checkbox.type = "checkbox"; checkbox.disabled = true; checkbox.checked = title === "锁定"; label.append(checkbox, title); panels.get("保护")!.append(label); }

  const put = (properties: Record<string, string>, name: string, field: HTMLInputElement | HTMLSelectElement) => { if (valueOf(field).trim()) properties[name] = valueOf(field).trim(); };
  function fontProperties(): Record<string, string> {
    const properties: Record<string, string> = {};
    const family = (field: HTMLInputElement) => '"' + valueOf(field).trim().replaceAll("\\", "\\\\").replaceAll('"', '\\"') + '"';
    if (valueOf(latin).trim()) { properties["--office-font-ascii"] = family(latin); properties["--office-font-h-ansi"] = family(latin); properties["--office-font-east-asia"] = family(latin); }
    if (valueOf(eastAsian).trim()) properties["--office-font-east-asia"] = family(eastAsian);
    if (size.validity.badInput) throw new DocumentError("InvalidOperation", "请输入有效字号。");
    if (valueOf(size)) { const value = Number(valueOf(size)); if (!Number.isFinite(value) || value <= 0 || value > 409) throw new DocumentError("InvalidOperation", "字号须大于 0 且不超过 409 磅。"); properties["font-size"] = value + "pt"; }
    put(properties, "color", foreground); if (valueOf(weight)) properties["font-weight"] = ["bold", "boldItalic"].includes(weight.value) ? "bold" : "normal"; put(properties, "font-style", italic); put(properties, "--office-font-underline", underline); put(properties, "--office-font-strike", strike);
    return properties;
  }
  function operations(snapshot: WorkbookSnapshot, selection: DocumentSelection | undefined): DocumentOperation[] {
    if (!selection || selection.anchor.kind !== "cells") throw new DocumentError("InvalidOperation", "请先选择要设置格式的单元格。");
    const manifest = snapshot.manifest;
    if (selection.documentId !== manifest.documentId || selection.baseRevision !== manifest.revision) throw new DocumentError("RevisionConflict", "格式选区已过期。");
    const properties = fontProperties();
    put(properties, "--office-cell-horizontal", horizontal); put(properties, "--office-cell-vertical", vertical); put(properties, "--office-cell-wrap", wrap); put(properties, "--office-cell-shrink-to-fit", shrink);
    if (rotation.validity.badInput || degree.validity.badInput) throw new DocumentError("InvalidOperation", "请输入有效角度。");
    if (rotation.value && stacked.value === "vertical") throw new DocumentError("InvalidOperation", "请在旋转角度和竖排之间选择一项。");
    if (rotation.value) { const value = Number(rotation.value); if (!Number.isInteger(value) || Math.abs(value) > 90) throw new DocumentError("InvalidOperation", "旋转角度须为 -90 到 90 的整数。"); properties["--office-cell-text-rotation"] = String(value < 0 ? 90 - value : value); }
    else if (stacked.value) properties["--office-cell-text-rotation"] = stacked.value === "vertical" ? "255" : "0";
    const paint = (prefix: string, color: string) => { properties[prefix + "-color"] = color; properties[prefix + "-tint"] = "0"; properties[prefix + "-native-alpha"] = "initial"; };
    if (pattern.value) {
      const gradient = ["linear", "path"].includes(pattern.value);
      properties["--office-cell-fill-kind"] = gradient ? "gradient" : pattern.value;
      if (pattern.value !== "none") {
        if (!fillColor.value.trim()) throw new DocumentError("InvalidOperation", "请选择填充前景色。");
        if (gradient) {
          if (!background.value.trim()) throw new DocumentError("InvalidOperation", "渐变需要两个颜色。");
          properties["--office-cell-gradient-kind"] = pattern.value; properties["--office-cell-gradient-stops"] = "2";
          const angle = degree.value ? Number(degree.value) : 0; if (!Number.isFinite(angle)) throw new DocumentError("InvalidOperation", "渐变角度须为有限数字。");
          properties["--office-cell-gradient-degree"] = String(angle);
          for (const side of ["left", "right", "top", "bottom"]) properties["--office-cell-gradient-" + side] = "0.5";
          for (const [index, color] of [fillColor.value.trim(), background.value.trim()].entries()) { properties[`--office-cell-gradient-stop-${index}-position`] = String(index); paint(`--office-cell-gradient-stop-${index}`, color); }
        } else { paint("--office-cell-fill", fillColor.value.trim()); paint("--office-cell-fill-background", background.value.trim() || "#ffffff"); }
      }
    } else if (fillColor.value.trim() || background.value.trim() || degree.value) throw new DocumentError("InvalidOperation", "请先选择填充方式。");
    if (borderTarget.value) {
      if (!borderKind.value) throw new DocumentError("InvalidOperation", "请选择边框线型。");
      const diagonal = borderTarget.value.startsWith("diagonal"), sides = borderTarget.value === "all" ? ["left", "right", "top", "bottom"] : [diagonal ? "diagonal" : borderTarget.value];
      for (const side of sides) { properties[`--office-cell-border-${side}-kind`] = borderKind.value; paint(`--office-cell-border-${side}`, borderColor.value.trim() || "#000000"); }
      if (diagonal) { properties["--office-cell-border-diagonal-up"] = String(borderTarget.value !== "diagonalDown"); properties["--office-cell-border-diagonal-down"] = String(borderTarget.value !== "diagonalUp"); }
    } else if (borderKind.value || borderColor.value.trim()) throw new DocumentError("InvalidOperation", "请先选择边框位置。");
    const operations: DocumentOperation[] = []; let numberFormatId: string | undefined;
    if (valueOf(numberCode).trim()) {
      const code = valueOf(numberCode).trim(), formats = snapshot.metadata.workbook.numberFormats;
      numberFormatId = formats.find(format => "code" in format && format.code === code)?.formatId;
      if (!numberFormatId) { let serial = 0; const ids = new Set(formats.map(format => format.formatId)); do { numberFormatId = "workbook-format-" + ++serial; } while (ids.has(numberFormatId)); operations.push({ op: "setWorkbookNumberFormat", format: { formatId: numberFormatId, code } }); }
    }
    if (Object.keys(properties).length || numberFormatId) operations.push({ op: "formatWorkbookRanges", sheetId: selection.anchor.sheetId, ranges: selection.anchor.ranges.map(range => ({ from: { ...range.from }, to: { ...range.to } })),
      properties, ...(numberFormatId ? { numberFormatId } : {}) });
    return operations;
  }
  function open(state = readState?.()): void {
    reset();
    if (state) {
      const match = /^\s*(?:"([^"]*)"|'([^']*)'|([^,]+))/.exec(state.fontFamily ?? "");
      latin.value = match?.[1] ?? match?.[2] ?? match?.[3] ?? ""; size.value = state.fontSize === undefined ? "" : String(state.fontSize);
      weight.value = state.bold === undefined ? "" : state.bold ? state.italic ? "boldItalic" : "bold" : state.italic ? "italic" : "normal"; italic.value = state.italic === undefined ? "" : state.italic ? "italic" : "normal";
      const rgb = /^rgb\(\s*(\d+)[, ]+\s*(\d+)[, ]+\s*(\d+)\s*\)$/.exec(state.color ?? "");
      foreground.value = /^#[0-9a-f]{6}$/i.test(state.color ?? "") ? state.color! : rgb ? "#" + rgb.slice(1).map(value => Number(value).toString(16).padStart(2, "0")).join("") : "#000000";
      numberCode.value = state.numberFormatCode ?? (state.numberFormatBuiltinId === 0 ? "General" : "");
      const category = codes.find(([_, code]) => code === numberCode.value)?.[0] ?? "自定义";
      for (const node of categoryList.children) node.setAttribute("aria-selected", String(node.textContent === category));
      numberCode.parentElement!.hidden = category !== "自定义";
      categoryInfo.textContent = category === "常规" ? "常规格式的单元格没有特定的数字格式。" : ""; sample.textContent = state.numberPreview ?? "";
    }
    for (const field of fields) initial.set(field, field.value);
    element.open = true; element.style.display = "contents"; summary.hidden = true; ui.show();
  }
  return { element, tools, fields, operations, fontProperties, open, close, error: ui.error };
}
