import { DocumentError } from "@doffi/core";
import type { DocumentOperation, DocumentSelection, WorkbookSnapshot } from "@doffi/core";

export interface WorkbookFormatUi {
  readonly element: HTMLDetailsElement; readonly tools: HTMLDivElement; readonly fields: readonly (HTMLInputElement | HTMLSelectElement)[];
  operations(snapshot: WorkbookSnapshot, selection: DocumentSelection | undefined): DocumentOperation[];
}
/** Plain form values are transient UI state. The command contains only the
 * chosen appearance changes and is validated/applied by the owning session. */
export function createWorkbookFormatUi(document: Document): WorkbookFormatUi {
  const element = document.createElement("details"), summary = document.createElement("summary"), tools = document.createElement("div");
  summary.textContent = "设置单元格格式"; tools.className = "tools"; element.append(summary);
  const fields: (HTMLInputElement | HTMLSelectElement)[] = [];
  const section = (title: string) => { const group = document.createElement("fieldset"), legend = document.createElement("legend"); legend.textContent = title; group.className = "bar"; group.append(legend); element.append(group); return group; };
  const font = section("字体"), alignment = section("对齐"), fill = section("填充"), border = section("边框"), number = section("数字");
  const input = (parent: HTMLElement, title: string, type = "text") => {
    const label = document.createElement("label"), field = document.createElement("input"); field.type = type; field.setAttribute("aria-label", title); field.maxLength = 2048;
    label.append(title, field); parent.append(label); fields.push(field); return field;
  };
  const select = (parent: HTMLElement, title: string, items: readonly (readonly [string, string])[]) => {
    const label = document.createElement("label"), field = document.createElement("select"); field.setAttribute("aria-label", title);
    for (const [value, text] of [["", "保持原设置"], ...items]) { const option = document.createElement("option"); option.value = value!; option.textContent = text!; field.append(option); }
    label.append(title, field); parent.append(label); fields.push(field); return field;
  };
  const latin = input(font, "西文字体"), eastAsian = input(font, "中文字体"), size = input(font, "字号（磅）", "number"), foreground = input(font, "文字颜色");
  size.min = "1"; size.max = "409"; size.step = "0.5"; foreground.placeholder = "#112233";
  const weight = select(font, "字重", [["bold", "加粗"], ["normal", "常规"]]), italic = select(font, "斜体", [["italic", "斜体"], ["normal", "常规"]]);
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
  const hint = document.createElement("p"); hint.textContent = "空白或“保持原设置”保留各单元格原格式；颜色可输入十六进制值。"; element.append(hint, tools);
  function operations(snapshot: WorkbookSnapshot, selection: DocumentSelection | undefined): DocumentOperation[] {
    if (!selection || selection.anchor.kind !== "cells") throw new DocumentError("InvalidOperation", "请先选择要设置格式的单元格。");
    const manifest = snapshot.manifest;
    if (selection.documentId !== manifest.documentId || selection.baseRevision !== manifest.revision) throw new DocumentError("RevisionConflict", "格式选区已过期。");
    const properties: Record<string, string> = {}, put = (name: string, field: HTMLInputElement | HTMLSelectElement) => { if (field.value.trim()) properties[name] = field.value.trim(); };
    const family = (field: HTMLInputElement) => '"' + field.value.trim().replaceAll("\\", "\\\\").replaceAll('"', '\\"') + '"';
    if (latin.value.trim()) { properties["--office-font-ascii"] = family(latin); properties["--office-font-h-ansi"] = family(latin); }
    if (eastAsian.value.trim()) properties["--office-font-east-asia"] = family(eastAsian);
    if (size.value) { const value = Number(size.value); if (!Number.isFinite(value) || value <= 0 || value > 409) throw new DocumentError("InvalidOperation", "字号须大于 0 且不超过 409 磅。"); properties["font-size"] = value + "pt"; }
    put("color", foreground); put("font-weight", weight); put("font-style", italic); put("--office-font-underline", underline); put("--office-font-strike", strike);
    put("--office-cell-horizontal", horizontal); put("--office-cell-vertical", vertical); put("--office-cell-wrap", wrap); put("--office-cell-shrink-to-fit", shrink);
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
    if (numberCode.value.trim()) {
      const code = numberCode.value.trim(), formats = snapshot.metadata.workbook.numberFormats;
      numberFormatId = formats.find(format => "code" in format && format.code === code)?.formatId;
      if (!numberFormatId) { let serial = 0; const ids = new Set(formats.map(format => format.formatId)); do { numberFormatId = "workbook-format-" + ++serial; } while (ids.has(numberFormatId)); operations.push({ op: "setWorkbookNumberFormat", format: { formatId: numberFormatId, code } }); }
    }
    if (Object.keys(properties).length || numberFormatId) operations.push({ op: "formatWorkbookRanges", sheetId: selection.anchor.sheetId, ranges: selection.anchor.ranges.map(range => ({ from: { ...range.from }, to: { ...range.to } })),
      properties, ...(numberFormatId ? { numberFormatId } : {}) });
    return operations;
  }
  return { element, tools, fields, operations };
}
