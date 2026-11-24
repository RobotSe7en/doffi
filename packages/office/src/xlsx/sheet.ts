import type { WorkbookHyperlink, WorkbookSheet } from "@doffi/core";
import { attr, check, elements, fail, R, S, unsupported } from "../ooxml/xml.js";
import type { XmlNode } from "../ooxml/xml.js";
import type { OpcPackage } from "../ooxml/package.js";
import { address, attributes, boolean, integer, number, one, range, spreadsheetText, value } from "./values.js";
import { readPrint } from "./print.js";
import type { NativeSheet } from "./workbook.js";
import { excelRootAttributes } from "./excel-metadata.js";
import { validatePhoneticProperties } from "./strings.js";

export interface CellStyleReference { readonly cssClass: string; readonly numberFormatId: string }
export interface SheetStyles {
  cell(index: number, node: XmlNode): CellStyleReference;
  defaultCell(): CellStyleReference;
  tabStyle(node: XmlNode): string;
}

function readOutline(node: XmlNode, dimension: Pick<WorkbookSheet["rowDimensions"][number], "outlineLevel" | "collapsed">): void {
  const level = integer(node, "outlineLevel", 0, 7, 0);
  if (level) dimension.outlineLevel = level;
  if (boolean(node, "collapsed")) dimension.collapsed = true;
}

/** Only fixed row heights are source dimensions; automatic native heights are
 * layout caches. Rows without an explicit override inherit sheet defaults. */
export function rowDimension(node: XmlNode, row: number, styles: SheetStyles): WorkbookSheet["rowDimensions"][number] | undefined {
  attributes(node, ["r", "spans", "s", "customFormat", "ht", "hidden", "customHeight", "outlineLevel", "collapsed", "thickTop", "thickBot", "ph"]);
  if (boolean(node, "thickTop") || boolean(node, "thickBot") || boolean(node, "ph")) unsupported(node, "Thick row spacing and phonetic display are not supported.");
  const result: WorkbookSheet["rowDimensions"][number] = { from: row, to: row };
  readOutline(node, result);
  if (value(node, "customHeight") !== undefined && !boolean(node, "customHeight")) result.autoHeight = true;
  if (value(node, "hidden") !== undefined) result.hidden = boolean(node, "hidden");
  if (value(node, "ht") !== undefined) {
    const height = number(node, "ht"); if (height < 0 || height > 409) fail(node, "A worksheet row height is invalid.");
    if (boolean(node, "customHeight")) result.heightPt = height;
  } else if (boolean(node, "customHeight")) fail(node, "A custom row height is missing.");
  if (value(node, "s") !== undefined || boolean(node, "customFormat")) {
    const style = styles.cell(integer(node, "s", 0, 0xffff_ffff, 0), node);
    if (boolean(node, "customFormat")) Object.assign(result, style);
  }
  return Object.keys(result).length === 2 ? undefined : result;
}

function columns(root: XmlNode, styles: SheetStyles): WorkbookSheet["columnDimensions"] {
  const parents = elements(root, "cols", S), result: WorkbookSheet["columnDimensions"] = [];
  for (const parent of parents) {
    attributes(parent, []); check(parent, ["col"], S);
    for (const node of elements(parent, "col", S)) {
      attributes(node, ["min", "max", "width", "style", "hidden", "bestFit", "customWidth", "phonetic", "outlineLevel", "collapsed"]); check(node, [], S);
      if (boolean(node, "phonetic") || boolean(node, "bestFit")) unsupported(node, "Automatic column sizing and phonetic display are not supported.");
      if (value(node, "customWidth") !== undefined) boolean(node, "customWidth");
      const from = integer(node, "min", 1, 16384) - 1, to = integer(node, "max", 1, 16384) - 1;
      if (from > to) fail(node, "A worksheet column span is reversed.");
      const dimension: WorkbookSheet["columnDimensions"][number] = { from, to, unit: "characters" };
      readOutline(node, dimension);
      if (value(node, "width") !== undefined) { dimension.width = number(node, "width"); if (dimension.width < 0 || dimension.width > 255) fail(node, "A column width exceeds native bounds."); }
      if (value(node, "hidden") !== undefined) dimension.hidden = boolean(node, "hidden");
      if (value(node, "style") !== undefined) Object.assign(dimension, styles.cell(integer(node, "style"), node));
      result.push(dimension);
    }
  }
  // Overlap is not flattened: it makes inherited formatting ambiguous.
  result.sort((a, b) => a.from - b.from);
  for (let i = 1; i < result.length; i++) if (result[i]!.from <= result[i - 1]!.to) fail(root, "Worksheet column definitions overlap.");
  return result;
}
function hyperlinks(root: XmlNode, opc: OpcPackage): WorkbookHyperlink[] {
  const parent = one(root, "hyperlinks"), result: WorkbookHyperlink[] = [];
  if (!parent) return result;
  attributes(parent, []); check(parent, ["hyperlink"], S);
  for (const node of elements(parent, "hyperlink", S)) {
    check(node, [], S);
    for (const [key] of node.attrs) if (!key.startsWith("http://www.w3.org/2000/xmlns/ ") && ![" ref", " location", " tooltip", " display", R + " id"].includes(key)) unsupported(node);
    const target: WorkbookHyperlink = { range: range(value(node, "ref") ?? "", node) }, id = attr(node, "id", R);
    if (id !== undefined) {
      const relation = opc.relationships(root.part).get(id);
      if (!relation || relation.type !== R + "/hyperlink" || !relation.external) fail(node, "A worksheet hyperlink relationship is invalid.");
      target.url = relation.target;
    }
    for (const key of ["location", "tooltip", "display"] as const) if (value(node, key) !== undefined) target[key] = spreadsheetText(value(node, key)!, node, key === "location" ? 8192 : 32767);
    if (!target.url && !target.location) fail(node, "A worksheet hyperlink has no target.");
    result.push(target);
  }
  return result;
}

/** Called after streaming cell/row records have been discarded. The remaining
 * header tree also carries bounded R3-07 table/rule/object references. */
export function readSheet(root: XmlNode, source: NativeSheet, rows: WorkbookSheet["rowDimensions"], opc: OpcPackage, styles: SheetStyles): WorkbookSheet {
  if (root.uri !== S || root.name !== "worksheet") fail(root, "The worksheet root is invalid.");
  excelRootAttributes(root, true);
  check(root, ["sheetPr", "dimension", "sheetViews", "sheetFormatPr", "cols", "sheetData", "autoFilter", "mergeCells", "phoneticPr", "conditionalFormatting", "dataValidations", "hyperlinks", "printOptions", "pageMargins", "pageSetup", "headerFooter", "rowBreaks", "colBreaks", "drawing", "legacyDrawing", "tableParts", "extLst"], S);
  validatePhoneticProperties(one(root, "phoneticPr"));
  const format = one(root, "sheetFormatPr"), properties = one(root, "sheetPr"), dimension = one(root, "dimension"), views = one(root, "sheetViews"), mergeCells = one(root, "mergeCells");
  if (dimension) { attributes(dimension, ["ref"]); check(dimension, [], S); range(value(dimension, "ref") ?? "", dimension); }
  if (format) {
    attributes(format, ["baseColWidth", "defaultColWidth", "defaultRowHeight", "customHeight", "zeroHeight", "thickTop", "thickBottom", "outlineLevelRow", "outlineLevelCol"]); check(format, [], S);
    if (boolean(format, "thickTop") || boolean(format, "thickBottom")) unsupported(format, "Thick row spacing is not supported.");
    integer(format, "outlineLevelRow", 0, 7, 0);
    integer(format, "outlineLevelCol", 0, 7, 0);
    if (value(format, "customHeight") !== undefined) boolean(format, "customHeight");
  }
  if (properties) {
    attributes(properties, ["codeName", "filterMode", "published", "enableFormatConditionsCalculation", "syncHorizontal", "syncVertical", "syncRef", "transitionEvaluation", "transitionEntry"]);
    check(properties, ["tabColor", "pageSetUpPr"], S);
    // A VBA code name can be present on an ordinary macro-free worksheet.
    const codeName = value(properties, "codeName");
    if (codeName !== undefined) spreadsheetText(codeName, properties, 31);
    if (boolean(properties, "transitionEvaluation") || boolean(properties, "transitionEntry") || !boolean(properties, "enableFormatConditionsCalculation", true)) unsupported(properties, "This worksheet calculation dialect is not supported.");
    for (const key of ["filterMode", "published", "syncHorizontal", "syncVertical"]) if (value(properties, key) !== undefined) boolean(properties, key);
    if (value(properties, "syncRef") !== undefined) address(value(properties, "syncRef")!, properties);
  }
  if (views) { attributes(views, []); check(views, ["sheetView"], S); }
  const viewList = elements(views, "sheetView", S); if (viewList.length > 1) unsupported(views!, "Multiple worksheet view definitions are not implemented.");
  const view = viewList[0];
  if (view) {
    attributes(view, ["windowProtection", "showFormulas", "showGridLines", "showRowColHeaders", "showZeros", "rightToLeft", "tabSelected", "showRuler", "showOutlineSymbols", "defaultGridColor", "showWhiteSpace", "view", "topLeftCell", "colorId", "zoomScale", "zoomScaleNormal", "zoomScaleSheetLayoutView", "zoomScalePageLayoutView", "workbookViewId"]);
    check(view, ["pane", "selection"], S);
    if (boolean(view, "windowProtection") || boolean(view, "showFormulas") || !boolean(view, "showZeros", true) || !boolean(view, "showRowColHeaders", true) || !boolean(view, "defaultGridColor", true)) unsupported(view, "This worksheet view display mode is not implemented.");
    integer(view, "workbookViewId");
    if (value(view, "topLeftCell") !== undefined) address(value(view, "topLeftCell")!, view);
  }
  const pane = one(view, "pane"); let frozen = { rows: 0, columns: 0 };
  if (pane) {
    attributes(pane, ["xSplit", "ySplit", "topLeftCell", "activePane", "state"]); check(pane, [], S);
    if (!["frozen", "frozenSplit"].includes(value(pane, "state") ?? "split")) unsupported(pane, "Unfrozen split panes are not implemented.");
    frozen = { rows: integer(pane, "ySplit", 0, 1_048_575, 0), columns: integer(pane, "xSplit", 0, 16383, 0) };
    if (value(pane, "topLeftCell") !== undefined) address(value(pane, "topLeftCell")!, pane);
  }
  const merges: WorkbookSheet["merges"] = [];
  if (mergeCells) {
    attributes(mergeCells, ["count"]); check(mergeCells, ["mergeCell"], S);
    for (const merge of elements(mergeCells, "mergeCell", S)) { attributes(merge, ["ref"]); check(merge, [], S); merges.push(range(value(merge, "ref") ?? "", merge)); }
    if (value(mergeCells, "count") !== undefined && integer(mergeCells, "count") !== merges.length) fail(mergeCells, "The merge count is inconsistent.");
  }
  const tab = one(properties, "tabColor"), defaultStyle = styles.defaultCell(), nativeWidth = value(format, "defaultColWidth");
  return {
    sheetId: source.sheetId, name: source.name, visibility: source.visibility,
    defaultRowHeightPt: format ? number(format, "defaultRowHeight") : 15,
    ...(value(format, "customHeight") !== undefined ? { defaultRowHeightCustom: boolean(format, "customHeight") } : {}),
    defaultColumnWidth: nativeWidth !== undefined ? number(format!, "defaultColWidth") : format ? integer(format, "baseColWidth", 1, 255, 8) : 8,
    defaultColumnWidthUnit: nativeWidth === undefined ? "baseCharacters" : "characters", defaultRowsHidden: boolean(format, "zeroHeight"),
    defaultCssClass: defaultStyle.cssClass, defaultNumberFormatId: defaultStyle.numberFormatId,
    rowDimensions: rows, columnDimensions: columns(root, styles), merges, frozen, showGridLines: boolean(view, "showGridLines", true), rightToLeft: boolean(view, "rightToLeft"),
    ...(tab ? { tabCssClass: styles.tabStyle(tab) } : {}), hyperlinks: hyperlinks(root, opc), print: readPrint(root, source.printNames),
  };
}
