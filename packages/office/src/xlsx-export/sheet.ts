import { iterateWorkbookRows, resolveWorkbookCellFormat } from "@doffi/core";
import type { DeepReadonly, WorkbookCalculatedCell, WorkbookCell, WorkbookModel, WorkbookSheet, WorkbookValue } from "@doffi/core";
import type { OutputOpc } from "../export/opc.js";
import { attributes, declaration, empty } from "../export/xml.js";
import { printSettings } from "./print.js";
import { conditionalRule, dataValidation } from "./rules.js";
import { colorTag } from "./style-values.js";
import type { WorkbookStylesWriter } from "./styles.js";
import { autoFilter, writeTable } from "./tables.js";
import { richCellText } from "./text.js";
import { address, bool, element, formulaText, rangeRef, sheetContentType, sheetNamespaces, textElement, xstring } from "./xml.js";
import { FILTER_HIDDEN_NAMESPACE } from "../xlsx/row-filter-extension.js";

type Sheet = DeepReadonly<WorkbookSheet>;
type Scalar = DeepReadonly<WorkbookValue>;
export const cellKey = (sheetId: string, row: number, column: number): string => JSON.stringify([sheetId, row, column]);
export interface CacheCounts { written: number; omitted: number }
export interface SheetWriteOptions {
  readonly model: WorkbookModel;
  readonly sheet: Sheet;
  readonly active: boolean;
  readonly partName: string;
  readonly opc: OutputOpc;
  readonly styles: WorkbookStylesWriter;
  readonly calculated: ReadonlyMap<string, WorkbookCalculatedCell>;
  readonly caches: CacheCounts;
  readonly firstTable: number;
  readonly filteredRows?: ReadonlySet<number>;
  readonly checkpoint: () => Promise<void>;
  readonly tail?: string;
}

function formulaCache(value: Scalar | undefined): { type?: string; xml: string } {
  switch (value?.type) {
    case "number": return { xml: textElement("v", String(value.value)) };
    case "boolean": return { type: "b", xml: textElement("v", value.value ? "1" : "0") };
    case "error": return { type: "e", xml: textElement("v", value.value) };
    case "text": return { type: "str", xml: textElement("v", xstring(value.value)) };
    default: return { xml: "" };
  }
}
function cellXml(cell: WorkbookCell, options: SheetWriteOptions): string {
  const { sheet, styles, caches, calculated } = options;
  let style: number | undefined;
  if (cell.cssClass !== undefined || cell.formatId !== undefined) {
    const format = resolveWorkbookCellFormat(sheet, cell, cell);
    style = styles.cell(format.cssClass, format.numberFormatId);
  }
  const attrs = { r: address(cell), s: style };
  if (cell.formula) {
    const current = calculated.get(cellKey(cell.sheetId, cell.row, cell.column));
    const cache = formulaCache(current?.state === "valid" ? current.value : undefined);
    if (cache.xml) caches.written++; else caches.omitted++;
    return element("c", element("f", formulaText(cell.formula.expression)) + cache.xml, { ...attrs, t: cache.type });
  }
  if (cell.value.type === "text")
    return element("c", element("is", richCellText(cell.html, cell.value.value, styles)), { ...attrs, t: "inlineStr" });
  const value = formulaCache(cell.value);
  return value.xml ? element("c", value.xml, { ...attrs, t: value.type }) : empty("c", attrs);
}

function sheetViews(sheet: Sheet, active: boolean): string {
  const { rows, columns } = sheet.frozen;
  const pane = rows && columns ? "bottomRight" : rows ? "bottomLeft" : "topRight";
  const topLeftCell = address({ row: rows, column: columns });
  const body = rows || columns ? empty("pane", {
    xSplit: columns || undefined, ySplit: rows || undefined, topLeftCell, activePane: pane, state: "frozen",
  }) + empty("selection", { pane, activeCell: topLeftCell, sqref: topLeftCell }) : "";
  return element("sheetViews", element("sheetView", body, {
    workbookViewId: 0, showGridLines: bool(sheet.showGridLines), rightToLeft: bool(sheet.rightToLeft), tabSelected: active ? 1 : undefined,
  }));
}
function sheetHeader(sheet: Sheet, active: boolean, styles: WorkbookStylesWriter, filterMode: boolean): string {
  const tab = sheet.tabCssClass;
  const tabColor = tab ? colorTag("tabColor", styles.raw(tab), styles.css(tab), "--office-tab", sheet.tabColor)
    : sheet.tabColor ? colorTag("tabColor", {}, {}, "--office-tab", sheet.tabColor) : "";
  const properties = tabColor + (sheet.print?.scale.kind === "fit" ? empty("pageSetUpPr", { fitToPage: 1 }) : "");
  const standardWidth = sheet.defaultColumnWidthUnit !== "baseCharacters";
  const baseWidth = Math.max(1, Math.round(sheet.defaultColumnWidth));
  if (!standardWidth && baseWidth !== sheet.defaultColumnWidth)
    styles.reportFormatLoss("A fractional base column width is rounded to a whole character in XLSX.");
  const format = empty("sheetFormatPr", { defaultRowHeight: sheet.defaultRowHeightPt,
    customHeight: bool(sheet.defaultRowHeightCustom), zeroHeight: bool(sheet.defaultRowsHidden),
    baseColWidth: standardWidth ? undefined : baseWidth, defaultColWidth: standardWidth ? sheet.defaultColumnWidth : undefined });
  return (properties || filterMode ? element("sheetPr", properties, { filterMode: filterMode ? 1 : undefined }) : "") + sheetViews(sheet, active) + format;
}

export async function writeSheet(options: SheetWriteOptions): Promise<void> {
  const { sheet, opc, partName, styles } = options;
  const part = opc.part(partName, sheetContentType("worksheet"));
  const filtered = !!options.filteredRows?.size;
  const extensions = filtered ? attributes({ "xmlns:doffi": FILTER_HIDDEN_NAMESPACE }) : "";
  const filterHiddenRow = element("extLst", element("ext", "<doffi:filterHidden/>", { uri: FILTER_HIDDEN_NAMESPACE }));
  part.add(declaration + "<worksheet" + sheetNamespaces + extensions + ">" + sheetHeader(sheet, options.active, styles, filtered));
  if (sheet.columnDimensions.length) {
    part.add("<cols>");
    for (const column of sheet.columnDimensions) part.add(empty("col", {
      min: column.from + 1, max: column.to + 1, width: column.width, customWidth: column.width === undefined ? undefined : 1,
      hidden: bool(column.hidden), style: column.cssClass || column.numberFormatId
        ? styles.cell(column.cssClass ?? sheet.defaultCssClass, column.numberFormatId ?? sheet.defaultNumberFormatId) : undefined,
    }));
    part.add("</cols>");
  }
  part.add("<sheetData>");
  const rows = iterateWorkbookRows(options.model, sheet.sheetId);
  let current = rows.next(), index = 0, rowsWritten = 0;
  const filteredRows = [...options.filteredRows ?? []].sort((left, right) => left - right)[Symbol.iterator]();
  let filteredRow = filteredRows.next();
  let dimensionRow = sheet.rowDimensions[0]?.from ?? Infinity;
  while (!current.done || dimensionRow !== Infinity || !filteredRow.done) {
    const row = Math.min(current.done ? Infinity : current.value[0]!.row, dimensionRow, filteredRow.done ? Infinity : filteredRow.value);
    const hiddenByFilter = !filteredRow.done && row === filteredRow.value;
    if (hiddenByFilter) filteredRow = filteredRows.next();
    const dimension = row === dimensionRow ? sheet.rowDimensions[index] : undefined;
    const format = dimension && (dimension.cssClass || dimension.numberFormatId)
      ? styles.cell(dimension.cssClass ?? sheet.defaultCssClass, dimension.numberFormatId ?? sheet.defaultNumberFormatId) : undefined;
    const rowAttributes = attributes({ r: row + 1, ht: dimension?.heightPt,
      customHeight: dimension?.heightPt !== undefined ? 1 : dimension?.autoHeight ? 0 : undefined,
      hidden: hiddenByFilter ? 1 : bool(dimension?.hidden), s: format, customFormat: format === undefined ? undefined : 1 });
    const solelyFiltered = hiddenByFilter && !(dimension?.hidden ?? sheet.defaultRowsHidden) && dimension?.heightPt !== 0;
    if (!current.done && row === current.value[0]!.row) {
      part.add("<row" + rowAttributes + ">");
      for (const cell of current.value) part.add(cellXml(cell, options));
      part.add((solelyFiltered ? filterHiddenRow : "") + "</row>");
      current = rows.next();
    } else part.add(solelyFiltered ? "<row" + rowAttributes + ">" + filterHiddenRow + "</row>" : "<row" + rowAttributes + "/>");
    if (dimension) dimensionRow = row < dimension.to ? row + 1 : sheet.rowDimensions[++index]?.from ?? Infinity;
    if (++rowsWritten % 256 === 0) await options.checkpoint();
  }
  part.add("</sheetData>");
  if (sheet.autoFilter) part.add(autoFilter(sheet.autoFilter));
  if (sheet.merges.length) {
    part.add("<mergeCells" + attributes({ count: sheet.merges.length }) + ">");
    for (const merge of sheet.merges) part.add(empty("mergeCell", { ref: rangeRef(merge) }));
    part.add("</mergeCells>");
  }
  for (const rule of sheet.conditionalRules ?? []) part.add(conditionalRule(rule, styles));
  const validations = sheet.dataValidations ?? [];
  if (validations.length) {
    part.add("<dataValidations" + attributes({ count: validations.length, disablePrompts: bool(sheet.disableValidationPrompts) }) + ">");
    for (const validation of validations) part.add(dataValidation(validation));
    part.add("</dataValidations>");
  } else if (sheet.disableValidationPrompts)
    styles.reportFormatLoss("XLSX cannot store the worksheet's validation prompt setting without any validation rules.");
  if (sheet.hyperlinks?.length) {
    part.add("<hyperlinks>");
    for (const link of sheet.hyperlinks) part.add(empty("hyperlink", {
      ref: rangeRef(link.range), "r:id": link.url ? opc.relate(partName, "hyperlink", link.url, true) : undefined,
      location: link.location === undefined ? undefined : xstring(link.location), tooltip: link.tooltip === undefined ? undefined : xstring(link.tooltip),
      display: link.display === undefined ? undefined : xstring(link.display),
    }));
    part.add("</hyperlinks>");
  }
  if (sheet.print) part.add(printSettings(sheet.print));
  if (options.tail) part.add(options.tail);
  if (sheet.tables?.length) {
    part.add("<tableParts" + attributes({ count: sheet.tables.length }) + ">");
    for (const [index, table] of sheet.tables.entries())
      part.add(empty("tablePart", { "r:id": writeTable(opc, styles, partName, table, options.firstTable + index) }));
    part.add("</tableParts>");
  }
  part.add("</worksheet>");
  part.finish();
}
