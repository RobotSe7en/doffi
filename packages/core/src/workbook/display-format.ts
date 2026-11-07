import type { DeepReadonly, WorkbookSnapshot } from "../runtime.js";
import type { WorkbookCoordinate, WorkbookSheet } from "../contracts/workbook.js";
import type { WorkbookCellFormat } from "./cell-format.js";
import type { WorkbookConditionalEvaluation } from "./rule-reader.js";
import { AxisDimensions } from "./axis-dimensions.js";
import { WorkbookTableRegions } from "./table-regions.js";
import { WorkbookRegionIndex } from "./region-index.js";

/** Number-format composition shares the same sparse table regions and named
 * style propagation as rendering, without requiring a DOM or enumerating cells. */
export function createWorkbookTableFormatResolver(snapshot: WorkbookSnapshot, sheet: DeepReadonly<WorkbookSheet>): (at: WorkbookCoordinate, source: WorkbookCellFormat) => WorkbookCellFormat {
  const rows = new AxisDimensions(1_048_576, sheet.defaultRowsHidden ? 0 : 1, sheet.rowDimensions.map(run => ({ from: run.from, to: run.to, sizePt: (run.hidden ?? sheet.defaultRowsHidden) || run.heightPt === 0 ? 0 : 1 })));
  const columns = new AxisDimensions(16_384, 1, sheet.columnDimensions.map(run => ({ from: run.from, to: run.to, sizePt: run.hidden || run.width === 0 ? 0 : 1 })));
  const workbook = snapshot.metadata.workbook, definitions = new Map(workbook.cellStyles?.definitions.map(style => [style.styleId, style])), named = new Map(workbook.cellStyles?.names.map(style => [style.name.toUpperCase(), style.styleId]));
  const tables = (sheet.tables ?? []).map(table => {
    const reference = table.style, elements = reference?.kind === "custom" ? workbook.tableStyles?.find(style => style.styleId === reference.styleId)?.elements ?? [] : [];
    return { table, regions: new WorkbookTableRegions(table, elements, rows) };
  });
  const index = new WorkbookRegionIndex(tables.map(entry => entry.table.range)), byRange = new Map(tables.map(entry => [entry.table.range, entry]));
  return (at, source) => {
    if (!rows.size(at.row) || !columns.size(at.column)) return source;
    const region = index.at(at.row, at.column), entry = region ? byRange.get(region) : undefined;
    if (!entry) return source;
    const { table, regions } = entry; let numberFormatId: string | undefined;
    for (const region of regions.at(at)) if (region.style.numberFormatId !== undefined) numberFormatId = region.style.numberFormatId;
    const header = table.header && at.row === table.range.from.row, total = table.totals && at.row === table.range.to.row;
    const key = header ? "headerStyle" : total ? "totalStyle" : "dataStyle", name = header ? "headerCellStyleName" : total ? "totalCellStyleName" : "dataCellStyleName";
    for (const layer of [table, table.columns[at.column - table.range.from.column]!]) {
      const definition = definitions.get(named.get(layer[name]?.toUpperCase() ?? "") ?? "");
      if (definition?.apply.numberFormat && definition.numberFormatId !== undefined) numberFormatId = definition.numberFormatId;
      if (layer[key]?.numberFormatId !== undefined) numberFormatId = layer[key]!.numberFormatId;
    }
    if (source.numberFormatId !== sheet.defaultNumberFormatId) numberFormatId = source.numberFormatId;
    const selected = numberFormatId ?? source.numberFormatId;
    return selected === source.numberFormatId ? source : Object.freeze({ cssClass: source.cssClass, ...(selected === undefined ? {} : { numberFormatId: selected }) });
  };
}

/** An unavailable condition claims its format slot without guessing a lower
 * rule. Callers can show the base appearance and report an incomplete display. */
export function createWorkbookConditionalFormatResolver(sheet: DeepReadonly<WorkbookSheet>): (at: WorkbookCoordinate, evaluations: readonly WorkbookConditionalEvaluation[], source: WorkbookCellFormat) => { readonly format: WorkbookCellFormat; readonly complete: boolean } {
  const definitions = new Map(sheet.conditionalRules?.map(rule => [rule.ruleId, rule]));
  return (at, evaluations, source) => {
    for (const result of evaluations) {
      if (result.state === "miss") continue;
      const rule = definitions.get(result.ruleId);
      if (!rule) continue;
      if ("numberFormatId" in rule && rule.numberFormatId !== undefined) return Object.freeze({ format: result.state === "match" ? Object.freeze({ ...source, numberFormatId: rule.numberFormatId }) : source, complete: result.state !== "unavailable" });
      if (result.state === "unavailable" && rule.stopIfTrue && sheet.conditionalRules?.some(lower => lower.priority > rule.priority && "numberFormatId" in lower && lower.numberFormatId !== undefined
        && lower.ranges.some(range => range.from.row <= at.row && at.row <= range.to.row && range.from.column <= at.column && at.column <= range.to.column))) return Object.freeze({ format: source, complete: false });
    }
    return Object.freeze({ format: source, complete: true });
  };
}
