import type { DeepReadonly, WorkbookAutoFilter, WorkbookFilterCriterion, WorkbookSort, WorkbookTable } from "@doffi/core";
import type { OutputOpc } from "../export/opc.js";
import { declaration, empty } from "../export/xml.js";
import { S } from "../ooxml/xml.js";
import type { WorkbookStylesWriter } from "./styles.js";
import { bool, element, formulaText, rangeRef, sheetContentType, xstring } from "./xml.js";

function filterCriterion(criterion: DeepReadonly<WorkbookFilterCriterion>): string {
  switch (criterion.kind) {
    case "values":
      return element("filters", criterion.values.map(value => empty("filter", { val: xstring(value) })).join("") +
        criterion.dates.map(({ grouping, ...parts }) => empty("dateGroupItem", { ...parts, dateTimeGrouping: grouping })).join(""),
      { blank: bool(criterion.blank) });
    case "custom":
      return element("customFilters", criterion.tests.map(test =>
        empty("customFilter", { operator: test.operator, val: xstring(test.value) })).join(""), { and: bool(criterion.and) });
    case "top":
      return empty("top10", { top: bool(criterion.top), percent: bool(criterion.percent), val: criterion.count });
    case "dynamic":
      return empty("dynamicFilter", { type: criterion.type });
  }
}
function sortState(sort: DeepReadonly<WorkbookSort>): string {
  return element("sortState", sort.keys.map(key => empty("sortCondition", { descending: bool(key.descending), ref: rangeRef({
    from: { row: sort.range.from.row, column: key.column }, to: { row: sort.range.to.row, column: key.column },
  }) })).join(""), { ref: rangeRef(sort.range), caseSensitive: bool(sort.caseSensitive), sortMethod: sort.method });
}
export function autoFilter(filter: DeepReadonly<WorkbookAutoFilter>): string {
  return element("autoFilter", filter.columns.map(column => element("filterColumn", column.criterion ? filterCriterion(column.criterion) : "", {
    colId: column.column - filter.range.from.column, hiddenButton: bool(column.hiddenButton), showButton: bool(column.showButton),
  })).join("") + (filter.sort ? sortState(filter.sort) : ""), { ref: rangeRef(filter.range) });
}

type RegionStyles = Pick<DeepReadonly<WorkbookTable>, "headerStyle" | "dataStyle" | "totalStyle" | "headerCellStyleName" | "dataCellStyleName" | "totalCellStyleName">;
function regionStyles(table: RegionStyles, styles: WorkbookStylesWriter): Record<string, string | number | undefined> {
  return {
    headerRowDxfId: styles.differential(table.headerStyle?.cssClass, table.headerStyle?.numberFormatId),
    dataDxfId: styles.differential(table.dataStyle?.cssClass, table.dataStyle?.numberFormatId),
    totalsRowDxfId: styles.differential(table.totalStyle?.cssClass, table.totalStyle?.numberFormatId),
    headerRowCellStyle: table.headerCellStyleName === undefined ? undefined : xstring(table.headerCellStyleName),
    dataCellStyle: table.dataCellStyleName === undefined ? undefined : xstring(table.dataCellStyleName),
    totalsRowCellStyle: table.totalCellStyleName === undefined ? undefined : xstring(table.totalCellStyleName),
  };
}
export function writeTable(opc: OutputOpc, styles: WorkbookStylesWriter, sheetPart: string, table: DeepReadonly<WorkbookTable>, id: number): string {
  const part = opc.part("xl/tables/table" + id + ".xml", sheetContentType("table"));
  const attrs = { id, name: xstring(table.objectName ?? table.name), displayName: xstring(table.name),
    comment: table.description === undefined ? undefined : xstring(table.description), ref: rangeRef(table.range),
    headerRowCount: bool(table.header), totalsRowCount: bool(table.totals), totalsRowShown: bool(table.totalsPreviouslyShown),
    ...regionStyles(table, styles), tableBorderDxfId: styles.differential(table.borderStyle?.cssClass, table.borderStyle?.numberFormatId),
    headerRowBorderDxfId: styles.differential(table.headerBorderStyle?.cssClass, table.headerBorderStyle?.numberFormatId),
    totalsRowBorderDxfId: styles.differential(table.totalBorderStyle?.cssClass, table.totalBorderStyle?.numberFormatId) };
  const body = (table.autoFilter ? autoFilter(table.autoFilter) : "") + (table.sort ? sortState(table.sort) : "") +
    element("tableColumns", table.columns.map((column, index) => element("tableColumn",
      (column.calculatedFormula ? element("calculatedColumnFormula", formulaText(column.calculatedFormula)) : "") +
      (column.totalFormula ? element("totalsRowFormula", formulaText(column.totalFormula)) : ""), {
        id: index + 1, name: xstring(column.name), totalsRowFunction: column.totalFunction,
        totalsRowLabel: column.totalLabel === undefined ? undefined : xstring(column.totalLabel), ...regionStyles(column, styles),
      })).join(""), { count: table.columns.length }) +
    empty("tableStyleInfo", { name: styles.tableName(table.style), showFirstColumn: bool(table.showFirstColumn),
      showLastColumn: bool(table.showLastColumn), showRowStripes: bool(table.rowStripes), showColumnStripes: bool(table.columnStripes) });
  part.add(declaration + element("table", body, { xmlns: S, ...attrs }));
  part.finish();
  return opc.relate(sheetPart, "table", "../tables/table" + id + ".xml");
}
