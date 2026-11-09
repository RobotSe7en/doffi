import type { WorkbookAutoFilter, WorkbookRange, WorkbookSheet, WorkbookSort } from "../contracts/workbook.js";
import { DocumentError } from "../errors.js";
import { rangeContains, rangesIntersect } from "./workbook-cells.js";

const contains = (outer: WorkbookRange, inner: WorkbookRange) => rangeContains(outer, inner.from.row, inner.from.column) && rangeContains(outer, inner.to.row, inner.to.column);
function invalid(message: string): never { throw new DocumentError("InvalidOperation", message); }

/** Definitions relocate only after all formulas have used their old geometry.
 * Native identities, styles and column recipes stay on the same source objects. */
export function planWorkbookDefinitionMove(source: WorkbookSheet, target: WorkbookSheet, range: WorkbookRange, destination: WorkbookRange): () => void {
  const intersected = source.tables?.filter(table => rangesIntersect(table.range, range)) ?? [], tables = intersected.filter(table => contains(range, table.range)), moved = new Set(tables);
  const header = (table: NonNullable<WorkbookSheet["tables"]>[number]): WorkbookRange => ({ from: table.range.from, to: { row: table.range.from.row, column: table.range.to.column } });
  if (intersected.some(table => !moved.has(table) && table.header && rangesIntersect(header(table), range)))
    invalid("Moving a table header requires its complete table or an explicit column definition change.");
  for (const table of target.tables ?? []) if (!moved.has(table)) {
    if (tables.length && rangesIntersect(table.range, destination)) invalid("A cut cannot overwrite a native table definition.");
    if (table.header && rangesIntersect(header(table), destination)) invalid("A cell move cannot overwrite a surviving table header.");
  }
  // Complete definitions relocate; partial data moves retain their source and
  // target table/filter identities at their addresses. Actual cell formulas
  // (including calculated exceptions and totals) use the shared move mapper.
  const filter = source.autoFilter && contains(range, source.autoFilter.range) ? source.autoFilter : undefined;
  if (target.autoFilter && target.autoFilter !== filter && (filter || tables.length && rangesIntersect(target.autoFilter.range, destination)))
    invalid("A cut cannot replace a worksheet auto-filter definition.");
  const delta = { row: destination.from.row - range.from.row, column: destination.from.column - range.from.column };
  const rectangle = (value: WorkbookRange): WorkbookRange => ({ from: { row: value.from.row + delta.row, column: value.from.column + delta.column },
    to: { row: value.to.row + delta.row, column: value.to.column + delta.column } });
  function sort(value: WorkbookSort): WorkbookSort { return { ...value, range: rectangle(value.range), keys: value.keys.map(key => ({ ...key, column: key.column + delta.column })) }; }
  function autoFilter(value: WorkbookAutoFilter): WorkbookAutoFilter { return { ...value, range: rectangle(value.range), columns: value.columns.map(column => ({ ...column, column: column.column + delta.column })),
    ...(value.sort ? { sort: sort(value.sort) } : {}) }; }
  return () => {
    if (source !== target && tables.length) {
      source.tables = source.tables!.filter(table => !tables.includes(table)); (target.tables ??= []).push(...tables);
    }
    for (const table of tables) {
      table.range = rectangle(table.range);
      if (table.autoFilter) table.autoFilter = autoFilter(table.autoFilter);
      if (table.sort) table.sort = sort(table.sort);
    }
    if (filter) { delete source.autoFilter; target.autoFilter = autoFilter(filter); }
  };
}
