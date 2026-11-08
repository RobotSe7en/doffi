import type { WorkbookAutoFilter, WorkbookRange, WorkbookSheet, WorkbookSort } from "../contracts/workbook.js";
import { DocumentError } from "../errors.js";
import { rangeContains, rangesIntersect } from "./workbook-cells.js";

const contains = (outer: WorkbookRange, inner: WorkbookRange) => rangeContains(outer, inner.from.row, inner.from.column) && rangeContains(outer, inner.to.row, inner.to.column);
function invalid(message: string): never { throw new DocumentError("InvalidOperation", message); }

/** Definitions relocate only after all formulas have used their old geometry.
 * Native identities, styles and column recipes stay on the same source objects. */
export function planWorkbookDefinitionMove(source: WorkbookSheet, target: WorkbookSheet, range: WorkbookRange, destination: WorkbookRange): () => void {
  const tables = source.tables?.filter(table => rangesIntersect(table.range, range)) ?? [];
  if (tables.some(table => !contains(range, table.range))) invalid("A cut must contain each intersected native table completely.");
  if (target.tables?.some(table => !tables.includes(table) && rangesIntersect(table.range, destination))) invalid("A cut cannot overwrite a native table definition.");
  const filter = source.autoFilter && rangesIntersect(source.autoFilter.range, range) ? source.autoFilter : undefined;
  if (filter && !contains(range, filter.range)) invalid("A cut must contain an intersected auto-filter completely.");
  if (target.autoFilter && target.autoFilter !== filter && (filter || rangesIntersect(target.autoFilter.range, destination))) invalid("A cut cannot replace a worksheet auto-filter definition.");
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
