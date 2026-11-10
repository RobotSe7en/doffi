import type { WorkbookAutoFilter, WorkbookMetadata, WorkbookRange, WorkbookSheet, WorkbookSort, WorkbookTable } from "../contracts/workbook.js";
import type { WorkbookOperation } from "../contracts/workbook-operations.js";
import { DocumentError } from "../errors.js";
import { copyFormula } from "../workbook/formula-rewrite.js";
import { WorkbookCellEdit } from "./workbook-cells.js";
import { createWorkbookReferenceRewriter, rewriteWorkbookReferenceSteps } from "./workbook-references.js";

type Axis = Extract<WorkbookOperation, { op: "editWorkbookAxis" }>;
interface Transform {
  point(value: number): number | undefined; range(range: WorkbookRange): WorkbookRange | undefined;
  filter(value: WorkbookAutoFilter): WorkbookAutoFilter | undefined; sort(value: WorkbookSort | undefined): WorkbookSort | undefined;
}
interface Plan { readonly table: WorkbookTable; readonly range: WorkbookRange; readonly columns: WorkbookTable["columns"];
  readonly calculated: ReadonlyMap<string, string>; readonly totals: ReadonlyMap<string, string>; readonly addedColumns: ReadonlySet<string>;
  readonly insertedRows?: { readonly from: number; readonly to: number }; readonly totalsDeleted: boolean }
function invalid(message: string): never { throw new DocumentError("InvalidOperation", message); }

/** Keep original table context until all bound reference rewrites finish.
 * Templates are recipes at the data-body origin, separately from actual cell
 * formulas/exceptions. A new row materializes only its own calculated cells. */
export function* planWorkbookAxisTables(workbook: WorkbookMetadata, sheet: WorkbookSheet, cells: WorkbookCellEdit, operation: Axis, transform: Transform,
  allocate: () => string): Generator<{ completed: number; total: number }, Plan[]> {
  const { axis, action, index, count } = operation, plans: Plan[] = [];
  const change = { kind: "axis" as const, sheet: sheet.name, axis, action, index, count };
  // Complete every bound table/column change before capturing any template:
  // one surviving table may refer to a later table's deleted column.
  for (const table of [...sheet.tables ?? []]) {
    const range = transform.range(table.range);
    if (!range) {
      yield* rewriteWorkbookReferenceSteps(workbook, cells, { kind: "convertTable", table: table.name });
      sheet.tables = sheet.tables!.filter(value => value !== table); continue;
    }
    if (axis === "row" && table.header && transform.point(table.range.from.row) === undefined)
      invalid("A surviving table header must be changed explicitly before deleting its worksheet row.");
    if (axis === "column" && action === "delete") for (const [offset, column] of table.columns.entries())
      if (transform.point(table.range.from.column + offset) === undefined)
        yield* rewriteWorkbookReferenceSteps(workbook, cells, { kind: "deleteColumn", table: table.name, column: column.name });
  }
  const rewrite = createWorkbookReferenceRewriter(workbook, change), recipe = createWorkbookReferenceRewriter(workbook, change, { preserveRelativeAxis: true });
  for (const table of sheet.tables ?? []) {
    const range = transform.range(table.range)!;
    const columns = table.columns.filter((_, offset) => axis !== "column" || transform.point(table.range.from.column + offset) !== undefined), addedColumns = new Set<string>();
    if (axis === "column" && action === "insert" && index > table.range.from.column && index <= table.range.to.column) {
      const names = new Set(columns.map(column => column.name.toUpperCase())); let serial = 0;
      const added = Array.from({ length: count }, () => {
        let name: string; do { name = "Column" + ++serial; } while (names.has(name.toUpperCase())); names.add(name.toUpperCase());
        const columnId = allocate(); addedColumns.add(columnId); return { columnId, name, totalFunction: "none" as const };
      });
      columns.splice(index - table.range.from.column, 0, ...added);
    }
    const totalsDeleted = axis === "row" && table.totals && transform.point(table.range.to.row) === undefined;
    const first = table.range.from.row + Number(table.header), last = table.range.to.row - Number(table.totals), targetFirst = range.from.row + Number(table.header);
    const candidate = action === "delete" && axis === "row" && first >= index && first < index + count ? index + count : first;
    const survivor = first <= last && candidate <= last ? candidate : undefined;
    const calculated = new Map<string, string>(), totals = new Map<string, string>();
    for (const [offset, column] of table.columns.entries()) {
      const oldColumn = table.range.from.column + offset, targetColumn = axis === "column" ? transform.point(oldColumn) : oldColumn;
      if (targetColumn === undefined) continue;
      if (column.calculatedFormula) {
        const target = { sheet: sheet.name, row: targetFirst, column: targetColumn };
        if (axis === "row" && survivor === undefined) calculated.set(column.columnId, recipe(column.calculatedFormula, sheet.name, { row: first, column: oldColumn }, table.name, target));
        else {
          const origin = { row: survivor ?? first, column: oldColumn }, moved = { row: axis === "row" ? transform.point(origin.row)! : origin.row, column: targetColumn };
          const source = copyFormula(column.calculatedFormula, { row: first, column: oldColumn }, origin);
          calculated.set(column.columnId, copyFormula(rewrite(source, sheet.name, origin, table.name, { sheet: sheet.name, ...moved }), moved, target));
        }
      }
      if (column.totalFormula) totals.set(column.columnId, rewrite(column.totalFormula, sheet.name, { row: table.range.to.row, column: oldColumn }, table.name,
        { sheet: sheet.name, row: range.to.row, column: targetColumn }));
    }
    const insertedRows = axis === "row" && action === "insert" && index > table.range.from.row && index <= table.range.to.row
      ? { from: index, to: index + count - 1 } : undefined;
    plans.push({ table, range, columns, calculated, totals, addedColumns, ...(insertedRows ? { insertedRows } : {}), totalsDeleted });
  }
  return plans;
}

export function* applyWorkbookAxisTables(sheet: WorkbookSheet, cells: WorkbookCellEdit, plans: readonly Plan[], transform: Transform): Generator<{ completed: number; total: number }, void> {
  let completed = 0, added = 0;
  const expansion = plans.reduce((count, plan) => count + (plan.table.header ? plan.addedColumns.size : 0)
    + (plan.insertedRows ? (plan.insertedRows.to - plan.insertedRows.from + 1) * plan.calculated.size : 0), 0);
  cells.checkExpandedCells(expansion);
  // Preflight the cumulative actual expansion before writing any table cells.
  for (const plan of plans) {
    if (plan.table.header) for (const [offset, column] of plan.columns.entries()) if (plan.addedColumns.has(column.columnId) && !cells.get(sheet.sheetId, plan.range.from.row, plan.range.from.column + offset)) added++;
    if (plan.insertedRows) for (const [offset, column] of plan.columns.entries()) if (plan.calculated.has(column.columnId)) for (let row = plan.insertedRows.from; row <= plan.insertedRows.to; row++) {
      if (!cells.get(sheet.sheetId, row, plan.range.from.column + offset)) added++;
      if (++completed % 256 === 0) { cells.checkAdditionalCells(added); yield { completed, total: 0 }; }
    }
  }
  cells.checkAdditionalCells(added); completed = 0;
  for (const plan of plans) {
    const table = plan.table, includesTotals = table.autoFilter?.range.to.row === table.range.to.row,
      filter = table.autoFilter && transform.filter(table.autoFilter), sort = transform.sort(table.sort);
    table.range = plan.range; table.columns = plan.columns;
    if (plan.totalsDeleted) table.totals = false;
    if (filter) filter.range = { from: { ...table.range.from }, to: { ...table.range.to, row: table.range.to.row - Number(table.totals && !includesTotals) } };
    if (filter) table.autoFilter = filter; else delete table.autoFilter;
    if (sort) table.sort = sort; else delete table.sort;
    const first = table.range.from.row + Number(table.header);
    for (const [offset, column] of table.columns.entries()) {
      const at = table.range.from.column + offset, formula = plan.calculated.get(column.columnId), total = plan.totals.get(column.columnId);
      if (formula !== undefined) column.calculatedFormula = formula;
      if (total !== undefined) column.totalFormula = total;
      if (table.header && plan.addedColumns.has(column.columnId)) cells.set(sheet.sheetId, { row: table.range.from.row, column: at, source: { kind: "value", value: { type: "text", value: column.name } } });
      if (plan.insertedRows && formula) for (let row = plan.insertedRows.from; row <= plan.insertedRows.to; row++) {
        cells.set(sheet.sheetId, { row, column: at, source: { kind: "formula", expression: copyFormula(formula, { row: first, column: at }, { row, column: at }) } });
        if (++completed % 256 === 0) yield { completed, total: 0 };
      }
    }
  }
  yield { completed, total: 0 };
}
