import type { WorkbookMetadata, WorkbookTable } from "../contracts/workbook.js";
import type { WorkbookOperation } from "../contracts/workbook-operations.js";
import type { DeepReadonly } from "../runtime.js";
import { DocumentError } from "../errors.js";
import { copyFormula } from "../workbook/formula-rewrite.js";
import { checkRange } from "../workbook/ranges.js";
import { WorkbookCellEdit } from "./workbook-cells.js";
import { rewriteWorkbookReferenceSteps } from "./workbook-references.js";

type DefinitionOperation = Extract<WorkbookOperation, { op:
  | "setWorkbookDefinedName" | "removeWorkbookDefinedName" | "setWorkbookTable" | "removeWorkbookTable"
  | "setWorkbookTableStyle" | "removeWorkbookTableStyle" | "setWorkbookConditionalRule" | "removeWorkbookConditionalRule"
  | "setWorkbookDataValidation" | "removeWorkbookDataValidation" }>;
export function isWorkbookDefinitionOperation(operation: WorkbookOperation): operation is DefinitionOperation {
  return ["setWorkbookDefinedName", "removeWorkbookDefinedName", "setWorkbookTable", "removeWorkbookTable", "setWorkbookTableStyle", "removeWorkbookTableStyle",
    "setWorkbookConditionalRule", "removeWorkbookConditionalRule", "setWorkbookDataValidation", "removeWorkbookDataValidation"].includes(operation.op);
}
function invalid(message: string): never { throw new DocumentError("InvalidOperation", message); }
function replace<T>(values: T[], value: T, matches: (entry: T) => boolean): void {
  const index = values.findIndex(matches); if (index < 0) values.push(value); else values[index] = value;
}
function remove<T>(values: T[] | undefined, matches: (entry: T) => boolean): void {
  const index = values?.findIndex(matches) ?? -1; if (index < 0) invalid("The workbook definition does not exist."); values!.splice(index, 1);
}

/** Definitions and their actual cell sources share the candidate/history and
 * final validation boundary. A table never owns a second mutable data store. */
export function* editWorkbookDefinition(workbook: WorkbookMetadata, cells: WorkbookCellEdit, operation: DefinitionOperation,
  allocate: () => string): Generator<{ completed: number; total: number }, void> {
  if (operation.op === "setWorkbookDefinedName") {
    const incoming = operation.name, previous = workbook.names?.find(name => name.nameId === incoming.nameId);
    if (previous && previous.sheetId !== incoming.sheetId) invalid("An existing name cannot silently change scope.");
    if (incoming.sheetId !== undefined && !workbook.sheets.some(sheet => sheet.sheetId === incoming.sheetId)) invalid("A defined name requires an existing worksheet scope.");
    if (workbook.names?.some(name => name.nameId !== incoming.nameId && name.sheetId === incoming.sheetId && name.name.toUpperCase() === incoming.name.toUpperCase())) invalid("Defined names must remain unique in their scope.");
    if (previous && previous.name !== incoming.name) {
      const expression = previous.expression;
      yield* rewriteWorkbookReferenceSteps(workbook, cells, { kind: "renameName", nameId: previous.nameId, to: incoming.name });
      if (incoming.expression === expression) incoming.expression = previous.expression;
    }
    replace(workbook.names ??= [], incoming, name => name.nameId === incoming.nameId); return;
  }
  if (operation.op === "removeWorkbookDefinedName") { remove(workbook.names, name => name.nameId === operation.nameId); return; }
  if (operation.op === "setWorkbookTableStyle") { replace(workbook.tableStyles ??= [], operation.style, style => style.styleId === operation.style.styleId); return; }
  if (operation.op === "removeWorkbookTableStyle") { remove(workbook.tableStyles, style => style.styleId === operation.styleId); return; }
  const sheet = workbook.sheets.find(sheet => sheet.sheetId === operation.sheetId) ?? invalid("An earlier operation removed the worksheet.");
  if (operation.op === "setWorkbookConditionalRule") { replace(sheet.conditionalRules ??= [], operation.rule, rule => rule.ruleId === operation.rule.ruleId); return; }
  if (operation.op === "removeWorkbookConditionalRule") { remove(sheet.conditionalRules, rule => rule.ruleId === operation.ruleId); return; }
  if (operation.op === "setWorkbookDataValidation") { replace(sheet.dataValidations ??= [], operation.rule, rule => rule.ruleId === operation.rule.ruleId); return; }
  if (operation.op === "removeWorkbookDataValidation") { remove(sheet.dataValidations, rule => rule.ruleId === operation.ruleId); return; }
  const previous = sheet.tables?.find(table => table.tableId === (operation.op === "setWorkbookTable" ? operation.table.tableId : operation.tableId));
  if (operation.op === "removeWorkbookTable") {
    if (!previous) invalid("The workbook table does not exist.");
    yield* rewriteWorkbookReferenceSteps(workbook, cells, { kind: "convertTable", table: previous.name });
    remove(sheet.tables, table => table === previous); return;
  }
  const table = operation.table;
  checkRange(table.range);
  if (table.columns.length !== table.range.to.column - table.range.from.column + 1) invalid("Table columns must match the range width.");
  if (workbook.sheets.some(owner => owner.tables?.some(other => other !== previous && (other.tableId === table.tableId || other.name.toUpperCase() === table.name.toUpperCase())))) invalid("Table identities and names must remain unique.");
  const original = previous ? JSON.parse(JSON.stringify(previous)) as WorkbookTable : undefined;
  if (previous) {
    for (const removed of previous.columns.filter(column => !table.columns.some(next => next.columnId === column.columnId)))
      yield* rewriteWorkbookReferenceSteps(workbook, cells, { kind: "deleteColumn", table: previous.name, column: removed.name });
    // Resolve renames by stable column identity through unique temporary names,
    // so a simultaneous A/B swap cannot redirect both references to one column.
    const renamed = previous.columns.flatMap(column => {
      const next = table.columns.find(next => next.columnId === column.columnId);
      return next && next.name !== column.name ? [{ column, next, temporary: "_doffi_" + allocate() }] : [];
    });
    for (const item of renamed) {
      yield* rewriteWorkbookReferenceSteps(workbook, cells, { kind: "renameColumn", table: previous.name, from: item.column.name, to: item.temporary });
      item.column.name = item.temporary;
    }
    for (const item of renamed) {
      yield* rewriteWorkbookReferenceSteps(workbook, cells, { kind: "renameColumn", table: previous.name, from: item.temporary, to: item.next.name });
      item.column.name = item.next.name;
    }
    if (previous.name !== table.name) {
      yield* rewriteWorkbookReferenceSteps(workbook, cells, { kind: "renameTable", from: previous.name, to: table.name });
      previous.name = table.name;
    }
    for (const column of table.columns) for (const key of ["calculatedFormula", "totalFormula"] as const) {
      const before = original!.columns.find(value => value.columnId === column.columnId), rewritten = previous.columns.find(value => value.columnId === column.columnId);
      if (before?.[key] !== undefined && column[key] === before[key]) column[key] = rewritten![key];
    }
  }
  replace(sheet.tables ??= [], table, value => value.tableId === table.tableId);
  if (operation.cellPolicy !== "preserve") yield* materializeTable(cells, sheet.sheetId, table, previous);
}

/** Replacing/removing a definition and restoring it in one batch must not
 * manufacture a revision because of property order or an absent empty list. */
export function restoreWorkbookDefinitionSource(workbook: WorkbookMetadata, before: DeepReadonly<WorkbookMetadata>): void {
  function same(a: unknown, b: unknown): boolean {
    if (a === b) return true;
    if (!a || !b || typeof a !== "object" || typeof b !== "object" || Array.isArray(a) !== Array.isArray(b)) return false;
    const left = Object.keys(a), right = Object.keys(b);
    return left.length === right.length && left.every(key => Object.hasOwn(b, key) && same((a as Record<string, unknown>)[key], (b as Record<string, unknown>)[key]));
  }
  for (const key of ["names", "tableStyles"] as const) {
    if (before[key] === undefined && workbook[key]?.length === 0) delete workbook[key];
    else if (same(workbook[key], before[key])) Object.assign(workbook, { [key]: before[key] });
  }
  for (const sheet of workbook.sheets) {
    const previous = before.sheets.find(value => value.sheetId === sheet.sheetId); if (!previous) continue;
    if (previous.tables?.length === sheet.tables?.length && previous.tables && sheet.tables) {
      const byId = new Map(sheet.tables.map(table => [table.tableId, table]));
      if (previous.tables.every(table => same(byId.get(table.tableId), table))) sheet.tables = previous.tables.map(table => byId.get(table.tableId)!);
    }
    for (const key of ["tables", "conditionalRules", "dataValidations"] as const) {
      if (previous[key] === undefined && sheet[key]?.length === 0) delete sheet[key];
      else if (same(sheet[key], previous[key])) Object.assign(sheet, { [key]: previous[key] });
    }
  }
}

function* materializeTable(cells: WorkbookCellEdit, sheetId: string, table: WorkbookTable, previous: WorkbookTable | undefined): Generator<{ completed: number; total: number }, void> {
  const first = table.range.from.row + Number(table.header), last = table.range.to.row - Number(table.totals);
  const expansion = table.columns.reduce((count, column) => count + (column.calculatedFormula ? Math.max(0, last - first + 1) : 0), 0) + table.columns.length * (Number(table.header) + Number(table.totals));
  // Conservative before allocation; the final model independently checks the
  // actual count. Existing cells do not consume an additional cell allowance.
  let added = 0;
  for (const [offset, column] of table.columns.entries()) {
    const at = table.range.from.column + offset;
    for (let row = first; column.calculatedFormula && row <= last; row++) {
      if (!cells.get(sheetId, row, at)) added++;
      if ((row - first + 1) % 256 === 0) { cells.checkAdditionalCells(added); yield { completed: row - first + 1, total: expansion }; }
    }
    if (table.header && !cells.get(sheetId, table.range.from.row, at)) added++;
    if (table.totals && !cells.get(sheetId, table.range.to.row, at)) added++;
  }
  cells.checkAdditionalCells(added);
  let completed = 0;
  for (const [offset, column] of table.columns.entries()) {
    const at = table.range.from.column + offset;
    if (table.header) cells.set(sheetId, { row: table.range.from.row, column: at, source: { kind: "value", value: { type: "text", value: column.name } } });
    const old = previous?.columns.find(item => item.columnId === column.columnId);
    for (let row = first; column.calculatedFormula && row <= last; row++) {
      // Preserve deliberate exceptions in the old calculated column. Changing
      // its template or adding rows explicitly fills the affected source cells.
      if (!previous || old?.calculatedFormula !== column.calculatedFormula || at !== previous.range.from.column + previous.columns.indexOf(old!)
        || row < previous.range.from.row + Number(previous.header) || row > previous.range.to.row - Number(previous.totals))
        cells.set(sheetId, { row, column: at, source: { kind: "formula", expression: copyFormula(column.calculatedFormula, { row: first, column: at }, { row, column: at }) } });
      if (++completed % 256 === 0) yield { completed, total: expansion };
    }
    if (table.totals && !(previous?.totals && table.range.to.row === previous.range.to.row
      && at === previous.range.from.column + previous.columns.indexOf(old!) && column.totalFormula === old?.totalFormula
      && column.totalFunction === old?.totalFunction && column.totalLabel === old?.totalLabel)) {
      const functions = { sum: 109, min: 105, max: 104, average: 101, count: 103, countNums: 102, stdDev: 107, var: 110 };
      const escaped = column.name.replace(/['#\[\]@]/g, character => "'" + character);
      const formula = column.totalFormula ?? (column.totalFunction !== "none" && column.totalFunction !== "custom" ? "=SUBTOTAL(" + functions[column.totalFunction] + "," + table.name + "[" + escaped + "])" : undefined);
      if (formula) cells.set(sheetId, { row: table.range.to.row, column: at, source: { kind: "formula", expression: formula } });
      else if (column.totalLabel !== undefined) cells.set(sheetId, { row: table.range.to.row, column: at, source: { kind: "value", value: { type: "text", value: column.totalLabel } } });
      else if (previous?.totals && (old?.totalFormula || old?.totalLabel !== undefined || old?.totalFunction !== "none"))
        cells.set(sheetId, { row: table.range.to.row, column: at, source: { kind: "value", value: { type: "blank" } } });
    }
  }
  yield { completed, total: expansion };
}
