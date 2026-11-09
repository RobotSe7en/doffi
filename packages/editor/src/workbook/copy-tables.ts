import { DocumentError, rewriteWorkbookFormula, WorkbookRegionIndex } from "@doffi/core";
import type { DeepReadonly, WorkbookCoordinate, WorkbookFormulaRewriteBindings, WorkbookMetadata, WorkbookRange, WorkbookSheet, WorkbookTable } from "@doffi/core";

interface Budget { guard(): void; step(): void; account(units: number): void; yield?: () => Promise<void> }
interface WorkbookTableCopyPlan {
  bind(expression: string, from: WorkbookCoordinate, to: WorkbookCoordinate): string;
  finish(copy: (expression: string, from: WorkbookCoordinate, to: WorkbookCoordinate) => Promise<string>): Promise<WorkbookTable[]>;
}
const contains = (a: DeepReadonly<WorkbookRange>, b: DeepReadonly<WorkbookRange>) => a.from.row <= b.from.row && a.from.column <= b.from.column && a.to.row >= b.to.row && a.to.column >= b.to.column;
const overlap = (a: DeepReadonly<WorkbookRange>, b: DeepReadonly<WorkbookRange>) => a.from.row <= b.to.row && b.from.row <= a.to.row && a.from.column <= b.to.column && b.from.column <= a.to.column;
/** Complete table copies own definitions independently of their actual cells.
 * Recipes are copied as metadata; exceptions, rich headers and blank overrides
 * come exclusively from the ordinary cell-copy transaction. */
export async function createWorkbookTableCopyPlan(workbook: DeepReadonly<WorkbookMetadata>, source: DeepReadonly<WorkbookSheet>, target: DeepReadonly<WorkbookSheet>,
  range: DeepReadonly<WorkbookRange>, destination: DeepReadonly<WorkbookRange>, budget: Budget): Promise<WorkbookTableCopyPlan> {
  const selected = (source.tables ?? []).filter(table => contains(range, table.range));
  const ids = new Set<string>(), names = new Set<string>(); let serial = 0, steps = 0;
  const tick = async () => { budget.guard(); budget.step(); if (++steps % 256 === 0 && budget.yield) { await budget.yield(); budget.guard(); } };
  for (const definition of workbook.names ?? []) { await tick(); ids.add(definition.nameId); names.add(definition.name.toUpperCase()); }
  for (const sheet of workbook.sheets) for (const table of sheet.tables ?? []) {
    await tick(); ids.add(table.tableId); names.add(table.name.toUpperCase()); names.add((table.objectName ?? table.name).toUpperCase());
    for (const column of table.columns) { await tick(); ids.add(column.columnId); }
  }
  async function allocate(): Promise<string> { let id: string; do { await tick(); id = "copied-table-" + ++serial; } while (ids.has(id)); ids.add(id); return id; }
  const delta = { row: destination.from.row - range.from.row, column: destination.from.column - range.from.column };
  const point = (at: DeepReadonly<WorkbookCoordinate>): WorkbookCoordinate => ({ row: at.row + delta.row, column: at.column + delta.column });
  const rectangle = (value: DeepReadonly<WorkbookRange>): WorkbookRange => ({ from: point(value.from), to: point(value.to) });
  const plans: { original: DeepReadonly<WorkbookTable>; table: WorkbookTable; bindings: WorkbookFormulaRewriteBindings }[] = [];
  for (const original of selected) {
    await tick(); const serialized = JSON.stringify(original); budget.account(serialized.length);
    const table = JSON.parse(serialized) as WorkbookTable; table.range = rectangle(original.range);
    for (const existing of target.tables ?? []) { await tick(); if (overlap(table.range, existing.range)) throw new DocumentError("InvalidOperation", "A copied table cannot overlap an existing table definition."); }
    let suffix = 2, name: string;
    do {
      await tick(); const tail = "_" + suffix++; let stem = original.name.slice(0, 255 - tail.length);
      if (/[\uD800-\uDBFF]$/.test(stem)) stem = stem.slice(0, -1); name = stem + tail;
    } while (names.has(name.toUpperCase()));
    names.add(name.toUpperCase()); table.name = name; table.tableId = await allocate(); if (table.objectName !== undefined) table.objectName = name;
    for (const column of table.columns) { await tick(); column.columnId = await allocate(); }
    const sort = async (value: WorkbookTable["sort"]) => { if (value) { value.range = rectangle(value.range); for (const key of value.keys) { await tick(); key.column += delta.column; } } };
    if (table.autoFilter) { table.autoFilter.range = rectangle(table.autoFilter.range); for (const column of table.autoFilter.columns) { await tick(); column.column += delta.column; } await sort(table.autoFilter.sort); }
    await sort(table.sort); budget.account(JSON.stringify(table).length);
    const tokens = new Map([[original.name.toUpperCase(), table.name]]);
    plans.push({ original, table, bindings: { copiedTables: tokens, isTableName: name => tokens.has(name.toUpperCase()) } });
  }
  const index = new WorkbookRegionIndex(plans.map(plan => plan.original.range)), byRange = new Map(plans.map(plan => [plan.original.range, plan]));
  function rebind(expression: string, plan: typeof plans[number], to: WorkbookCoordinate): string {
    const place = { ...to, sheet: target.name }; return rewriteWorkbookFormula(expression, place, place, { kind: "copy" }, plan.bindings);
  }
  function bind(expression: string, from: WorkbookCoordinate, to: WorkbookCoordinate): string {
    const range = index.at(from.row, from.column), plan = range && byRange.get(range); if (!plan) return expression;
    // Copying a complete table changes its self identity. Other tables and
    // formulas outside this copied table retain ordinary copy bindings.
    return rebind(expression, plan, to);
  }
  return {
    bind,
    async finish(copy: (expression: string, from: WorkbookCoordinate, to: WorkbookCoordinate) => Promise<string>): Promise<WorkbookTable[]> {
      for (const plan of plans) for (const [index, column] of plan.table.columns.entries()) {
        await tick();
        for (const key of ["calculatedFormula", "totalFormula"] as const) if (column[key] !== undefined) {
          const from = { row: key === "calculatedFormula" ? plan.original.range.from.row + Number(plan.original.header) : plan.original.range.to.row, column: plan.original.range.from.column + index }, to = point(from);
          column[key] = rebind(await copy(column[key], from, to), plan, to); budget.account(column[key]!.length);
        }
      }
      budget.guard(); return plans.map(plan => plan.table);
    },
  };
}
