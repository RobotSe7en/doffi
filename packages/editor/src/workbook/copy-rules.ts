import { DocumentError } from "@doffi/core";
import type { DeepReadonly, WorkbookConditionalRule, WorkbookCoordinate, WorkbookDataValidation, WorkbookMetadata, WorkbookOperation, WorkbookRange, WorkbookSheet } from "@doffi/core";

type Rule = WorkbookConditionalRule | WorkbookDataValidation;
/** Consecutive copied cells become row runs and then maximal vertical bands.
 * A blank mask is bounded by fragments rather than one operation per cell. */
export function createWorkbookCopyFootprint(): { cell(row: number, column: number, copied: boolean): void; finish(): WorkbookRange[] } {
  const ranges: WorkbookRange[] = []; let rowNumber = -1, run: WorkbookRange | undefined;
  let previous = new Map<string, WorkbookRange>(), current = new Map<string, WorkbookRange>();
  function flush(): void {
    if (!run) return;
    const key = run.from.column + ":" + run.to.column, prior = previous.get(key);
    if (prior && prior.to.row === rowNumber - 1) { prior.to.row = rowNumber; current.set(key, prior); }
    else { ranges.push(run); current.set(key, run); }
    if (ranges.length > 4096) throw new DocumentError("ResourceLimitExceeded", "Rule copy exceeds its footprint fragment budget.");
    run = undefined;
  }
  return { cell(row, column, copied) {
    if (row !== rowNumber) { flush(); previous = current; current = new Map(); rowNumber = row; }
    if (run && (!copied || run.to.column !== column - 1)) flush();
    if (copied) { if (run) run.to.column = column; else run = { from: { row, column }, to: { row, column } }; }
  }, finish() { flush(); return ranges; } };
}

export async function copyWorkbookRules(workbook: DeepReadonly<WorkbookMetadata>, source: DeepReadonly<WorkbookSheet>, target: DeepReadonly<WorkbookSheet>, from: DeepReadonly<WorkbookRange>, to: DeepReadonly<WorkbookRange>, ranges: WorkbookRange[], validations: boolean,
  copy: (expression: string, from: WorkbookCoordinate, to: WorkbookCoordinate) => Promise<string>, budget: { step(): void; account(units: number): void; guard(): void; yield?: () => Promise<void> }): Promise<Extract<WorkbookOperation, { op: "replaceWorkbookRules" }> | undefined> {
  if (!ranges.length) return;
  const conditionalRules: WorkbookConditionalRule[] = [], dataValidations: WorkbookDataValidation[] = [];
  const used = new Set<string>(), priorities = new Set<number>(); let serial = 0, work = 0;
  async function tick(): Promise<void> { budget.step(); if (++work % 256 === 0 && budget.yield) { await budget.yield(); budget.guard(); } }
  for (const sheet of workbook.sheets) for (const rule of [...sheet.conditionalRules ?? [], ...sheet.dataValidations ?? []]) { used.add(rule.ruleId); await tick(); }
  let priority = 0;
  for (const rule of target.conditionalRules ?? []) { priorities.add(rule.priority); priority = Math.max(priority, rule.priority); await tick(); }
  const rowDelta = to.from.row - from.from.row, columnDelta = to.from.column - from.from.column;
  if (priority + (source.conditionalRules?.length ?? 0) > 2147483647) priority = 0;
  let fragments = 0;
  for (const key of ["conditionalRules", "dataValidations"] as const) {
    if (key === "dataValidations" && !validations) continue;
    const originals = [...source[key] ?? []].sort((a, b) => ("priority" in a ? a.priority : 0) - ("priority" in b ? b.priority : 0));
    for (const rule of originals) {
      const parts: WorkbookRange[] = [];
      for (const area of rule.ranges) for (const cover of ranges) {
        await tick(); const firstRow = Math.max(area.from.row + rowDelta, cover.from.row), lastRow = Math.min(area.to.row + rowDelta, cover.to.row);
        const firstColumn = Math.max(area.from.column + columnDelta, cover.from.column), lastColumn = Math.min(area.to.column + columnDelta, cover.to.column);
        if (firstRow <= lastRow && firstColumn <= lastColumn) { parts.push({ from: { row: firstRow, column: firstColumn }, to: { row: lastRow, column: lastColumn } }); if (++fragments > 4096) throw new DocumentError("ResourceLimitExceeded", "Rule copy exceeds its rule fragment budget."); }
      }
      if (!parts.length) continue;
      const serialized = JSON.stringify({ ...rule, ranges: parts, origin: parts[0]!.from }); budget.account(serialized.length);
      const next = JSON.parse(serialized) as Rule;
      do { next.ruleId = "copied-rule-" + ++serial; await tick(); } while (used.has(next.ruleId)); used.add(next.ruleId);
      const update = (expression: string) => copy(expression, rule.origin, next.origin);
      if ("expression" in next) next.expression = await update(next.expression);
      if ("formulas" in next) { const formulas: string[] = []; for (const expression of next.formulas) formulas.push(await update(expression)); next.formulas = formulas; }
      if (next.kind === "colorScale") for (const stop of next.stops) if (stop.kind === "formula") stop.expression = await update(stop.expression);
      if ("priority" in next) {
        do { ++priority; await tick(); } while (priorities.has(priority)); priorities.add(priority); next.priority = priority; conditionalRules.push(next);
      } else dataValidations.push(next);
    }
  }
  budget.guard(); return { op: "replaceWorkbookRules", sheetId: target.sheetId, ranges, conditionalRules, ...(validations ? { dataValidations } : {}) };
}
