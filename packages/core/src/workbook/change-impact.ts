import { DocumentError } from "../errors.js";
import { freeze } from "../document/validation.js";
import type { WorkbookModel, WorkbookRangeChange } from "../runtime.js";
import type { WorkbookCalculationResult } from "./calculation.js";
import { workbookCalculationSource } from "./calculation-source.js";
import type { ScalarResolved } from "./scalar.js";
import { compareWorkbooks } from "./source-changes.js";
import { currentWorkbookModel, workbookSource } from "./source.js";

/** Cell-level invalidation, never an authorization token or an editable
 * source. Layout owners must separately compare metadata, styles and content. */
export interface WorkbookChangeImpact {
  readonly documentId: string;
  readonly fromRevision: number;
  readonly revision: number;
  readonly sourceRanges: readonly WorkbookRangeChange[];
  readonly calculatedRanges: readonly WorkbookRangeChange[];
}
function sameValue(a: ScalarResolved | undefined, b: ScalarResolved | undefined): boolean {
  if (a === b) return true;
  if (!a || !b) return false;
  if ("type" in a) return "type" in b && a.type === b.type && (a.type === "blank" || "value" in b && Object.is(a.value, b.value));
  return !("type" in b) && a.reason === b.reason && a.functionName === b.functionName;
}
/** Compare actual immutable models and optional completed batches. Imported
 * caches cannot stand in for a missing batch. Equal revision numbers do not
 * replace the private model/batch ownership checks. */
export function readWorkbookChangeImpact(previous: WorkbookModel, model: WorkbookModel, previousCalculation?: WorkbookCalculationResult, calculation?: WorkbookCalculationResult): WorkbookChangeImpact {
  previous = currentWorkbookModel(previous); model = currentWorkbookModel(model);
  if (!workbookSource(previous) || !workbookSource(model)) throw new DocumentError("InvalidOperation", "Workbook change impact requires validated first-party models.");
  const before = previous.snapshot().manifest, after = model.snapshot().manifest;
  if (before.documentId !== after.documentId) throw new DocumentError("RevisionConflict", "Workbook change impact cannot switch document identity.");
  const left = previousCalculation === undefined ? undefined : workbookCalculationSource(previous, previousCalculation).values;
  const right = calculation === undefined ? undefined : workbookCalculationSource(model, calculation).values;
  const points = new Map<string, { sheetId: string; row: number; column: number }>();
  if (left || right) for (const batch of [previousCalculation, calculation]) for (const cell of batch?.cells ?? []) {
    const key = JSON.stringify([cell.sheetId, cell.row, cell.column]);
    if (!left || !right || !sameValue(left.get(key), right.get(key))) points.set(key, cell);
  }
  const sheets = new Map<string, number>();
  for (const sheet of [...model.snapshot().metadata.workbook.sheets, ...previous.snapshot().metadata.workbook.sheets]) if (!sheets.has(sheet.sheetId)) sheets.set(sheet.sheetId, sheets.size);
  const calculatedRanges: WorkbookRangeChange[] = [];
  for (const cell of [...points.values()].sort((a, b) => sheets.get(a.sheetId)! - sheets.get(b.sheetId)! || a.row - b.row || a.column - b.column)) {
    const last = calculatedRanges.at(-1);
    if (last?.sheetId === cell.sheetId && last.range.from.row === cell.row && last.range.to.column + 1 === cell.column) {
      calculatedRanges[calculatedRanges.length - 1] = { sheetId: cell.sheetId, range: { from: last.range.from, to: { row: cell.row, column: cell.column } } };
    } else calculatedRanges.push({ sheetId: cell.sheetId, range: { from: { row: cell.row, column: cell.column }, to: { row: cell.row, column: cell.column } } });
  }
  return freeze({ documentId: after.documentId, fromRevision: before.revision, revision: after.revision,
    sourceRanges: compareWorkbooks(previous, model).changedRanges, calculatedRanges });
}
