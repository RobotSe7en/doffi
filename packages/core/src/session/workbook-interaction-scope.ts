import { DocumentError } from "../errors.js";
import type { DeepReadonly, WorkbookModel } from "../runtime.js";
import type { Annotation } from "../contracts/annotations.js";
import type { WorkbookRange } from "../contracts/workbook.js";
import { WorkbookInteractionScopeSchema } from "../contracts/workbook-interaction.js";
import type { WorkbookInteractionRange, WorkbookInteractionScope } from "../contracts/workbook-interaction.js";
import { checkRange } from "../workbook/ranges.js";

type Range = DeepReadonly<WorkbookRange>;
type Located = DeepReadonly<WorkbookInteractionRange>;
export function workbookContextIntersects(left: Range, right: Range): boolean {
  return left.from.row <= right.to.row && right.from.row <= left.to.row && left.from.column <= right.to.column && right.from.column <= left.to.column;
}
export function workbookContextContains(ranges: readonly Located[], sheetId: string, row: number, column: number, step: () => void): boolean {
  return ranges.some(entry => { step(); return entry.sheetId === sheetId && entry.range.from.row <= row && entry.range.to.row >= row && entry.range.from.column <= column && entry.range.to.column >= column; });
}
/** Rectangle subtraction checks the complete union, including sparse blanks.
 * Every split consumes the caller's work budget; no grid cells are expanded. */
export function workbookContextCovers(ranges: readonly Located[], target: Located, step: () => void): boolean {
  let remaining: Range[] = [target.range];
  for (const entry of ranges) {
    step(); if (entry.sheetId !== target.sheetId) continue;
    const cover = entry.range, next: Range[] = [];
    for (const range of remaining) {
      step(); if (!workbookContextIntersects(range, cover)) { next.push(range); continue; }
      const top = Math.max(range.from.row, cover.from.row), bottom = Math.min(range.to.row, cover.to.row), left = Math.max(range.from.column, cover.from.column), right = Math.min(range.to.column, cover.to.column);
      if (range.from.row < top) next.push({ from: range.from, to: { row: top - 1, column: range.to.column } });
      if (bottom < range.to.row) next.push({ from: { row: bottom + 1, column: range.from.column }, to: range.to });
      if (range.from.column < left) next.push({ from: { row: top, column: range.from.column }, to: { row: bottom, column: left - 1 } });
      if (right < range.to.column) next.push({ from: { row: top, column: right + 1 }, to: { row: bottom, column: range.to.column } });
    }
    remaining = next; if (!remaining.length) return true;
  }
  return false;
}
export interface WorkbookWritableContext {
  readonly scope: WorkbookInteractionScope;
  readonly ranges: readonly WorkbookInteractionRange[];
  readonly objectIds: readonly string[];
  readonly newObjectSheetIds: readonly string[];
}
/** The host declares scope; identities are validated against this immutable
 * source and never infer additional permission from a formula or native ID. */
export function workbookWritableContext(model: WorkbookModel, annotations: readonly DeepReadonly<Annotation>[], input: DeepReadonly<WorkbookInteractionScope> | undefined, step: () => void): WorkbookWritableContext {
  const workbook = model.snapshot().metadata.workbook;
  const anchors = annotations.flatMap(annotation => {
    const anchor = annotation.anchor;
    if (anchor.kind !== "cells") throw new DocumentError("AnchorNotFound", "Workbook requests require current worksheet range anchors.");
    return anchor.ranges.map(range => ({ sheetId: anchor.sheetId, range }));
  });
  const parsed = WorkbookInteractionScopeSchema.safeParse(input ?? { kind: "ranges", ranges: anchors });
  if (!parsed.success) throw new DocumentError("InvalidOperation", "Workbook request scope requires supported ranges, tables or explicit global access.");
  const scope = parsed.data, sheetIds = new Set(workbook.sheets.map(sheet => sheet.sheetId)), objects = new Set<string>();
  for (const object of workbook.objects ?? []) { step(); objects.add(object.nodeId); }
  const objectIds = scope.kind === "global" ? [...objects] : scope.objectIds ?? [];
  const newObjectSheetIds = scope.kind === "global" ? [...sheetIds] : scope.newObjectSheetIds ?? [];
  if (new Set(newObjectSheetIds).size !== newObjectSheetIds.length || newObjectSheetIds.some(id => { step(); return !sheetIds.has(id); })) throw new DocumentError("InvalidOperation", "New workbook drawings require unique current worksheet permissions.");
  if (new Set(objectIds).size !== objectIds.length || objectIds.some(id => { step(); return !objects.has(id); })) throw new DocumentError("InvalidOperation", "Workbook request objects must have unique identities from the current source.");
  let ranges: WorkbookInteractionRange[];
  if (scope.kind === "global") ranges = workbook.sheets.map(sheet => ({ sheetId: sheet.sheetId, range: { from: { row: 0, column: 0 }, to: { row: 1_048_575, column: 16_383 } } }));
  else if (scope.kind === "tables") {
    if (new Set(scope.tableIds).size !== scope.tableIds.length) throw new DocumentError("InvalidOperation", "Workbook request tables must be unique.");
    const tables = new Map<string, WorkbookInteractionRange>();
    for (const sheet of workbook.sheets) for (const table of sheet.tables ?? []) { step(); tables.set(table.tableId, { sheetId: sheet.sheetId, range: table.range }); }
    ranges = scope.tableIds.map(id => { step(); const entry = tables.get(id); if (!entry) throw new DocumentError("InvalidOperation", "A workbook request table no longer exists."); return { sheetId: entry.sheetId, range: { from: { ...entry.range.from }, to: { ...entry.range.to } } }; });
  } else ranges = scope.ranges;
  const unique = new Map<string, WorkbookInteractionRange>();
  for (const entry of ranges) { step(); checkRange(entry.range); if (!sheetIds.has(entry.sheetId)) throw new DocumentError("InvalidOperation", "A workbook request worksheet no longer exists."); unique.set(JSON.stringify(entry), entry); }
  ranges = [...unique.values()];
  for (const anchor of anchors) if (!workbookContextCovers(ranges, anchor, step)) throw new DocumentError("AnchorNotFound", "The declared writable scope must cover every selected annotation anchor.");
  return { scope, ranges, objectIds, newObjectSheetIds };
}
