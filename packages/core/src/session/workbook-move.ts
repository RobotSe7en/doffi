import type { Annotation } from "../contracts/annotations.js";
import type { WorkbookMetadata, WorkbookRange } from "../contracts/workbook.js";
import type { WorkbookOperation } from "../contracts/workbook-operations.js";
import type { WorkbookCell, WorkbookStructureMapping } from "../runtime.js";
import { DocumentError } from "../errors.js";
import { resolveWorkbookCellFormat } from "../workbook/cell-format.js";
import { mapWorkbookRanges } from "../workbook/range-mapping.js";
import { mapWorkbookCommentAnchors } from "./workbook-comment-mapping.js";
import { checkRange } from "../workbook/ranges.js";
import { rangeContains, rangesIntersect, WorkbookCellEdit } from "./workbook-cells.js";
import { createWorkbookReferenceRewriter, rewriteWorkbookReferenceSteps } from "./workbook-references.js";
import { planWorkbookRuleMove } from "./workbook-rule-move.js";
import { planWorkbookDefinitionMove } from "./workbook-move-definitions.js";

type Move = Extract<WorkbookOperation, { op: "moveWorkbookRange" }>;
function invalid(message: string): never { throw new DocumentError("InvalidOperation", message); }
const contains = (a: WorkbookRange, b: WorkbookRange) => rangeContains(a, b.from.row, b.from.column) && rangeContains(a, b.to.row, b.to.column);
/** Captures old source before clearing either side. Formula rewrites see the
 * old containing sheet and final destination; overlapping cuts never consume
 * cells produced earlier in the same operation. */
export function* moveWorkbookRange(workbook: WorkbookMetadata, cells: WorkbookCellEdit, annotations: Annotation[], operation: Move,
  mappings: WorkbookStructureMapping[]): Generator<{ completed: number; total: number }, void> {
  const source = workbook.sheets.find(sheet => sheet.sheetId === operation.sheetId) ?? invalid("The cut source worksheet does not exist."),
    target = workbook.sheets.find(sheet => sheet.sheetId === operation.targetSheetId) ?? invalid("The cut destination worksheet does not exist.");
  const range = operation.range; checkRange(range);
  const rows = range.to.row - range.from.row + 1, columns = range.to.column - range.from.column + 1;
  const destination = { from: { ...operation.target }, to: { row: operation.target.row + rows - 1, column: operation.target.column + columns - 1 } }; checkRange(destination);
  if (source === target && range.from.row === destination.from.row && range.from.column === destination.from.column) return;
  cells.checkExpandedCells(rows * columns);
  const moveDefinitions = planWorkbookDefinitionMove(source, target, range, destination);
  const covered = (sheetId: string, value: WorkbookRange) => sheetId === source.sheetId && rangesIntersect(value, range) || sheetId === target.sheetId && rangesIntersect(value, destination);
  for (const sheet of workbook.sheets) {
    for (const merge of sheet.merges) if (sheet === source && rangesIntersect(merge, range) && !contains(range, merge)
      || sheet === target && rangesIntersect(merge, destination) && !contains(destination, merge)) invalid("A cut cannot split a merged region.");
  }
  for (const object of workbook.objects ?? []) if (object.anchor.kind !== "absolute" && covered(object.sheetId, { from: object.anchor.from, to: object.anchor.from })) invalid("Cutting cell-bound drawings is not yet connected.");
  for (const binding of workbook.comments ?? []) if (binding.format === "classic" && [binding.box, binding.commentPrBox].some(box => box && box.anchor.kind !== "absolute")) {
    const anchor = annotations.find(annotation => annotation.annotationId === binding.annotationId)?.anchor;
    if (anchor?.kind === "cells" && anchor.ranges.some(value => covered(anchor.sheetId, value))) invalid("Cutting positioned native notes is not yet connected.");
  }
  let removedDestination = 0;
  for (const cell of cells.entries(target.sheetId, false, [destination])) {
    if (source === target && rangeContains(range, cell.row, cell.column)) continue;
    removedDestination++;
    if (operation.overwrittenContent !== "clear" && (cell.formula || cell.value.type !== "blank")) throw new DocumentError("InvalidOperation", "The cut destination contains data; choose explicit clearing or cancel.",
      [{ code: "move-covered-content", path: [target.sheetId, cell.row, cell.column], message: "Overwriting destination content requires an explicit clearing choice." }]);
  }
  const delta = { row: destination.from.row - range.from.row, column: destination.from.column - range.from.column },
    step: WorkbookStructureMapping = { kind: "move", source: { sheetId: source.sheetId, range }, target: { sheetId: target.sheetId, range: destination } },
    change = { kind: "move" as const, sheet: source.name, range, target: { sheet: target.name, ...destination.from } };
  const links = workbook.sheets.flatMap(sheet => (sheet.hyperlinks ?? []).map(link => ({ sheet, link: { ...link, range: { from: { ...link.range.from }, to: { ...link.range.to } } } }))), rewrite = createWorkbookReferenceRewriter(workbook, change);
  const rules = planWorkbookRuleMove(workbook, operation, destination, rewrite);
  const merges = source.merges.filter(merge => rangesIntersect(merge, range)).map(merge => ({ from: { row: merge.from.row + delta.row, column: merge.from.column + delta.column }, to: { row: merge.to.row + delta.row, column: merge.to.column + delta.column } }));
  yield* rewriteWorkbookReferenceSteps(workbook, cells, change);
  for (const [sheet, lists] of rules) {
    if (lists.conditionalRules) sheet.conditionalRules = lists.conditionalRules;
    if (lists.dataValidations) sheet.dataValidations = lists.dataValidations;
  }
  const moving: { cell?: WorkbookCell; row: number; column: number; cssClass: string; formatId?: string }[] = [];
  let completed = 0, additional = 0;
  for (let row = range.from.row; row <= range.to.row; row++) for (let column = range.from.column; column <= range.to.column; column++) {
    const cell = cells.get(source.sheetId, row, column), at = { row: row + delta.row, column: column + delta.column }, format = resolveWorkbookCellFormat(source, { row, column }, cell), base = resolveWorkbookCellFormat(target, at);
    if (cell || format.cssClass !== base.cssClass || format.numberFormatId !== base.numberFormatId) {
      if (!cell) additional++;
      moving.push({ ...(cell ? { cell } : {}), ...at, cssClass: format.cssClass, ...(format.numberFormatId ? { formatId: format.numberFormatId } : {}) });
    }
    if (++completed % 256 === 0) { cells.checkAdditionalCells(additional - removedDestination); yield { completed, total: rows * columns }; }
  }
  cells.checkAdditionalCells(additional - removedDestination);
  yield* cells.clearSteps(source.sheetId, [range], "all"); yield* cells.clearSteps(target.sheetId, [destination], "all");
  let general = workbook.numberFormats.find(format => "builtinId" in format ? format.builtinId === 0 : format.code.toUpperCase() === "GENERAL")?.formatId;
  completed = 0;
  for (const item of moving) {
    if (item.cell) cells.relocate(item.cell, item.row, item.column, target.sheetId);
    const base = resolveWorkbookCellFormat(target, item);
    if (!item.formatId && base.numberFormatId && !general) {
      const ids = new Set(workbook.numberFormats.map(format => format.formatId)); let serial = 0;
      do { general = "move-general-" + ++serial; } while (ids.has(general)); workbook.numberFormats.push({ formatId: general, builtinId: 0 });
    }
    cells.format(target.sheetId, { row: item.row, column: item.column, cssClass: item.cell?.cssClass === undefined && base.cssClass === item.cssClass ? null : item.cssClass, formatId: item.formatId ?? (base.numberFormatId ? general! : null) });
    if (++completed % 256 === 0) yield { completed, total: moving.length };
  }
  for (const sheet of workbook.sheets) {
    if (sheet === source) sheet.merges = sheet.merges.filter(merge => !rangesIntersect(merge, range));
    if (sheet === target) sheet.merges = sheet.merges.filter(merge => !rangesIntersect(merge, destination));
    if (sheet.hyperlinks) sheet.hyperlinks = [];
  }
  target.merges.push(...merges);
  moveDefinitions();
  for (const { sheet, link } of links) for (const fragment of mapWorkbookRanges([{ sheetId: sheet.sheetId, range: link.range }], [step])) {
    const owner = workbook.sheets.find(sheet => sheet.sheetId === fragment.sheetId)!;
    (owner.hyperlinks ??= []).push({ ...link, range: fragment.range,
      ...(link.location ? { location: rewrite("=" + link.location, sheet.name, { row: 0, column: 0 }, undefined, { sheet: owner.name, row: 0, column: 0 }).slice(1) } : {}) });
  }
  yield* mapWorkbookCommentAnchors(workbook, annotations, step);
  mappings.push(step); yield { completed, total: moving.length };
}
