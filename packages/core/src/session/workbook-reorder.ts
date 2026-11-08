import type { Annotation } from "../contracts/annotations.js";
import type { WorkbookMetadata, WorkbookRange } from "../contracts/workbook.js";
import type { WorkbookOperation } from "../contracts/workbook-operations.js";
import type { WorkbookCell, WorkbookStructureMapping } from "../runtime.js";
import { DocumentError } from "../errors.js";
import { copyFormula } from "../workbook/formula-rewrite.js";
import { resolveWorkbookCellFormat } from "../workbook/cell-format.js";
import { mapWorkbookRanges } from "../workbook/range-mapping.js";
import { checkRange } from "../workbook/ranges.js";
import { mapWorkbookCommentAnchors } from "./workbook-comment-mapping.js";
import { rangeContains, rangesIntersect, WorkbookCellEdit } from "./workbook-cells.js";

type Reorder = Extract<WorkbookOperation, { op: "reorderWorkbookRows" }>;
function invalid(message: string): never { throw new DocumentError("InvalidOperation", message); }
const contains = (a: WorkbookRange, b: WorkbookRange) => rangeContains(a, b.from.row, b.from.column) && rangeContains(a, b.to.row, b.to.column);

/** Explicit old-row -> new-row permutation. Human sorting computes this from a
 * revision-bound current calculation; proposals use the same typed operation.
 * Only formulas travelling with records are copied. Outside references keep
 * their original addresses, including absolute references inside the region. */
export function* reorderWorkbookRows(workbook: WorkbookMetadata, cells: WorkbookCellEdit, annotations: Annotation[], operation: Reorder,
  mappings: WorkbookStructureMapping[]): Generator<{ completed: number; total: number }, void> {
  const sheet = workbook.sheets.find(sheet => sheet.sheetId === operation.sheetId) ?? invalid("An earlier operation removed the worksheet.");
  const { range, rows } = operation; checkRange(range);
  const height = range.to.row - range.from.row + 1, width = range.to.column - range.from.column + 1, area = height * width;
  cells.checkExpandedCells(area);
  if (rows.length !== height || new Set(rows).size !== height || rows.some(row => row < range.from.row || row > range.to.row)) invalid("Row reordering requires a complete permutation of the selected rows.");
  if (rows.every((row, offset) => row === range.from.row + offset)) return;
  if (sheet.autoFilter && rangesIntersect(sheet.autoFilter.range, range) && range.from.row <= sheet.autoFilter.range.from.row)
    invalid("Row sorting cannot move an auto-filter header.");
  for (const table of sheet.tables ?? []) if (rangesIntersect(table.range, range)) {
    const body = { from: { ...table.range.from, row: table.range.from.row + Number(table.header) }, to: { ...table.range.to, row: table.range.to.row - Number(table.totals) } };
    if (!contains(body, range) || range.from.column !== table.range.from.column || range.to.column !== table.range.to.column)
      invalid("Table row sorting must move complete records within the data body.");
  }
  const merges = sheet.merges.filter(merge => rangesIntersect(merge, range));
  if (merges.length) {
    const shapes = new Map<number, string[]>();
    for (const merge of merges) {
      if (!contains(range, merge) || merge.from.row !== merge.to.row) invalid("Row sorting cannot split a merged region.");
      const shape = shapes.get(merge.from.row) ?? []; shape.push(merge.from.column + ":" + merge.to.column); shapes.set(merge.from.row, shape);
    }
    const first = JSON.stringify((shapes.get(range.from.row) ?? []).sort());
    for (let row = range.from.row; row <= range.to.row; row++) if (JSON.stringify((shapes.get(row) ?? []).sort()) !== first) invalid("Sorted records must have the same merge topology.");
  }
  for (const object of workbook.objects ?? []) if (object.sheetId === sheet.sheetId && object.anchor.kind !== "absolute" && rangeContains(range, object.anchor.from.row, object.anchor.from.column))
    invalid("Sorting cell-bound drawings is not yet connected.");
  for (const comment of workbook.comments ?? []) if (comment.format === "classic" && [comment.box, comment.commentPrBox].some(box => box && box.anchor.kind !== "absolute")) {
    const owner = annotations.find(annotation => annotation.annotationId === comment.annotationId)?.anchor;
    if (owner?.kind === "cells" && owner.sheetId === sheet.sheetId && owner.ranges.some(value => rangesIntersect(value, range))) invalid("Sorting positioned native note boxes is not yet connected.");
  }
  const step: WorkbookStructureMapping = { kind: "sort", sheetId: sheet.sheetId, range, rows };
  const moving: { cell?: WorkbookCell; row: number; column: number; cssClass: string; formatId?: string }[] = [];
  let completed = 0, additional = 0;
  for (let row = range.from.row; row <= range.to.row; row++) for (let column = range.from.column; column <= range.to.column; column++) {
    const target = rows[row - range.from.row]!, raw = cells.get(sheet.sheetId, row, column), sourceFormat = resolveWorkbookCellFormat(sheet, { row, column }, raw);
    const targetFormat = resolveWorkbookCellFormat(sheet, { row: target, column });
    if (raw || sourceFormat.cssClass !== targetFormat.cssClass || sourceFormat.numberFormatId !== targetFormat.numberFormatId) {
      let cell = raw;
      if (cell?.formula) {
        cells.rewriteFormula(sheet.sheetId, row, column, copyFormula(cell.formula.expression, { row, column }, { row: target, column }));
        cell = cells.get(sheet.sheetId, row, column)!;
      }
      if (!raw) additional++;
      moving.push({ ...(cell ? { cell } : {}), row: target, column, cssClass: sourceFormat.cssClass, ...(sourceFormat.numberFormatId ? { formatId: sourceFormat.numberFormatId } : {}) });
    }
    if (++completed % 256 === 0) { cells.checkAdditionalCells(additional); yield { completed, total: area }; }
  }
  cells.checkAdditionalCells(additional);
  yield* cells.clearSteps(sheet.sheetId, [range], "all");
  let general = workbook.numberFormats.find(format => "builtinId" in format ? format.builtinId === 0 : format.code.toUpperCase() === "GENERAL")?.formatId;
  completed = 0;
  for (const item of moving) {
    if (item.cell) cells.relocate(item.cell, item.row, item.column);
    const targetFormat = resolveWorkbookCellFormat(sheet, item);
    if (!item.formatId && targetFormat.numberFormatId && !general) {
      const ids = new Set(workbook.numberFormats.map(format => format.formatId)); let serial = 0;
      do { general = "sort-general-" + ++serial; } while (ids.has(general)); workbook.numberFormats.push({ formatId: general, builtinId: 0 });
    }
    cells.format(sheet.sheetId, { row: item.row, column: item.column, cssClass: item.cell?.cssClass === undefined && targetFormat.cssClass === item.cssClass ? null : item.cssClass,
      formatId: item.formatId ?? (targetFormat.numberFormatId ? general! : null) });
    if (++completed % 256 === 0) yield { completed, total: moving.length };
  }
  for (const merge of merges) merge.from.row = merge.to.row = rows[merge.from.row - range.from.row]!;
  if (sheet.hyperlinks) sheet.hyperlinks = sheet.hyperlinks.flatMap(link => mapWorkbookRanges([{ sheetId: sheet.sheetId, range: link.range }], [step]).map(value => ({ ...link, range: value.range })));
  for (const rule of [...sheet.conditionalRules ?? [], ...sheet.dataValidations ?? []]) {
    rule.ranges = mapWorkbookRanges(rule.ranges.map(range => ({ sheetId: sheet.sheetId, range })), [step]).map(value => ({ from: { ...value.range.from }, to: { ...value.range.to } }));
    if (rangeContains(range, rule.origin.row, rule.origin.column)) {
      const origin = { ...rule.origin, row: rows[rule.origin.row - range.from.row]! }, copy = (expression: string) => copyFormula(expression, rule.origin, origin);
      if ("expression" in rule) rule.expression = copy(rule.expression);
      if ("formulas" in rule) rule.formulas = rule.formulas.map(copy);
      if (rule.kind === "colorScale") for (const stop of rule.stops) if (stop.kind === "formula") stop.expression = copy(stop.expression);
      rule.origin = origin;
    }
  }
  yield* mapWorkbookCommentAnchors(workbook, annotations, step);
  mappings.push(step); yield { completed, total: moving.length };
}
