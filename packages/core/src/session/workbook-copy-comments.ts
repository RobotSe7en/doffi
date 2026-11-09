import type { Annotation } from "../contracts/annotations.js";
import type { WorkbookCoordinate, WorkbookMetadata, WorkbookNoteBox, WorkbookRange } from "../contracts/workbook.js";
import { WorkbookCoordinateSchema } from "../contracts/workbook.js";
import type { WorkbookOperation } from "../contracts/workbook-operations.js";
import type { DeepReadonly, WorkbookModel } from "../runtime.js";
import { DocumentError } from "../errors.js";
import { subtractWorkbookRange } from "../workbook/range-mapping.js";
import { checkNonOverlappingRanges, checkRange } from "../workbook/ranges.js";
import { WorkbookDependencyIndex } from "../workbook/dependency-index.js";
import { createWorkbookCommentCloner } from "./workbook-comment-copy.js";
import { removeWorkbookCommentSources, reanchorWorkbookCommentSources } from "./workbook-comments.js";
import type { WorkbookMeasuredLayout } from "./workbook-layout.js";
import { WorkbookContentEdit } from "./workbook-content.js";
import type { SourceStyleEdit } from "./source-styles.js";
import { translateWorkbookAnchor } from "./workbook-anchor-translation.js";

type Copy = Extract<WorkbookOperation, { op: "copyWorkbookComments" }>;
const invalid = (message: string): never => { throw new DocumentError("InvalidOperation", message); };
const overlap = (a: WorkbookRange, b: WorkbookRange): boolean => a.from.row <= b.to.row && b.from.row <= a.to.row && a.from.column <= b.to.column && b.from.column <= a.to.column;
const intersection = (a: WorkbookRange, b: WorkbookRange): WorkbookRange | undefined => overlap(a, b) ? {
  from: { row: Math.max(a.from.row, b.from.row), column: Math.max(a.from.column, b.from.column) },
  to: { row: Math.min(a.to.row, b.to.row), column: Math.min(a.to.column, b.to.column) },
} : undefined;
export function workbookCommentCopyPlan(operation: Copy) {
  checkRange(operation.range);
  const delta = { row: operation.target.row - operation.range.from.row, column: operation.target.column - operation.range.from.column };
  const translate = (range: WorkbookRange, sign = 1): WorkbookRange => ({ from: { row: range.from.row + delta.row * sign, column: range.from.column + delta.column * sign }, to: { row: range.to.row + delta.row * sign, column: range.to.column + delta.column * sign } });
  const target = translate(operation.range);
  if (!WorkbookCoordinateSchema.safeParse(target.to).success) invalid("Comment copying exceeds the destination worksheet bounds.");
  const footprint = operation.ranges ?? [target];
  for (const range of footprint) {
    checkRange(range);
    if (range.from.row < target.from.row || range.to.row > target.to.row || range.from.column < target.from.column || range.to.column > target.to.column)
      invalid("The comment-copy footprint must lie inside its destination rectangle.");
  }
  checkNonOverlappingRanges(footprint);
  const source = footprint.map(range => translate(range, -1));
  let work = 0;
  const step = (): void => { if (++work > 10_000_000) throw new DocumentError("ResourceLimitExceeded", "Comment copying exceeds its range-work budget."); };
  const index = (ranges: readonly WorkbookRange[], sheetId: string) => new WorkbookDependencyIndex(ranges.map((range, index) => ({ area: { ...range, sheetId }, owner: String(index) })), step);
  const sourceIndex = index(source, operation.sheetId), targetIndex = index(footprint, operation.targetSheetId);
  const query = (range: DeepReadonly<WorkbookRange>, ranges: readonly WorkbookRange[]) => (ranges === source ? sourceIndex : targetIndex).query({ ...range, sheetId: ranges === source ? operation.sheetId : operation.targetSheetId }, step);
  const affected = (annotation: DeepReadonly<Annotation>, sheetId: string, ranges: readonly WorkbookRange[]): boolean => annotation.anchor.kind === "cells" && annotation.anchor.sheetId === sheetId && annotation.anchor.ranges.some(range => query(range, ranges).next().done === false);
  const clipped = (ranges: readonly DeepReadonly<WorkbookRange>[]): WorkbookRange[] => {
    const result: WorkbookRange[] = [];
    for (const range of ranges) for (const key of query(range, source)) {
      step(); result.push(translate(intersection(range, source[Number(key)]!)!));
      if (result.length > 4096) throw new DocumentError("ResourceLimitExceeded", "Comment copying exceeds its anchor-fragment budget.");
    }
    return result;
  };
  const remainder = (ranges: readonly WorkbookRange[]): WorkbookRange[] => {
    const result: WorkbookRange[] = [];
    for (const range of ranges) {
      let pieces = [range];
      for (const key of query(range, footprint)) {
        const mask = footprint[Number(key)]!; pieces = pieces.flatMap(piece => { step(); return subtractWorkbookRange(piece, mask); });
        if (pieces.length + result.length > 4096) throw new DocumentError("ResourceLimitExceeded", "Comment replacement exceeds its anchor-fragment budget.");
      }
      result.push(...pieces);
      if (result.length > 4096) throw new DocumentError("ResourceLimitExceeded", "Comment replacement exceeds its anchor-fragment budget.");
    }
    return result;
  };
  return { delta, footprint, source, translate, affected, clipped, remainder, unchanged: operation.sheetId === operation.targetSheetId && delta.row === 0 && delta.column === 0 };
}
export function workbookCommentCopySheets(workbook: WorkbookMetadata, annotations: readonly Annotation[], operation: Copy): readonly string[] {
  const plan = workbookCommentCopyPlan(operation);
  if (plan.unchanged || !plan.footprint.length) return [];
  const source = new Set(annotations.filter(annotation => plan.affected(annotation, operation.sheetId, plan.source)).map(annotation => annotation.annotationId));
  return workbook.comments?.some(binding => source.has(binding.annotationId) && binding.format === "classic" && (binding.box || binding.commentPrBox))
    ? [...new Set([operation.sheetId, operation.targetSheetId])] : [];
}
/** All source clones exist before destination removal, so overlapping copies
 * never consume bodies or anchors produced earlier in this same operation. */
export function* copyWorkbookComments(workbook: WorkbookMetadata, annotations: Annotation[], content: WorkbookContentEdit, base: WorkbookModel, operation: Copy,
  allocate: () => string, styles: SourceStyleEdit, layout?: WorkbookMeasuredLayout): Generator<{ completed: number; total: number }, void> {
  if (![operation.sheetId, operation.targetSheetId].every(id => workbook.sheets.some(sheet => sheet.sheetId === id))) invalid("Comment copying requires existing worksheets.");
  const plan = workbookCommentCopyPlan(operation);
  if (plan.unchanged || !plan.footprint.length) return;
  const clone = createWorkbookCommentCloner(workbook, content, allocate), identities = new Map<string, string>(), added: Annotation[] = [];
  let completed = 0;
  for (const original of annotations) {
    if (original.anchor.kind === "cells" && plan.affected(original, operation.sheetId, plan.source)) {
      const ranges = plan.clipped(original.anchor.ranges);
      const point = plan.translate({ from: ranges[0]!.from, to: ranges[0]!.from }, -1).from;
      const coordinate = (at: WorkbookCoordinate): WorkbookCoordinate => {
        const next = { row: at.row + plan.delta.row, column: at.column + plan.delta.column };
        if (!WorkbookCoordinateSchema.safeParse(next).success) invalid("Copying this retained native thread location exceeds worksheet bounds."); return next;
      };
      const box = (value: WorkbookNoteBox): WorkbookNoteBox => {
        const measured = layout ?? invalid("Copying positioned notes requires measured worksheet layouts.");
        return { ...value, anchor: translateWorkbookAnchor(value.anchor, { sheetId: operation.sheetId, at: point, layout: measured },
          { sheetId: operation.targetSheetId, at: coordinate(point), layout: measured }) };
      };
      const value = clone(original, { kind: "cells", sheetId: operation.targetSheetId, ranges }, base.snapshot().manifest.revision, { box, coordinate });
      added.push(value.annotation); for (const [from, to] of value.identities) identities.set(from, to);
    }
    if (++completed % 64 === 0) yield { completed, total: annotations.length };
  }
  for (let index = annotations.length - 1; index >= 0; index--) {
    const annotation = annotations[index]!;
    if (annotation.anchor.kind === "cells" && plan.affected(annotation, operation.targetSheetId, plan.footprint)) {
      const ranges = plan.remainder(annotation.anchor.ranges);
      if (!ranges.length) { removeWorkbookCommentSources(workbook, content, annotation); annotations.splice(index, 1); }
      else { annotation.anchor.ranges = ranges; if (annotation.status !== "needs_reanchor") reanchorWorkbookCommentSources(workbook, annotation); }
    }
    if (++completed % 64 === 0) yield { completed, total: 0 };
  }
  annotations.push(...added); styles.aliasNodeIds(identities); yield { completed, total: completed };
}
