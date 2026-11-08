import type { Annotation } from "../contracts/annotations.js";
import type { WorkbookDrawingAnchor, WorkbookMetadata } from "../contracts/workbook.js";
import type { WorkbookOperation } from "../contracts/workbook-operations.js";
import { DocumentError } from "../errors.js";
import { MAX_SHEET_COLUMNS, MAX_SHEET_ROWS } from "../workbook/chunk-index.js";

/** One-cell extents are already physical source dimensions. Both moving,
 * resizing markers of a two-cell anchor follow the changed axis directly;
 * fixed/mixed two-cell modes still need measured layout compensation. */
export function* mapWorkbookAxisAnchors(workbook: WorkbookMetadata, annotations: readonly Annotation[], operation: Extract<WorkbookOperation, { op: "editWorkbookAxis" }>): Generator<{ completed: number; total: number }, void> {
  const owners = new Set<string>(), anchors: WorkbookDrawingAnchor[] = []; let completed = 0;
  for (const annotation of annotations) {
    if (annotation.anchor.kind === "cells" && annotation.anchor.sheetId === operation.sheetId) owners.add(annotation.annotationId);
    if (++completed % 256 === 0) yield { completed, total: 0 };
  }
  for (const object of workbook.objects ?? []) {
    if (object.sheetId === operation.sheetId) anchors.push(object.anchor);
    if (++completed % 256 === 0) yield { completed, total: 0 };
  }
  for (const comment of workbook.comments ?? []) {
    if (comment.format === "classic" && owners.has(comment.annotationId)) for (const box of [comment.box, comment.commentPrBox]) if (box) anchors.push(box.anchor);
    if (++completed % 256 === 0) yield { completed, total: 0 };
  }
  if (!anchors.length) return;
  for (const anchor of anchors) {
    if (anchor.kind === "two-cell" && (!anchor.moveWithCells || !anchor.sizeWithCells)) throw new DocumentError("InvalidOperation", "Axis editing of fixed or mixed cell-bound native drawings and note boxes requires their measured layout mapping.");
    if (++completed % 256 === 0) yield { completed, total: 0 };
  }
  const bound = operation.axis === "row" ? MAX_SHEET_ROWS : MAX_SHEET_COLUMNS;
  const point = (value: number): number => {
    if (value < operation.index) return value;
    if (operation.action === "delete" && value < operation.index + operation.count) return operation.index;
    const mapped = value + (operation.action === "insert" ? operation.count : -operation.count);
    if (mapped >= bound) throw new DocumentError("InvalidOperation", "The worksheet edit would discard an object anchor outside its native bounds.");
    return mapped;
  };
  for (const anchor of anchors) {
    if (anchor.kind !== "absolute") { anchor.from[operation.axis] = point(anchor.from[operation.axis]); if (anchor.kind === "two-cell") anchor.to[operation.axis] = point(anchor.to[operation.axis]); }
    if (++completed % 256 === 0) yield { completed, total: 0 };
  }
  yield { completed, total: completed };
}
