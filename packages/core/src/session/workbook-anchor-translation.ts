import type { WorkbookCoordinate, WorkbookDrawingAnchor } from "../contracts/workbook.js";
import { DocumentError } from "../errors.js";
import { MAX_SHEET_COLUMNS, MAX_SHEET_ROWS } from "../workbook/chunk-index.js";
import type { WorkbookMeasuredLayout } from "./workbook-layout.js";

/** Explicitly transporting a cell-owned object retains its physical extents
 * and displacement from that cell. Grid response flags remain source metadata
 * for subsequent dimension edits; both native marker identities stay bounded. */
export function translateWorkbookAnchor(anchor: WorkbookDrawingAnchor, source: { sheetId: string; at: WorkbookCoordinate; layout: WorkbookMeasuredLayout },
  target: { sheetId: string; at: WorkbookCoordinate; layout: WorkbookMeasuredLayout }): WorkbookDrawingAnchor {
  const from = source.layout.sheets.get(source.sheetId), to = target.layout.sheets.get(target.sheetId);
  if (!from || !to) throw new DocumentError("InvalidOperation", "Transporting native geometry requires both measured worksheet layouts.");
  const delta = { row: target.at.row - source.at.row, column: target.at.column - source.at.column };
  const marker = (value: Extract<WorkbookDrawingAnchor, { kind: "one-cell" | "two-cell" }>["from"]) => {
    const row = Math.max(0, Math.min(MAX_SHEET_ROWS - 1, value.row + delta.row)), column = Math.max(0, Math.min(MAX_SHEET_COLUMNS - 1, value.column + delta.column));
    // Subtract cell origins first, so distant destination coordinates do not
    // unnecessarily erase small original offsets through cancellation.
    return { row, column,
      offsetXPt: value.offsetXPt + ((from.columns.position(value.column) - from.columns.position(source.at.column)) - (to.columns.position(column) - to.columns.position(target.at.column))),
      offsetYPt: value.offsetYPt + ((from.rows.position(value.row) - from.rows.position(source.at.row)) - (to.rows.position(row) - to.rows.position(target.at.row))) };
  };
  if (anchor.kind === "absolute") return { ...anchor,
    xPt: anchor.xPt + (to.columns.position(target.at.column) - from.columns.position(source.at.column)),
    yPt: anchor.yPt + (to.rows.position(target.at.row) - from.rows.position(source.at.row)) };
  return anchor.kind === "one-cell" ? { ...anchor, from: marker(anchor.from) } : { ...anchor, from: marker(anchor.from), to: marker(anchor.to) };
}
