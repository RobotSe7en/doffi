import type { WorkbookCoordinate, WorkbookDrawingAnchor } from "../contracts/workbook.js";
import { DocumentError } from "../errors.js";
import { MAX_SHEET_COLUMNS, MAX_SHEET_ROWS } from "../workbook/chunk-index.js";
import type { WorkbookMeasuredLayout } from "./workbook-layout.js";

export interface WorkbookAnchorPlace { readonly sheetId: string; readonly at: WorkbookCoordinate; readonly layout: WorkbookMeasuredLayout }
export interface WorkbookAnchorTransport {
  project(target: WorkbookAnchorPlace): WorkbookDrawingAnchor;
  apply(target: WorkbookAnchorPlace): WorkbookDrawingAnchor;
}
interface Origin {
  readonly sheetId: string; readonly at: WorkbookCoordinate; readonly anchor: WorkbookDrawingAnchor; readonly layout: WorkbookMeasuredLayout;
  readonly lastSheetId: string; readonly lastAt: WorkbookCoordinate; readonly lastAnchor: string; readonly lastGeometry: string;
}
function geometry(anchor: WorkbookDrawingAnchor, place: WorkbookAnchorPlace): string {
  const axes = place.layout.sheets.get(place.sheetId); if (!axes) throw new DocumentError("InvalidOperation", "Native geometry requires its measured worksheet.");
  const ox = axes.columns.position(place.at.column), oy = axes.rows.position(place.at.row);
  if (anchor.kind === "absolute") return JSON.stringify([anchor.xPt - ox, anchor.yPt - oy, anchor.widthPt, anchor.heightPt]);
  const x = axes.columns.position(anchor.from.column) - ox + anchor.from.offsetXPt, y = axes.rows.position(anchor.from.row) - oy + anchor.from.offsetYPt;
  return JSON.stringify([x, y, anchor.kind === "one-cell" ? anchor.widthPt : axes.columns.position(anchor.to.column) - ox + anchor.to.offsetXPt - x,
    anchor.kind === "one-cell" ? anchor.heightPt : axes.rows.position(anchor.to.row) - oy + anchor.to.offsetYPt - y]);
}
/** Private to one preparation. Original marker identities survive only while
 * the actual current anchor and relative geometry prove the previous move.
 * Projection can prepare an intermediate model without advancing that proof. */
export class WorkbookAnchorTransportTracker {
  private readonly history = new Map<string, Origin>();
  capture(key: string, anchor: WorkbookDrawingAnchor, source: WorkbookAnchorPlace): WorkbookAnchorTransport {
    const previous = this.history.get(key), origin = previous && previous.lastSheetId === source.sheetId && previous.lastAt.row === source.at.row && previous.lastAt.column === source.at.column
      && previous.lastAnchor === JSON.stringify(anchor) && previous.lastGeometry === geometry(anchor, source)
      ? previous : { sheetId: source.sheetId, at: { ...source.at }, anchor: JSON.parse(JSON.stringify(anchor)) as WorkbookDrawingAnchor, layout: source.layout };
    const project = (target: WorkbookAnchorPlace) => translateWorkbookAnchor(origin.anchor, origin, target);
    return { project, apply: target => {
      const mapped = project(target); this.history.set(key, { sheetId: origin.sheetId, at: origin.at, anchor: origin.anchor, layout: origin.layout,
        lastSheetId: target.sheetId, lastAt: { ...target.at }, lastAnchor: JSON.stringify(mapped), lastGeometry: geometry(mapped, target) }); return mapped;
    } };
  }
}

/** Explicitly transporting a cell-owned object retains its physical extents
 * and displacement from that cell. Grid response flags remain source metadata
 * for subsequent dimension edits; both native marker identities stay bounded. */
export function translateWorkbookAnchor(anchor: WorkbookDrawingAnchor, source: WorkbookAnchorPlace, target: WorkbookAnchorPlace): WorkbookDrawingAnchor {
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
