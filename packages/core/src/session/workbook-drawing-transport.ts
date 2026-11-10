import type { WorkbookCoordinate, WorkbookDrawingAnchor, WorkbookMetadata, WorkbookObject } from "../contracts/workbook.js";
import type { DocumentOperation } from "../contracts/operations.js";
import { DocumentError } from "../errors.js";
import { rangeContains } from "./workbook-cells.js";
import { WorkbookAnchorTransportTracker, type WorkbookAnchorTransport } from "./workbook-anchor-translation.js";
import type { WorkbookMeasuredLayout } from "./workbook-layout.js";
import { checkRange } from "../workbook/ranges.js";

interface Move {
  readonly object: WorkbookObject; readonly sheetId: string; readonly at: WorkbookCoordinate; readonly anchor: WorkbookDrawingAnchor;
  readonly targetSheetId: string; readonly target: WorkbookCoordinate; transport?: WorkbookAnchorTransport;
}
export interface WorkbookDrawingTransport {
  readonly sheetIds: readonly string[];
  readonly destinations: ReadonlyMap<string, string>;
  project(before: WorkbookMeasuredLayout): Generator<{ completed: number; total: number }, void>;
  apply(after: WorkbookMeasuredLayout): Generator<{ completed: number; total: number }, ReadonlySet<string>>;
}
/** Geometry follows a moving anchor cell. Absolute/nonmoving drawings remain
 * independent of cell replacement; all source object/paragraph IDs survive. */
export class WorkbookDrawingTransportTracker {
  private readonly anchors = new WorkbookAnchorTransportTracker();
  *capture(workbook: WorkbookMetadata, operation: DocumentOperation): Generator<{ completed: number; total: number }, WorkbookDrawingTransport | undefined> {
    if (operation.op !== "moveWorkbookRange" && operation.op !== "reorderWorkbookRows") return;
    if (!workbook.sheets.some(sheet => sheet.sheetId === operation.sheetId)
      || operation.op === "moveWorkbookRange" && !workbook.sheets.some(sheet => sheet.sheetId === operation.targetSheetId))
      throw new DocumentError("InvalidOperation", "Drawing transport requires existing source and target worksheets.");
    checkRange(operation.range);
    if (operation.op === "moveWorkbookRange") checkRange({ from: operation.target, to: {
      row: operation.target.row + operation.range.to.row - operation.range.from.row,
      column: operation.target.column + operation.range.to.column - operation.range.from.column,
    } });
    const moves: Move[] = [], anchors = this.anchors; let completed = 0;
    for (const object of workbook.objects ?? []) {
      const anchor = object.anchor;
      if (object.sheetId === operation.sheetId && anchor.kind !== "absolute" && (anchor.kind === "one-cell" || anchor.moveWithCells)
        && rangeContains(operation.range, anchor.from.row, anchor.from.column)) {
        const at = { row: anchor.from.row, column: anchor.from.column }, targetSheetId = operation.op === "moveWorkbookRange" ? operation.targetSheetId : object.sheetId;
        const target = operation.op === "moveWorkbookRange" ? { row: at.row + operation.target.row - operation.range.from.row, column: at.column + operation.target.column - operation.range.from.column }
          : { row: operation.rows[at.row - operation.range.from.row] ?? NaN, column: at.column };
        if (!Number.isInteger(target.row) || !Number.isInteger(target.column)) throw new DocumentError("InvalidOperation", "Drawing transport requires a complete cell mapping.");
        if (targetSheetId !== object.sheetId || target.row !== at.row || target.column !== at.column) moves.push({ object, sheetId: object.sheetId, at,
          anchor: JSON.parse(JSON.stringify(anchor)) as WorkbookDrawingAnchor, targetSheetId, target });
      }
      if (++completed % 256 === 0) yield { completed, total: 0 };
    }
    if (!moves.length) return;
    return {
      sheetIds: [...new Set(moves.flatMap(move => [move.sheetId, move.targetSheetId]))],
      destinations: new Map(moves.map(move => [move.object.nodeId, move.targetSheetId])),
      *project(before) {
        let completed = 0;
        for (const move of moves) {
          move.transport = anchors.capture(JSON.stringify(["object", move.object.nodeId]), move.anchor, { sheetId: move.sheetId, at: move.at, layout: before });
          move.object.sheetId = move.targetSheetId;
          move.object.anchor = move.transport.project({ sheetId: move.targetSheetId, at: move.target, layout: before });
          if (++completed % 256 === 0) yield { completed, total: moves.length };
        }
      },
      *apply(after) {
        const explicit = new Set<string>(); let completed = 0;
        for (const move of moves) {
          if (!move.transport) throw new DocumentError("InvalidOperation", "Drawing transport requires its original measured projection.");
          move.object.anchor = move.transport.apply({ sheetId: move.targetSheetId, at: move.target, layout: after }); explicit.add(JSON.stringify(["object", move.object.nodeId]));
          if (++completed % 256 === 0) yield { completed, total: moves.length };
        }
        return explicit;
      },
    };
  }
}
