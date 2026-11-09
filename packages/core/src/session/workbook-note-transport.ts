import type { Annotation } from "../contracts/annotations.js";
import type { WorkbookCoordinate, WorkbookDrawingAnchor, WorkbookMetadata, WorkbookNoteBox, WorkbookRange } from "../contracts/workbook.js";
import type { DocumentOperation } from "../contracts/operations.js";
import type { WorkbookMeasuredLayout } from "./workbook-layout.js";
import { translateWorkbookAnchor } from "./workbook-anchor-translation.js";
import { rangesIntersect } from "./workbook-cells.js";
import { DocumentError } from "../errors.js";
import type { WorkbookStructureMapping } from "../runtime.js";
import { mapWorkbookRanges } from "../workbook/range-mapping.js";

interface CapturedNote {
  readonly annotationId: string; readonly sheetId: string; readonly at: WorkbookCoordinate;
  readonly boxes: ReadonlyMap<"box" | "commentPrBox", WorkbookNoteBox>;
}
interface Origin {
  readonly sheetId: string; readonly at: WorkbookCoordinate; readonly anchor: WorkbookDrawingAnchor; readonly layout: WorkbookMeasuredLayout;
  readonly lastSheetId: string; readonly lastAt: WorkbookCoordinate; readonly lastAnchor: string; readonly lastGeometry: string;
}
type History = Map<string, Map<"box" | "commentPrBox", Origin>>;
function geometry(anchor: WorkbookDrawingAnchor, sheetId: string, at: WorkbookCoordinate, layout: WorkbookMeasuredLayout): string {
  const axes = layout.sheets.get(sheetId); if (!axes) throw new DocumentError("InvalidOperation", "Native note geometry requires its measured worksheet.");
  const ox = axes.columns.position(at.column), oy = axes.rows.position(at.row);
  if (anchor.kind === "absolute") return JSON.stringify([anchor.xPt - ox, anchor.yPt - oy, anchor.widthPt, anchor.heightPt]);
  const x = axes.columns.position(anchor.from.column) - ox + anchor.from.offsetXPt, y = axes.rows.position(anchor.from.row) - oy + anchor.from.offsetYPt;
  return JSON.stringify([x, y, anchor.kind === "one-cell" ? anchor.widthPt : axes.columns.position(anchor.to.column) - ox + anchor.to.offsetXPt - x,
    anchor.kind === "one-cell" ? anchor.heightPt : axes.rows.position(anchor.to.row) - oy + anchor.to.offsetYPt - y]);
}
/** Private to one preparation. Carry original marker identities only while
 * the actual current anchor and measured relative geometry still prove the
 * previous transport. Explicit geometry/grid changes establish a new origin. */
export class WorkbookNoteTransportTracker {
  private readonly history: History = new Map();
  capture(workbook: WorkbookMetadata, annotations: Annotation[], operation: DocumentOperation): Generator<{ completed: number; total: number }, WorkbookNoteTransport | undefined> {
    return captureWorkbookNoteTransport(workbook, annotations, operation, this.history);
  }
}
export interface WorkbookNoteTransport {
  readonly sheetIds: readonly string[];
  apply(before: WorkbookMeasuredLayout, after: WorkbookMeasuredLayout): Generator<{ completed: number; total: number }, ReadonlySet<string>>;
}
/** Capture before structural anchor mapping. Recovery-only notes retain their
 * old boxes; a valid moved representative transports both native sources. */
function* captureWorkbookNoteTransport(workbook: WorkbookMetadata, annotations: Annotation[], operation: DocumentOperation, history: History): Generator<{ completed: number; total: number }, WorkbookNoteTransport | undefined> {
  if (operation.op !== "moveWorkbookRange" && operation.op !== "reorderWorkbookRows") return;
  if (operation.op === "moveWorkbookRange" && operation.sheetId === operation.targetSheetId && operation.range.from.row === operation.target.row && operation.range.from.column === operation.target.column
    || operation.op === "reorderWorkbookRows" && operation.rows.every((row, index) => row === operation.range.from.row + index)) return;
  const affected = operation.op === "moveWorkbookRange" ? new Set([operation.sheetId, operation.targetSheetId]) : new Set([operation.sheetId]);
  const ranges = [{ sheetId: operation.sheetId, range: operation.range }];
  if (operation.op === "moveWorkbookRange") ranges.push({ sheetId: operation.targetSheetId, range: { from: operation.target, to: {
    row: operation.target.row + operation.range.to.row - operation.range.from.row, column: operation.target.column + operation.range.to.column - operation.range.from.column } } });
  const step: WorkbookStructureMapping = operation.op === "moveWorkbookRange"
    ? { kind: "move", source: ranges[0]!, target: ranges[1]! }
    : { kind: "sort", sheetId: operation.sheetId, range: operation.range, rows: operation.rows };
  const covered = (sheetId: string, values: WorkbookRange[]) => ranges.some(entry => entry.sheetId === sheetId && values.some(value => rangesIntersect(entry.range, value)));
  const entries = new Map(annotations.map(annotation => [annotation.annotationId, annotation])), notes: CapturedNote[] = [];
  let completed = 0;
  for (const binding of workbook.comments ?? []) {
    const annotation = entries.get(binding.annotationId);
    if (binding.format === "classic" && annotation?.status !== "needs_reanchor" && annotation?.anchor.kind === "cells" && covered(annotation.anchor.sheetId, annotation.anchor.ranges)) {
      const anchor = annotation.anchor, at = anchor.ranges[0]!.from, mapped = mapWorkbookRanges(anchor.ranges.map(range => ({ sheetId: anchor.sheetId, range })), [step]);
      // Lost or mixed-sheet targets retain recovery geometry. They require no
      // transport measurements; only a valid changed representative moves.
      if (mapped.length && mapped.every(value => value.sheetId === mapped[0]!.sheetId)
        && (mapped[0]!.sheetId !== anchor.sheetId || mapped[0]!.range.from.row !== at.row || mapped[0]!.range.from.column !== at.column)) {
        const boxes = new Map<"box" | "commentPrBox", WorkbookNoteBox>();
        for (const key of ["box", "commentPrBox"] as const) if (binding[key]) boxes.set(key, JSON.parse(JSON.stringify(binding[key])) as WorkbookNoteBox);
        if (boxes.size) notes.push({ annotationId: binding.annotationId, sheetId: anchor.sheetId, at: { ...at }, boxes });
      }
    }
    if (++completed % 256 === 0) yield { completed, total: 0 };
  }
  if (!notes.length) return;
  return {
    sheetIds: [...affected],
    *apply(before, after) {
      const explicit = new Set<string>(), bindings = new Map(workbook.comments?.map(binding => [binding.annotationId, binding])); let completed = 0;
      for (const note of notes) {
        const annotation = entries.get(note.annotationId), binding = bindings.get(note.annotationId);
        if (annotation?.status !== "needs_reanchor" && annotation?.anchor.kind === "cells" && binding?.format === "classic" && annotation.anchor.ranges.length) {
          const anchor = annotation.anchor, at = anchor.ranges[0]!.from;
          if (note.sheetId !== anchor.sheetId || note.at.row !== at.row || note.at.column !== at.column) {
            const retained = history.get(note.annotationId) ?? new Map<"box" | "commentPrBox", Origin>(); history.set(note.annotationId, retained);
            for (const [key, box] of note.boxes) {
              const previous = retained.get(key), origin = previous && previous.lastSheetId === note.sheetId && previous.lastAt.row === note.at.row && previous.lastAt.column === note.at.column
                && previous.lastAnchor === JSON.stringify(box.anchor) && previous.lastGeometry === geometry(box.anchor, note.sheetId, note.at, before)
                ? previous : { sheetId: note.sheetId, at: note.at, anchor: box.anchor, layout: before };
              const mapped = translateWorkbookAnchor(origin.anchor, { sheetId: origin.sheetId, at: origin.at, layout: origin.layout }, { sheetId: anchor.sheetId, at, layout: after });
              binding[key] = { ...box, anchor: mapped }; explicit.add(JSON.stringify([key, note.annotationId]));
              retained.set(key, { sheetId: origin.sheetId, at: origin.at, anchor: origin.anchor, layout: origin.layout, lastSheetId: anchor.sheetId,
                lastAt: { ...at }, lastAnchor: JSON.stringify(mapped), lastGeometry: geometry(mapped, anchor.sheetId, at, after) });
            }
          }
        }
        if (++completed % 64 === 0) yield { completed, total: notes.length };
      }
      return explicit;
    },
  };
}
