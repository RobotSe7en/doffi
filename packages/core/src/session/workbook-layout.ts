import type { Annotation } from "../contracts/annotations.js";
import type { WorkbookDrawingAnchor, WorkbookMetadata } from "../contracts/workbook.js";
import type { DocumentOperation } from "../contracts/operations.js";
import type { WorkbookLayoutMeasurement, WorkbookLayoutMeasurer, WorkbookModel } from "../runtime.js";
import { DocumentError } from "../errors.js";
import { assetDigest } from "../document/assets.js";
import { encode } from "../document/package.js";
import { AxisDimensions } from "../workbook/axis-dimensions.js";
import { MAX_SHEET_COLUMNS, MAX_SHEET_ROWS } from "../workbook/chunk-index.js";
import type { WorkbookPreparation } from "./workbook-preparation.js";

type CellAnchor = Extract<WorkbookDrawingAnchor, { kind: "two-cell" }>;
interface Owner { readonly key: string; readonly sheetId: string; readonly anchor: CellAnchor }
export interface WorkbookMeasuredLayout {
  readonly fontEnvironment: string;
  readonly sheets: ReadonlyMap<string, { readonly rows: AxisDimensions; readonly columns: AxisDimensions }>;
}
const invalid = (message: string): never => { throw new DocumentError("InvalidOperation", message); };

/** Include both classic native box sources, even when VML wins the preview. */
export function* workbookLayoutOwners(workbook: WorkbookMetadata, annotations: readonly Annotation[]): WorkbookPreparation<Owner[]> {
  const owners: Owner[] = [], sheets = new Map<string, string>(); let completed = 0;
  const add = (key: string, sheetId: string, anchor: WorkbookDrawingAnchor): void => {
    if (anchor.kind === "two-cell" && (!anchor.moveWithCells || !anchor.sizeWithCells)) owners.push({ key, sheetId, anchor });
  };
  for (const item of annotations) { if (item.anchor.kind === "cells") sheets.set(item.annotationId, item.anchor.sheetId); if (++completed % 256 === 0) yield { kind: "progress", value: { phase: "validation", completed, total: 0 } }; }
  for (const object of workbook.objects ?? []) { add(JSON.stringify(["object", object.nodeId]), object.sheetId, object.anchor); if (++completed % 256 === 0) yield { kind: "progress", value: { phase: "validation", completed, total: 0 } }; }
  for (const comment of workbook.comments ?? []) {
    const sheetId = sheets.get(comment.annotationId);
    if (sheetId && comment.format === "classic") for (const key of ["box", "commentPrBox"] as const) {
      const box = comment[key]; if (box) add(JSON.stringify([key, comment.annotationId]), sheetId, box.anchor);
    }
    if (++completed % 256 === 0) yield { kind: "progress", value: { phase: "validation", completed, total: 0 } };
  }
  return owners;
}

function checkedMeasurement(value: WorkbookLayoutMeasurement, sourceDigest: string, sheetIds: readonly string[], maxNodes: number): WorkbookMeasuredLayout {
  if (!value || typeof value !== "object" || Object.keys(value).some(key => !["sourceDigest", "fontEnvironment", "sheets"].includes(key))
    || value.sourceDigest !== sourceDigest || typeof value.fontEnvironment !== "string" || !value.fontEnvironment.length || value.fontEnvironment.length > 1024
    || !Array.isArray(value.sheets) || value.sheets.length !== sheetIds.length) invalid("Workbook layout measurements must match the requested candidate source.");
  const sheets = new Map<string, { rows: AxisDimensions; columns: AxisDimensions }>(), requested = new Set(sheetIds); let nodes = 0;
  for (const sheet of value.sheets) {
    if (!sheet || Object.keys(sheet).some(key => !["sheetId", "rows", "columns"].includes(key)) || !requested.has(sheet.sheetId) || sheets.has(sheet.sheetId)) invalid("Workbook layout measurements must contain each requested worksheet once.");
    const axis = (input: typeof sheet.rows, count: number, maxSize: number): AxisDimensions => {
      if (!input || Object.keys(input).some(key => !["defaultSizePt", "overrides"].includes(key)) || !Array.isArray(input.overrides)
        || !Number.isFinite(input.defaultSizePt) || input.defaultSizePt < 0 || input.defaultSizePt > maxSize) invalid("Invalid measured workbook axis.");
      nodes += input.overrides.length;
      if (nodes > maxNodes) throw new DocumentError("ResourceLimitExceeded", "Workbook layout measurements exceed their dimension-run budget.");
      for (const run of input.overrides) if (!run || Object.keys(run).some(key => !["from", "to", "sizePt"].includes(key)) || run.sizePt > maxSize) invalid("Invalid measured workbook dimension run.");
      return new AxisDimensions(count, input.defaultSizePt, input.overrides);
    };
    sheets.set(sheet.sheetId, { rows: axis(sheet.rows, MAX_SHEET_ROWS, 409), columns: axis(sheet.columns, MAX_SHEET_COLUMNS, 27273042316900 / 12700) });
  }
  return { fontEnvironment: value.fontEnvironment, sheets };
}

export function* measureWorkbookCandidate(model: WorkbookModel, sheetIds: readonly string[], measure: WorkbookLayoutMeasurer, maxNodes: number, calculationContext?: import("../workbook/scalar.js").WorkbookScalarContext): WorkbookPreparation<WorkbookMeasuredLayout> {
  const sourceDigest = assetDigest(encode(JSON.stringify([model.snapshot(), calculationContext ?? null]))); let result: WorkbookMeasuredLayout | undefined;
  yield { kind: "layout", measure, request: Object.freeze({ model, sourceDigest, sheetIds: Object.freeze([...sheetIds]), ...(calculationContext ? { calculationContext } : {}) }),
    accept: value => { result = checkedMeasurement(value, sourceDigest, sheetIds, maxNodes); } };
  return result ?? invalid("Workbook layout measurement did not complete.");
}

function rectangle(anchor: CellAnchor, layout: WorkbookMeasuredLayout, sheetId: string, validate = true) {
  const axes = layout.sheets.get(sheetId) ?? invalid("The measured worksheet is unavailable.");
  const x = axes.columns.position(anchor.from.column) + anchor.from.offsetXPt, y = axes.rows.position(anchor.from.row) + anchor.from.offsetYPt;
  const right = axes.columns.position(anchor.to.column) + anchor.to.offsetXPt, bottom = axes.rows.position(anchor.to.row) + anchor.to.offsetYPt;
  if (validate && (right < x || bottom < y)) invalid("A cell-bound drawing has negative extents in the measured worksheet.");
  return { x, y, width: right - x, height: bottom - y };
}
export function workbookLayoutBounds(owners: readonly Owner[], layout: WorkbookMeasuredLayout): ReadonlyMap<string, ReturnType<typeof rectangle>> {
  return new Map(owners.map(owner => [owner.key, rectangle(owner.anchor, layout, owner.sheetId)]));
}

/** Response flags affect changes to the grid. Explicitly replacing an anchor
 * or note box supplies new geometry, which is already the user's intent. */
export function* compensateWorkbookLayout(owners: readonly Owner[], before: ReturnType<typeof workbookLayoutBounds>, layout: WorkbookMeasuredLayout, operation: DocumentOperation): Generator<{ completed: number; total: number }, void> {
  let completed = 0;
  const explicit = operation.op === "updateWorkbookObject" && operation.changes.anchor ? JSON.stringify(["object", operation.nodeId]) : undefined;
  const note = operation.op === "setWorkbookNoteBox" ? operation.annotationId : undefined;
  const same = (a: number, b: number): boolean => Math.abs(a - b) <= Number.EPSILON * Math.max(1, Math.abs(a), Math.abs(b)) * 8;
  for (const owner of owners) {
    const old = before.get(owner.key);
    if (old && owner.key !== explicit && (!note || owner.key !== JSON.stringify(["box", note]) && owner.key !== JSON.stringify(["commentPrBox", note]))) {
      const anchor = owner.anchor, current = rectangle(anchor, layout, owner.sheetId, false);
      const x = anchor.moveWithCells ? current.x : old.x, y = anchor.moveWithCells ? current.y : old.y;
      const width = anchor.sizeWithCells ? current.width : old.width, height = anchor.sizeWithCells ? current.height : old.height;
      if (width < 0 || height < 0) invalid("Resizing the cell-bound drawing would produce negative measured extents.");
      const place = (point: number, currentPoint: number, marker: typeof anchor.from, axis: "row" | "column"): void => {
        if (same(point, currentPoint)) return;
        // Keep the source cell identities. Rebinding to the cell currently
        // under a fixed origin breaks later inverse structural operations.
        marker[axis === "row" ? "offsetYPt" : "offsetXPt"] += point - currentPoint;
      };
      place(x, current.x, anchor.from, "column"); place(y, current.y, anchor.from, "row");
      place(x + width, current.x + current.width, anchor.to, "column"); place(y + height, current.y + current.height, anchor.to, "row");
    }
    if (++completed % 256 === 0) yield { completed, total: owners.length };
  }
}
