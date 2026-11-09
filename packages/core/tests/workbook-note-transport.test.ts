import { afterEach, expect, test } from "vitest";
import { createDocumentSession, createWorkbook, openDocument, WorkbookAxisDimensions } from "../src/index.js";
import type { DeepReadonly, DocumentOperation, WorkbookDrawingAnchor, WorkbookLayoutMeasurement, WorkbookLayoutRequest, WorkbookModel, WorkbookNoteBox, WorkbookSession } from "../src/index.js";

const sessions: WorkbookSession[] = [], sourceId = "sheet-1";
afterEach(() => sessions.splice(0).forEach(s => s.dispose()));
const point = (row: number, column: number) => ({ row, column });
const range = (row: number, column: number, lastRow = row, lastColumn = column) => ({ from: point(row, column), to: point(lastRow, lastColumn) });
const marker = (row: number, column: number) => ({ row, column, offsetXPt: 1.25, offsetYPt: 2.5 });
const anchors: WorkbookDrawingAnchor[] = [
  { kind: "absolute", xPt: 81.25, yPt: 32.5, widthPt: 82.5, heightPt: 47.5 },
  { kind: "one-cell", from: marker(2, 2), widthPt: 82.5, heightPt: 47.5 },
  ...[false, true].flatMap(moveWithCells => [false, true].map(sizeWithCells => ({ kind: "two-cell" as const, from: marker(2, 2), to: { ...marker(5, 4), offsetXPt: 3.75, offsetYPt: 5 }, moveWithCells, sizeWithCells }))),
];
function transaction(s: WorkbookSession, operations: DocumentOperation[]) { const m = s.snapshot().manifest; return { protocolVersion: 1 as const, documentId: m.documentId, baseRevision: m.revision, operations }; }
function metrics(request: WorkbookLayoutRequest): WorkbookLayoutMeasurement {
  return { sourceDigest: request.sourceDigest, fontEnvironment: "explicit-two-sheet-point-metrics", sheets: request.sheetIds.map(sheetId => {
    const sheet = request.model.snapshot().metadata.workbook.sheets.find(sheet => sheet.sheetId === sheetId)!;
    return { sheetId, rows: { defaultSizePt: sheet.defaultRowHeightPt, overrides: sheet.rowDimensions.map(run => ({ from: run.from, to: run.to, sizePt: run.hidden ? 0 : run.heightPt ?? sheet.defaultRowHeightPt })) },
      columns: { defaultSizePt: sheet.name === "Other" ? 70 : 40, overrides: [] } };
  }) };
}
function apply(s: WorkbookSession, operations: DocumentOperation[]) { return s.commit(s.prepare(transaction(s, operations), { measureWorkbookLayout: metrics })); }
function example(anchor: WorkbookDrawingAnchor) {
  const original = createDocumentSession(openDocument(createWorkbook({ documentId: "note-transport" }), { kind: "workbook" })); sessions.push(original);
  const box: WorkbookNoteBox = { anchor, visible: true, printable: true, locked: true, lockText: false, defaultSize: false, disabled: false, autoFill: false, autoLine: false, autoScale: false, justifyLastLine: false };
  apply(original, [{ op: "insertWorkbookSheet", name: "Other" }, { op: "addWorkbookComment", anchor: { kind: "cells", sheetId: sourceId, ranges: [range(1, 1)] }, format: "classic", html: "<p>禰󠄀😀<strong>预算</strong></p>", box }]);
  const targetId = original.snapshot().metadata.workbook.sheets[1]!.sheetId;
  apply(original, [{ op: "updateWorkbookSheet", sheetId: targetId, changes: { defaultRowHeightPt: 30 } }, { op: "setWorkbookCells", sheetId: sourceId, cells: [{ row: 1, column: 1, source: { kind: "value", value: { type: "number", value: 3 } } }] }]);
  const files = new Map(original.serialize()), metadata = JSON.parse(new TextDecoder().decode(files.get("metadata.json")!));
  Object.assign(metadata.workbook.comments[0], { nativeShapeId: 143, boxSource: "vml", commentPrBox: { ...box, anchor: anchors[1] }, boxDifferences: ["anchor"] });
  files.set("metadata.json", new TextEncoder().encode(JSON.stringify(metadata)));
  const s = createDocumentSession(openDocument(files, { kind: "workbook" })); sessions.push(s); return { s, targetId };
}
function rectangles(model: WorkbookModel) {
  const a = model.snapshot().annotations[0]!.anchor; if (a.kind !== "cells") throw Error("Expected native anchor");
  const measurement = metrics({ model, sourceDigest: "assertion", sheetIds: [a.sheetId] }).sheets[0]!, rows = new WorkbookAxisDimensions(1_048_576, measurement.rows.defaultSizePt, measurement.rows.overrides),
    columns = new WorkbookAxisDimensions(16_384, measurement.columns.defaultSizePt, measurement.columns.overrides);
  const rect = (anchor: DeepReadonly<WorkbookDrawingAnchor>) => {
    if (anchor.kind === "absolute") return { x: anchor.xPt, y: anchor.yPt, width: anchor.widthPt, height: anchor.heightPt };
    const x = columns.position(anchor.from.column) + anchor.from.offsetXPt, y = rows.position(anchor.from.row) + anchor.from.offsetYPt;
    return { x, y, width: anchor.kind === "one-cell" ? anchor.widthPt : columns.position(anchor.to.column) + anchor.to.offsetXPt - x, height: anchor.kind === "one-cell" ? anchor.heightPt : rows.position(anchor.to.row) + anchor.to.offsetYPt - y };
  };
  const binding = model.snapshot().metadata.workbook.comments![0]!; if (binding.format !== "classic") throw Error("Expected classic binding");
  return { box: rect(binding.box!.anchor), commentPrBox: rect(binding.commentPrBox!.anchor) };
}
const move = (targetId: string, back = false): DocumentOperation => ({ op: "moveWorkbookRange", sheetId: back ? targetId : sourceId, range: back ? range(259, 7) : range(1, 1), targetSheetId: back ? sourceId : targetId, target: back ? point(1, 1) : point(259, 7) });
test.each(anchors)("cross-sheet note transport preserves both physical boxes and native/body identity for $kind", anchor => {
  const { s, targetId } = example(anchor), before = s.snapshot(), geometry = rectangles(s), body = s.getNode(before.annotations[0]!.bodyNodeId!)!.html;
  const candidate = s.prepare(transaction(s, [move(targetId)]), { measureWorkbookLayout: metrics }); expect(s.snapshot()).toEqual(before); s.commit(candidate);
  const result = rectangles(s);
  for (const key of ["box", "commentPrBox"] as const) {
    expect(result[key].x).toBeCloseTo(geometry[key].x + 450, 9); expect(result[key].y).toBeCloseTo(geometry[key].y + 7755, 9);
    expect(result[key].width).toBeCloseTo(geometry[key].width, 9); expect(result[key].height).toBeCloseTo(geometry[key].height, 9);
  }
  expect(s.snapshot().annotations[0]).toMatchObject({ annotationId: before.annotations[0]!.annotationId, bodyNodeId: before.annotations[0]!.bodyNodeId, anchor: { sheetId: targetId, ranges: [range(259, 7)] } });
  expect(s.snapshot().metadata.workbook.comments![0]).toMatchObject({ nativeShapeId: 143, boxSource: "vml", boxDifferences: ["anchor"] });
  expect(s.getNode(before.annotations[0]!.bodyNodeId!)!.html).toBe(body); expect(s.history().undoSteps).toBe(1);
  expect(openDocument(s.serialize(), { kind: "workbook" }).snapshot()).toEqual(s.snapshot());
  s.undo(); expect(s.snapshot().metadata).toEqual(before.metadata); expect(s.snapshot().manifest.chunks).toEqual(before.manifest.chunks);
});
test("inverse cuts preserve exact source and history while only affected positioned notes require measurement", () => {
  const { s, targetId } = example(anchors[2]!), before = s.serialize();
  expect(apply(s, [move(targetId), move(targetId, true)]).changed).toBe(false); expect(s.serialize()).toEqual(before); expect(s.history().undoSteps).toBe(0);
  expect(() => s.prepare(transaction(s, [move(targetId)]))).toThrowError(expect.objectContaining({ code: "InvalidOperation" })); expect(s.serialize()).toEqual(before);
  expect(s.commit(s.prepare(transaction(s, [{ op: "moveWorkbookRange", sheetId: sourceId, range: range(1, 1), targetSheetId: sourceId, target: point(1, 1) }]))).changed).toBe(false);
  expect(s.commit(s.prepare(transaction(s, [{ op: "moveWorkbookRange", sheetId: sourceId, range: range(30, 0), targetSheetId: sourceId, target: point(40, 0) }]))).changed).toBe(false);
});
test("sort transports boxes by data identity, preserves physical extents across unequal row heights and composes its inverse", () => {
  const { s } = example(anchors[2]!);
  apply(s, [{ op: "updateWorkbookSheet", sheetId: sourceId, changes: { rowDimensions: [{ from: 2, to: 2, heightPt: 45 }] } }]); s.clearHistory();
  const before = s.snapshot(), geometry = rectangles(s), operation: DocumentOperation = { op: "reorderWorkbookRows", sheetId: sourceId, range: range(1, 0, 3, 1), rows: [3, 2, 1] };
  apply(s, [operation]); expect(s.snapshot().annotations[0]!.anchor).toMatchObject({ ranges: [range(3, 1)] });
  const current = rectangles(s);
  for (const key of ["box", "commentPrBox"] as const) { expect(current[key].y).toBeCloseTo(geometry[key].y + 60, 9); expect(current[key].height).toBeCloseTo(geometry[key].height, 9); }
  s.undo(); expect(s.snapshot().metadata).toEqual(before.metadata);
  const serialized = s.serialize(); expect(apply(s, [operation, operation]).changed).toBe(false); expect(s.serialize()).toEqual(serialized);
});
test("cancelling after structural measurements discards geometry and stories along with moved cells", async () => {
  const { s, targetId } = example(anchors[2]!), before = s.serialize(), signal = { aborted: false }; let calls = 0;
  await expect(s.prepareAsync(transaction(s, [move(targetId)]), { measureWorkbookLayout: request => { const value = metrics(request); if (++calls === 2) signal.aborted = true; return value; } }, { signal, yield: async () => {} })).rejects.toMatchObject({ code: "OperationCancelled" });
  expect(calls).toBe(2); expect(s.serialize()).toEqual(before); expect(s.history().undoSteps).toBe(0);
});

test("mixed-sheet fragments retain original note geometry for recovery and a composed inverse restores the active source", () => {
  const { s, targetId } = example(anchors[2]!), original = s.snapshot().annotations[0]!;
  apply(s, [{ op: "updateWorkbookComment", annotationId: original.annotationId, changes: { anchor: { kind: "cells", sheetId: sourceId, ranges: [range(1, 1), range(3, 1)] } } }]); s.clearHistory();
  const before = s.snapshot(); apply(s, [move(targetId)]);
  expect(s.snapshot().annotations[0]).toMatchObject({ status: "needs_reanchor", anchor: before.annotations[0]!.anchor });
  expect(s.snapshot().metadata.workbook.comments![0]).toMatchObject({ recoveryStatus: "open", box: before.metadata.workbook.comments![0]!.format === "classic" ? before.metadata.workbook.comments![0]!.box : undefined });
  s.undo(); const bytes = s.serialize(); expect(apply(s, [move(targetId), move(targetId, true)]).changed).toBe(false);
  expect(s.serialize()).toEqual(bytes); expect(s.snapshot().annotations[0]!.status).toBe("open");
});
test("transport to the last native cell clamps box markers while preserving physical extents beyond the grid", () => {
  const { s, targetId } = example(anchors[2]!), before = rectangles(s), original = s.serialize();
  const operation: DocumentOperation = { op: "moveWorkbookRange", sheetId: sourceId, range: range(1, 1), targetSheetId: targetId, target: point(1_048_575, 16_383) };
  apply(s, [operation]); const after = rectangles(s);
  for (const key of ["box", "commentPrBox"] as const) { expect(after[key].width).toBeCloseTo(before[key].width, 8); expect(after[key].height).toBeCloseTo(before[key].height, 8); }
  const native = s.snapshot().metadata.workbook.comments![0]!; if (native.format !== "classic" || native.box!.anchor.kind !== "two-cell") throw Error("Missing native cell box");
  expect(native.box!.anchor.from).toMatchObject({ row: 1_048_575, column: 16_383 }); expect(native.box!.anchor.to).toMatchObject({ row: 1_048_575, column: 16_383 });
  expect(openDocument(s.serialize(), { kind: "workbook" }).snapshot()).toEqual(s.snapshot());
  s.undo(); const restored = s.serialize();
  const inverse: DocumentOperation = { op: "moveWorkbookRange", sheetId: targetId, range: range(1_048_575, 16_383), targetSheetId: sourceId, target: point(1, 1) };
  expect(apply(s, [operation, inverse]).changed).toBe(false); expect(s.serialize()).toEqual(restored);
  expect(s.snapshot().metadata).toEqual(openDocument(original, { kind: "workbook" }).snapshot().metadata);
});

test("explicit replacement of one far note box resets only its own origin and retains the untouched native box markers", () => {
  const { s, targetId } = example(anchors[2]!), before = s.snapshot(), binding = before.metadata.workbook.comments![0]!;
  if (binding.format !== "classic") throw Error("Expected classic note");
  const target = point(1_048_575, 16_383), forward: DocumentOperation = { op: "moveWorkbookRange", sheetId: sourceId, range: range(1, 1), targetSheetId: targetId, target };
  const back: DocumentOperation = { op: "moveWorkbookRange", sheetId: targetId, range: { from: target, to: target }, targetSheetId: sourceId, target: point(1, 1) };
  apply(s, [forward, { op: "setWorkbookNoteBox", annotationId: binding.annotationId, box: { ...binding.box!, visible: false,
    anchor: { kind: "absolute", xPt: target.column * 70 + 70, yPt: target.row * 30 + 30, widthPt: 150, heightPt: 70 } } }, back]);
  const current = s.snapshot().metadata.workbook.comments![0]!; if (current.format !== "classic") throw Error("Missing classic note");
  expect(current.commentPrBox).toEqual(binding.commentPrBox);
  expect(current.box).toMatchObject({ visible: false, anchor: { kind: "absolute", xPt: 110, yPt: 45, widthPt: 150, heightPt: 70 } });
});

test("overwriting the sole destination note retains recovery geometry without demanding transport measurements", () => {
  const { s } = example(anchors[0]!), before = s.snapshot(), body = s.getNode(before.annotations[0]!.bodyNodeId!)!.html;
  const candidate = s.prepare(transaction(s, [{ op: "moveWorkbookRange", sheetId: sourceId, range: range(30, 0), targetSheetId: sourceId,
    target: point(1, 1), overwrittenContent: "clear" }]));
  expect(s.snapshot()).toEqual(before); s.commit(candidate);
  expect(s.snapshot().annotations[0]).toMatchObject({ status: "needs_reanchor", anchor: before.annotations[0]!.anchor });
  expect(s.snapshot().metadata.workbook.comments![0]).toMatchObject({ recoveryStatus: "open", box: before.metadata.workbook.comments![0]!.format === "classic" ? before.metadata.workbook.comments![0]!.box : undefined });
  expect(s.getNode(before.annotations[0]!.bodyNodeId!)!.html).toBe(body);
  s.undo(); expect(s.snapshot().metadata).toEqual(before.metadata);
});
