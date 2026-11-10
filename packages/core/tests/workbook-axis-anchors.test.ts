import { afterEach, expect, test } from "vitest";
import { createDocumentSession, openDocument } from "../src/index.js";
import type { DocumentOperation, WorkbookDocumentMetadata, WorkbookDrawingAnchor, WorkbookLayoutMeasurer, WorkbookNoteBox, WorkbookSession } from "../src/index.js";
import { scalarWorkbook } from "./workbook-scalar-fixture.js";

const owned: WorkbookSession[] = [];
afterEach(() => owned.splice(0).forEach(session => session.dispose()));
const marker = (row: number, column = 1) => ({ row, column, offsetXPt: 2, offsetYPt: 3 });
const one = { kind: "one-cell", from: marker(2), widthPt: 120, heightPt: 90 } satisfies WorkbookDrawingAnchor;
const both = { kind: "two-cell", from: marker(2), to: { ...marker(6, 4), offsetXPt: 5, offsetYPt: 7 }, moveWithCells: true, sizeWithCells: true } satisfies WorkbookDrawingAnchor;
const absolute = { kind: "absolute", xPt: 20, yPt: 30, widthPt: 160, heightPt: 80 } satisfies WorkbookDrawingAnchor;
const box: WorkbookNoteBox = { anchor: { kind: "one-cell", from: marker(3, 2), widthPt: 180, heightPt: 90 }, visible: true, printable: true, locked: true,
  lockText: false, defaultSize: false, disabled: false, autoFill: false, autoLine: false, autoScale: false, justifyLastLine: false };
const axis = (direction: "row" | "column", action: "insert" | "delete", index: number, count = 1) => ({ op: "editWorkbookAxis" as const, sheetId: "sheet-1", axis: direction, action, index, count });
const set: DocumentOperation = { op: "setWorkbookCells", sheetId: "sheet-1", cells: [{ row: 0, column: 0, source: { kind: "value", value: { type: "number", value: 99 } } }] };
function tx(s: WorkbookSession, operations: DocumentOperation[]) { const m = s.snapshot().manifest; return { protocolVersion: 1 as const, documentId: m.documentId, baseRevision: m.revision, operations }; }
const metrics: WorkbookLayoutMeasurer = request => ({ sourceDigest: request.sourceDigest, fontEnvironment: 'axis-source-test-points',
  sheets: request.sheetIds.map(sheetId => ({ sheetId, rows: { defaultSizePt: 15, overrides: [] }, columns: { defaultSizePt: 40, overrides: [] } })) });
function apply(s: WorkbookSession, operations: DocumentOperation[]) { return s.commit(s.prepare(tx(s, operations), { measureWorkbookLayout: metrics })); }
function fixture() {
  const setup = createDocumentSession(scalarWorkbook([{ row: 255, column: 31, value: { type: "number", value: 7 } }, { row: 4, column: 3, value: { type: "number", value: 17 }, formula: "=$AF$256" }]));
  apply(setup, [{ op: "insertWorkbookSheet", name: "Other" }]); const other = setup.snapshot().metadata.workbook.sheets[1]!.sheetId;
  apply(setup, [one, both, absolute].map((anchor, index) => ({ op: "insertWorkbookObject", sheetId: "sheet-1", anchor, name: "Object " + index, object: { kind: "geometry", preset: "rect", html: "<p><strong>禰󠄀😀</strong>&#13;<br>尾</p>" } })));
  apply(setup, [{ op: "insertWorkbookObject", sheetId: other, anchor: { ...both, kind: "two-cell", moveWithCells: false, sizeWithCells: false }, object: { kind: "text", html: "<p>Other</p>" } },
    { op: "addWorkbookComment", anchor: { kind: "cells", sheetId: "sheet-1", ranges: [{ from: { row: 4, column: 2 }, to: { row: 4, column: 2 } }] }, html: "<p>备注<strong>禰󠄀😀</strong></p>", format: "classic", box }]);
  const files = new Map(setup.serialize()), metadata = JSON.parse(new TextDecoder().decode(files.get("metadata.json")!)) as WorkbookDocumentMetadata;
  const note = metadata.workbook.comments![0]!;
  if (note.format !== "classic") throw Error("Missing classic note.");
  note.boxSource = "vml"; note.commentPrBox = { ...structuredClone(box), visible: false, anchor: { kind: "two-cell", from: marker(3, 2), to: marker(9, 5), moveWithCells: true, sizeWithCells: true } }; note.boxDifferences = ["anchor", "visible"];
  files.set("metadata.json", new TextEncoder().encode(JSON.stringify(metadata))); setup.dispose();
  const s = createDocumentSession(openDocument(files, { kind: "workbook" }), { sessionLimits: { maxCandidates: 1 } }); owned.push(s); return s;
}
function classic(s: WorkbookSession) { const value = s.snapshot().metadata.workbook.comments![0]!; if (value.format !== "classic") throw Error("Missing classic note."); return value; }

test("axis insertions move one-cell objects and both resizing markers, retain physical extents, rich identities and both note sources", () => {
  const s = fixture(), before = s.snapshot(), candidate = s.prepare(tx(s, [axis("row", "insert", 1, 2)]), { measureWorkbookLayout: metrics });
  expect(s.snapshot()).toEqual(before); expect(candidate.preview.metadata.workbook.objects![0]!.anchor).toEqual({ ...one, from: marker(4) }); s.commit(candidate);
  const objects = s.snapshot().metadata.workbook.objects!;
  expect(objects[0]!.anchor).toEqual({ ...one, from: marker(4) }); expect(objects[1]!.anchor).toEqual({ ...both, from: marker(4), to: { ...marker(8, 4), offsetXPt: 5, offsetYPt: 7 } });
  expect(objects.slice(2)).toEqual(before.metadata.workbook.objects!.slice(2));
  expect(classic(s).box!.anchor).toEqual({ ...box.anchor, from: marker(5, 2) }); expect(classic(s).commentPrBox!.anchor).toMatchObject({ from: marker(5, 2), to: marker(11, 5) });
  expect(classic(s)).toMatchObject({ boxSource: "vml", boxDifferences: ["anchor", "visible"] });
  expect(s.snapshot().annotations[0]).toMatchObject({ status: "open", anchor: { ranges: [{ from: { row: 6, column: 2 }, to: { row: 6, column: 2 } }] } });
  expect(s.getCell("sheet-1", 6, 3)).toMatchObject({ value: { type: "number", value: 17 }, formula: { cacheState: "stale" } }); s.calculateSync(); expect(s.getCalculatedCell("sheet-1", 6, 3)).toMatchObject({ value: { type: "number", value: 7 } });
  expect(s.snapshot().content).toEqual(before.content); expect(s.snapshot().styles).toEqual(before.styles); expect(s.history().undoSteps).toBe(1);
  expect(openDocument(s.serialize(), { kind: "workbook" }).snapshot()).toEqual(s.snapshot()); s.undo(); expect(s.snapshot().metadata).toEqual(before.metadata); expect(s.snapshot().manifest.chunks).toEqual(before.manifest.chunks);
});

test("an internal insertion grows only two-cell extents and its same-batch reversal preserves exact source and redo", () => {
  const s = fixture(); apply(s, [set]); s.undo(); const before = s.serialize(), metadata = s.snapshot().metadata;
  expect(apply(s, [axis("row", "insert", 4, 2), axis("row", "delete", 4, 2)]).changed).toBe(false); expect(s.serialize()).toEqual(before); expect(s.history()).toMatchObject({ undoSteps: 0, redoSteps: 1 });
  apply(s, [axis("row", "insert", 4, 2)]); expect(s.snapshot().metadata.workbook.objects![0]!.anchor).toEqual(one);
  expect(s.snapshot().metadata.workbook.objects![1]!.anchor).toEqual({ ...both, to: { ...marker(8, 4), offsetXPt: 5, offsetYPt: 7 } });
  expect(classic(s).box!.anchor).toEqual(box.anchor); expect(classic(s).commentPrBox!.anchor).toMatchObject({ from: marker(3, 2), to: marker(11, 5) });
  s.undo(); expect(s.snapshot().metadata).toEqual(metadata);
});

test("deleted anchor positions clamp to the surviving boundary, while invalidated note cells retain their body and recovery status", () => {
  const s = fixture(), before = s.snapshot(); apply(s, [axis("row", "delete", 2, 4)]);
  expect(s.snapshot().metadata.workbook.objects![0]!.anchor).toEqual(one); expect(s.snapshot().metadata.workbook.objects![1]!.anchor).toMatchObject({ from: marker(2), to: { row: 2, column: 4, offsetXPt: 5, offsetYPt: 7 } });
  expect(classic(s).box!.anchor).toEqual({ ...box.anchor, from: marker(2, 2) }); expect(classic(s).commentPrBox!.anchor).toMatchObject({ from: marker(2, 2), to: marker(5, 5) });
  expect(s.snapshot().annotations[0]).toMatchObject({ annotationId: before.annotations[0]!.annotationId, bodyNodeId: before.annotations[0]!.bodyNodeId, status: "needs_reanchor" });
  expect(classic(s).recoveryStatus).toBe("open"); expect(s.snapshot().content).toEqual(before.content); expect(openDocument(s.serialize(), { kind: "workbook" }).snapshot()).toEqual(s.snapshot());
  s.undo(); expect(s.snapshot().metadata).toEqual(before.metadata); expect(s.snapshot().annotations.map(value => value.status)).toEqual(["open"]);
});

test("fixed or mixed two-cell modes and native-bound overflow reject every earlier write in the batch", () => {
  for (const flags of [{ moveWithCells: true, sizeWithCells: false }, { moveWithCells: false, sizeWithCells: false }, { moveWithCells: false, sizeWithCells: true }]) {
    const s = fixture(), nodeId = s.snapshot().metadata.workbook.objects![0]!.nodeId; apply(s, [{ op: "updateWorkbookObject", nodeId, changes: { anchor: { ...both, kind: "two-cell", ...flags } } }]); s.clearHistory(); const before = s.serialize();
    expect(() => s.prepare(tx(s, [set, axis("column", "insert", 0)]))).toThrowError(expect.objectContaining({ code: "InvalidOperation" })); expect(s.serialize()).toEqual(before); expect(s.history().undoSteps).toBe(0);
  }
  const s = fixture(), nodeId = s.snapshot().metadata.workbook.objects![0]!.nodeId; apply(s, [{ op: "updateWorkbookObject", nodeId, changes: { anchor: { ...one, kind: "one-cell", from: marker(1_048_575) } } }]); s.clearHistory(); const before = s.serialize();
  expect(() => apply(s, [set, axis("row", "insert", 0)])).toThrowError(expect.objectContaining({ code: "InvalidOperation" })); expect(s.serialize()).toEqual(before); expect(s.history().undoSteps).toBe(0);
});

test("cancellation after private anchor mapping releases its sole candidate slot, and a later revision rejects a prepared mapping", async () => {
  const s = fixture(), before = s.serialize(), signal = { aborted: false }; let reached = false;
  await expect(s.prepareAsync(tx(s, [axis("row", "insert", 1)]), { measureWorkbookLayout: metrics }, { signal, yield: async () => {}, onProgress(value) { if (value.phase === "cells" && value.completed >= 8) { signal.aborted = true; reached = true; } } })).rejects.toMatchObject({ code: "OperationCancelled" });
  expect(reached).toBe(true); expect(s.serialize()).toEqual(before); expect(s.history().undoSteps).toBe(0);
  s.discard(s.prepare(tx(s, [axis("row", "insert", 1)]), { measureWorkbookLayout: metrics })); apply(s, [set]); const candidate = s.prepare(tx(s, [axis("row", "insert", 1)]), { measureWorkbookLayout: metrics }); s.undo();
  expect(() => s.commit(candidate)).toThrowError(expect.objectContaining({ code: "RevisionConflict" })); s.discard(candidate); expect(s.snapshot().metadata.workbook.objects![0]!.anchor).toEqual(one);
});
