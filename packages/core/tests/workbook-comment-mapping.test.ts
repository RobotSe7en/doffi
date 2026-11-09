import { afterEach, expect, test } from "vitest";
import { createDocumentSession, openDocument } from "../src/index.js";
import type { DocumentOperation, WorkbookLayoutMeasurement, WorkbookLayoutRequest, WorkbookNoteBox, WorkbookRange, WorkbookSession } from "../src/index.js";
import { scalarWorkbook } from "./workbook-scalar-fixture.js";

const owned: WorkbookSession[] = [], sheetId = "sheet-1", range = (from: number, to = from, column = 0): WorkbookRange => ({ from: { row: from, column }, to: { row: to, column } });
afterEach(() => owned.splice(0).forEach(session => session.dispose()));
function workbook() { const s = createDocumentSession(scalarWorkbook([1, 2, 3, 4, 5].map(value => ({ row: value, column: 0, value: { type: "number", value } })))); owned.push(s); return s; }
function transaction(s: WorkbookSession, operations: DocumentOperation[]) { const m = s.snapshot().manifest; return { protocolVersion: 1 as const, documentId: m.documentId, baseRevision: m.revision, operations }; }
const metrics = (request: WorkbookLayoutRequest): WorkbookLayoutMeasurement => ({ sourceDigest: request.sourceDigest, fontEnvironment: "comment-mapping-point-metrics",
  sheets: request.sheetIds.map(sheetId => ({ sheetId, rows: { defaultSizePt: 15, overrides: [] }, columns: { defaultSizePt: 45, overrides: [] } })) });
const commit = (s: WorkbookSession, operations: DocumentOperation[]) => s.commit(s.prepare(transaction(s, operations), { measureWorkbookLayout: metrics }));
const binding = (s: WorkbookSession, id: string) => s.snapshot().metadata.workbook.comments!.find(comment => comment.annotationId === id)!;
const absoluteBox: WorkbookNoteBox = { anchor: { kind: "absolute", xPt: 20, yPt: 30, widthPt: 160, heightPt: 90 }, visible: true, printable: true, locked: true, lockText: false,
  defaultSize: false, disabled: false, autoFill: false, autoLine: false, autoScale: false, justifyLastLine: false };
function add(s: WorkbookSession, area: WorkbookRange, format: "classic" | "modern" = "modern", atSheet = sheetId) {
  commit(s, [{ op: "addWorkbookComment", anchor: { kind: "cells", sheetId: atSheet, ranges: [area] }, format, html: "<p>源<strong>禰󠄀😀</strong></p>", ...(format === "classic" ? { box: absoluteBox } : {}) }]); return s.getAnnotations().at(-1)!;
}
const axis = (action: "insert" | "delete", index: number, count = 1): DocumentOperation => ({ op: "editWorkbookAxis", sheetId, axis: "row", action, index, count, ...(action === "insert" ? { formatSource: "none" as const } : {}) });

test("axis changes fragment native/host anchors without selecting new blank cells, preserve absolute note boxes and retain deleted resolved bodies for recovery", () => {
  const s = workbook(), modern = add(s, range(1, 3)), classic = add(s, range(4, 4, 2), "classic");
  commit(s, [{ op: "addWorkbookCommentReply", annotationId: modern.annotationId, html: "<p>回复禰󠄀</p>" }, { op: "updateWorkbookComment", annotationId: modern.annotationId, changes: { status: "resolved" } }]); s.clearHistory();
  const before = s.snapshot(), body = s.getNode(modern.bodyNodeId!)!.html;
  commit(s, [axis("insert", 2, 2)]);
  expect(s.getAnnotations()[0]!.anchor).toEqual({ kind: "cells", sheetId, ranges: [range(1), range(4, 5)] });
  expect(s.getAnnotations()[1]!.anchor).toEqual({ kind: "cells", sheetId, ranges: [range(6, 6, 2)] }); expect(binding(s, classic.annotationId)).toMatchObject({ box: absoluteBox });
  const inserted = s.snapshot(); commit(s, [axis("delete", 1, 5)]);
  expect(s.getAnnotations()[0]).toMatchObject({ status: "needs_reanchor", anchor: inserted.annotations[0]!.anchor }); expect(binding(s, modern.annotationId)).toMatchObject({ recoveryStatus: "resolved" });
  expect(s.getNode(modern.bodyNodeId!)!.html).toBe(body); expect(s.getAnnotations()[0]!.replies).toHaveLength(1);
  expect(s.getAnnotations()[1]!.anchor).toEqual({ kind: "cells", sheetId, ranges: [range(1, 1, 2)] }); expect(openDocument(s.serialize(), { kind: "workbook" }).snapshot()).toEqual(s.snapshot());
  s.undo(); expect(s.snapshot().metadata).toEqual(inserted.metadata); s.undo(); expect(s.snapshot().metadata).toEqual(before.metadata); expect(s.snapshot().content).toEqual(before.content); expect(s.history()).toMatchObject({ undoSteps: 0, redoSteps: 2 });
});

test("sorting follows complete records and exact fragments; a reversed two-sort batch restores original grouping, cells, source bytes and redo", () => {
  const s = workbook(), root = add(s, range(1, 2)), classic = add(s, range(4), "classic"); s.clearHistory();
  const sort: DocumentOperation = { op: "reorderWorkbookRows", sheetId, range: { from: { row: 1, column: 0 }, to: { row: 4, column: 1 } }, rows: [4, 1, 3, 2] };
  commit(s, [sort]); expect(s.getAnnotations()[0]!.anchor).toEqual({ kind: "cells", sheetId, ranges: [range(1), range(4)] });
  expect(binding(s, classic.annotationId)).toMatchObject({ box: { anchor: { kind: "absolute", xPt: 20, yPt: 0, widthPt: 160, heightPt: 90 } } });
  expect(s.getCell(sheetId, 1, 0)?.value).toEqual({ type: "number", value: 2 }); expect(s.getCell(sheetId, 4, 0)?.value).toEqual({ type: "number", value: 1 });
  s.undo(); const before = s.serialize(), history = s.history();
  expect(commit(s, [sort, { ...sort, rows: [2, 4, 3, 1] }]).changed).toBe(false); expect(s.serialize()).toEqual(before); expect(s.history()).toEqual(history); expect(s.getAnnotations()[0]!.annotationId).toBe(root.annotationId);
});

test("cross-sheet cuts move native rich stories with an intact thread while split or overwritten annotations require reanchoring and retain bodies", () => {
  const s = workbook(); commit(s, [{ op: "insertWorkbookSheet", name: "Other" }]); const other = s.snapshot().metadata.workbook.sheets[1]!.sheetId;
  const whole = add(s, range(2)), partial = add(s, range(1, 3)), target = add(s, range(255, 255, 31), "classic", other);
  commit(s, [{ op: "addWorkbookCommentReply", annotationId: whole.annotationId, html: "<p>存续回复禰󠄀</p>" },
    { op: "updateWorkbookComment", annotationId: partial.annotationId, changes: { status: "resolved" } }]); s.clearHistory();
  const before = s.snapshot(), bodies = before.annotations.flatMap(annotation => [annotation, ...annotation.replies ?? []]).flatMap(message => message.bodyNodeId ? [[message.bodyNodeId, s.getNode(message.bodyNodeId)!.html] as const] : []);
  commit(s, [{ op: "moveWorkbookRange", sheetId, range: { from: { row: 2, column: 0 }, to: { row: 3, column: 1 } }, targetSheetId: other, target: { row: 255, column: 31 } }]);
  const [moved, split, overwritten] = s.getAnnotations(); expect(moved).toMatchObject({ status: "open", anchor: { kind: "cells", sheetId: other, ranges: [range(255, 255, 31)] } });
  for (const message of [moved!, ...moved!.replies ?? []]) expect(s.snapshot().metadata.workbook.stories!.find(story => story.rootNodeId === message.bodyNodeId)!.sheetId).toBe(other);
  expect(split).toMatchObject({ status: "needs_reanchor", anchor: partial.anchor }); expect(binding(s, partial.annotationId)).toMatchObject({ recoveryStatus: "resolved" });
  expect(overwritten).toMatchObject({ status: "needs_reanchor", anchor: target.anchor }); expect(binding(s, target.annotationId)).toMatchObject({ recoveryStatus: "open", box: absoluteBox });
  for (const [id, html] of bodies) expect(s.getNode(id)!.html).toBe(html); expect(openDocument(s.serialize(), { kind: "workbook" }).snapshot()).toEqual(s.snapshot());
  expect(s.history().undoSteps).toBe(1); s.undo(); expect(s.snapshot().metadata).toEqual(before.metadata); expect(s.snapshot().manifest.chunks).toEqual(before.manifest.chunks); expect(s.snapshot().content).toEqual(before.content);
});

test("cancellation while mapping many source-owned threads cannot publish moved stories, current cells or history", async () => {
  const s = workbook(); commit(s, Array.from({ length: 130 }, (_, index): DocumentOperation => ({ op: "addWorkbookComment", anchor: { kind: "cells", sheetId, ranges: [range(index)] }, html: "<p>原源</p>" })));
  s.clearHistory(); const before = s.serialize(), signal = { aborted: false }; let mapping = false;
  await expect(s.prepareAsync(transaction(s, [axis("insert", 0)]), {}, { signal, onProgress(progress) { if (progress.phase === "cells" && progress.completed === 64) mapping = true; },
    yield: async () => { if (mapping) signal.aborted = true; } })).rejects.toMatchObject({ code: "OperationCancelled" });
  expect(mapping).toBe(true);
  expect(s.serialize()).toEqual(before); expect(s.history().undoSteps).toBe(0); expect(s.getAnnotations()).toHaveLength(130);
});

test("absolute drawings keep fixed document geometry while axis edits update their current chart references and source cells", () => {
  const s = workbook(); commit(s, [{ op: "insertWorkbookChart", sheetId, anchor: { kind: "absolute", xPt: 40, yPt: 50, widthPt: 180, heightPt: 90 }, kind: "column",
    data: { series: [{ values: { expression: "=Sheet1!$A$2:$A$3", valueType: "number" } }] }, title: "<p>固定位置</p>" }]);
  const before = s.snapshot(), chart = before.metadata.workbook.charts![0]!, body = s.getNode(chart.nodeId)!.html;
  commit(s, [axis("insert", 1)]); expect(s.snapshot().metadata.workbook.objects).toEqual(before.metadata.workbook.objects);
  expect(s.snapshot().metadata.workbook.charts![0]!.series[0]!.values).toMatchObject({ expression: "='Sheet1'!$A$3:$A$4" }); expect(s.getNode(chart.nodeId)!.html).toBe(body);
  expect(s.getCell(sheetId, 2, 0)?.value).toEqual({ type: "number", value: 1 }); s.undo(); expect(s.snapshot().metadata).toEqual(before.metadata);
});

test("a cut split is private to the batch: inverse moves restore exact active anchors, while an explicit invalidation remains a user decision", () => {
  const s = workbook(); commit(s, [{ op: "insertWorkbookSheet", name: "Other" }]); const other = s.snapshot().metadata.workbook.sheets[1]!.sheetId, root = add(s, range(1, 3));
  commit(s, [{ op: "updateWorkbookComment", annotationId: root.annotationId, changes: { status: "resolved" } }]); s.clearHistory();
  const move: DocumentOperation = { op: "moveWorkbookRange", sheetId, range: { from: { row: 2, column: 0 }, to: { row: 3, column: 1 } }, targetSheetId: other, target: { row: 255, column: 31 } };
  const back: DocumentOperation = { op: "moveWorkbookRange", sheetId: other, range: { from: { row: 255, column: 31 }, to: { row: 256, column: 32 } }, targetSheetId: sheetId, target: { row: 2, column: 0 } };
  const before = s.serialize(); expect(commit(s, [move, back]).changed).toBe(false); expect(s.serialize()).toEqual(before); expect(s.getAnnotations()[0]!.status).toBe("resolved");
  expect(commit(s, [move, { op: "updateWorkbookComment", annotationId: root.annotationId, changes: { status: "needs_reanchor" } }, back]).changed).toBe(true);
  expect(s.getAnnotations()[0]!.status).toBe("needs_reanchor"); expect(binding(s, root.annotationId)).toMatchObject({ recoveryStatus: "resolved" });
});
