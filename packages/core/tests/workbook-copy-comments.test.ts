import { afterEach, expect, test } from "vitest";
import { createDocumentSession, createWorkbook, openDocument } from "../src/index.js";
import type { DocumentOperation, TransactionPrepareOptions, WorkbookLayoutMeasurement, WorkbookLayoutRequest, WorkbookSession } from "../src/index.js";

const sessions: WorkbookSession[] = [], sheetId = "sheet-1";
const range = (row: number, column: number, lastRow = row, lastColumn = column) => ({ from: { row, column }, to: { row: lastRow, column: lastColumn } });
const anchor = (row: number, column: number, lastRow = row, lastColumn = column) => ({ kind: "cells" as const, sheetId, ranges: [range(row, column, lastRow, lastColumn)] });
afterEach(() => { for (const session of sessions.splice(0)) session.dispose(); });
function blank() { const s = createDocumentSession(openDocument(createWorkbook({ documentId: "copy-comments" }), { kind: "workbook" })); sessions.push(s); return s; }
function transaction(s: WorkbookSession, operations: DocumentOperation[]) { const m = s.snapshot().manifest; return { protocolVersion: 1 as const, documentId: m.documentId, baseRevision: m.revision, operations }; }
function commit(s: WorkbookSession, operations: DocumentOperation[], options?: TransactionPrepareOptions) { return s.commit(s.prepare(transaction(s, operations), options)); }
const operation = (fromRow: number, toRow: number, source = range(fromRow, 0)): DocumentOperation => ({ op: "copyWorkbookComments", sheetId, range: source, targetSheetId: sheetId, target: { row: toRow, column: 0 } });
const metrics = (request: WorkbookLayoutRequest): WorkbookLayoutMeasurement => ({ sourceDigest: request.sourceDigest, fontEnvironment: "explicit-test-metrics", sheets: request.sheetIds.map(id => ({ sheetId: id,
  rows: { defaultSizePt: 10, overrides: [{ from: 2, to: 100, sizePt: 20 }] }, columns: { defaultSizePt: 20, overrides: [{ from: 2, to: 100, sizePt: 30 }] } })) });

test("overlapping comment copies freeze every original body before clearing destinations and restore exact source in one undo", () => {
  const s = blank();
  for (const row of [0, 1]) commit(s, [{ op: "addWorkbookComment", anchor: anchor(row, 0), html: `<p>thread-${row}禰󠄀😀</p>`, author: { name: "王" } }]);
  s.clearHistory(); const before = s.snapshot(), original = s.getAnnotations(), texts = original.map(item => s.getNode(item.bodyNodeId!)!.text);
  const candidate = s.prepare(transaction(s, [operation(0, 1, range(0, 0, 1, 0))])); expect(s.snapshot()).toEqual(before); s.commit(candidate);
  const after = s.getAnnotations(); expect(after).toHaveLength(3);
  for (const row of [1, 2]) {
    const item = after.find(item => item.anchor.kind === "cells" && item.anchor.ranges[0]!.from.row === row)!;
    expect(s.getNode(item.bodyNodeId!)!.text).toBe(texts[row - 1]); expect(original.some(old => old.annotationId === item.annotationId || old.bodyNodeId === item.bodyNodeId)).toBe(false);
  }
  expect(s.getNode(original[1]!.bodyNodeId!)).toBeUndefined(); expect(s.history().undoSteps).toBe(1);
  expect(openDocument(s.serialize(), { kind: "workbook" }).snapshot()).toEqual(s.snapshot());
  s.undo(); expect(s.snapshot().metadata).toEqual(before.metadata); expect(s.snapshot().annotations).toMatchObject(original.map(item => ({ ...item, baseRevision: s.snapshot().manifest.revision })));
  expect(s.snapshot().content).toEqual(before.content);
});
test("copied native reply graphs, display authors, exact Unicode mentions and resolved state receive independent identities", () => {
  const s = blank(); commit(s, [{ op: "addWorkbookComment", anchor: anchor(0, 0), html: "<p>王😀禰󠄀</p>", author: { name: "王" } }]);
  const root = s.getAnnotations()[0]!, author = s.snapshot().metadata.workbook.commentAuthors![0]!;
  commit(s, [{ op: "addWorkbookCommentReply", annotationId: root.annotationId, html: "<p>第一回复</p>", author: { name: "李" } },
    { op: "setWorkbookCommentBody", annotationId: root.annotationId, html: "<p>王😀禰󠄀</p>", mentions: [{ mentionId: "original-mention", authorId: author.authorId, startIndex: 0, length: 1 }] }]);
  const first = s.getAnnotations()[0]!.replies![0]!;
  commit(s, [{ op: "addWorkbookCommentReply", annotationId: root.annotationId, parentReplyId: first.replyId, html: "<p>子回复😀</p>" }, { op: "updateWorkbookComment", annotationId: root.annotationId, changes: { status: "resolved" } }]);
  s.clearHistory(); const before = s.snapshot(); commit(s, [operation(0, 4)]);
  const copied = s.getAnnotations().at(-1)!, binding = s.snapshot().metadata.workbook.comments!.find(value => value.annotationId === copied.annotationId)!;
  expect(copied).toMatchObject({ status: "resolved", author: { name: "王" }, anchor: anchor(4, 0) });
  expect(binding).toMatchObject({ format: "modern", mentions: [{ authorId: author.authorId, startIndex: 0, length: 1 }], replies: [{ replyId: copied.replies![0]!.replyId }, { replyId: copied.replies![1]!.replyId, parentReplyId: copied.replies![0]!.replyId }] });
  if (binding.format === "modern") expect(binding.mentions![0]!.mentionId).not.toBe("original-mention");
  expect(copied.replies!.every(reply => !s.getAnnotations()[0]!.replies!.some(old => old.replyId === reply.replyId || old.bodyNodeId === reply.bodyNodeId))).toBe(true);
  expect(s.getNode(copied.replies![1]!.bodyNodeId!)!.text).toBe("子回复😀"); s.undo(); expect(s.snapshot().metadata).toEqual(before.metadata);
});
test("exact destination masks remove only copied comment fragments, including blank source cells, and same-position copies are no-ops", () => {
  const s = blank(); commit(s, [{ op: "addWorkbookComment", anchor: anchor(3, 0, 4, 1), html: "<p>retained</p>" }]); s.clearHistory();
  const before = s.snapshot(), id = s.getAnnotations()[0]!.annotationId, body = s.getAnnotations()[0]!.bodyNodeId;
  expect(commit(s, [operation(3, 3, range(3, 0, 4, 1))]).changed).toBe(false);
  commit(s, [{ op: "copyWorkbookComments", sheetId, range: range(0, 0, 1, 1), targetSheetId: sheetId, target: { row: 3, column: 0 }, ranges: [range(3, 0)] }]);
  expect(s.getAnnotations()[0]).toMatchObject({ annotationId: id, bodyNodeId: body, anchor: { kind: "cells", ranges: [range(4, 0, 4, 1), range(3, 1)] } });
  expect(s.getNode(body!)!.text).toBe("retained"); s.undo(); expect(s.snapshot().content).toEqual(before.content);
});
test("classic note copying uses measured positions across unequal axes, keeps fixed sizes, and requires a measurer", () => {
  const s = blank();
  commit(s, [{ op: "addWorkbookComment", anchor: anchor(0, 0), format: "classic", html: "<p>note</p>", box: {
    anchor: { kind: "two-cell", from: { row: 0, column: 1, offsetXPt: 2, offsetYPt: 3 }, to: { row: 2, column: 3, offsetXPt: 4, offsetYPt: 5 }, moveWithCells: false, sizeWithCells: false },
    visible: true, printable: true, locked: false, lockText: false, defaultSize: false, disabled: false, autoFill: false, autoLine: false, autoScale: false, justifyLastLine: false,
  } }]); s.clearHistory(); const before = s.snapshot();
  expect(() => commit(s, [operation(0, 3)])).toThrowError(/layout measurer/); expect(s.snapshot()).toEqual(before);
  commit(s, [operation(0, 3)], { measureWorkbookLayout: metrics });
  const copied = s.getAnnotations().at(-1)!, binding = s.snapshot().metadata.workbook.comments!.find(binding => binding.annotationId === copied.annotationId)!;
  expect(binding).toMatchObject({ format: "classic", box: { anchor: { from: { row: 3, column: 1, offsetXPt: 2, offsetYPt: 3 }, to: { row: 5, column: 3, offsetXPt: 4, offsetYPt: -15 }, moveWithCells: false, sizeWithCells: false } } });
  if (binding.format === "classic" && binding.box?.anchor.kind === "two-cell") expect(80 + binding.box.anchor.to.offsetYPt - (40 + binding.box.anchor.from.offsetYPt)).toBe(22);
  s.undo(); expect(s.snapshot().metadata).toEqual(before.metadata); expect(s.snapshot().content).toEqual(before.content);
});
test("absolute note copies recheck the measured font environment before a private candidate can publish", () => {
  const s = blank(); commit(s, [{ op: "addWorkbookComment", anchor: anchor(0, 0), format: "classic", html: "<p>absolute</p>", box: {
    anchor: { kind: "absolute", xPt: 2, yPt: 3, widthPt: 40, heightPt: 20 }, visible: false, printable: true, locked: false, lockText: false,
    defaultSize: false, disabled: false, autoFill: false, autoLine: false, autoScale: false, justifyLastLine: false,
  } }]); const before = s.serialize(); let calls = 0;
  expect(() => commit(s, [operation(0, 3)], { measureWorkbookLayout: request => ({ ...metrics(request), fontEnvironment: "fonts-" + ++calls }) })).toThrow();
  expect(calls).toBe(2); expect(s.serialize()).toEqual(before);
});
test("range proposals copy only declared source comment bodies and destination ranges", () => {
  const s = blank();
  for (const column of [0, 1]) commit(s, [{ op: "addWorkbookComment", anchor: anchor(0, column), html: `<p>${column ? "undisclosed-secret" : "source"}</p>` }]);
  const m = s.snapshot().manifest;
  commit(s, [{ op: "addAnnotation", annotation: { annotationId: "request", documentId: m.documentId, baseRevision: m.revision, status: "open", instruction: "复制已声明范围的批注", anchor: { kind: "cells", sheetId, ranges: [range(0, 0), range(2, 0)] } } }]);
  const request = s.buildInteractionRequest({ kind: "workbook", annotationIds: ["request"] }), before = s.serialize();
  const proposal = { protocolVersion: 1 as const, documentId: request.documentId, baseRevision: request.baseRevision, requestId: request.requestId, proposalId: "copy", annotationIds: ["request"], summary: "copy comments", operations: [{ op: "copyWorkbookComments" as const, sheetId, range: range(0, 0), targetSheetId: sheetId, target: { row: 2, column: 0 } }] };
  expect(JSON.stringify(request)).not.toContain("undisclosed-secret");
  const candidate = s.prepareProposal(request, proposal); expect(s.serialize()).toEqual(before); expect(candidate.preview.annotations.length).toBeGreaterThan(s.getAnnotations().length); s.discard(candidate);
  expect(() => s.prepareProposal(request, { ...proposal, operations: [{ ...proposal.operations[0]!, range: range(0, 1) }] })).toThrowError(expect.objectContaining({ code: "InvalidProposal" }));
  expect(s.serialize()).toEqual(before);
});
test("invalid comment footprints fail atomically, while empty footprints leave comments and redo untouched", () => {
  const s = blank(); commit(s, [{ op: "addWorkbookComment", anchor: anchor(0, 0), html: "<p>source</p>" }]);
  commit(s, [operation(0, 4)]); s.undo(); const before = s.serialize(), history = s.history();
  for (const ranges of [[range(3, 0)], [range(4, 0), range(4, 0)]]) {
    expect(() => commit(s, [{ op: "copyWorkbookComments", sheetId, range: range(0, 0), targetSheetId: sheetId, target: { row: 4, column: 0 }, ranges }])).toThrow();
    expect(s.serialize()).toEqual(before); expect(s.history()).toEqual(history);
  }
  expect(commit(s, [{ op: "copyWorkbookComments", sheetId, range: range(0, 0), targetSheetId: sheetId, target: { row: 4, column: 0 }, ranges: [] }]).changed).toBe(false);
  expect(s.serialize()).toEqual(before); expect(s.history()).toEqual(history);
});
test("fragmented source threads clone only the exact selected intersections and keep a single independent reply graph", () => {
  const s = blank(); commit(s, [{ op: "addWorkbookComment", anchor: { kind: "cells", sheetId, ranges: [range(0, 0, 2, 0), range(0, 2, 2, 2)] }, html: "<p>fragmented禰󠄀</p>" }]);
  const original = s.getAnnotations()[0]!;
  commit(s, [{ op: "addWorkbookCommentReply", annotationId: original.annotationId, html: "<p>reply😀</p>" }]); s.clearHistory(); const before = s.snapshot();
  commit(s, [{ op: "copyWorkbookComments", sheetId, range: range(0, 0, 2, 2), targetSheetId: sheetId, target: { row: 10, column: 4 }, ranges: [range(10, 4, 11, 4), range(12, 6)] }]);
  expect(s.getAnnotations()).toHaveLength(2); const copied = s.getAnnotations()[1]!;
  expect(copied.anchor).toEqual({ kind: "cells", sheetId, ranges: [range(10, 4, 11, 4), range(12, 6)] });
  expect(copied.replies).toHaveLength(1); expect(s.getNode(copied.replies![0]!.bodyNodeId!)!.text).toBe("reply😀");
  expect(s.history().undoSteps).toBe(1); s.undo(); const restored = s.snapshot();
  expect(restored.content).toEqual(before.content); expect(restored.metadata).toEqual(before.metadata);
});
test("cancelling after private comment bodies were cloned releases the candidate, and stale copy candidates cannot commit", async () => {
  const s = blank(); commit(s, Array.from({ length: 70 }, (_, row) => ({ op: "addWorkbookComment", anchor: anchor(row, 0), html: `<p>source-${row}😀</p>` })));
  s.clearHistory(); const before = s.serialize(), history = s.history(), abort = new AbortController(); let cloned = false;
  const copy = transaction(s, [operation(0, 100, range(0, 0, 69, 0))]);
  await expect(s.prepareAsync(copy, undefined, { signal: abort.signal, yield: async () => {}, onProgress(progress) {
    if (progress.phase === "cells" && progress.completed === 64 && progress.total === 70) { cloned = true; abort.abort(); }
  } })).rejects.toMatchObject({ code: "OperationCancelled" });
  expect(cloned).toBe(true); expect(s.serialize()).toEqual(before); expect(s.history()).toEqual(history);
  const candidate = await s.prepareAsync(copy); expect(s.serialize()).toEqual(before);
  commit(s, [{ op: "setWorkbookCells", sheetId, cells: [{ row: 0, column: 3, source: { kind: "value", value: { type: "number", value: 7 } } }] }]); const edited = s.serialize();
  expect(() => s.commit(candidate)).toThrowError(expect.objectContaining({ code: "RevisionConflict" }));
  expect(s.serialize()).toEqual(edited); expect(s.history().undoSteps).toBe(1); expect(s.getAnnotations()).toHaveLength(70);
});
