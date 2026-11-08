import { afterEach, expect, test } from "vitest";
import { createDocumentSession } from "../src/index.js";
import type { DocumentOperation, WorkbookEditProposal, WorkbookHyperlink, WorkbookOperation, WorkbookSession } from "../src/index.js";
import { scalarWorkbook } from "./workbook-scalar-fixture.js";
const sessions: WorkbookSession[] = [];
afterEach(() => { for (const session of sessions.splice(0)) session.dispose(); });
const range = (row: number, column: number, lastRow = row, lastColumn = column) => ({ from: { row, column }, to: { row: lastRow, column: lastColumn } });
function fixture() { const s = createDocumentSession(scalarWorkbook([], metadata => { metadata.workbook.sheets[0]!.hyperlinks = [{ range: range(0, 0, 5, 5), url: "https://example.org/old", location: "Sheet1!A1", display: "显示\r\n禰󠄀😀", tooltip: "提示\r\n尾" }]; }), { sessionLimits: { maxCandidates: 1 } }); sessions.push(s); return s; }
const links = (s: WorkbookSession) => s.snapshot().metadata.workbook.sheets[0]!.hyperlinks!;
const transaction = (s: WorkbookSession, operations: readonly DocumentOperation[]) => ({ protocolVersion: 1 as const, documentId: s.snapshot().manifest.documentId, baseRevision: s.snapshot().manifest.revision, operations: [...operations] });
const apply = (s: WorkbookSession, operations: readonly DocumentOperation[]) => s.commit(s.prepare(transaction(s, operations)));
test("one regional replacement keeps all outside fragments, source text attributes and one history step", () => {
  const s = fixture(), original = links(s)[0]!, before = s.snapshot();
  apply(s, [{ op: "replaceWorkbookHyperlinks", sheetId: "sheet-1", range: range(1, 1, 4, 4), links: [{ range: range(2, 2, 3, 3), url: "mailto:new@example.org", display: "新😀\r\n文本", tooltip: "新提示" }] }]);
  expect(links(s)).toHaveLength(5); for (const value of links(s).slice(0, 4)) expect(value).toMatchObject({ url: original.url, location: original.location, display: original.display, tooltip: original.tooltip });
  expect(links(s).at(-1)).toMatchObject({ url: "mailto:new@example.org", display: "新😀\r\n文本" }); expect(s.snapshot().manifest.chunks).toEqual(before.manifest.chunks); expect(s.history().undoSteps).toBe(1);
  s.undo(); expect(links(s)).toEqual(before.metadata.workbook.sheets[0]!.hyperlinks); s.redo(); expect(links(s)).toHaveLength(5);
});
test("semantically equal fragments retain original source bytes and no history", () => {
  const s = fixture(), original = links(s)[0]!, before = s.serialize();
  const { range: originalRange, ...target } = original;
  expect(apply(s, [{ op: "replaceWorkbookHyperlinks", sheetId: "sheet-1", range: originalRange, links: [{ ...target, range: range(0, 0, 2, 5) }, { ...target, range: range(3, 0, 5, 5) }] }]).changed).toBe(false);
  expect(s.serialize()).toEqual(before); expect(s.history().undoSteps).toBe(0);
});
test("out-of-scope, overlapping and unsafe targets reject preceding cell writes atomically", () => {
  const s = fixture(), before = s.serialize();
  for (const values of [[{ range: range(0, 0), url: "https://example.org" }], [{ range: range(1, 1, 2, 2), url: "https://example.org" }, { range: range(2, 2), location: "A1" }], [{ range: range(1, 1), url: "javascript:alert(1)" }]]) {
    expect(() => apply(s, [{ op: "setWorkbookCells", sheetId: "sheet-1", cells: [{ row: 20, column: 0, source: { kind: "value", value: { type: "number", value: 99 } } }] }, { op: "replaceWorkbookHyperlinks", sheetId: "sheet-1", range: range(1, 1, 3, 3), links: values }])).toThrow();
    expect(s.serialize()).toEqual(before); expect(s.history().undoSteps).toBe(0);
  }
});
test("thousands of links use one bounded command, support cancellation and release the sole candidate", async () => {
  const s = fixture(), before = s.serialize(), values: WorkbookHyperlink[] = Array.from({ length: 2048 }, (_, row) => ({ range: range(row, 10), url: "https://example.org/" + row }));
  const tx = transaction(s, [{ op: "replaceWorkbookHyperlinks", sheetId: "sheet-1", range: range(0, 10, 2047, 10), links: values }]), abort = new AbortController();
  await expect(s.prepareAsync(tx, {}, { signal: abort.signal, yield: async () => {}, onProgress(progress) { if (progress.phase === "cells" && progress.completed >= 256) abort.abort(); } })).rejects.toThrowError(expect.objectContaining({ code: "OperationCancelled" })); expect(s.serialize()).toEqual(before);
  s.commit(await s.prepareAsync(tx)); expect(links(s)).toHaveLength(2049); expect(s.history().undoSteps).toBe(1); s.undo(); expect(links(s)).toHaveLength(1);
});

test("range proposals replace only owned hyperlink coordinates and retain the outside source complement", () => {
  const s = fixture(), manifest = s.snapshot().manifest, owned = range(1, 1, 4, 4);
  apply(s, [{ op: "addAnnotation", annotation: { annotationId: "review", documentId: manifest.documentId, baseRevision: manifest.revision, status: "open", instruction: "复制链接", anchor: { kind: "cells", sheetId: "sheet-1", ranges: [owned] } } }]); s.clearHistory();
  const request = s.buildInteractionRequest({ kind: "workbook", annotationIds: ["review"], scope: { kind: "ranges", ranges: [{ sheetId: "sheet-1", range: owned }] } });
  const proposal = (operations: WorkbookOperation[]): WorkbookEditProposal => ({ protocolVersion: 1, documentId: request.documentId, baseRevision: request.baseRevision, requestId: request.requestId, proposalId: "link-proposal", annotationIds: ["review"], summary: "更新链接", operations });
  const original = s.serialize(); expect(() => s.prepareProposal(request, proposal([{ op: "replaceWorkbookHyperlinks", sheetId: "sheet-1", range: range(0, 0, 5, 5), links: [] }]))).toThrowError(expect.objectContaining({ code: "InvalidProposal" })); expect(s.serialize()).toEqual(original);
  const candidate = s.prepareProposal(request, proposal([{ op: "replaceWorkbookHyperlinks", sheetId: "sheet-1", range: owned, links: [{ range: range(2, 2), url: "https://example.org/new" }] }]));
  expect(candidate.affectedSheetIds).toContain("sheet-1"); s.commit(candidate); expect(links(s)).toHaveLength(5); expect(s.getAnnotations()[0]!.status).toBe("resolved"); expect(s.history().undoSteps).toBe(1); s.undo(); expect(s.getAnnotations()[0]!.status).toBe("open"); expect(links(s)).toHaveLength(1);
});
