import { afterEach, expect, test } from "vitest";
import { createDocumentSession } from "../src/index.js";
import type { DocumentOperation, WorkbookConditionalRule, WorkbookEditProposal, WorkbookOperation, WorkbookSession } from "../src/index.js";
import { scalarWorkbook } from "./workbook-scalar-fixture.js";
const sessions: WorkbookSession[] = [];
afterEach(() => { for (const session of sessions.splice(0)) session.dispose(); });
const range = (row: number, column: number, lastRow = row, lastColumn = column) => ({ from: { row, column }, to: { row: lastRow, column: lastColumn } });
const rule = (ruleId: string, ranges = [range(1, 1)]): Extract<WorkbookConditionalRule, { kind: "expression" }> => ({ ruleId, ranges, origin: { row: 0, column: 0 }, kind: "expression", expression: "=TRUE()", priority: 1, stopIfTrue: true });
function fixture() { const s = createDocumentSession(scalarWorkbook([], metadata => { metadata.workbook.sheets[0]!.conditionalRules = [rule("old", [range(0, 0, 5, 5)])]; }), { sessionLimits: { maxCandidates: 1 } }); sessions.push(s); return s; }
const tx = (s: WorkbookSession, operations: readonly DocumentOperation[]) => ({ protocolVersion: 1 as const, documentId: s.snapshot().manifest.documentId, baseRevision: s.snapshot().manifest.revision, operations: [...operations] });
const apply = (s: WorkbookSession, operations: readonly DocumentOperation[]) => s.commit(s.prepare(tx(s, operations)));
test("regional rules clip retained templates, preserve empty categories and undo atomically", () => {
  const s = fixture(), before = s.snapshot();
  apply(s, [{ op: "replaceWorkbookRules", sheetId: "sheet-1", ranges: [range(1, 1, 4, 4)], conditionalRules: [{ ...rule("new", [range(2, 2)]), priority: 2 }] }]);
  const sheet = s.snapshot().metadata.workbook.sheets[0]!;
  expect(sheet.conditionalRules![0]).toMatchObject({ ruleId: "old", origin: { row: 0, column: 0 }, expression: "=TRUE()", priority: 1 });
  expect(sheet.conditionalRules![0]!.ranges).toHaveLength(4); expect(sheet.conditionalRules![1]!.ruleId).toBe("new"); expect(sheet.dataValidations).toBeUndefined();
  expect(s.history().undoSteps).toBe(1); s.undo(); expect(s.snapshot().metadata.workbook).toEqual(before.metadata.workbook);
});
test("malformed footprints and out-of-footprint rules reject preceding source edits", () => {
  const s = fixture(), before = s.serialize();
  const variants: WorkbookOperation[] = [
    { op: "replaceWorkbookRules", sheetId: "sheet-1", ranges: [range(1, 1)] },
    { op: "replaceWorkbookRules", sheetId: "sheet-1", ranges: [range(2, 2, 1, 1)], conditionalRules: [] },
    { op: "replaceWorkbookRules", sheetId: "sheet-1", ranges: [range(1, 1), range(1, 1)], conditionalRules: [] },
    { op: "replaceWorkbookRules", sheetId: "sheet-1", ranges: [range(1, 1)], conditionalRules: [rule("new", [range(1, 1, 2, 1)])] },
  ];
  for (const operation of variants) { expect(() => apply(s, [{ op: "setWorkbookCells", sheetId: "sheet-1", cells: [{ row: 20, column: 0, source: { kind: "value", value: { type: "number", value: 99 } } }] }, operation])).toThrow(); expect(s.serialize()).toEqual(before); }
});
test("large sparse rule batches yield, cancel and release the sole candidate lease", async () => {
  const s = fixture(), before = s.serialize(), rules = Array.from({ length: 1500 }, (_, row) => ({ ...rule("new-" + row, [range(row, 10)]), priority: row + 2 }));
  const operation: WorkbookOperation = { op: "replaceWorkbookRules", sheetId: "sheet-1", ranges: [range(0, 10, 1499, 10)], conditionalRules: rules };
  const abort = new AbortController(); await expect(s.prepareAsync(tx(s, [operation]), {}, { signal: abort.signal, yield: async () => {}, onProgress(progress) { if (progress.phase === "cells" && progress.completed >= 256) abort.abort(); } })).rejects.toThrowError(expect.objectContaining({ code: "OperationCancelled" })); expect(s.serialize()).toEqual(before);
  s.commit(await s.prepareAsync(tx(s, [operation]))); expect(s.snapshot().metadata.workbook.sheets[0]!.conditionalRules).toHaveLength(1501); expect(s.history().undoSteps).toBe(1);
});
test("scoped proposals may clip a crossing rule but cannot change its outside template or priority", () => {
  const s = fixture(), m = s.snapshot().manifest, owned = range(1, 1, 4, 4);
  apply(s, [{ op: "addAnnotation", annotation: { annotationId: "review", documentId: m.documentId, baseRevision: m.revision, status: "open", instruction: "复制规则", anchor: { kind: "cells", sheetId: "sheet-1", ranges: [owned] } } }]); s.clearHistory();
  const request = s.buildInteractionRequest({ kind: "workbook", annotationIds: ["review"], scope: { kind: "ranges", ranges: [{ sheetId: "sheet-1", range: owned }] } });
  const proposal = (operations: WorkbookOperation[]): WorkbookEditProposal => ({ protocolVersion: 1, documentId: request.documentId, baseRevision: request.baseRevision, requestId: request.requestId, proposalId: "rules", annotationIds: ["review"], summary: "更新规则", operations });
  expect(() => s.prepareProposal(request, proposal([{ op: "replaceWorkbookRules", sheetId: "sheet-1", ranges: [range(0, 0, 5, 5)], conditionalRules: [] }]))).toThrowError(expect.objectContaining({ code: "InvalidProposal" }));
  expect(() => s.prepareProposal(request, proposal([{ op: "setWorkbookConditionalRule", sheetId: "sheet-1", rule: { ...rule("old", [range(0, 0, 5, 5)]), expression: "=FALSE()" } }]))).toThrowError(expect.objectContaining({ code: "InvalidProposal" }));
  const candidate = s.prepareProposal(request, proposal([{ op: "replaceWorkbookRules", sheetId: "sheet-1", ranges: [owned], conditionalRules: [{ ...rule("new", [range(2, 2)]), priority: 2 }] }]));
  s.commit(candidate); expect(s.snapshot().metadata.workbook.sheets[0]!.conditionalRules).toHaveLength(2); expect(s.getAnnotations()[0]!.status).toBe("resolved"); s.undo(); expect(s.snapshot().metadata.workbook.sheets[0]!.conditionalRules).toHaveLength(1);
});
