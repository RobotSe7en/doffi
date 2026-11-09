import { afterEach, expect, test } from "vitest";
import { createDocumentSession, createWorkbook, openDocument } from "@doffi/core";
import type { DocumentOperation, WorkbookSession } from "@doffi/core";
import { createWorkbookRangeCopyTransaction } from "../src/index.js";

const sessions: WorkbookSession[] = [], sheetId = "sheet-1";
afterEach(() => { for (const session of sessions.splice(0)) session.dispose(); });
const range = (row: number, column: number, lastRow = row, lastColumn = column) => ({ from: { row, column }, to: { row: lastRow, column: lastColumn } });
function commit(s: WorkbookSession, operations: DocumentOperation[]) { const m = s.snapshot().manifest; s.commit(s.prepare({ protocolVersion: 1, documentId: m.documentId, baseRevision: m.revision, operations })); }
function setup() {
  const s = createDocumentSession(openDocument(createWorkbook({ documentId: "copy-comment-editor" }), { kind: "workbook" })); sessions.push(s);
  commit(s, [{ op: "setWorkbookCells", sheetId, cells: [{ row: 0, column: 0, source: { kind: "value", value: { type: "number", value: 1 } } }] }]);
  for (const [row, column, format, text] of [[0, 0, "modern", "source😀"], [0, 1, "classic", "blank-note" + "x".repeat(20_000)], [2, 0, "modern", "target-A"], [2, 1, "modern", "target-B"]] as const)
    commit(s, [{ op: "addWorkbookComment", anchor: { kind: "cells", sheetId, ranges: [range(row, column)] }, format, html: `<p>${text}</p>` }]);
  s.clearHistory(); return s;
}
function comments(s: WorkbookSession, row: number) {
  return s.getAnnotations().filter(item => item.anchor.kind === "cells" && item.anchor.ranges[0]!.from.row === row).map(item => ({ id: item.annotationId, text: s.getNode(item.bodyNodeId!)!.text }));
}
test("ordinary range copy clones both native comment formats and replaces destination comments in the same cell transaction", async () => {
  const s = setup(), before = s.snapshot(), originals = comments(s, 0), oldTargets = comments(s, 2);
  const tx = await createWorkbookRangeCopyTransaction(s, sheetId, range(0, 0, 0, 1), { row: 2, column: 0 });
  expect(tx.operations[0]?.op).toBe("copyWorkbookComments"); expect(s.snapshot()).toEqual(before);
  s.commit(s.prepare(tx)); expect(comments(s, 2).map(item => item.text)).toEqual(originals.map(item => item.text));
  expect(comments(s, 2).every(item => ![...originals, ...oldTargets].some(old => old.id === item.id))).toBe(true);
  expect(s.getCell(sheetId, 2, 0)?.value).toEqual({ type: "number", value: 1 }); expect(s.history().undoSteps).toBe(1);
  const formats = s.snapshot().metadata.workbook.comments!.filter(item => comments(s, 2).some(comment => comment.id === item.annotationId)).map(item => item.format);
  expect(formats).toEqual(["modern", "classic"]);
  s.undo(); expect(s.snapshot().metadata).toEqual(before.metadata); expect(s.snapshot().content).toEqual(before.content); expect(s.snapshot().manifest.chunks).toEqual(before.manifest.chunks);
});
test.each(["values", "formulas", "formats"] as const)("%s copies retain destination comment bodies and identities", async mode => {
  const s = setup(), before = comments(s, 2);
  const tx = await createWorkbookRangeCopyTransaction(s, sheetId, range(0, 0, 0, 1), { row: 2, column: 0 }, { mode });
  expect(tx.operations.some(op => op.op === "copyWorkbookComments")).toBe(false); s.commit(s.prepare(tx)); expect(comments(s, 2)).toEqual(before);
});
test("skip-blanks applies the same exact mask to comments without charging skipped large bodies", async () => {
  const s = setup(), before = comments(s, 2);
  const tx = await createWorkbookRangeCopyTransaction(s, sheetId, range(0, 0, 0, 1), { row: 2, column: 0 }, { skipBlanks: true, maxCodeUnits: 2000 });
  expect(tx.operations[0]).toMatchObject({ op: "copyWorkbookComments", ranges: [range(2, 0)] });
  s.commit(s.prepare(tx)); expect(comments(s, 2).map(item => item.text)).toEqual(["target-B", "source😀"]);
  expect(comments(s, 2).find(item => item.text === "target-B")!.id).toBe(before[1]!.id);
});
test("cancelled or stale range-copy preparation publishes no cloned comments or stories", async () => {
  const s = setup(), before = s.snapshot(), signal = { aborted: false };
  await expect(createWorkbookRangeCopyTransaction(s, sheetId, range(0, 0, 0, 1), { row: 2, column: 0 }, { signal, onProgress: () => { signal.aborted = true; } })).rejects.toMatchObject({ code: "OperationCancelled" });
  expect(s.snapshot()).toEqual(before);
  const tx = await createWorkbookRangeCopyTransaction(s, sheetId, range(0, 0, 0, 1), { row: 2, column: 0 }), candidate = s.prepare(tx);
  commit(s, [{ op: "setWorkbookCells", sheetId, cells: [{ row: 9, column: 0, source: { kind: "value", value: { type: "number", value: 2 } } }] }]);
  expect(() => s.commit(candidate)).toThrowError(expect.objectContaining({ code: "RevisionConflict" }));
  expect(s.snapshot().content).toEqual(before.content); expect(s.snapshot().metadata.workbook.comments).toEqual(before.metadata.workbook.comments);
});
