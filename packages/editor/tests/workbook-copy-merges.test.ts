import { afterEach, expect, test } from "vitest";
import { createDocumentSession, createWorkbook, openDocument } from "@doffi/core";
import type { DocumentOperation, WorkbookSession } from "@doffi/core";
import { connectWorkbookSession, createWorkbookFillTransaction, createWorkbookRangeCopyTransaction, serveWorkbookSession } from "../src/index.js";
import { workbookFillSeedCount } from "../src/workbook/fill-selection.js";

const range = (row: number, column: number, lastRow = row, lastColumn = column) => ({ from: { row, column }, to: { row: lastRow, column: lastColumn } });
const cleanups: (() => void)[] = [];
afterEach(() => { for (const cleanup of cleanups.splice(0).reverse()) cleanup(); });
function apply(session: WorkbookSession, operations: DocumentOperation[]) {
  const manifest = session.snapshot().manifest;
  return session.commit(session.prepare({ protocolVersion: 1, documentId: manifest.documentId, baseRevision: manifest.revision, operations }));
}
function workbook() {
  const session = createDocumentSession(openDocument(createWorkbook({ documentId: "copy-merges" }), { kind: "workbook" })); cleanups.push(() => session.dispose());
  apply(session, [{ op: "setWorkbookCells", sheetId: "sheet-1", cells: [
    { row: 255, column: 31, source: { kind: "richText", html: '<span class="cell-default">合并&#13;😀禰󠄀</span>' } },
    { row: 257, column: 31, source: { kind: "formula", expression: "=LEN(AF256)+$A$1" } },
  ] }, { op: "mergeWorkbookCells", sheetId: "sheet-1", range: range(255, 31, 256, 32) },
  { op: "insertWorkbookSheet", name: "Other" }]); session.clearHistory(); return session;
}

test("fill seed bands close across staggered merges and ignore unrelated columns", () => {
  const merges = [range(2, 4, 3, 5), range(1, 2, 2, 3), range(0, 0, 1, 1), range(1, 6, 8, 7)];
  expect(workbookFillSeedCount(range(0, 0, 11, 5), merges, "row")).toBe(4);
  const transpose = (value: ReturnType<typeof range>) => ({ from: { row: value.from.column, column: value.from.row }, to: { row: value.to.column, column: value.to.row } });
  expect(workbookFillSeedCount(transpose(range(0, 0, 11, 5)), merges.map(transpose), "column")).toBe(4);
  expect(() => workbookFillSeedCount(range(0, 0, 2, 5), merges, "row")).toThrowError(expect.objectContaining({ code: "InvalidOperation" }));
  expect(() => workbookFillSeedCount(range(0, 0, 11, 0), merges, "row")).toThrowError(expect.objectContaining({ code: "InvalidOperation" }));
});

test.each([false, true])("complete cross-chunk merged copy keeps rich source and formula rebasing with one undo (cross-sheet %s)", async crossSheet => {
  const session = workbook(), targetSheetId = crossSheet ? session.snapshot().metadata.workbook.sheets[1]!.sheetId : "sheet-1";
  const before = session.serialize(), source = session.getCell("sheet-1", 255, 31)!;
  const transaction = await createWorkbookRangeCopyTransaction(session, "sheet-1", range(255, 31, 257, 32), { row: 511, column: 63 }, { targetSheetId });
  expect(session.serialize()).toEqual(before); apply(session, transaction.operations);
  expect(session.getCell(targetSheetId, 511, 63)).toMatchObject({ value: source.value, text: source.text });
  expect(session.getCell(targetSheetId, 511, 63)!.html).toContain('<span class="cell-default">合并&#13;😀禰󠄀</span>');
  expect(session.getCell(targetSheetId, 513, 63)!.formula!.expression).toBe("=(LEN(BL512)+$A$1)");
  expect(session.snapshot().metadata.workbook.sheets.find(sheet => sheet.sheetId === targetSheetId)!.merges).toContainEqual(range(511, 63, 512, 64));
  expect(openDocument(session.serialize(), { kind: "workbook" }).getCell(targetSheetId, 511, 63)!.text).toBe(source.text);
  expect(session.history().undoSteps).toBe(1); session.undo(); expect(session.snapshot().manifest.chunks).toEqual(openDocument(before, { kind: "workbook" }).snapshot().manifest.chunks);
  expect(session.snapshot().metadata).toEqual(openDocument(before, { kind: "workbook" }).snapshot().metadata);
  session.redo(); expect(session.getCell(targetSheetId, 511, 63)!.text).toBe(source.text);
});

test("copy uses immutable overlapping merged seeds and identical copies remain no-ops", async () => {
  const session = workbook();
  apply(session, [{ op: "mergeWorkbookCells", sheetId: "sheet-1", range: range(257, 31, 258, 32) }]); session.clearHistory();
  const before = session.serialize();
  apply(session, (await createWorkbookRangeCopyTransaction(session, "sheet-1", range(255, 31, 258, 32), { row: 255, column: 31 })).operations);
  expect(session.history().undoSteps).toBe(0); expect(session.serialize()).toEqual(before);
  const copied = await createWorkbookRangeCopyTransaction(session, "sheet-1", range(255, 31, 258, 32), { row: 257, column: 31 });
  // The existing destination merge is completely covered; captured sources are
  // read before any unmerge or cell write can affect an overlapping seed.
  apply(session, copied.operations);
  expect(session.snapshot().metadata.workbook.sheets[0]!.merges).toEqual([range(255, 31, 256, 32), range(257, 31, 258, 32), range(259, 31, 260, 32)]);
  expect(session.getCell("sheet-1", 257, 31)!.text).toBe("合并\r😀禰󠄀");
  expect(session.getCell("sheet-1", 259, 31)!.formula!.expression).toBe("=(LEN(AF258)+$A$1)");
  session.undo(); expect(session.snapshot().metadata).toEqual(openDocument(before, { kind: "workbook" }).snapshot().metadata);
});

test("partial merges, unsupported selective modes, failed validation, cancellation and conflicts cannot publish topology", async () => {
  const session = workbook(), before = session.serialize();
  await expect(createWorkbookRangeCopyTransaction(session, "sheet-1", range(255, 31), { row: 0, column: 0 })).rejects.toMatchObject({ code: "InvalidOperation" });
  await expect(createWorkbookRangeCopyTransaction(session, "sheet-1", range(0, 0), { row: 255, column: 31 })).rejects.toMatchObject({ code: "InvalidOperation" });
  for (const options of [{ mode: "values" as const }, { mode: "formats" as const }, { mode: "formulas" as const }, { skipBlanks: true }]) {
    await expect(createWorkbookRangeCopyTransaction(session, "sheet-1", range(255, 31, 256, 32), { row: 0, column: 0 }, options)).rejects.toMatchObject({ code: "InvalidOperation" });
  }
  await expect(createWorkbookRangeCopyTransaction(session, "sheet-1", range(255, 31, 512, 32), { row: 1000, column: 0 }, { signal: { aborted: true } })).rejects.toMatchObject({ code: "OperationCancelled" });
  await expect(createWorkbookRangeCopyTransaction(session, "sheet-1", range(255, 31, 256, 32), { row: 0, column: 0 }, { maxCodeUnits: 1 })).rejects.toMatchObject({ code: "ResourceLimitExceeded" });
  expect(session.serialize()).toEqual(before); expect(session.history().undoSteps).toBe(0);
  apply(session, [{ op: "mergeWorkbookCells", sheetId: "sheet-1", range: range(0, 0, 1, 0) }]);
  const incompatible = session.serialize();
  await expect(createWorkbookRangeCopyTransaction(session, "sheet-1", range(255, 31, 256, 32), { row: 0, column: 0 })).rejects.toMatchObject({ code: "InvalidOperation" });
  await expect(createWorkbookFillTransaction(session, "sheet-1", range(255, 31, 256, 32), range(0, 0, 1, 1))).rejects.toMatchObject({ code: "InvalidOperation" });
  expect(session.serialize()).toEqual(incompatible);
  session.undo(); session.clearHistory();
  const generated = await createWorkbookRangeCopyTransaction(session, "sheet-1", range(255, 31, 256, 32), { row: 0, column: 0 });
  apply(session, [{ op: "setWorkbookDataValidation", sheetId: "sheet-1", rule: { ruleId: "stop", kind: "textLength", operator: "lessThan", origin: { row: 0, column: 0 }, ranges: [range(0, 0)], formulas: ["=1"], allowBlank: true, showInputMessage: false, showErrorMessage: true, errorStyle: "stop", errorTitle: "", error: "", promptTitle: "", prompt: "" } }]);
  expect(() => session.prepare(generated)).toThrowError(expect.objectContaining({ code: "RevisionConflict" }));
  const protectedSource = session.serialize(), blocked = await createWorkbookRangeCopyTransaction(session, "sheet-1", range(255, 31, 256, 32), { row: 0, column: 0 });
  expect(() => session.prepare(blocked)).toThrowError(expect.objectContaining({ code: "InvalidOperation" })); expect(session.serialize()).toEqual(protectedSource);
});

test("the sole Worker owns merged copy candidates and discards their structure without touching source", async () => {
  const session = workbook(), channel = new MessageChannel(), server = serveWorkbookSession(channel.port1, session); cleanups.push(() => server.dispose());
  const client = await connectWorkbookSession(channel.port2); cleanups.push(() => client.dispose());
  const before = session.serialize(), command = { kind: "copy" as const, documentId: client.state().documentId, revision: client.state().revision, sheetId: "sheet-1", source: range(255, 31, 256, 32), target: { row: 0, column: 0 } };
  const candidate = await client.prepare(command), view = await client.view(candidate);
  expect(view.cells.snapshot.metadata.workbook.sheets[0]!.merges).toContainEqual(range(0, 0, 1, 1)); view.dispose();
  expect(session.serialize()).toEqual(before); await client.discard(candidate); expect(session.serialize()).toEqual(before);
  await client.commit(await client.prepare(command)); expect(session.history().undoSteps).toBe(1);
  await client.undo(); expect(session.snapshot().manifest.chunks).toEqual(openDocument(before, { kind: "workbook" }).snapshot().manifest.chunks);
});

test.each([253, 257])("fill repeats merged rich seeds in both directions without reading generated cells (row %s)", async row => {
  const session = workbook(), before = session.serialize(), target = range(row, 31, row + 3, 32);
  const transaction = await createWorkbookFillTransaction(session, "sheet-1", range(255, 31, 256, 32), target);
  expect(session.serialize()).toEqual(before); apply(session, transaction.operations);
  for (const offset of [0, 2]) {
    expect(session.snapshot().metadata.workbook.sheets[0]!.merges).toContainEqual(range(row + offset, 31, row + offset + 1, 32));
    expect(session.getCell("sheet-1", row + offset, 31)!.text).toBe("合并\r😀禰󠄀");
  }
  expect(session.history().undoSteps).toBe(1); session.undo();
  expect(session.snapshot().metadata).toEqual(openDocument(before, { kind: "workbook" }).snapshot().metadata);
  expect(session.snapshot().manifest.chunks).toEqual(openDocument(before, { kind: "workbook" }).snapshot().manifest.chunks);
});

test("fill rejects clipped repeated merges and cancels during topology preparation before publishing any cells", async () => {
  const session = workbook(), before = session.serialize();
  await expect(createWorkbookFillTransaction(session, "sheet-1", range(255, 31, 256, 32), range(257, 31, 259, 32))).rejects.toMatchObject({ code: "InvalidOperation" });
  await expect(createWorkbookFillTransaction(session, "sheet-1", range(255, 31, 256, 32), range(257, 31, 258, 32), { format: "preserve" })).rejects.toMatchObject({ code: "InvalidOperation" });
  const signal = { aborted: false }; let yields = 0;
  await expect(createWorkbookFillTransaction(session, "sheet-1", range(255, 31, 256, 32), range(257, 31, 1280, 32), { signal, yield: async () => { yields++; signal.aborted = true; } })).rejects.toMatchObject({ code: "OperationCancelled" });
  expect(yields).toBe(1); expect(session.serialize()).toEqual(before); expect(session.history().undoSteps).toBe(0);
});

test("Worker fill publishes repeated merged topology through the same candidate and history", async () => {
  const session = workbook(), channel = new MessageChannel(), server = serveWorkbookSession(channel.port1, session); cleanups.push(() => server.dispose());
  const client = await connectWorkbookSession(channel.port2); cleanups.push(() => client.dispose()); const before = session.serialize();
  const candidate = await client.prepare({ kind: "fill", documentId: client.state().documentId, revision: client.state().revision, sheetId: "sheet-1", source: range(255, 31, 256, 32), target: range(257, 31, 260, 32) });
  expect(session.serialize()).toEqual(before); await client.commit(candidate);
  expect(session.snapshot().metadata.workbook.sheets[0]!.merges).toEqual([range(255, 31, 256, 32), range(257, 31, 258, 32), range(259, 31, 260, 32)]);
  await client.undo(); expect(session.snapshot().manifest.chunks).toEqual(openDocument(before, { kind: "workbook" }).snapshot().manifest.chunks);
});
