import { expect, test } from "vitest";
import { createDocumentSession, createWorkbook, openDocument, readWorkbookCellSource } from "@doffi/core";
import type { WorkbookDocumentMetadata, WorkbookSession } from "@doffi/core";
import { createWorkbookFillTransaction } from "../src/workbook/fill.js";
const number = (value: number) => ({ type: "number" as const, value });
const range = (fromRow: number, fromColumn = 0, toRow = fromRow, toColumn = fromColumn) => ({ from: { row: fromRow, column: fromColumn }, to: { row: toRow, column: toColumn } });
function session(): WorkbookSession { return createDocumentSession(openDocument(createWorkbook({ documentId: "fill" }), { kind: "workbook" })); }
function seed(source: WorkbookSession, values: readonly number[], firstRow = 0): void {
  source.commit(source.prepare({ protocolVersion: 1, documentId: "fill", baseRevision: source.snapshot().manifest.revision,
    operations: [{ op: "setWorkbookCells", sheetId: "sheet-1", cells: values.map((value, index) => ({ row: firstRow + index, column: 0, source: { kind: "value" as const, value: number(value) } })) }] })); source.clearHistory();
}
test("fill copies rich spans, inherited formats and mixed formula references in one cross-chunk undo step", async () => {
  const files = new Map(createWorkbook({ documentId: "fill" })), metadata = JSON.parse(new TextDecoder().decode(files.get("metadata.json")!)) as WorkbookDocumentMetadata;
  metadata.workbook.numberFormats.push({ formatId: "date", code: "yyyy-mm-dd" });
  metadata.workbook.sheets[0]!.rowDimensions = [{ from: 255, to: 255, cssClass: "strong", numberFormatId: "date" }];
  files.set("metadata.json", new TextEncoder().encode(JSON.stringify(metadata)));
  files.set("styles.css", new TextEncoder().encode(new TextDecoder().decode(files.get("styles.css")!) + ".strong{font-weight:700}"));
  const source = createDocumentSession(openDocument(files, { kind: "workbook" }));
  source.commit(source.prepare({ protocolVersion: 1, documentId: "fill", baseRevision: 0, operations: [{ op: "setWorkbookCells", sheetId: "sheet-1", cells: [
    { row: 0, column: 0, source: { kind: "value", value: { type: "text", value: "末尾" } } },
    { row: 255, column: 31, source: { kind: "richText", html: '中文&#13;😀<span class="cell-default">禰󠄀</span><br>' } },
    { row: 255, column: 32, source: { kind: "formula", expression: "=AF256&$A$1" } },
  ] }] })); source.clearHistory(); const before = source.serialize();
  const rich = readWorkbookCellSource(source, { sheetId: "sheet-1", row: 255, column: 31 });
  expect(rich).toEqual({ kind: "richText", html: '中文&#13;😀<span class="cell-default">禰󠄀</span><br>' });
  const transaction = await createWorkbookFillTransaction(source, "sheet-1", range(255, 31, 255, 32), range(256, 31, 256, 32));
  expect(source.serialize()).toEqual(before); expect(source.history().undoSteps).toBe(0); source.commit(source.prepare(transaction));
  expect(source.getCell("sheet-1", 256, 31)).toMatchObject({ cssClass: "strong", formatId: "date", text: "中文\r😀禰󠄀\n" });
  expect(source.getCell("sheet-1", 256, 31)?.html).toContain('<span class="cell-default">禰󠄀</span>');
  expect(source.getCell("sheet-1", 256, 32)?.formula?.expression).toContain("AF257"); expect(source.getCell("sheet-1", 256, 32)?.formula?.expression).toContain("$A$1");
  source.calculateSync(); expect(source.getCalculatedCell("sheet-1", 256, 32)).toMatchObject({ value: { type: "text", value: "中文\r😀禰󠄀\n末尾" } });
  expect(source.history().undoSteps).toBe(1); source.undo(); expect(source.getCell("sheet-1", 256, 31)).toBeUndefined(); expect(source.getCell("sheet-1", 255, 31)?.html).toContain("禰󠄀");
});
test("arithmetic fill extends decimals, reverses direction and never mutates overlapping seeds", async () => {
  const source = session(); seed(source, [.1, .2]);
  const transaction = await createWorkbookFillTransaction(source, "sheet-1", range(0, 0, 1), range(0, 0, 260), { mode: "series" });
  source.commit(source.prepare(transaction)); expect(source.getCell("sheet-1", 260, 0)?.value).toMatchObject({ value: expect.closeTo(26.1, 12) });
  expect(source.getCell("sheet-1", 0, 0)?.value).toEqual(number(.1)); expect(source.getCell("sheet-1", 1, 0)?.value).toEqual(number(.2)); expect(source.history().undoSteps).toBe(1);
  const reverse = session(); seed(reverse, [10, 12], 10);
  reverse.commit(reverse.prepare(await createWorkbookFillTransaction(reverse, "sheet-1", range(10, 0, 11), range(7, 0, 9), { mode: "series" })));
  expect([7, 8, 9].map(row => reverse.getCell("sheet-1", row, 0)?.value)).toEqual([4, 6, 8].map(number));
  const copy = session(); seed(copy, [1, 2]); copy.commit(copy.prepare(await createWorkbookFillTransaction(copy, "sheet-1", range(0, 0, 1), range(1, 0, 4), { format: "preserve" })));
  expect([0, 1, 2, 3, 4].map(row => copy.getCell("sheet-1", row, 0)?.value)).toEqual([1, 2, 1, 2, 1].map(number));
});
test("bad series, merged regions and budgets reject before a fill can publish source", async () => {
  const source = session(); seed(source, [1, 2, 4]); const before = source.serialize();
  await expect(createWorkbookFillTransaction(source, "sheet-1", range(0, 0, 2), range(3, 0, 5), { mode: "series" })).rejects.toMatchObject({ code: "InvalidOperation" });
  await expect(createWorkbookFillTransaction(source, "sheet-1", range(0), range(1, 0, 11), { maxCells: 10 })).rejects.toMatchObject({ code: "ResourceLimitExceeded" });
  await expect(createWorkbookFillTransaction(source, "sheet-1", range(0), range(1, 0, 2), { maxCodeUnits: 1 })).rejects.toMatchObject({ code: "ResourceLimitExceeded" });
  expect(source.serialize()).toEqual(before); expect(source.history().undoSteps).toBe(0);
  const files = new Map(createWorkbook({ documentId: "merged" })), metadata = JSON.parse(new TextDecoder().decode(files.get("metadata.json")!)) as WorkbookDocumentMetadata;
  metadata.workbook.sheets[0]!.merges = [range(2, 0, 3, 1)]; files.set("metadata.json", new TextEncoder().encode(JSON.stringify(metadata)));
  await expect(createWorkbookFillTransaction(openDocument(files, { kind: "workbook" }), "sheet-1", range(0), range(2, 0, 4))).rejects.toMatchObject({ code: "InvalidOperation" });
});
test("progress cancellation and source replacement abandon the entire uncommitted fill", async () => {
  const source = session(); seed(source, [1]); const before = source.serialize(), signal = { aborted: false }, progress: number[] = [];
  await expect(createWorkbookFillTransaction(source, "sheet-1", range(0), range(1, 0, 300), { signal, onProgress: value => progress.push(value.completed), yield: async () => { signal.aborted = true; } })).rejects.toMatchObject({ code: "OperationCancelled" });
  expect(progress).toEqual([256]); expect(source.serialize()).toEqual(before); expect(source.history().undoSteps).toBe(0);
  await expect(createWorkbookFillTransaction(source, "sheet-1", range(0), range(1, 0, 300), { yield: async () => seed(source, [5]) })).rejects.toMatchObject({ code: "RevisionConflict" });
  expect(source.getCell("sheet-1", 1, 0)).toBeUndefined(); expect(source.getCell("sheet-1", 0, 0)?.value).toEqual(number(5));
});
test("blank fill remains sparse and stop validation rejects the complete generated batch", async () => {
  const empty = session(), before = empty.serialize();
  const noChange = await createWorkbookFillTransaction(empty, "sheet-1", range(0), range(1, 0, 260));
  expect(empty.commit(empty.prepare(noChange))).toMatchObject({ changed: false, revision: 0 }); expect(empty.serialize()).toEqual(before);
  const files = new Map(createWorkbook({ documentId: "fill" })), metadata = JSON.parse(new TextDecoder().decode(files.get("metadata.json")!)) as WorkbookDocumentMetadata;
  metadata.workbook.sheets[0]!.dataValidations = [{ ruleId: "stop", kind: "whole", operator: "between", origin: { row: 10, column: 0 }, ranges: [range(10, 0, 20)], formulas: ["=1", "=3"],
    allowBlank: true, showInputMessage: true, showErrorMessage: true, errorStyle: "stop", errorTitle: "", error: "", promptTitle: "", prompt: "" }];
  files.set("metadata.json", new TextEncoder().encode(JSON.stringify(metadata)));
  const source = createDocumentSession(openDocument(files, { kind: "workbook" })); seed(source, [5]); const original = source.serialize();
  const filled = await createWorkbookFillTransaction(source, "sheet-1", range(0), range(10, 0, 20));
  expect(() => source.prepare(filled)).toThrowError(expect.objectContaining({ code: "InvalidOperation" })); expect(source.serialize()).toEqual(original); expect(source.history().undoSteps).toBe(0);
});
test("fill registers an absent General definition atomically while preserving sparse empty fills and undo", async () => {
  const files = new Map(createWorkbook({ documentId: "fill" })), metadata = JSON.parse(new TextDecoder().decode(files.get("metadata.json")!)) as WorkbookDocumentMetadata;
  metadata.workbook.numberFormats = [{ formatId: "date", code: "yyyy-mm-dd" }, { formatId: "fill-general-1", code: "0.00" }];
  delete metadata.workbook.sheets[0]!.defaultNumberFormatId;
  metadata.workbook.sheets[0]!.rowDimensions = [{ from: 1, to: 2, numberFormatId: "date" }]; files.set("metadata.json", new TextEncoder().encode(JSON.stringify(metadata)));
  const source = createDocumentSession(openDocument(files, { kind: "workbook" })); seed(source, [5, 22]);
  const before = source.serialize(), transaction = await createWorkbookFillTransaction(source, "sheet-1", range(0), range(1, 0, 2));
  expect(transaction.operations[0]).toEqual({ op: "setWorkbookNumberFormat", format: { formatId: "fill-general-2", builtinId: 0 } });
  expect(source.serialize()).toEqual(before); source.commit(source.prepare(transaction));
  for (const row of [1, 2]) expect(source.getCell("sheet-1", row, 0)).toMatchObject({ value: number(5), formatId: "fill-general-2" });
  expect(source.history().undoSteps).toBe(1); source.undo(); expect(source.getCell("sheet-1", 1, 0)?.value).toEqual(number(22));
  expect(source.snapshot().metadata.workbook.numberFormats).toEqual(metadata.workbook.numberFormats); expect(source.getCell("sheet-1", 2, 0)).toBeUndefined();
  const emptyFiles = new Map(files), emptyMetadata = structuredClone(metadata); emptyMetadata.workbook.sheets[0]!.rowDimensions = [];
  emptyFiles.set("metadata.json", new TextEncoder().encode(JSON.stringify(emptyMetadata))); const empty = createDocumentSession(openDocument(emptyFiles, { kind: "workbook" }));
  const unchanged = await createWorkbookFillTransaction(empty, "sheet-1", range(0), range(1, 0, 260));
  expect(unchanged.operations).toEqual([]); expect(empty.commit(empty.prepare(unchanged))).toMatchObject({ changed: false, revision: 0 });
  expect(empty.snapshot().metadata.workbook.numberFormats).toEqual(metadata.workbook.numberFormats);
});
