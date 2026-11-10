import { expect, test } from "vitest";
import { createDocumentSession, createWorkbookCalculatedEvaluator, openDocument } from "@doffi/core";
import type { DocumentOperation, WorkbookDocumentMetadata, WorkbookSession } from "@doffi/core";
import { scalarWorkbook } from "../../core/tests/workbook-scalar-fixture.js";
import { createWorkbookRangeCopyTransaction } from "../src/workbook/copy-range.js";
const number = (value: number) => ({ type: "number" as const, value });
const text = (value: string) => ({ type: "text" as const, value });
const range = (row: number, column = 0, lastRow = row, lastColumn = column) => ({ from: { row, column }, to: { row: lastRow, column: lastColumn } });
function apply(session: WorkbookSession, operations: DocumentOperation[]) {
  const manifest = session.snapshot().manifest; return session.commit(session.prepare({ protocolVersion: 1, documentId: manifest.documentId, baseRevision: manifest.revision, operations }));
}
function styled(): WorkbookSession {
  const files = new Map(scalarWorkbook([{ row: 0, column: 0, value: number(7) }, { row: 255, column: 32, value: number(999), formula: "=LEN(AF256)+$A$1" }, { row: 256, column: 32, value: number(9) }], metadata => {
    metadata.workbook.numberFormats = [{ formatId: "date", code: "yyyy-mm-dd" }]; delete metadata.workbook.sheets[0]!.defaultNumberFormatId;
    metadata.workbook.sheets[0]!.rowDimensions = [{ from: 255, to: 255, numberFormatId: "date" }];
  }).serialize());
  files.set("styles.css", new TextEncoder().encode(new TextDecoder().decode(files.get("styles.css")!) + ".source-font{font-weight:700}.target-font{font-style:italic}"));
  const metadata = JSON.parse(new TextDecoder().decode(files.get("metadata.json")!)) as WorkbookDocumentMetadata;
  metadata.workbook.sheets[0]!.rowDimensions[0]!.cssClass = "source-font"; files.set("metadata.json", new TextEncoder().encode(JSON.stringify(metadata)));
  return createDocumentSession(openDocument(files, { kind: "workbook" }));
}
test("copy snapshots overlapping cells, rich font spans, inherited formats and mixed references in one cross-chunk undo", async () => {
  const session = styled(); apply(session, [{ op: "setWorkbookCells", sheetId: "sheet-1", cells: [{ row: 255, column: 31, source: { kind: "richText", html: '<span class="source-font">中文&#13;</span><span class="cell-default">😀禰󠄀</span>' } }] }]);
  session.clearHistory(); const before = session.serialize(), rich = session.getCell("sheet-1", 255, 31)!;
  const transaction = await createWorkbookRangeCopyTransaction(session, "sheet-1", range(255, 31, 256, 32), { row: 256, column: 32 });
  expect(session.serialize()).toEqual(before); session.commit(session.prepare(transaction));
  expect(session.getCell("sheet-1", 256, 32)).toMatchObject({ value: rich.value, cssClass: "source-font", formatId: "date" });
  expect(session.getCell("sheet-1", 256, 32)?.html).toContain('<span class="cell-default">😀禰󠄀</span>');
  expect(session.getCell("sheet-1", 256, 33)?.formula).toMatchObject({ expression: expect.stringContaining("AG257"), cacheState: "stale" });
  expect(session.getCell("sheet-1", 256, 33)?.formula?.expression).toContain("$A$1");
  expect(session.getCell("sheet-1", 257, 33)?.value).toEqual(number(9)); expect(session.getCell("sheet-1", 257, 32)).toBeUndefined();
  expect(session.history().undoSteps).toBe(1); session.undo();
  expect(session.getCell("sheet-1", 256, 32)?.value).toEqual(number(9)); expect(session.getCell("sheet-1", 256, 33)).toBeUndefined();
  expect(session.getCell("sheet-1", 255, 31)?.html).toBe(rich.html);
});
test("paste values uses a completed long-chain batch and supports in-place formula conversion without imported caches", async () => {
  const session = createDocumentSession(scalarWorkbook(Array.from({ length: 1200 }, (_, row) => row === 0 ? { row, column: 0, value: number(1) } : { row, column: 0, value: number(999), formula: `=A${row}+1` })));
  const before = session.serialize();
  await expect(createWorkbookRangeCopyTransaction(session, "sheet-1", range(1199), { row: 1300, column: 0 }, { mode: "values" })).rejects.toMatchObject({ code: "InvalidOperation" });
  expect(session.serialize()).toEqual(before);
  const calculation = session.calculateSync();
  await expect(createWorkbookRangeCopyTransaction(session, "sheet-1", range(1199), { row: 1300, column: 0 }, { mode: "values", calculation: { ...calculation } })).rejects.toMatchObject({ code: "InvalidOperation" });
  session.commit(session.prepare(await createWorkbookRangeCopyTransaction(session, "sheet-1", range(1199), { row: 1199, column: 0 }, { mode: "values", calculation })));
  expect(session.getCell("sheet-1", 1199, 0)?.value).toEqual(number(1200)); expect(session.getCell("sheet-1", 1199, 0)?.formula).toBeUndefined();
  expect(session.history().undoSteps).toBe(1); session.undo(); expect(session.getCell("sheet-1", 1199, 0)).toMatchObject({ value: number(999), formula: { expression: "=A1199+1", cacheState: "imported" } });
  await expect(createWorkbookRangeCopyTransaction(session, "sheet-1", range(1199), { row: 1300, column: 0 }, { mode: "values", calculation })).rejects.toThrow();
});
test("formulas and values keep target formats and target rich-text roles, including literal formula-like text", async () => {
  const session = styled();
  apply(session, [{ op: "setWorkbookCells", sheetId: "sheet-1", cells: [
    { row: 0, column: 0, source: { kind: "richText", html: '<span class="source-font">=raw禰󠄀</span>' } },
    { row: 0, column: 1, source: { kind: "formula", expression: "=LEN(A1)" } },
    { row: 1, column: 0, source: { kind: "richText", html: '<span class="target-font">old禰󠄀</span>' }, cssClass: "target-font", formatId: "date" },
    { row: 1, column: 1, source: { kind: "formula", expression: "=UNIMPLEMENTED()" }, cssClass: "target-font", formatId: "date" },
  ] }]); session.clearHistory();
  session.commit(session.prepare(await createWorkbookRangeCopyTransaction(session, "sheet-1", range(0, 0, 0, 1), { row: 1, column: 0 }, { mode: "formulas" })));
  expect(session.getCell("sheet-1", 1, 0)).toMatchObject({ value: text("=raw禰󠄀"), cssClass: "target-font", formatId: "date" });
  expect(session.getCell("sheet-1", 1, 0)?.html).toContain('<span class="target-font">=raw禰󠄀</span>');
  expect(session.getCell("sheet-1", 1, 0)?.html).not.toContain('class="source-font"');
  expect(session.getCell("sheet-1", 1, 1)).toMatchObject({ formula: { expression: "=LEN(A2)" }, cssClass: "target-font", formatId: "date" });
  expect(session.history().undoSteps).toBe(1);
});
test("format paste does not copy content or invalidate an imported unsupported formula cache", async () => {
  const files = new Map(scalarWorkbook([{ row: 0, column: 0, value: text("长".repeat(3000)) }, { row: 1, column: 0, value: number(777), formula: "=UNIMPLEMENTED()" }], metadata => {
    metadata.workbook.numberFormats.push({ formatId: "date", code: "yyyy-mm-dd" }); metadata.workbook.sheets[0]!.rowDimensions = [{ from: 0, to: 0, numberFormatId: "date" }];
  }).serialize()), session = createDocumentSession(openDocument(files, { kind: "workbook" })), original = session.getCell("sheet-1", 1, 0)!;
  const transaction = await createWorkbookRangeCopyTransaction(session, "sheet-1", range(0), { row: 1, column: 0 }, { mode: "formats", maxCodeUnits: 512 });
  expect(transaction.operations[0]?.op).toBe("setWorkbookCellFormats"); session.commit(session.prepare(transaction));
  expect(session.getCell("sheet-1", 1, 0)).toMatchObject({ value: original.value, formula: original.formula, text: original.text, formatId: "date" });
  expect(session.history().undoSteps).toBe(1); session.undo(); expect(session.getCell("sheet-1", 1, 0)?.html).toBe(original.html);
});
test("skip blanks distinguishes empty text from absent cells and unavailable current values fail without partial source", async () => {
  const session = createDocumentSession(scalarWorkbook([{ row: 0, column: 1, value: text("") }, { row: 0, column: 2, value: number(999), formula: "=UNIMPLEMENTED()" },
    { row: 1, column: 0, value: text("keep") }, { row: 1, column: 1, value: text("replace") }]));
  session.commit(session.prepare(await createWorkbookRangeCopyTransaction(session, "sheet-1", range(0, 0, 0, 1), { row: 1, column: 0 }, { mode: "values", skipBlanks: true })));
  expect(session.getCell("sheet-1", 1, 0)?.value).toEqual(text("keep")); expect(session.getCell("sheet-1", 1, 1)?.value).toEqual(text(""));
  const before = session.serialize(), calculation = session.calculateSync();
  await expect(createWorkbookRangeCopyTransaction(session, "sheet-1", range(0, 0, 0, 2), { row: 2, column: 0 }, { mode: "values", calculation })).rejects.toMatchObject({ code: "UnsupportedFeature" });
  expect(session.serialize()).toEqual(before); expect(session.history().undoSteps).toBe(1);
});
test("range copy cancellation, source replacement, bounds and budgets leave no generated edits", async () => {
  const session = createDocumentSession(scalarWorkbook([{ row: 0, column: 0, value: number(1) }])), before = session.serialize(), signal = { aborted: false };
  await expect(createWorkbookRangeCopyTransaction(session, "sheet-1", range(0, 0, 300), { row: 400, column: 0 }, { signal, yield: async () => { signal.aborted = true; } })).rejects.toMatchObject({ code: "OperationCancelled" });
  for (const options of [{ maxCells: 1 }, { maxCodeUnits: 1 }]) await expect(createWorkbookRangeCopyTransaction(session, "sheet-1", range(0, 0, 2), { row: 400, column: 0 }, options)).rejects.toMatchObject({ code: "ResourceLimitExceeded" });
  await expect(createWorkbookRangeCopyTransaction(session, "sheet-1", range(0, 0, 2), { row: 1_048_575, column: 0 })).rejects.toMatchObject({ code: "InvalidOperation" });
  expect(session.serialize()).toEqual(before); expect(session.history().undoSteps).toBe(0);
  let changed = false;
  await expect(createWorkbookRangeCopyTransaction(session, "sheet-1", range(0, 0, 300), { row: 400, column: 0 }, { yield: async () => {
    if (changed) return; changed = true; apply(session, [{ op: "setWorkbookCells", sheetId: "sheet-1", cells: [{ row: 0, column: 0, source: { kind: "value", value: number(2) } }] }]);
  } })).rejects.toMatchObject({ code: "RevisionConflict" });
  expect(session.getCell("sheet-1", 400, 0)).toBeUndefined(); expect(session.getCell("sheet-1", 0, 0)?.value).toEqual(number(2));
});

test("cross-sheet copy at identical coordinates preserves source rich spans and inherited formats in one undo", async () => {
  const session = styled(); apply(session, [{ op: "insertWorkbookSheet", name: "Destination" }]);
  const targetSheetId = session.snapshot().metadata.workbook.sheets.find(sheet => sheet.name === "Destination")!.sheetId;
  apply(session, [{ op: "setWorkbookCells", sheetId: "sheet-1", cells: [{ row: 255, column: 31, source: { kind: "richText", html: '<span class="source-font">甲&#13;</span><span class="cell-default">禰󠄀😀</span>' } }] },
    { op: "setWorkbookCells", sheetId: targetSheetId, cells: [{ row: 255, column: 31, source: { kind: "value", value: text("old") }, cssClass: "target-font" }] }]);
  session.clearHistory(); const before = session.serialize(), rich = session.getCell("sheet-1", 255, 31)!;
  const command = await createWorkbookRangeCopyTransaction(session, "sheet-1", range(255, 31, 256, 32), { row: 255, column: 31 }, { targetSheetId });
  expect(session.serialize()).toEqual(before); session.commit(session.prepare(command));
  expect(session.getCell(targetSheetId, 255, 31)).toMatchObject({ value: rich.value, cssClass: "source-font", formatId: "date" });
  expect(session.getCell(targetSheetId, 255, 31)?.html).toContain('<span class="cell-default">禰󠄀😀</span>');
  expect(session.getCell(targetSheetId, 255, 32)?.formula?.expression).toBe("=LEN(AF256)+$A$1");
  expect(session.getCell(targetSheetId, 256, 32)?.value).toEqual(number(9)); expect(session.getCell("sheet-1", 255, 31)?.html).toBe(rich.html);
  expect(openDocument(session.serialize(), { kind: "workbook" }).getCell(targetSheetId, 255, 31)?.value).toEqual(rich.value);
  expect(session.history().undoSteps).toBe(1); session.undo(); expect(session.getCell(targetSheetId, 255, 31)?.value).toEqual(text("old"));
  expect(session.snapshot().manifest.chunks).toEqual(openDocument(before, { kind: "workbook" }).snapshot().manifest.chunks);
});

test("cross-sheet selective values read the source calculation and retain destination content formatting", async () => {
  const session = styled(); apply(session, [{ op: "insertWorkbookSheet", name: "Destination" }]);
  const targetSheetId = session.snapshot().metadata.workbook.sheets.find(sheet => sheet.name === "Destination")!.sheetId;
  apply(session, [{ op: "setWorkbookCells", sheetId: "sheet-1", cells: [{ row: 0, column: 1, source: { kind: "formula", expression: "=A1+1" } }] },
    { op: "setWorkbookCells", sheetId: targetSheetId, cells: [{ row: 256, column: 32, source: { kind: "value", value: number(100) }, cssClass: "target-font", formatId: "date" }] }]);
  session.clearHistory(); const calculation = session.calculateSync();
  session.commit(session.prepare(await createWorkbookRangeCopyTransaction(session, "sheet-1", range(0, 1), { row: 256, column: 32 }, { targetSheetId, mode: "values", calculation })));
  expect(session.getCell(targetSheetId, 256, 32)).toMatchObject({ value: number(8), cssClass: "target-font", formatId: "date" });
  expect(session.getCell(targetSheetId, 256, 32)?.formula).toBeUndefined(); expect(session.getCell("sheet-1", 0, 1)?.formula?.expression).toBe("=A1+1");
  expect(session.history().undoSteps).toBe(1); session.undo(); expect(session.getCell(targetSheetId, 256, 32)?.value).toEqual(number(100));
});

test("copied formula tokens retain explicit sheets while relative cells, local names and unqualified tables use destination context", async () => {
  const expression = "=A1+Sheet1!$A$1+Local+Sheet1!Local+Shared+SUM(SourceTable[Amount])+SUM([Amount])";
  const model = scalarWorkbook([{ row: 0, column: 0, value: number(5) }, { row: 1, column: 0, value: number(7) }, { row: 0, column: 1, value: number(999), formula: expression }], metadata => {
    const source = metadata.workbook.sheets[0]!, target = { ...structuredClone(source), sheetId: "destination", name: "Destination" }; metadata.workbook.sheets.push(target);
    for (const [sheet, name] of [[source, "SourceTable"], [target, "DestinationTable"]] as const) sheet.tables = [{ tableId: "table-" + sheet.sheetId, name, range: range(0, 0, 2, 1), header: false, totals: false, totalsPreviouslyShown: false,
      showFirstColumn: false, showLastColumn: false, rowStripes: false, columnStripes: false,
      columns: [{ columnId: "amount-" + sheet.sheetId, name: "Amount", totalFunction: "none" }, { columnId: "formula-" + sheet.sheetId, name: "Formula", totalFunction: "none" }] }];
    metadata.workbook.names = [{ nameId: "local-source", name: "Local", sheetId: source.sheetId, origin: { row: 0, column: 0 }, expression: "=3" },
      { nameId: "local-destination", name: "Local", sheetId: target.sheetId, origin: { row: 0, column: 0 }, expression: "=17" },
      { nameId: "shared", name: "Shared", origin: { row: 0, column: 0 }, expression: "=2" }];
  }), session = createDocumentSession(model);
  apply(session, [{ op: "setWorkbookCells", sheetId: "destination", cells: [{ row: 0, column: 0, source: { kind: "value", value: number(11) } }, { row: 1, column: 0, source: { kind: "value", value: number(13) } }] }]); session.clearHistory();
  session.commit(session.prepare(await createWorkbookRangeCopyTransaction(session, "sheet-1", range(0, 1), { row: 1, column: 1 }, { targetSheetId: "destination", mode: "formulas" })));
  const copied = session.getCell("destination", 1, 1)!.formula!.expression;
  expect(copied).toContain("A2"); expect(copied).toContain("'Sheet1'!$A$1"); expect(copied).toContain("SourceTable[Amount]"); expect(copied).toContain("SUM([Amount])");
  const evaluator = createWorkbookCalculatedEvaluator(session, session.calculateSync());
  try { expect(evaluator.cell({ sheetId: "destination", row: 1, column: 1 })).toMatchObject({ kind: "value", value: number(76) }); }
  finally { evaluator.dispose(); }
  session.undo(); expect(session.getCell("destination", 1, 1)).toBeUndefined(); expect(session.getCell("sheet-1", 0, 1)?.formula?.expression).toBe(expression);
});

test("cross-sheet format paste leaves formula caches intact and invalid destinations fail atomically", async () => {
  const session = styled(); apply(session, [{ op: "insertWorkbookSheet", name: "Destination" }]);
  const targetSheetId = session.snapshot().metadata.workbook.sheets.find(sheet => sheet.name === "Destination")!.sheetId;
  apply(session, [{ op: "setWorkbookCells", sheetId: targetSheetId, cells: [{ row: 0, column: 0, source: { kind: "formula", expression: "=UNIMPLEMENTED()" }, cssClass: "target-font" }] }]);
  session.clearHistory(); const previous = session.getCell(targetSheetId, 0, 0)!;
  session.commit(session.prepare(await createWorkbookRangeCopyTransaction(session, "sheet-1", range(255, 32), { row: 0, column: 0 }, { targetSheetId, mode: "formats" })));
  expect(session.getCell(targetSheetId, 0, 0)).toMatchObject({ value: previous.value, formula: previous.formula, formatId: "date", cssClass: "source-font" });
  const before = session.serialize();
  for (const destination of ["missing", null]) await expect(createWorkbookRangeCopyTransaction(session, "sheet-1", range(0), { row: 0, column: 0 }, { targetSheetId: destination as string })).rejects.toMatchObject({ code: "InvalidOperation" });
  expect(session.serialize()).toEqual(before); expect(session.history().undoSteps).toBe(1);
});
