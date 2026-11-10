import { expect, test } from "vitest";
import { createDocumentSession, createWorkbookCalculator, openDocument } from "@doffi/core";
import type { DocumentOperation, WorkbookDocumentMetadata, WorkbookRangeChange, WorkbookSession } from "@doffi/core";
import { scalarWorkbook as sourceWorkbook } from "../../core/tests/workbook-scalar-fixture.js";
import { findWorkbookText } from "../src/workbook/find.js";
import { createWorkbookReplaceTransaction } from "../src/workbook/replace.js";
const text = (value: string) => ({ type: "text" as const, value });
const number = (value: number) => ({ type: "number" as const, value });
const scalarWorkbook = (...args: Parameters<typeof sourceWorkbook>) => openDocument(sourceWorkbook(...args).serialize(), { kind: "workbook" });
const area = (row: number, column: number, toRow = row, toColumn = column, sheetId = "sheet-1"): WorkbookRangeChange => ({ sheetId, range: { from: { row, column }, to: { row: toRow, column: toColumn } } });
function commit(session: WorkbookSession, operations: DocumentOperation[]): void {
  const manifest = session.snapshot().manifest; session.commit(session.prepare({ protocolVersion: 1, documentId: manifest.documentId, baseRevision: manifest.revision, operations }));
}
test("find is literal Unicode search with exact offsets and row-major order across horizontal chunks", async () => {
  const model = scalarWorkbook([
    { row: 255, column: 0, value: text("match") }, { row: 0, column: 32, value: text("MATCH") }, { row: 256, column: 0, value: text("match") },
    { row: 0, column: 0, value: text("İ😀禰󠄀 match .*$& MATCH") }, { row: 1_048_575, column: 16383, value: text("match") },
  ]);
  expect((await findWorkbookText(model, { text: "match" })).matches.map(hit => [hit.row, hit.column, hit.from, hit.text])).toEqual([
    [0, 0, 7, "match"], [0, 0, 18, "MATCH"], [0, 32, 0, "MATCH"], [255, 0, 0, "match"], [256, 0, 0, "match"], [1_048_575, 16383, 0, "match"],
  ]);
  expect((await findWorkbookText(model, { text: ".*$&" })).matches).toMatchObject([{ from: 13, to: 17, text: ".*$&" }]);
  expect((await findWorkbookText(model, { text: "禰" })).matches).toEqual([]);
  expect((await findWorkbookText(model, { text: "禰󠄀" })).matches).toMatchObject([{ from: 3, to: 6 }]);
  expect((await findWorkbookText(model, { text: "MATCH", matchCase: true, entireCell: true })).matches).toMatchObject([{ row: 0, column: 32 }]);
  expect((await findWorkbookText(scalarWorkbook([{ row: 0, column: 0, value: text("match\n") }]), { text: "match", entireCell: true })).matches).toEqual([]);
});
test("scope is frozen before yielding, overlapping ranges do not duplicate results and empty source stays sparse", async () => {
  const session = createDocumentSession(scalarWorkbook([{ row: 255, column: 31, value: text("hit") }, { row: 256, column: 32, value: text("hit") }]));
  const scope = [area(255, 31, 256, 32), area(255, 31)], query = { text: "hit", scope };
  const result = await findWorkbookText(session, query, { yield: async () => { scope.length = 0; } });
  expect(result.matches).toHaveLength(2); expect(result.complete).toBe(true); expect(result.cells).toBe(2);
  expect((await findWorkbookText(session, { text: "hit", scope: [] })).matches).toEqual([]);
  expect((await findWorkbookText(session, { text: "hit", scope: [area(255, 31)] }, { windowCells: 1 })).matches).toHaveLength(1);
  expect((await findWorkbookText(scalarWorkbook([]), { text: "hit" })).windows).toBe(0);
  expect(session.history().undoSteps).toBe(0);
});
test("display find requires current calculation and format context, never imported caches or forged batches", async () => {
  const model = scalarWorkbook([{ row: 0, column: 0, value: number(1.25) }, { row: 0, column: 1, value: number(999), formula: "=A1+1" },
    { row: 0, column: 2, value: text("ignored"), formula: "=UNIMPLEMENTED(1)" }], metadata => {
    metadata.workbook.numberFormats.push({ formatId: "fixed", code: '"EUR "0.00' }); metadata.workbook.sheets[0]!.defaultNumberFormatId = "fixed";
  });
  const incomplete = await findWorkbookText(model, { text: "999", field: "display" });
  expect(incomplete).toMatchObject({ matches: [], complete: false, diagnostics: [{ reason: "calculation-required" }, { reason: "calculation-required" }] });
  const calculator = createWorkbookCalculator(model), calculation = calculator.calculateSync(); calculator.dispose();
  const result = await findWorkbookText(model, { text: "EUR 2.25", field: "display" }, { calculation });
  expect(result.matches).toMatchObject([{ column: 1, from: 0, to: 8 }]); expect(result.diagnostics).toMatchObject([{ column: 2, reason: "unavailable" }]);
  expect((await findWorkbookText(model, { text: "A1+1", field: "formula" })).matches).toMatchObject([{ column: 1 }]);
  await expect(findWorkbookText(model, { text: "2", field: "display" }, { calculation: { ...calculation } })).rejects.toMatchObject({ code: "InvalidOperation" });
});
test("display search shares table, named, direct and conditional number-format precedence", async () => {
  const model = scalarWorkbook([
    { row: 0, column: 0, value: number(1) }, { row: 0, column: 1, value: number(2) }, { row: 1, column: 0, value: number(3) },
    { row: 1, column: 1, value: number(4) }, { row: 2, column: 1, value: number(5) }, { row: 3, column: 1, value: number(6) },
  ], metadata => {
    const workbook = metadata.workbook, sheet = workbook.sheets[0]!;
    workbook.numberFormats.push(...["base", "table", "named", "column", "direct", "condition"].map((formatId, index) => ({ formatId, code: '"' + ["B", "T", "N", "C", "D", "CF"][index] + '"0' })));
    sheet.defaultNumberFormatId = "base";
    workbook.cellStyles = { definitions: [{ styleId: "named", cssClass: "cell-default", numberFormatId: "named", apply: { font: false, fill: false, border: false, alignment: false, numberFormat: true, protection: false } }],
      names: [{ name: "Native", styleId: "named", hidden: false, customBuiltin: false }], bindings: [] };
    sheet.tables = [{ tableId: "table", name: "TableOne", range: area(0, 0, 3, 1).range, header: false, totals: false, totalsPreviouslyShown: false,
      showFirstColumn: false, showLastColumn: false, rowStripes: false, columnStripes: false, dataStyle: { cssClass: "cell-default", numberFormatId: "table" },
      columns: [{ columnId: "a", name: "A", totalFunction: "none", dataCellStyleName: "Native" }, { columnId: "b", name: "B", totalFunction: "none", dataStyle: { cssClass: "cell-default", numberFormatId: "column" } }] }];
    sheet.conditionalRules = [
      { ruleId: "known", kind: "expression", expression: "=TRUE", origin: { row: 1, column: 1 }, ranges: [area(1, 1).range], priority: 1, stopIfTrue: false, numberFormatId: "condition" },
      { ruleId: "unknown-stop", kind: "expression", expression: "=MISSING(1)", origin: { row: 2, column: 1 }, ranges: [area(2, 1).range], priority: 2, stopIfTrue: true },
      { ruleId: "uncertain", kind: "expression", expression: "=TRUE", origin: { row: 2, column: 1 }, ranges: [area(2, 1).range], priority: 3, stopIfTrue: false, numberFormatId: "condition" },
    ];
  });
  const session = createDocumentSession(model);
  commit(session, [{ op: "setWorkbookCells", sheetId: "sheet-1", cells: [{ row: 3, column: 1, source: { kind: "value", value: number(6) }, formatId: "direct" }] }]);
  const calculation = session.calculateSync();
  for (const [text, row, column] of [["N1", 0, 0], ["C2", 0, 1], ["N3", 1, 0], ["CF4", 1, 1], ["D6", 3, 1]] as const) {
    const result = await findWorkbookText(session, { text, field: "display" }, { calculation });
    expect(result.matches).toMatchObject([{ row, column }]); expect(result.diagnostics).toMatchObject([{ row: 2, column: 1, reason: "unavailable" }]); expect(result.complete).toBe(false);
  }
  expect((await findWorkbookText(session, { text: "CF5", field: "display" }, { calculation })).matches).toEqual([]);
});
test("replace all preserves rich runs, CR, IVS, literal replacements and one cross-chunk undo", async () => {
  const files = new Map(scalarWorkbook([]).serialize()); files.set("styles.css", new TextEncoder().encode(new TextDecoder().decode(files.get("styles.css")!) + ".em{font-weight:700}"));
  const session = createDocumentSession(openDocument(files, { kind: "workbook" }));
  commit(session, [{ op: "setWorkbookCells", sheetId: "sheet-1", cells: [
    { row: 255, column: 31, source: { kind: "richText", html: '<span class="em">猫😀</span><span class="cell-default">&#13;禰󠄀</span><span class="em">猫</span>' } },
    { row: 256, column: 32, source: { kind: "value", value: text("猫猫") } },
  ] }]); session.clearHistory();
  const before = session.serialize(), transaction = await createWorkbookReplaceTransaction(session, { text: "猫" }, "$&=1");
  expect(session.serialize()).toEqual(before); expect(session.history().undoSteps).toBe(0); session.commit(session.prepare(transaction));
  expect(session.getCell("sheet-1", 255, 31)?.value).toEqual(text("$&=1😀\r禰󠄀$&=1"));
  expect(session.getCell("sheet-1", 255, 31)?.html).toContain('<span class="cell-default">&#13;禰󠄀</span>');
  expect(session.getCell("sheet-1", 255, 31)?.html).toContain('<span class="em">$&amp;=1😀</span>');
  expect(session.getCell("sheet-1", 256, 32)?.value).toEqual(text("$&=1$&=1")); expect(session.history().undoSteps).toBe(1);
  session.undo(); expect(session.getCell("sheet-1", 255, 31)?.value).toEqual(text("猫😀\r禰󠄀猫")); expect(session.getCell("sheet-1", 256, 32)?.value).toEqual(text("猫猫"));
});
test("single replacement validates query, revision and exact occurrence; formula source remains a formula", async () => {
  const session = createDocumentSession(scalarWorkbook([{ row: 0, column: 0, value: text("A A") }, { row: 0, column: 1, value: number(88), formula: "=A1&A1" }]));
  const result = await findWorkbookText(session, { text: "A" }), only = { documentId: result.documentId, baseRevision: result.baseRevision, match: result.matches[1]! };
  session.commit(session.prepare(await createWorkbookReplaceTransaction(session, { text: "A" }, "=B1", { only, maxMatches: 1 })));
  expect(session.getCell("sheet-1", 0, 0)?.value).toEqual(text("A =B1")); expect(session.getCell("sheet-1", 0, 1)?.formula?.expression).toBe("=A1&A1");
  await expect(createWorkbookReplaceTransaction(session, { text: "A" }, "B", { only })).rejects.toMatchObject({ code: "RevisionConflict" });
  await expect(createWorkbookReplaceTransaction(session, { text: "B" }, "C", { only: { ...only, baseRevision: 1 } })).rejects.toMatchObject({ code: "InvalidOperation" });
  session.commit(session.prepare(await createWorkbookReplaceTransaction(session, { text: "A1", field: "formula" }, "A2")));
  expect(session.getCell("sheet-1", 0, 1)?.formula?.expression).toBe("=A2&A2");
  await expect(createWorkbookReplaceTransaction(session, { text: "=A2&A2", field: "formula" }, "literal")).rejects.toMatchObject({ code: "InvalidOperation" });
});
test("constant replacement requires explicit type conversion and repeated identical replacements are no-ops", async () => {
  const session = createDocumentSession(scalarWorkbook([{ row: 0, column: 0, value: number(12) }, { row: 0, column: 1, value: text("12") }]));
  await expect(createWorkbookReplaceTransaction(session, { text: "12" }, "text")).rejects.toMatchObject({ code: "InvalidOperation" });
  expect(session.history().undoSteps).toBe(0);
  session.commit(session.prepare(await createWorkbookReplaceTransaction(session, { text: "12" }, "13")));
  expect(session.getCell("sheet-1", 0, 0)?.value).toEqual(number(13)); expect(session.getCell("sheet-1", 0, 1)?.value).toEqual(text("13"));
  expect(session.commit(session.prepare(await createWorkbookReplaceTransaction(session, { text: "13" }, "13")))).toMatchObject({ changed: false, revision: 1 });
  session.commit(session.prepare(await createWorkbookReplaceTransaction(session, { text: "13" }, "=1+1", { valueMode: "auto" })));
  expect(session.getCell("sheet-1", 0, 0)?.formula?.expression).toBe("=1+1"); expect(session.getCell("sheet-1", 0, 1)?.formula?.expression).toBe("=1+1");
  await expect(createWorkbookReplaceTransaction(session, { text: "2", field: "display" }, "3")).rejects.toMatchObject({ code: "InvalidOperation" });
});
test("limits, cancellation, changed source and stop validation cannot publish partial replacements", async () => {
  const session = createDocumentSession(scalarWorkbook([{ row: 0, column: 0, value: text("aaa") }, { row: 0, column: 32, value: text("aaa") }]));
  const before = session.serialize();
  for (const limits of [{ maxCells: 1 }, { maxMatches: 1 }, { maxCodeUnits: 1 }, { maxWindows: 1 }]) await expect(createWorkbookReplaceTransaction(session, { text: "a" }, "b", limits)).rejects.toMatchObject({ code: "ResourceLimitExceeded" });
  await expect(createWorkbookReplaceTransaction(session, { text: "a" }, "b", { signal: { aborted: true } })).rejects.toMatchObject({ code: "OperationCancelled" });
  const signal = { aborted: false }; await expect(createWorkbookReplaceTransaction(session, { text: "a" }, "b", { onProgress: () => { signal.aborted = true; }, signal })).rejects.toMatchObject({ code: "OperationCancelled" });
  expect(session.serialize()).toEqual(before); expect(session.history().undoSteps).toBe(0);
  let changed = false; await expect(createWorkbookReplaceTransaction(session, { text: "a" }, "b", { yield: async () => {
    if (changed) return; changed = true; commit(session, [{ op: "setWorkbookCells", sheetId: "sheet-1", cells: [{ row: 5, column: 5, source: { kind: "value", value: text("external") } }] }]);
  } })).rejects.toMatchObject({ code: "RevisionConflict" });
  expect(session.getCell("sheet-1", 0, 0)?.value).toEqual(text("aaa"));
  const stop = createDocumentSession(scalarWorkbook([{ row: 0, column: 0, value: number(1) }], (metadata: WorkbookDocumentMetadata) => {
    metadata.workbook.sheets[0]!.dataValidations = [{ ruleId: "stop", kind: "whole", operator: "between", origin: { row: 0, column: 0 }, ranges: [area(0, 0).range], formulas: ["=1", "=3"], allowBlank: true,
      showInputMessage: true, showErrorMessage: true, errorStyle: "stop", errorTitle: "", error: "", promptTitle: "", prompt: "" }];
  }));
  const candidate = await createWorkbookReplaceTransaction(stop, { text: "1" }, "5"); expect(() => stop.prepare(candidate)).toThrowError(expect.objectContaining({ code: "InvalidOperation" })); expect(stop.history().undoSteps).toBe(0);
  await expect(findWorkbookText(session, { text: "\ud800" })).rejects.toMatchObject({ code: "InvalidOperation" });
});
