import { afterEach, expect, test } from "vitest";
import { createDocumentSession, openDocument } from "../src/index.js";
import type { DocumentOperation, WorkbookDataValidation, WorkbookSession, WorkbookTable } from "../src/index.js";
import { scalarWorkbook } from "./workbook-scalar-fixture.js";

const sheetId = "sheet-1", sessions: WorkbookSession[] = [];
const range = (row: number, column: number, lastRow = row, lastColumn = column) => ({ from: { row, column }, to: { row: lastRow, column: lastColumn } });
const number = (value: number) => ({ type: "number" as const, value });
function session(model = scalarWorkbook([])) { const s = createDocumentSession(model); sessions.push(s); return s; }
afterEach(() => { for (const s of sessions.splice(0)) s.dispose(); });
function transaction(s: WorkbookSession, operations: DocumentOperation[]) { const m = s.snapshot().manifest; return { protocolVersion: 1 as const, documentId: m.documentId, baseRevision: m.revision, operations }; }
function apply(s: WorkbookSession, operations: DocumentOperation[]) { return s.commit(s.prepare(transaction(s, operations))); }
function table(changes: Partial<WorkbookTable> = {}): WorkbookTable {
  return { tableId: "sales", name: "Sales", range: range(0, 0, 3, 1), header: true, totals: false, totalsPreviouslyShown: false,
    columns: [{ columnId: "a", name: "Amount", totalFunction: "none" }, { columnId: "b", name: "Double", calculatedFormula: "=A2*2", totalFunction: "none" }],
    rowStripes: true, columnStripes: false, showFirstColumn: false, showLastColumn: false, ...changes };
}
const validation: WorkbookDataValidation = { ruleId: "positive", kind: "whole", operator: "greaterThan", formulas: ["=0"], ranges: [range(0, 0)], origin: { row: 0, column: 0 },
  allowBlank: true, showErrorMessage: true, showInputMessage: true, errorStyle: "stop", error: "positive", errorTitle: "", prompt: "", promptTitle: "" };

test("name rename rewrites only bound global/local references and keeps historical caches, one revision and undo", () => {
  const model = scalarWorkbook([{ row: 0, column: 0, value: number(7) }, { row: 0, column: 1, value: number(999), formula: "=Rate" }], metadata => {
    metadata.workbook.names = [{ nameId: "global", name: "Rate", expression: "=Sheet1!$A$1", origin: { row: 0, column: 0 } }];
  });
  const s = session(model), before = s.snapshot();
  apply(s, [{ op: "insertWorkbookSheet", name: "Local" }]); const local = s.snapshot().metadata.workbook.sheets[1]!.sheetId;
  apply(s, [{ op: "setWorkbookDefinedName", name: { nameId: "local", name: "Rate", sheetId: local, expression: "=13", origin: { row: 0, column: 0 } } },
    { op: "setWorkbookCells", sheetId: local, cells: [{ row: 0, column: 0, source: { kind: "formula", expression: "=Rate" } }] }]);
  s.clearHistory();
  apply(s, [{ op: "setWorkbookDefinedName", name: { ...before.metadata.workbook.names![0]!, name: "Interest" } }]);
  expect(s.getCell(sheetId, 0, 1)).toMatchObject({ value: number(999), formula: { expression: "=Interest", cacheState: "stale" } });
  expect(s.getCell(local, 0, 0)?.formula?.expression).toBe("=Rate");
  s.calculateSync(); expect(s.getCalculatedCell(sheetId, 0, 1)).toMatchObject({ value: number(7) });
  expect(s.getCalculatedCell(local, 0, 0)).toMatchObject({ value: number(13) });
  expect(s.history().undoSteps).toBe(1); s.undo(); expect(s.getCell(sheetId, 0, 1)?.html).toBe(before.manifest.chunks.length ? model.getCell(sheetId, 0, 1)?.html : "");
  s.redo(); expect(openDocument(s.serialize(), { kind: "workbook" }).snapshot()).toEqual(s.snapshot());
});

test("removing a name keeps its formula source and produces a current name error; colliding/scope changes fail atomically", () => {
  const s = session(scalarWorkbook([{ row: 0, column: 0, value: number(1), formula: "=Rate" }]));
  const name = { nameId: "rate", name: "Rate", expression: "=7", origin: { row: 0, column: 0 } };
  apply(s, [{ op: "setWorkbookDefinedName", name }]); s.clearHistory(); const before = s.serialize();
  for (const next of [{ ...name, nameId: "other" }, { ...name, sheetId }]) {
    expect(() => apply(s, [{ op: "setWorkbookDefinedName", name: next }])).toThrowError(expect.objectContaining({ code: "InvalidOperation" }));
    expect(s.serialize()).toEqual(before); expect(s.history().undoSteps).toBe(0);
  }
  apply(s, [{ op: "removeWorkbookDefinedName", nameId: "rate" }]); s.calculateSync();
  expect(s.getCell(sheetId, 0, 0)?.formula?.expression).toBe("=Rate");
  expect(s.getCalculatedCell(sheetId, 0, 0)).toMatchObject({ value: { type: "error", value: "#NAME?" } });
  s.undo(); s.calculateSync(); expect(s.getCalculatedCell(sheetId, 0, 0)).toMatchObject({ value: number(7) });
});

test("table creation materializes headers and actual calculated-column formulas, and expansion preserves deliberate old exceptions", () => {
  const s = session(scalarWorkbook([1, 2, 3, 4].map((value, row) => ({ row: row + 1, column: 0, value: number(value) }))));
  const before = s.snapshot(), definition = table();
  const candidate = s.prepare(transaction(s, [{ op: "setWorkbookTable", sheetId, table: definition }]));
  expect(s.snapshot()).toEqual(before); expect(candidate.model.getCell(sheetId, 0, 0)?.value).toEqual({ type: "text", value: "Amount" });
  s.commit(candidate); s.calculateSync(); expect(s.getCalculatedCell(sheetId, 3, 1)).toMatchObject({ value: number(6) });
  apply(s, [{ op: "setWorkbookCells", sheetId, cells: [{ row: 2, column: 1, source: { kind: "value", value: number(99) } }] }]);
  apply(s, [{ op: "setWorkbookTable", sheetId, table: { ...definition, range: range(0, 0, 4, 1) } }]); s.calculateSync();
  expect(s.getCell(sheetId, 2, 1)?.value).toEqual(number(99)); expect(s.getCalculatedCell(sheetId, 4, 1)).toMatchObject({ value: number(8) });
  expect(s.history().undoSteps).toBe(3); s.undo(); expect(s.getCell(sheetId, 4, 1)).toBeUndefined();
  s.undo(); s.undo(); expect(s.snapshot().metadata).toEqual(before.metadata); expect(s.snapshot().manifest.chunks).toEqual(before.manifest.chunks);
});

test("table column swaps and table rename follow stable identities across cells, rules and names", () => {
  const s = session(scalarWorkbook([{ row: 1, column: 0, value: number(3) }, { row: 8, column: 0, value: number(999), formula: "=SUM(Sales[Amount])" }]));
  const definition = table(); apply(s, [{ op: "setWorkbookTable", sheetId, table: definition },
    { op: "setWorkbookDefinedName", name: { nameId: "all", name: "Amounts", expression: "=Sales[Amount]", origin: { row: 0, column: 0 } } }]); s.clearHistory();
  apply(s, [{ op: "setWorkbookTable", sheetId, table: { ...definition, name: "Revenue", columns: [{ ...definition.columns[0]!, name: "Double" }, { ...definition.columns[1]!, name: "Amount" }] } }]);
  expect(s.getCell(sheetId, 8, 0)?.formula?.expression).toBe("=SUM(Revenue[Double])");
  expect(s.snapshot().metadata.workbook.names![0]!.expression).toBe("=Revenue[Double]");
  s.calculateSync(); expect(s.getCalculatedCell(sheetId, 8, 0)).toMatchObject({ value: number(3) });
  expect(s.getCell(sheetId, 0, 0)?.value).toEqual({ type: "text", value: "Double" });
  expect(s.history().undoSteps).toBe(1); s.undo(); expect(s.getCell(sheetId, 8, 0)?.formula?.expression).toBe("=SUM(Sales[Amount])");
});

test("converting a table to cells resolves structured references and preserves all actual cell content", () => {
  const s = session(scalarWorkbook([{ row: 1, column: 0, value: number(3) }, { row: 8, column: 0, value: number(999), formula: "=SUM(Sales[Amount])" }]));
  apply(s, [{ op: "setWorkbookTable", sheetId, table: table() }]); s.clearHistory(); const before = s.snapshot();
  apply(s, [{ op: "removeWorkbookTable", sheetId, tableId: "sales" }]);
  expect(s.snapshot().metadata.workbook.sheets[0]!.tables).toEqual([]);
  expect(s.getCell(sheetId, 8, 0)?.formula?.expression).toBe("=SUM($A$2:$A$4)");
  expect(s.getCell(sheetId, 1, 0)?.value).toEqual(number(3)); s.calculateSync();
  expect(s.getCalculatedCell(sheetId, 8, 0)).toMatchObject({ value: number(3) });
  s.undo(); expect(s.snapshot().metadata).toEqual(before.metadata); expect(s.snapshot().manifest.chunks).toEqual(before.manifest.chunks);
});

test("table totals use current SUBTOTAL results, style dependencies reject removal, and no-op definitions preserve history", () => {
  const s = session(scalarWorkbook([{ row: 1, column: 0, value: number(3) }, { row: 2, column: 0, value: number(4) }]));
  const style = { styleId: "custom", name: "My table", pivotCompatible: false, elements: [{ region: "wholeTable" as const, size: 1, style: { cssClass: "cell-default" } }] };
  const definition = table({ totals: true, totalsPreviouslyShown: true, style: { kind: "custom", styleId: "custom" },
    columns: [{ columnId: "a", name: "Amount", totalFunction: "sum" }, { columnId: "b", name: "Double", totalFunction: "none" }] });
  apply(s, [{ op: "setWorkbookTableStyle", style }, { op: "setWorkbookTable", sheetId, table: definition }]); s.calculateSync();
  expect(s.getCalculatedCell(sheetId, 3, 0)).toMatchObject({ value: number(7) });
  const before = s.serialize(), history = s.history();
  expect(apply(s, [{ op: "setWorkbookTable", sheetId, table: definition }, { op: "setWorkbookTableStyle", style }]).changed).toBe(false);
  expect(s.serialize()).toEqual(before); expect(s.history()).toEqual(history);
  expect(() => apply(s, [{ op: "removeWorkbookTableStyle", styleId: "custom" }])).toThrow(); expect(s.serialize()).toEqual(before);
});

test("rule insertion/edit/removal enforces stop and warning semantics for every write entry and preserves failed batches", () => {
  const s = session(scalarWorkbook([{ row: 0, column: 0, value: number(3) }]));
  apply(s, [{ op: "setWorkbookDataValidation", sheetId, rule: validation }]); s.clearHistory(); const before = s.serialize();
  expect(() => apply(s, [{ op: "setWorkbookDataValidation", sheetId, rule: { ...validation, formulas: ["=9"] } }])).toThrowError(expect.objectContaining({ code: "InvalidOperation" }));
  expect(s.serialize()).toEqual(before); expect(s.history().undoSteps).toBe(0);
  const warning = { ...validation, formulas: ["=9"], errorStyle: "warning" as const };
  expect(() => apply(s, [{ op: "setWorkbookDataValidation", sheetId, rule: warning }])).toThrowError(expect.objectContaining({ details: [expect.objectContaining({ code: "validation-confirmation-required" })] }));
  s.commit(s.prepare(transaction(s, [{ op: "setWorkbookDataValidation", sheetId, rule: warning }]), { validation: { acceptedRuleIds: ["positive"] } }));
  apply(s, [{ op: "removeWorkbookDataValidation", sheetId, ruleId: "positive" }, { op: "setWorkbookCells", sheetId, cells: [{ row: 0, column: 0, source: { kind: "value", value: number(-4) } }] }]);
  expect(s.getCell(sheetId, 0, 0)?.value).toEqual(number(-4)); s.undo(); expect(s.getCell(sheetId, 0, 0)?.value).toEqual(number(3));
});

test("conditional rule priorities, missing definitions and cancelled large calculated-column expansion leave the source intact", async () => {
  const s = session(), before = s.serialize();
  const rule = { ruleId: "red", kind: "expression" as const, expression: "=A1>0", origin: { row: 0, column: 0 }, ranges: [range(0, 0, 5)], priority: 1, stopIfTrue: true, cssClass: "cell-default" };
  expect(() => apply(s, [{ op: "setWorkbookConditionalRule", sheetId, rule }, { op: "setWorkbookConditionalRule", sheetId, rule: { ...rule, ruleId: "other" } }])).toThrow();
  expect(s.serialize()).toEqual(before);
  expect(() => apply(s, [{ op: "removeWorkbookConditionalRule", sheetId, ruleId: "missing" }])).toThrowError(expect.objectContaining({ code: "InvalidOperation" }));
  const abort = new AbortController();
  await expect(s.prepareAsync(transaction(s, [{ op: "setWorkbookTable", sheetId, table: table({ range: range(0, 0, 1024, 1) }) }]), undefined,
    { signal: abort.signal, yield: async () => {}, onProgress(value) { if (value.phase === "cells" && value.completed >= 256) abort.abort(); } })).rejects.toThrowError(expect.objectContaining({ code: "OperationCancelled" }));
  expect(s.serialize()).toEqual(before); expect(s.history().undoSteps).toBe(0);
  apply(s, [{ op: "setWorkbookConditionalRule", sheetId, rule }]); apply(s, [{ op: "removeWorkbookConditionalRule", sheetId, ruleId: "red" }]);
  expect(s.snapshot().metadata.workbook.sheets[0]!.conditionalRules).toEqual([]);
});

test("removing a table column invalidates only its structured references and shrinks endpoint column ranges", () => {
  const s = session(scalarWorkbook([{ row: 8, column: 0, value: number(99), formula: "=SUM(Sales[Amount])" },
    { row: 8, column: 1, value: number(99), formula: "=SUM(Sales[[Amount]:[Double]])" }]));
  const definition = table(); apply(s, [{ op: "setWorkbookTable", sheetId, table: definition }]); s.clearHistory(); const before = s.snapshot();
  apply(s, [{ op: "setWorkbookTable", sheetId, table: { ...definition, range: range(0, 1, 3, 1), columns: [{ ...definition.columns[1]!, calculatedFormula: undefined }] } }]);
  expect(s.getCell(sheetId, 8, 0)?.formula?.expression).toBe("=SUM(#REF!)");
  expect(s.getCell(sheetId, 8, 1)?.formula?.expression).toBe("=SUM(Sales[[Double]:[Double]])");
  s.undo(); expect(s.snapshot().metadata).toEqual(before.metadata); expect(s.snapshot().manifest.chunks).toEqual(before.manifest.chunks);
});

test("net-zero definitions preserve absent lists, redo and events; foreign local-qualified references are not rebound", () => {
  const s = session(scalarWorkbook([{ row: 0, column: 0, value: number(99), formula: "=Sheet1!Rate" }]));
  apply(s, [{ op: "setWorkbookCells", sheetId, cells: [{ row: 2, column: 0, source: { kind: "value", value: number(3) } }] }]); s.undo();
  const before = s.serialize(), history = s.history(), events: string[] = []; s.subscribe(event => events.push(event.type));
  const name = { nameId: "rate", name: "Rate", expression: "=7", origin: { row: 0, column: 0 } };
  expect(apply(s, [{ op: "setWorkbookDefinedName", name }, { op: "removeWorkbookDefinedName", nameId: "rate" },
    { op: "setWorkbookDataValidation", sheetId, rule: validation }, { op: "removeWorkbookDataValidation", sheetId, ruleId: "positive" }]).changed).toBe(false);
  expect(s.serialize()).toEqual(before); expect(s.history()).toEqual(history); expect(events).toEqual([]);
  apply(s, [{ op: "setWorkbookDefinedName", name }, { op: "setWorkbookDefinedName", name: { ...name, name: "Interest" } }]);
  expect(s.getCell(sheetId, 0, 0)?.formula?.expression).toBe("=Sheet1!Rate");
});
test("renaming a global name does not guess the caller scope of an ambiguous nested workbook definition", () => {
  const s = session(scalarWorkbook([{ row: 0, column: 0, value: number(99), formula: "=Parent" }], metadata => {
    metadata.workbook.sheets.push({ ...metadata.workbook.sheets[0]!, sheetId: "other", name: "Other" });
    metadata.workbook.names = [{ nameId: "global", name: "Rate", expression: "=7", origin: { row: 0, column: 0 } },
      { nameId: "parent", name: "Parent", expression: "=Rate", origin: { row: 0, column: 0 } },
      { nameId: "local", name: "Rate", sheetId: "other", expression: "=13", origin: { row: 0, column: 0 } }];
  }));
  s.calculateSync(); expect(s.getCalculatedCell(sheetId, 0, 0)).toMatchObject({ state: "unavailable", issue: { reason: "ambiguous-name-scope" } });
  apply(s, [{ op: "setWorkbookDefinedName", name: { nameId: "global", name: "Interest", expression: "=7", origin: { row: 0, column: 0 } } }]);
  expect(s.snapshot().metadata.workbook.names!.find(name => name.nameId === "parent")?.expression).toBe("=Rate");
  s.calculateSync(); expect(s.getCalculatedCell(sheetId, 0, 0)).toMatchObject({ state: "unavailable", issue: { reason: "ambiguous-name-scope" } });
});
