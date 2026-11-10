import { expect, test } from "vitest";
import { createDocumentSession, parseContract, parseWorkbookFormula } from "../src/index.js";
import type { DocumentOperation, WorkbookDefinedName, WorkbookEditProposal, WorkbookSession, WorkbookTable } from "../src/index.js";
import { scalarWorkbook } from "./workbook-scalar-fixture.js";

const value = (value: number) => ({ type: "number" as const, value });
const name = (name: string, expression: string, sheetId?: string): WorkbookDefinedName => ({ nameId: (sheetId ?? "global") + "-" + name, name, expression, origin: { row: 0, column: 0 }, ...(sheetId ? { sheetId } : {}) });
function transaction(s: WorkbookSession, operations: DocumentOperation[]) { const m = s.snapshot().manifest; return { protocolVersion: 1 as const, documentId: m.documentId, baseRevision: m.revision, operations }; }
function apply(s: WorkbookSession, operations: DocumentOperation[]) { return s.commit(s.prepare(transaction(s, operations))); }
function namesFixture() {
  const s = createDocumentSession(scalarWorkbook([{ row: 0, column: 0, value: value(3) }, { row: 0, column: 1, value: value(999), formula: "=Alias+[0]!DataValue+DataValue+Local+ConstX+OtherOnly+MixedX+Sheet1!Local+FromLocal" }, { row: 1, column: 1, value: value(999), formula: "=CycleA" }], metadata => {
    const source = metadata.workbook.sheets[0]!; metadata.workbook.sheets.push({ ...structuredClone(source), sheetId: "other", name: "Other" });
    metadata.workbook.names = [name("DataValue", "=Sheet1!$A$1"), name("Alias", "=[0]!DataValue"), name("ExplicitAlias", "=[0]!DataValue"), name("MixedX", "=Sheet1!$A$1+Other!$A$1"), name("ConstX", "=2"), name("OtherOnly", "=Other!$A$1"), name("FromLocal", "=Sheet1!Local"),
      name("CycleA", "=[0]!CycleB+Sheet1!$A$1"), name("CycleB", "=[0]!CycleA"), { ...name("HiddenX", "=SUM(Sheet1!$A$1:$A$2)"), hidden: true, description: "原名禰󠄀😀" }, name("DataValue", "=99", "sheet-1"), name("Local", "=Sheet1!$A$1*3", "sheet-1")];
  })); apply(s, [{ op: "setWorkbookCells", sheetId: "other", cells: [{ row: 0, column: 0, source: { kind: "value", value: value(5) } }] }]); s.clearHistory(); return s;
}
test.each([252, 251])("copied table names never split a supplementary ideograph at the suffix boundary: %s", prefix => {
  const original = "列".repeat(prefix) + "𠀀" + "a".repeat(253 - prefix), s = createDocumentSession(scalarWorkbook([{ row: 1, column: 0, value: value(3) }]));
  const table: WorkbookTable = { tableId: "long-table", name: original, objectName: original, range: { from: { row: 0, column: 0 }, to: { row: 1, column: 1 } }, header: true, totals: false, totalsPreviouslyShown: false,
    columns: [{ columnId: "price", name: "Price", totalFunction: "none" }, { columnId: "double", name: "Double", calculatedFormula: `=${original}[@Price]*2`, totalFunction: "none" }], rowStripes: true, columnStripes: false, showFirstColumn: false, showLastColumn: false };
  try {
    expect(original.length).toBe(255); apply(s, [{ op: "setWorkbookTable", sheetId: "sheet-1", table }]); s.clearHistory(); const before = s.snapshot();
    for (let index = 0; index < 9; index++) apply(s, [{ op: "copyWorkbookSheet", sheetId: "sheet-1", name: "Copy" + index }]);
    for (const copy of s.snapshot().metadata.workbook.sheets.slice(1)) {
      const copied = copy.tables![0]!; expect(copied.name.length).toBeLessThanOrEqual(255); expect(copied.name).not.toMatch(/[\uD800-\uDFFF]/u);
      expect(copied.objectName).toBe(copied.name); expect(parseWorkbookFormula(`=${copied.name}[Price]`, { row: 0, column: 0 })).toMatchObject({ kind: "structured", table: copied.name });
      expect(copied.columns[1]!.calculatedFormula).toContain(copied.name);
    }
    s.calculateSync(); const copy = s.snapshot().metadata.workbook.sheets.at(-1)!;
    expect(s.getCalculatedCell(copy.sheetId, 1, 1)).toMatchObject({ value: value(6) });
    for (let index = 0; index < 9; index++) s.undo();
    expect(s.snapshot().metadata).toEqual(before.metadata); expect(s.snapshot().manifest.chunks).toEqual(before.manifest.chunks);
  } finally { s.dispose(); }
});
test("whole-sheet named graphs clone reachable definitions, resolve explicit conflicts and keep unrelated globals and cycles", () => {
  const s = namesFixture();
  try {
    const before = s.snapshot(), op = { op: "copyWorkbookSheet" as const, sheetId: "sheet-1", name: "Copied", nameConflicts: [{ nameId: "global-DataValue", name: "CopyData" }] };
    const tx = parseContract("document-transaction", transaction(s, [op])), candidate = s.prepare(tx); expect(s.snapshot()).toEqual(before); s.commit(candidate);
    const copied = s.snapshot().metadata.workbook.sheets.at(-1)!, definitions = s.snapshot().metadata.workbook.names!.filter(name => name.sheetId === copied.sheetId);
    expect(definitions.map(name => name.name).sort()).toEqual(["Alias", "CopyData", "CycleA", "CycleB", "DataValue", "ExplicitAlias", "FromLocal", "HiddenX", "Local", "MixedX"].sort());
    expect(definitions.find(name => name.name === "HiddenX")).toMatchObject({ hidden: true, description: "原名禰󠄀😀", origin: { row: 0, column: 0 } });
    expect(parseWorkbookFormula(definitions.find(name => name.name === "Alias")!.expression, { row: 0, column: 0 })).toEqual(parseWorkbookFormula("=Copied!CopyData", { row: 0, column: 0 }));
    expect(parseWorkbookFormula(definitions.find(name => name.name === "ExplicitAlias")!.expression, { row: 0, column: 0 })).toEqual(parseWorkbookFormula("=Copied!CopyData", { row: 0, column: 0 }));
    apply(s, [{ op: "setWorkbookCells", sheetId: copied.sheetId, cells: [{ row: 0, column: 0, source: { kind: "value", value: value(100) } }] }]); s.calculateSync();
    expect(s.getCalculatedCell(copied.sheetId, 0, 1)).toMatchObject({ state: "valid", value: value(1311) });
    expect(s.getCalculatedCell(copied.sheetId, 1, 1)).toMatchObject({ state: "unavailable", issue: { reason: "circular-name" } });
    expect(s.snapshot().metadata.workbook.names!.filter(name => name.sheetId !== copied.sheetId)).toEqual(before.metadata.workbook.names);
    expect(s.snapshot().manifest.chunks.filter(chunk => chunk.sheetId !== copied.sheetId)).toEqual(before.manifest.chunks);
    expect(() => s.getCell("not-a-sheet", 0, 0)).toThrow(); s.undo(); s.undo(); expect(s.snapshot().metadata).toEqual(before.metadata);
  } finally { s.dispose(); }
});
test("copy name conflicts, duplicate sheet names and removed insertion/source targets fail the complete batch without history", () => {
  const s = namesFixture();
  try {
    const before = s.serialize(), copy = { op: "copyWorkbookSheet" as const, sheetId: "sheet-1", name: "Copied" };
    expect(() => s.prepare(transaction(s, [copy]))).toThrowError(expect.objectContaining({ code: "InvalidOperation", details: [expect.objectContaining({ code: "copy-name-conflict" })] }));
    for (const nameConflicts of [[{ nameId: "missing", name: "Chosen" }], [{ nameId: "global-DataValue", name: "ConstX" }], [{ nameId: "global-DataValue", name: "A1" }], [{ nameId: "global-DataValue", name: "Chosen" }, { nameId: "global-DataValue", name: "Again" }]]) expect(() => s.prepare(transaction(s, [{ ...copy, nameConflicts }]))).toThrow();
    for (const name of ["Sheet1", "'bad", "bad'"]) expect(() => s.prepare(transaction(s, [{ ...copy, name }]))).toThrow();
    expect(() => s.prepare(transaction(s, [{ op: "removeWorkbookSheet", sheetId: "other" }, { ...copy, beforeSheetId: "other" }]))).toThrow();
    expect(() => s.prepare(transaction(s, [{ op: "removeWorkbookSheet", sheetId: "sheet-1" }, copy]))).toThrow();
    expect(s.serialize()).toEqual(before); expect(s.history().undoSteps).toBe(0);
    expect(() => parseContract("document-transaction", transaction(s, [{ ...copy, unexpected: true } as DocumentOperation]))).toThrow();
  } finally { s.dispose(); }
});
test("copy preflights actual sparse cell capacity and can cancel after private destination cells have been produced", async () => {
  const s = createDocumentSession(scalarWorkbook(Array.from({ length: 1000 }, (_, row) => ({ row, column: 0, value: value(row) }))), { sessionLimits: { maxCandidates: 1 } });
  try {
    const before = s.serialize(), signal = { aborted: false }, tx = transaction(s, [{ op: "copyWorkbookSheet", sheetId: "sheet-1", name: "Copied" }]); let reached = 0;
    await expect(s.prepareAsync(tx, undefined, { signal, yield: async () => {}, onProgress(progress) { if (progress.phase === "cells" && progress.completed > 1300) { reached = progress.completed; signal.aborted = true; } } })).rejects.toMatchObject({ code: "OperationCancelled" });
    expect(reached).toBeGreaterThan(1300); expect(s.serialize()).toEqual(before); expect(s.history().undoSteps).toBe(0);
    const retry = s.prepare(tx); s.discard(retry); apply(s, [{ op: "setWorkbookCells", sheetId: "sheet-1", cells: [{ row: 0, column: 0, source: { kind: "value", value: value(7) } }] }]);
    const stale = s.prepare(transaction(s, tx.operations)); s.undo();
    expect(() => s.commit(stale)).toThrowError(expect.objectContaining({ code: "RevisionConflict" })); expect(s.snapshot().metadata.workbook.sheets).toHaveLength(1);
  } finally { s.dispose(); }
  const bounded = createDocumentSession(scalarWorkbook([{ row: 0, column: 0, value: value(1) }, { row: 999999, column: 100, value: value(2) }]), { workbookLimits: { maxCells: 3 } });
  try { const before = bounded.serialize(); expect(() => bounded.prepare(transaction(bounded, [{ op: "copyWorkbookSheet", sheetId: "sheet-1", name: "Copied" }]))).toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded" })); expect(bounded.serialize()).toEqual(before); } finally { bounded.dispose(); }
});
test("sheet copies are explicitly global proposal operations and release isolated candidate scopes after discard", () => {
  const s = createDocumentSession(scalarWorkbook([{ row: 0, column: 0, value: value(3) }]));
  try {
    const m = s.snapshot().manifest; apply(s, [{ op: "addAnnotation", annotation: { annotationId: "copy", documentId: m.documentId, baseRevision: m.revision, status: "open", instruction: "复制表", anchor: { kind: "cells", sheetId: "sheet-1", ranges: [{ from: { row: 0, column: 0 }, to: { row: 0, column: 0 } }] } } }]); s.clearHistory();
    const request = (global: boolean) => s.buildInteractionRequest({ kind: "workbook", annotationIds: ["copy"], ...(global ? { scope: { kind: "global" as const } } : {}) });
    for (const global of [false, true]) {
      const input = request(global), proposal: WorkbookEditProposal = { protocolVersion: 1, documentId: input.documentId, baseRevision: input.baseRevision, requestId: input.requestId, proposalId: "copy", annotationIds: ["copy"], summary: "复制工作表", operations: [{ op: "copyWorkbookSheet", sheetId: "sheet-1", name: "Copied" }] };
      if (!global) expect(() => s.prepareProposal(input, proposal)).toThrowError(expect.objectContaining({ code: "InvalidProposal" }));
      else { const before = s.snapshot(), candidate = s.prepareProposal(input, proposal); expect(candidate.preview.metadata.workbook.sheets).toHaveLength(2); expect(s.snapshot()).toEqual(before); s.discard(candidate); expect(s.snapshot()).toEqual(before); }
    }
  } finally { s.dispose(); }
});

test("copied rule identities, explicit self references and stop validation share the same final candidate", () => {
  const make = (input: number) => createDocumentSession(scalarWorkbook([{ row: 0, column: 0, value: value(input) }], metadata => {
    const sheet = metadata.workbook.sheets[0]!, at = { row: 0, column: 0 }, ranges = [{ from: at, to: at }];
    sheet.conditionalRules = [{ ruleId: "color", kind: "expression", expression: "=Sheet1!A1>0", origin: at, ranges, priority: 1, stopIfTrue: false, cssClass: "cell-default" }];
    sheet.dataValidations = [{ ruleId: "stop", kind: "custom", expression: "=Sheet1!A1>0", origin: at, ranges, allowBlank: false, showInputMessage: false, showErrorMessage: true, errorStyle: "stop", error: "positive", errorTitle: "", prompt: "", promptTitle: "" }];
  }));
  const s = make(3);
  try {
    const before = s.snapshot(); apply(s, [{ op: "copyWorkbookSheet", sheetId: "sheet-1", name: "Copied" }]); const copy = s.snapshot().metadata.workbook.sheets.at(-1)!;
    expect(copy.conditionalRules![0]!.ruleId).not.toBe("color"); expect(copy.dataValidations![0]!.ruleId).not.toBe("stop");
    expect(copy.conditionalRules![0]).toMatchObject({ expression: "=('Copied'!A1>0)" }); expect(copy.dataValidations![0]).toMatchObject({ expression: "=('Copied'!A1>0)" });
    const current = s.serialize(); expect(() => apply(s, [{ op: "setWorkbookCells", sheetId: copy.sheetId, cells: [{ row: 0, column: 0, source: { kind: "value", value: value(-1) } }] }])).toThrowError(expect.objectContaining({ code: "InvalidOperation" })); expect(s.serialize()).toEqual(current);
    s.undo(); expect(s.snapshot().metadata).toEqual(before.metadata);
  } finally { s.dispose(); }
  const invalid = make(-1);
  try { const before = invalid.serialize(); expect(() => apply(invalid, [{ op: "copyWorkbookSheet", sheetId: "sheet-1", name: "Copied" }])).toThrowError(expect.objectContaining({ code: "InvalidOperation" })); expect(invalid.serialize()).toEqual(before); expect(invalid.history().undoSteps).toBe(0); } finally { invalid.dispose(); }
});

test("copied aliases bind cloned definitions explicitly while outside global aliases retain their workbook binding", () => {
  const s = createDocumentSession(scalarWorkbook([{ row: 0, column: 0, value: value(3) }, { row: 0, column: 1, value: value(999), formula: "=Alias" }], metadata => {
    metadata.workbook.names = [name("Tax", "=Sheet1!$A$1"), name("Alias", "=Tax")];
  }));
  try {
    s.calculateSync(); expect(s.getCalculatedCell("sheet-1", 0, 1)).toMatchObject({ state: "valid", value: value(3) }); const before = s.snapshot();
    apply(s, [{ op: "copyWorkbookSheet", sheetId: "sheet-1", name: "Copied" }]);
    const copied = s.snapshot().metadata.workbook.sheets.at(-1)!;
    expect(s.snapshot().metadata.workbook.names!.find(name => name.nameId === "global-Alias")!.expression).toBe("=Tax");
    expect(s.snapshot().metadata.workbook.names!.find(name => name.sheetId === copied.sheetId && name.name === "Alias")!.expression).toBe("='Copied'!Tax");
    apply(s, [{ op: "setWorkbookCells", sheetId: copied.sheetId, cells: [{ row: 0, column: 0, source: { kind: "value", value: value(19) } }] }]); s.calculateSync();
    expect(s.getCalculatedCell("sheet-1", 0, 1)).toMatchObject({ state: "valid", value: value(3) });
    expect(s.getCalculatedCell(copied.sheetId, 0, 1)).toMatchObject({ state: "valid", value: value(19) }); expect(s.history().undoSteps).toBe(2);
    s.undo(); s.undo(); expect(s.snapshot().metadata).toEqual(before.metadata); expect(s.snapshot().manifest.chunks).toEqual(before.manifest.chunks);
  } finally { s.dispose(); }
});
