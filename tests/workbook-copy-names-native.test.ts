import { readFileSync } from "node:fs";
import { expect, test } from "vitest";
import { createDocumentSession, DEFAULT_WORKBOOK_RULE_LIMITS, openDocument, parseWorkbookFormula } from "@doffi/core";
import type { DocumentOperation, WorkbookDocumentMetadata, WorkbookSession } from "@doffi/core";
import { connectWorkbookSession, createWorkbookRangeCopyTransaction, serveWorkbookSession } from "@doffi/editor";
import { importXlsx, readDocumentZip, writeDocumentZip } from "@doffi/office";

const bytes = readFileSync("packages/office/tests/fixtures/workbook-name-prefix.xlsx"), area = { from: { row: 4, column: 1 }, to: { row: 7, column: 1 } };
const evidence = JSON.parse(readFileSync("docs/evidence/workbook-name-prefix.json", "utf8")) as { cases: { name: string; observed: { sheets: { address: string; expression: string; value: string }[][]; names: { name: string; sheet: string | null; expression: string }[] } }[] };
function apply(s: WorkbookSession, operations: DocumentOperation[]) { const m = s.snapshot().manifest; return s.commit(s.prepare({ protocolVersion: 1, documentId: m.documentId, baseRevision: m.revision, operations })); }
async function setup() {
  const s = createDocumentSession(openDocument(await importXlsx(bytes, { documentId: "native-copy-names" }), { kind: "workbook" })), [source, target] = s.snapshot().metadata.workbook.sheets;
  apply(s, [{ op: "setWorkbookCells", sheetId: target!.sheetId, cells: [{ row: 0, column: 9, source: { kind: "formula", expression: "=Rate" } }] }]); s.clearHistory();
  return { s, source: source!.sheetId, target: target!.sheetId };
}
test("native cross-sheet name copy clones only missing local definitions and keeps global/local associations through reopen and one undo", async () => {
  const { s, source, target } = await setup();
  try {
    const before = s.snapshot(), originalNames = before.metadata.workbook.names!;
    const tx = await createWorkbookRangeCopyTransaction(s, source, area, area.from, { targetSheetId: target }); expect(s.snapshot()).toEqual(before);
    const created = tx.operations.filter(op => op.op === "setWorkbookDefinedName"); expect(created).toHaveLength(1);
    s.commit(s.prepare(tx)); s.calculateSync();
    const native = evidence.cases.find(item => item.name === "copy-other")!, name = s.snapshot().metadata.workbook.names!.find(name => name.sheetId === target && name.name === "Rate")!;
    expect(name.nameId).not.toBe(originalNames.find(name => name.name === "Rate")!.nameId);
    expect(parseWorkbookFormula(name.expression, name.origin)).toEqual(parseWorkbookFormula("=" + native.observed.names.find(name => name.name === "Rate" && name.sheet === "1")!.expression, name.origin));
    for (const cell of native.observed.sheets[1]!) {
      const row = Number(cell.address.slice(1)) - 1; expect(s.getCalculatedCell(target, row, 1)).toMatchObject({ state: "valid", value: { type: "number", value: Number(cell.value) } });
    }
    expect(s.getCell(target, 6, 1)!.formula!.expression).toBe("=[0]!Tax"); expect(s.getCell(target, 7, 1)!.formula!.expression).toBe("=[0]!Tax");
    expect(s.getCalculatedCell(target, 0, 9)).toMatchObject({ value: { type: "number", value: 100 } }); expect(s.getCell(target, 0, 9)!.formula!.expression).toBe("=Rate");
    expect(s.snapshot().metadata.workbook.names!.filter(name => originalNames.some(old => old.nameId === name.nameId))).toEqual(originalNames);
    expect(s.snapshot().manifest.chunks.filter(chunk => chunk.sheetId === source)).toEqual(before.manifest.chunks.filter(chunk => chunk.sheetId === source));
    expect(openDocument(await readDocumentZip(await writeDocumentZip(s.serialize())), { kind: "workbook" }).snapshot()).toEqual(s.snapshot());
    expect(s.history().undoSteps).toBe(1); const repeated = await createWorkbookRangeCopyTransaction(s, source, area, area.from, { targetSheetId: target });
    expect(repeated.operations.some(op => op.op === "setWorkbookDefinedName")).toBe(false); expect(s.commit(s.prepare(repeated)).changed).toBe(false); expect(s.history().undoSteps).toBe(1);
    s.undo(); expect(s.snapshot().manifest.chunks).toEqual(before.manifest.chunks); expect(s.snapshot().metadata).toEqual(before.metadata);
  } finally { s.dispose(); }
});
test("selective native formula values do not introduce local definitions or alter unrelated name consumers", async () => {
  const { s, source, target } = await setup();
  try {
    const before = s.snapshot(), calculation = s.calculateSync();
    const tx = await createWorkbookRangeCopyTransaction(s, source, area, area.from, { targetSheetId: target, mode: "values", calculation });
    expect(tx.operations.some(op => op.op === "setWorkbookDefinedName")).toBe(false); s.commit(s.prepare(tx)); s.calculateSync();
    expect([4, 5, 6, 7].map(row => s.getCell(target, row, 1)!.value)).toEqual([17, 70, 7, 7].map(value => ({ type: "number", value })));
    expect(s.getCalculatedCell(target, 0, 9)).toMatchObject({ value: { type: "error", value: "#NAME?" } }); expect(s.snapshot().metadata.workbook.names).toEqual(before.metadata.workbook.names);
  } finally { s.dispose(); }
});
test("a newly bound outside consumer cannot bypass its stop validation when local definitions are copied", async () => {
  const original = await setup(), files = new Map(original.s.serialize()); original.s.dispose();
  const metadata = JSON.parse(new TextDecoder().decode(files.get("metadata.json")!)) as WorkbookDocumentMetadata;
  const sheet = metadata.workbook.sheets.find(sheet => sheet.sheetId === original.target)!;
  sheet.dataValidations = [{ ruleId: "outside-stop", kind: "whole", operator: "greaterThan", origin: { row: 0, column: 9 }, ranges: [{ from: { row: 0, column: 9 }, to: { row: 0, column: 9 } }], formulas: ["=200"], allowBlank: false, showInputMessage: false, showErrorMessage: true, errorStyle: "stop", errorTitle: "", error: "", promptTitle: "", prompt: "" }];
  files.set("metadata.json", new TextEncoder().encode(JSON.stringify(metadata))); const s = createDocumentSession(openDocument(files, { kind: "workbook" }));
  try {
    const before = s.serialize(), tx = await createWorkbookRangeCopyTransaction(s, original.source, area, area.from, { targetSheetId: original.target });
    expect(() => s.prepare(tx)).toThrowError(expect.objectContaining({ code: "InvalidOperation" })); expect(s.serialize()).toEqual(before); expect(s.history().undoSteps).toBe(0);
  } finally { s.dispose(); }
});
test("the sole Worker returns a private current name-copy preview and discards all cloned definitions before retry/commit", async () => {
  const { s, source, target } = await setup(), channel = new MessageChannel(), server = serveWorkbookSession(channel.port1, s), client = await connectWorkbookSession(channel.port2);
  try {
    const before = s.snapshot(), state = client.state(), command = { kind: "copy" as const, documentId: state.documentId, revision: state.revision, sheetId: source, source: area, target: area.from, targetSheetId: target };
    const candidate = await client.prepare(command), view = await client.view(candidate);
    const current = await view.rules!.read({ ...DEFAULT_WORKBOOK_RULE_LIMITS, documentId: state.documentId, revision: state.revision, sheetId: target, coordinates: [4, 5, 6, 7].map(row => ({ row, column: 1 })), includeCurrent: true }, new AbortController().signal);
    expect(current.cells.map(cell => cell.current)).toEqual([107, 700, 7, 7].map(value => ({ kind: "value", value: { type: "number", value } })));
    expect(s.snapshot()).toEqual(before); view.dispose(); await client.discard(candidate); expect(s.snapshot()).toEqual(before);
    await client.commit(await client.prepare(command)); expect(s.history().undoSteps).toBe(1); await client.undo(); expect(s.snapshot().metadata).toEqual(before.metadata); expect(s.snapshot().manifest.chunks).toEqual(before.manifest.chunks);
  } finally { client.dispose(); server.dispose(); s.dispose(); }
});
