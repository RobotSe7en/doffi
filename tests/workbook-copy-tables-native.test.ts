import { createHash } from "node:crypto";
import { execFileSync } from "node:child_process";
import { readFileSync } from "node:fs";
import { expect, test } from "vitest";
import { createDocumentSession, openDocument } from "@doffi/core";
import { connectWorkbookSession, createWorkbookRangeCopyTransaction, serveWorkbookSession } from "@doffi/editor";
import { importXlsx, readDocumentZip, writeDocumentZip } from "@doffi/office";

interface Producer { file: string; sha256: string; observed: { sheets: { formulas: { address: string; type: string; value: string }[] }[] } }
const evidence = JSON.parse(readFileSync("docs/evidence/workbook-copy-tables.json", "utf8")) as { sourceSha256: string; cases: { name: string; copied: Producer; reopened: Producer }[] };
const bytes = readFileSync("packages/office/tests/fixtures/workbook-axis-tables.xlsx");
const at = (address: string) => { const [, letters, digits] = /^([A-Z]+)(\d+)$/.exec(address)!; let column = 0; for (const letter of letters!) column = column * 26 + letter.charCodeAt(0) - 64; return { row: Number(digits) - 1, column: column - 1 }; };
const sourceRange = { from: at("A1"), to: at("D6") }, number = (value: number) => ({ type: "number" as const, value });
test("whole table copy evidence retains immutable source and producer bytes with its independent behavior", () => {
  expect(createHash("sha256").update(bytes).digest("hex")).toBe(evidence.sourceSha256);
  expect(execFileSync("python3", ["scripts/check-workbook-copy-tables.py"], { encoding: "utf8" })).toContain("two independent");
});
test.each(evidence.cases)("whole table contents copy $name preserves native source, copies identities and recalculates current targets", async item => {
  const s = createDocumentSession(openDocument(await importXlsx(bytes, { documentId: "table-copy-native" }), { kind: "workbook" }));
  try {
    const before = s.snapshot(), sheets = before.metadata.workbook.sheets, sheetId = sheets[0]!.sheetId, targetSheetId = sheets[item.name === "other-sheet" ? 1 : 0]!.sheetId;
    const transaction = await createWorkbookRangeCopyTransaction(s, sheetId, sourceRange, at("H260"), { targetSheetId });
    expect(s.snapshot()).toEqual(before); s.commit(s.prepare(transaction)); s.calculateSync();
    const copied = s.snapshot().metadata.workbook.sheets.find(sheet => sheet.sheetId === targetSheetId)!.tables!.at(-1)!, original = sheets[0]!.tables![0]!;
    expect(copied.name).toBe("SalesTable_2"); expect(copied.tableId).not.toBe(original.tableId); expect(copied.range).toEqual({ from: at("H260"), to: at("K265") });
    expect(copied.columns[3]!.calculatedFormula).toContain("SalesTable_2"); expect(copied.autoFilter?.range).toEqual({ from: at("H260"), to: at("K264") });
    expect(s.snapshot().metadata.workbook.sheets[0]!.tables![0]).toEqual(original);
    for (const [address, value] of [["K261", 20], ["K262", 60], ["K263", 120], ["K264", 40], ["K265", 240]] as const) { const p = at(address); expect(s.getCalculatedCell(targetSheetId, p.row, p.column)).toMatchObject({ state: "valid", value: number(value) }); }
    expect(openDocument(await readDocumentZip(await writeDocumentZip(s.serialize())), { kind: "workbook" }).snapshot()).toEqual(s.snapshot());
    expect(s.history().undoSteps).toBe(1);
    s.commit(s.prepare({ protocolVersion: 1, documentId: before.manifest.documentId, baseRevision: 1, operations: [{ op: "setWorkbookCells", sheetId: targetSheetId,
      cells: [{ ...at("I261"), source: { kind: "value", value: number(6) } }] }] })); s.calculateSync();
    expect(s.getCalculatedCell(targetSheetId, 260, 10)).toMatchObject({ value: number(60) }); expect(s.getCalculatedCell(targetSheetId, 264, 10)).toMatchObject({ value: number(280) });
    expect(s.getCalculatedCell(sheetId, 5, 3)).toMatchObject({ value: number(240) });
    s.undo(); s.undo(); expect(s.snapshot().metadata).toEqual(before.metadata); expect(s.snapshot().manifest.chunks).toEqual(before.manifest.chunks);
  } finally { s.dispose(); }
});
for (const step of ["copied", "reopened"] as const) test.each(evidence.cases)(`native producer ${step} $name retains its missing copy identity and original structured errors`, async item => {
  const producer = readFileSync(item[step].file); expect(createHash("sha256").update(producer).digest("hex")).toBe(item[step].sha256);
  const s = createDocumentSession(openDocument(await importXlsx(producer, { documentId: "table-copy-producer" }), { kind: "workbook" }));
  try {
    const before = s.snapshot(); expect(before.metadata.workbook.sheets.reduce((sum, sheet) => sum + (sheet.tables?.length ?? 0), 0)).toBe(1); s.calculateSync();
    for (const [index, sheet] of item[step].observed.sheets.entries()) for (const formula of sheet.formulas) {
      const point = at(formula.address), sheetId = before.metadata.workbook.sheets[index]!.sheetId, value = formula.type === "e" ? { type: "error", value: formula.value } : number(Number(formula.value));
      expect(s.getCell(sheetId, point.row, point.column)).toMatchObject({ value, formula: { cacheState: "imported" } });
      expect(s.getCalculatedCell(sheetId, point.row, point.column)).toMatchObject({ value });
    }
    expect(s.snapshot()).toEqual(before);
  } finally { s.dispose(); }
});
test("sole Worker table copy discards safely, commits one definition/cell batch and preserves target inputs and history", async () => {
  const s = createDocumentSession(openDocument(await importXlsx(bytes, { documentId: "worker-table-copy" }), { kind: "workbook" })),
    channel = new MessageChannel(), server = serveWorkbookSession(channel.port1, s), client = await connectWorkbookSession(channel.port2);
  try {
    const before = s.snapshot(), [source, target] = before.metadata.workbook.sheets, state = client.state();
    const command = { kind: "copy" as const, documentId: state.documentId, revision: state.revision, sheetId: source!.sheetId, source: sourceRange, target: at("H260"), targetSheetId: target!.sheetId };
    const candidate = await client.prepare(command), view = await client.view(candidate);
    expect(view.cells.snapshot.metadata.workbook.sheets[1]!.tables![0]!.name).toBe("SalesTable_2"); view.dispose(); await client.discard(candidate); expect(s.snapshot()).toEqual(before);
    await client.commit(await client.prepare(command)); expect(s.history().undoSteps).toBe(1);
    const current = client.state(); await client.commit(await client.prepare({ kind: "input", documentId: current.documentId, revision: current.revision, sheetId: target!.sheetId, at: at("I261"), text: "6" }));
    const updated = await client.view(), result = await updated.rules!.read({ documentId: current.documentId, revision: client.state().revision, sheetId: target!.sheetId, coordinates: [at("K261"), at("K265")], includeCurrent: true,
      maxCells: 100, maxRules: 100, maxValues: 100, maxSteps: 10000, maxCodeUnits: 10000 }, new AbortController().signal);
    expect(result.cells.map(cell => cell.current)).toEqual([60, 280].map(value => ({ kind: "value", value: number(value) }))); updated.dispose();
    await client.undo(); await client.undo(); expect(s.snapshot().metadata).toEqual(before.metadata); expect(s.snapshot().manifest.chunks).toEqual(before.manifest.chunks);
    await expect(client.prepare(command)).rejects.toMatchObject({ code: "RevisionConflict" });
  } finally { client.dispose(); server.dispose(); s.dispose(); }
});
