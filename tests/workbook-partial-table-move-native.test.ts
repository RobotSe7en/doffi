import { createHash } from "node:crypto";
import { execFileSync } from "node:child_process";
import { readFileSync } from "node:fs";
import { expect, test } from "vitest";
import { createDocumentSession, openDocument } from "@doffi/core";
import { connectWorkbookSession, serveWorkbookSession } from "@doffi/editor";
import { importXlsx, readDocumentZip, writeDocumentZip } from "@doffi/office";

interface Formula { address: string; expression: string; type: string; value: string }
interface Producer { file: string; sha256: string; observed: { tables: unknown[]; sheets: { tables: number; formulas: Formula[] }[] } }
const evidence = JSON.parse(readFileSync("docs/evidence/workbook-partial-table-move.json", "utf8")) as { sourceSha256: string; cases: {
  case: { name: string; source: string; target: string; sheet: number }; moved: Producer; reopened: Producer;
}[] };
const bytes = readFileSync("packages/office/tests/fixtures/workbook-axis-tables.xlsx");
const at = (address: string) => { const [, letters, digits] = /^([A-Z]+)(\d+)$/.exec(address)!; let column = 0; for (const letter of letters!) column = column * 26 + letter.charCodeAt(0) - 64; return { row: Number(digits) - 1, column: column - 1 }; };
const area = (address: string) => { const [from, to = from] = address.split(":"); return { from: at(from!), to: at(to!) }; };
const number = (value: number) => ({ type: "number" as const, value });
const expected: Record<string, { target: [string, number][]; observer: number[] }> = {
  "body-same": { target: [["K260", 20], ["K261", 60]], observer: [5, 160, 160] },
  "body-other": { target: [["K260", 20], ["K261", 60]], observer: [5, 160, 160] },
  "column-other": { target: [["H260", 20], ["H261", 60]], observer: [10, 160, 160] },
  "inside-body": { target: [["D2", 0], ["D3", 0], ["D4", 20], ["D5", 60]], observer: [5, 80, 80] },
  "total-other": { target: [["H260", 240]], observer: [10, 240, 240] },
};
test("five partial-table native cuts and actual reopen records preserve independent provenance", () => {
  expect(createHash("sha256").update(bytes).digest("hex")).toBe(evidence.sourceSha256);
  expect(execFileSync("python3", ["scripts/check-workbook-partial-table-move.py"], { encoding: "utf8" })).toContain("five independent");
});
test.each(evidence.cases)("partial native table cut $case.name retains definitions, maps actual inputs, reopens and undoes atomically", async item => {
  const source = openDocument(await importXlsx(bytes, { documentId: "partial-table-native" }), { kind: "workbook" }), s = createDocumentSession(source);
  try {
    const before = s.snapshot(), sheets = before.metadata.workbook.sheets, c = item.case, target = sheets[c.sheet]!.sheetId;
    const candidate = s.prepare({ protocolVersion: 1, documentId: before.manifest.documentId, baseRevision: 0, operations: [{ op: "moveWorkbookRange", sheetId: sheets[0]!.sheetId,
      range: area(c.source), targetSheetId: target, target: at(c.target), ...(c.name === "inside-body" ? { overwrittenContent: "clear" as const } : {}) }] });
    expect(s.snapshot()).toEqual(before); s.commit(candidate); s.calculateSync();
    expect(s.snapshot().metadata.workbook.sheets.map(sheet => sheet.tables)).toEqual(before.metadata.workbook.sheets.map(sheet => sheet.tables));
    for (const [address, value] of expected[c.name]!.target) { const point = at(address); expect(s.getCalculatedCell(target, point.row, point.column)).toMatchObject({ state: "valid", value: number(value) }); }
    for (const [column, value] of expected[c.name]!.observer.entries()) expect(s.getCalculatedCell(sheets[1]!.sheetId, 0, column)).toMatchObject({ value: number(value) });
    expect(openDocument(await readDocumentZip(await writeDocumentZip(s.serialize())), { kind: "workbook" }).snapshot()).toEqual(s.snapshot());
    expect(s.history().undoSteps).toBe(1); s.undo(); expect(s.snapshot().manifest.chunks).toEqual(before.manifest.chunks); expect(s.snapshot().metadata).toEqual(before.metadata);
    expect(source.snapshot()).toEqual(before);
  } finally { s.dispose(); }
});
for (const step of ["moved", "reopened"] as const) test.each(evidence.cases)(`actual producer ${step} $case.name preserves caches and recalculates to its independently reopened state`, async item => {
  const producer = readFileSync(item[step].file); expect(createHash("sha256").update(producer).digest("hex")).toBe(item[step].sha256);
  const s = createDocumentSession(openDocument(await importXlsx(producer, { documentId: "partial-table-producer" }), { kind: "workbook" }));
  try {
    const before = s.snapshot();
    for (const [index, sheet] of item[step].observed.sheets.entries()) for (const formula of sheet.formulas) {
      const point = at(formula.address), value = formula.type === "e" ? { type: "error", value: formula.value } : number(Number(formula.value));
      expect(s.getCell(before.metadata.workbook.sheets[index]!.sheetId, point.row, point.column)).toMatchObject({ value, formula: { cacheState: "imported" } });
    }
    s.calculateSync();
    for (const [index, sheet] of item.reopened.observed.sheets.entries()) for (const formula of sheet.formulas) {
      const point = at(formula.address), value = formula.type === "e" ? { type: "error", value: formula.value } : number(Number(formula.value));
      expect(s.getCalculatedCell(before.metadata.workbook.sheets[index]!.sheetId, point.row, point.column)).toMatchObject({ state: "valid", value });
    }
    s.commit(s.prepare({ protocolVersion: 1, documentId: before.manifest.documentId, baseRevision: 0, operations: [{ op: "setWorkbookCells", sheetId: before.metadata.workbook.sheets[1]!.sheetId,
      cells: [{ row: 400, column: 20, source: { kind: "value", value: { type: "text", value: "原生再编辑禰󠄀😀" } } }] }] }));
    expect(openDocument(await readDocumentZip(await writeDocumentZip(s.serialize())), { kind: "workbook" }).snapshot()).toEqual(s.snapshot());
    s.undo(); expect(s.snapshot().metadata).toEqual(before.metadata); expect(s.snapshot().manifest.chunks).toEqual(before.manifest.chunks);
  } finally { s.dispose(); }
});

test("the sole Worker discards, retries and commits partial table moves, then recalculates their moved inputs", async () => {
  const s = createDocumentSession(openDocument(await importXlsx(bytes, { documentId: "worker-partial-table" }), { kind: "workbook" })),
    channel = new MessageChannel(), server = serveWorkbookSession(channel.port1, s), client = await connectWorkbookSession(channel.port2);
  try {
    const before = s.snapshot(), [source, target] = before.metadata.workbook.sheets, state = client.state();
    const command = { kind: "transaction" as const, transaction: { protocolVersion: 1 as const, documentId: state.documentId, baseRevision: state.revision,
      operations: [{ op: "moveWorkbookRange" as const, sheetId: source!.sheetId, range: area("A2:D3"), targetSheetId: target!.sheetId, target: at("H260") }] } };
    const candidate = await client.prepare(command), view = await client.view(candidate);
    expect(s.snapshot()).toEqual(before); expect(view.cells.snapshot.metadata.workbook.sheets[0]!.tables).toEqual(before.metadata.workbook.sheets[0]!.tables);
    view.dispose(); await client.discard(candidate); expect(s.snapshot()).toEqual(before); await client.commit(await client.prepare(command));
    const current = client.state(); await client.commit(await client.prepare({ kind: "input", documentId: current.documentId, revision: current.revision, sheetId: target!.sheetId, at: at("I260"), text: "6" }));
    const updated = await client.view();
    const result = await updated.rules!.read({ documentId: current.documentId, revision: client.state().revision, sheetId: target!.sheetId, coordinates: [at("K260")], includeCurrent: true,
      maxCells: 100, maxRules: 100, maxValues: 100, maxSteps: 10000, maxCodeUnits: 10000 }, new AbortController().signal);
    expect(result.cells[0]!.current).toEqual({ kind: "value", value: number(60) }); updated.dispose(); expect(s.history().undoSteps).toBe(2);
    await client.undo(); await client.undo(); expect(s.snapshot().metadata).toEqual(before.metadata); expect(s.snapshot().manifest.chunks).toEqual(before.manifest.chunks);
    await expect(client.prepare(command)).rejects.toMatchObject({ code: "RevisionConflict" });
  } finally { client.dispose(); server.dispose(); s.dispose(); }
});
