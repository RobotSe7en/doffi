import { createHash } from "node:crypto";
import { execFileSync } from "node:child_process";
import { readFileSync } from "node:fs";
import { expect, test } from "vitest";
import { createDocumentSession, openDocument } from "@doffi/core";
import { connectWorkbookSession, createWorkbookFillTransaction, createWorkbookRangeCopyTransaction, serveWorkbookSession } from "@doffi/editor";
import { importXlsx, readDocumentZip, writeDocumentZip } from "@doffi/office";

const evidence = JSON.parse(readFileSync("docs/evidence/workbook-structured-fill.json", "utf8")) as {
  sourceSha256: string; producer: string; producerSha256: string;
  cases: { name: string; seed: string; range: string; formula: string; sdk: number[] }[];
  observed: { formulas: Record<string, Record<string, { formula: string; type: string; value: string }>> };
};
const bytes = readFileSync("packages/office/tests/fixtures/workbook-axis-tables.xlsx");
const at = (address: string) => { const [, letters, digits] = /^([A-Z]+)(\d+)$/.exec(address)!; let column = 0; for (const letter of letters!) column = column * 26 + letter.charCodeAt(0) - 64; return { row: Number(digits) - 1, column: column - 1 }; };
const area = (address: string) => { const [from, to = from] = address.split(":"); return { from: at(from!), to: at(to!) }; };
const number = (value: number) => ({ type: "number" as const, value });
test("structured-fill provenance retains native producer differences instead of relabeling them as Excel results", () => {
  expect(createHash("sha256").update(bytes).digest("hex")).toBe(evidence.sourceSha256);
  expect(execFileSync("python3", ["scripts/check-workbook-structured-fill.py"], { encoding: "utf8" })).toContain("seven independent");
});
test.each(evidence.cases)("native-table fill $name uses current values, remains revision bound and restores one-step source", async item => {
  const s = createDocumentSession(openDocument(await importXlsx(bytes, { documentId: "native-fill" }), { kind: "workbook" }));
  try {
    const sheetId = s.snapshot().metadata.workbook.sheets[1]!.sheetId, seed = at(item.seed), range = area(item.range);
    s.commit(s.prepare({ protocolVersion: 1, documentId: "native-fill", baseRevision: 0, operations: [{ op: "setWorkbookCells", sheetId, cells: [{ ...seed, source: { kind: "formula", expression: item.formula } }] }] }));
    s.clearHistory(); const before = s.snapshot(), transaction = await createWorkbookFillTransaction(s, sheetId, { from: seed, to: seed }, range);
    expect(s.snapshot()).toEqual(before); s.commit(s.prepare(transaction)); s.calculateSync();
    for (let index = 0; index < item.sdk.length; index++) {
      const row = range.from.row + (range.to.row > range.from.row ? index : 0), column = range.from.column + (range.to.column > range.from.column ? index : 0);
      expect(s.getCalculatedCell(sheetId, row, column)).toMatchObject({ state: "valid", value: number(item.sdk[index]!) });
    }
    expect(s.snapshot().metadata).toEqual(before.metadata); expect(s.history().undoSteps).toBe(1);
    expect(openDocument(await readDocumentZip(await writeDocumentZip(s.serialize())), { kind: "workbook" }).snapshot()).toEqual(s.snapshot());
    s.undo(); expect(s.snapshot().manifest.chunks).toEqual(before.manifest.chunks); expect(s.snapshot().metadata).toEqual(before.metadata);
    expect(() => s.prepare(transaction)).toThrowError(expect.objectContaining({ code: "RevisionConflict" }));
  } finally { s.dispose(); }
});
test("pinned Calc fill source retains its column and All-selector differences without rewriting native caches", async () => {
  const producer = readFileSync(evidence.producer); expect(createHash("sha256").update(producer).digest("hex")).toBe(evidence.producerSha256);
  const s = createDocumentSession(openDocument(await importXlsx(producer, { documentId: "native-fill-producer" }), { kind: "workbook" }));
  try {
    const before = s.snapshot(); s.calculateSync();
    for (const [name, formulas] of Object.entries(evidence.observed.formulas)) for (const [address, formula] of Object.entries(formulas)) {
      const point = at(address), sheetId = before.metadata.workbook.sheets.find(sheet => sheet.name === name)!.sheetId;
      const value = formula.type === "e" ? { type: "error", value: formula.value } : number(Number(formula.value));
      expect(s.getCell(sheetId, point.row, point.column)).toMatchObject({ formula: { cacheState: "imported" }, value });
      // The producer also records #N/A for SUM of #All/数量. SDK sums its
      // native data rectangle (text header and blank total do not add values).
      expect(s.getCalculatedCell(sheetId, point.row, point.column)).toMatchObject({ value: name === "Observer" && point.row === 21 ? number(10) : value });
    }
    expect(s.snapshot()).toEqual(before);
  } finally { s.dispose(); }
});
test("the sole Worker distinguishes structured fill from ordinary copy, discards candidates and preserves history", async () => {
  const s = createDocumentSession(openDocument(await importXlsx(bytes, { documentId: "worker-fill" }), { kind: "workbook" }));
  const channel = new MessageChannel(), server = serveWorkbookSession(channel.port1, s), client = await connectWorkbookSession(channel.port2);
  try {
    const sheetId = s.snapshot().metadata.workbook.sheets[1]!.sheetId, state = client.state();
    await client.commit(await client.prepare({ kind: "input", documentId: state.documentId, revision: state.revision, sheetId, at: at("H260"), text: "=SUM(SalesTable[数量])" }));
    const seeded = s.snapshot(), current = client.state(), command = { kind: "fill" as const, documentId: current.documentId, revision: current.revision, sheetId, source: area("H260"), target: area("H260:L260") };
    const candidate = await client.prepare(command); expect(s.snapshot()).toEqual(seeded); await client.discard(candidate); expect(s.snapshot()).toEqual(seeded);
    await client.commit(await client.prepare(command)); const view = await client.view();
    const values = await view.rules!.read({ documentId: current.documentId, revision: client.state().revision, sheetId, coordinates: [at("I260"), at("J260"), at("K260"), at("L260")], includeCurrent: true,
      maxCells: 100, maxRules: 100, maxValues: 100, maxSteps: 10000, maxCodeUnits: 10000 }, new AbortController().signal);
    expect(values.cells.map(cell => cell.current)).toEqual([100, 240, 0, 10].map(value => ({ kind: "value", value: number(value) }))); view.dispose();
    await client.undo(); expect(s.snapshot().manifest.chunks).toEqual(seeded.manifest.chunks);
    s.commit(s.prepare(await createWorkbookRangeCopyTransaction(s, sheetId, area("H260"), at("I260")))); s.calculateSync();
    expect(s.getCalculatedCell(sheetId, 259, 8)).toMatchObject({ value: number(10) });
    await expect(client.prepare(command)).rejects.toMatchObject({ code: "RevisionConflict" });
    s.undo(); s.undo(); expect(s.history().undoSteps).toBe(0);
  } finally { client.dispose(); server.dispose(); s.dispose(); }
});
