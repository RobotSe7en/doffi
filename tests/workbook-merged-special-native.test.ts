import { createHash } from "node:crypto";
import { execFileSync } from "node:child_process";
import { readFileSync } from "node:fs";
import { expect, test } from "vitest";
import { createDocumentSession, openDocument } from "@doffi/core";
import { connectWorkbookSession, createWorkbookFillTransaction, createWorkbookRangeCopyTransaction, serveWorkbookSession } from "@doffi/editor";
import { importXlsx, readDocumentZip, writeDocumentZip } from "@doffi/office";

interface Case { mode: "contents" | "values" | "formulas" | "formats" | "series"; target?: "plain" | "matching" | "blank"; skipBlanks?: boolean; axis?: "row" | "column" }
interface Observed { merges: string[]; cells: { address: string; type: string; value: string; formula: string | null }[] }
const evidence = JSON.parse(readFileSync("docs/evidence/workbook-merged-special.json", "utf8")) as { sourceSha256: string; cases: { case: Case; producerSha256: string; observed: Observed[] }[] };
const bytes = readFileSync("packages/office/tests/fixtures/workbook-merged-special.xlsx"), range = (row: number, column: number, lastRow = row, lastColumn = column) => ({ from: { row, column }, to: { row: lastRow, column: lastColumn } });
const at = (address: string) => { const [, letters, digits] = /^([A-Z]+)(\d+)$/.exec(address)!; let column = 0; for (const letter of letters!) column = column * 26 + letter.charCodeAt(0) - 64; return { row: Number(digits) - 1, column: column - 1 }; };
const area = (address: string) => { const [from, to = from] = address.split(":"); return { from: at(from!), to: at(to!) }; };
const key = (c: Case) => c.mode === "series" ? "series-" + c.axis : c.mode + "-" + Number(c.skipBlanks) + "-" + c.target;
const value = (cell: Observed["cells"][number]) => cell.type === "text" ? { type: "text", value: cell.value } : cell.type === "e" ? { type: "error", value: cell.value } : { type: "number", value: Number(cell.value) };

test("twenty merged-special native cases have independent source and producer provenance", () => {
  expect(createHash("sha256").update(bytes).digest("hex")).toBe(evidence.sourceSha256);
  expect(execFileSync("python3", ["scripts/check-workbook-merged-special.py"], { encoding: "utf8" })).toContain("twenty independent");
});
test.each(evidence.cases)("current merged special $case matches native content, preserves selected formatting and rejects implicit content loss", async item => {
  const source = openDocument(await importXlsx(bytes, { documentId: "merged-special-native" }), { kind: "workbook" }), s = createDocumentSession(source);
  try {
    const before = s.snapshot(), c = item.case, sheetId = before.metadata.workbook.sheets[0]!.sheetId;
    const transaction = c.mode === "series" ? await createWorkbookFillTransaction(s, sheetId,
      c.axis === "row" ? range(0, 11, 3, 12) : range(0, 14, 1, 17), c.axis === "row" ? range(4, 11, 9, 12) : range(0, 18, 1, 23), { mode: "series", axis: c.axis! })
      : await createWorkbookRangeCopyTransaction(s, sheetId, range(255, 0, 257, 1), c.target === "matching" ? at("D256") : c.target === "blank" ? at("H1") : at("D1"),
        { mode: c.mode, skipBlanks: c.skipBlanks!, calculation: s.calculateSync() });
    const rejects = c.target === "plain" && (c.mode === "formats" || c.mode === "contents" && c.skipBlanks);
    if (rejects) {
      expect(() => s.prepare(transaction)).toThrowError(expect.objectContaining({ code: "InvalidOperation" }));
      expect(s.snapshot()).toEqual(before); expect(s.history().undoSteps).toBe(0); return;
    }
    const candidate = s.prepare(transaction); expect(s.snapshot()).toEqual(before); s.commit(candidate); s.calculateSync();
    for (const cell of item.observed[0]!.cells) {
      const position = at(cell.address);
      expect(s.getCalculatedCell(sheetId, position.row, position.column)).toMatchObject({ state: "valid", value: value(cell) });
    }
    const merges = item.observed[0]!.merges.map(area);
    if (c.target === "matching" && (c.mode === "values" || c.mode === "formulas")) merges.push(range(255, 3, 256, 4));
    expect(s.snapshot().metadata.workbook.sheets[0]!.merges).toEqual(expect.arrayContaining(merges));
    expect(s.snapshot().metadata.workbook.sheets[0]!.merges).toHaveLength(merges.length);
    expect(s.getCell(sheetId, 257, 0)!.value).toEqual({ type: "number", value: 999 });
    expect(openDocument(await readDocumentZip(await writeDocumentZip(s.serialize())), { kind: "workbook" }).snapshot()).toEqual(s.snapshot());
    expect(s.history().undoSteps).toBe(1); s.undo(); expect(s.snapshot().metadata).toEqual(before.metadata); expect(s.snapshot().manifest.chunks).toEqual(before.manifest.chunks);
    expect(source.snapshot()).toEqual(before);
  } finally { s.dispose(); }
});

test.each(evidence.cases)("actual merged-special producer $case retains all native cells, supports editing, package reopen and undo", async item => {
  const producer = readFileSync("packages/office/tests/fixtures/workbook-merged-special-" + key(item.case) + "-calc.xlsx");
  expect(createHash("sha256").update(producer).digest("hex")).toBe(item.producerSha256);
  const s = createDocumentSession(openDocument(await importXlsx(producer, { documentId: "merged-special-producer" }), { kind: "workbook" }));
  try {
    const before = s.snapshot(), sheetId = before.metadata.workbook.sheets[0]!.sheetId;
    for (const cell of item.observed[0]!.cells) { const position = at(cell.address); expect(s.getCell(sheetId, position.row, position.column)).toMatchObject({ value: value(cell) }); }
    s.commit(s.prepare({ protocolVersion: 1, documentId: before.manifest.documentId, baseRevision: 0, operations: [{ op: "setWorkbookCells", sheetId,
      cells: [{ row: 270, column: 25, source: { kind: "value", value: { type: "text", value: "原生编辑禰󠄀😀" } } }] }] }));
    expect(openDocument(await readDocumentZip(await writeDocumentZip(s.serialize())), { kind: "workbook" }).snapshot()).toEqual(s.snapshot());
    s.undo(); expect(s.snapshot().manifest.chunks).toEqual(before.manifest.chunks); expect(s.snapshot().metadata).toEqual(before.metadata);
  } finally { s.dispose(); }
});

test("the sole Worker owns selective merged candidates and fill-format choices, with discard, conflict and one undo", async () => {
  const s = createDocumentSession(openDocument(await importXlsx(bytes, { documentId: "worker-merged-special" }), { kind: "workbook" })),
    channel = new MessageChannel(), server = serveWorkbookSession(channel.port1, s), client = await connectWorkbookSession(channel.port2);
  try {
    const before = s.snapshot(), state = client.state(), sheetId = before.metadata.workbook.sheets[0]!.sheetId;
    const command = { kind: "copy" as const, documentId: state.documentId, revision: state.revision, sheetId, source: range(255, 0, 257, 1), target: at("D256"), mode: "values" as const, skipBlanks: true };
    const candidate = await client.prepare(command), view = await client.view(candidate);
    expect(s.snapshot()).toEqual(before); expect(view.cells.snapshot.metadata.workbook.sheets[0]!.merges).toContainEqual(range(255, 3, 256, 4));
    view.dispose(); await client.discard(candidate); expect(s.snapshot()).toEqual(before);
    await client.commit(await client.prepare(command)); expect(s.getCell(sheetId, 257, 3)).toMatchObject({ value: { type: "number", value: 14 } });
    await expect(client.prepare(command)).rejects.toMatchObject({ code: "RevisionConflict" }); await client.undo();
    const current = client.state(), fill = await client.prepare({ kind: "fill", documentId: current.documentId, revision: current.revision, sheetId,
      source: range(0, 11, 3, 12), target: range(4, 11, 9, 12), mode: "series", axis: "row", format: "preserve" });
    await client.commit(fill); expect(s.getCell(sheetId, 8, 11)).toMatchObject({ value: { type: "number", value: 50 } });
    expect(s.snapshot().metadata.workbook.sheets[0]!.merges).toEqual(before.metadata.workbook.sheets[0]!.merges); expect(s.history().undoSteps).toBe(1);
    await client.undo(); expect(s.snapshot().metadata).toEqual(before.metadata); expect(s.snapshot().manifest.chunks).toEqual(before.manifest.chunks);
  } finally { client.dispose(); server.dispose(); s.dispose(); }
});
