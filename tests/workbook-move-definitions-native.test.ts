import { createHash } from "node:crypto";
import { execFileSync } from "node:child_process";
import { readFileSync } from "node:fs";
import { expect, test } from "vitest";
import { createDocumentSession, openDocument, parseWorkbookFormula } from "@doffi/core";
import { connectWorkbookSession, serveWorkbookSession } from "@doffi/editor";
import { importXlsx, readDocumentZip, writeDocumentZip } from "@doffi/office";

const bytes = readFileSync("packages/office/tests/fixtures/workbook-axis-tables.xlsx");
const evidence = JSON.parse(readFileSync("docs/evidence/workbook-move-definitions.json", "utf8")) as { sourceSha256: string; cases: { name: string; producerSha256: string;
  observed: { tables: { name: string; range: string; columns: string[] }[]; sheets: { tables: number; formulas: { address: string; expression: string; type: string; value: string }[] }[] } }[] };
function at(address: string) {
  const match = /^([A-Z]+)([1-9][0-9]*)$/.exec(address)!;
  let column = 0; for (const letter of match[1]!) column = column * 26 + letter.charCodeAt(0) - 64;
  return { row: Number(match[2]) - 1, column: column - 1 };
}
const area = { from: { row: 259, column: 7 }, to: { row: 264, column: 10 } };
test("native whole-table cut observations pin unchanged independent input and actual producer bytes", () => {
  expect(createHash("sha256").update(bytes).digest("hex")).toBe(evidence.sourceSha256);
  expect(execFileSync("python3", ["scripts/check-workbook-move-definitions.py"], { encoding: "utf8" })).toContain("two independent");
});
for (const transport of ["direct", "Worker"] as const) test.each(evidence.cases)(`${transport} $name retains native current formulas and table identity through isolated candidate and undo`, async item => {
  const source = openDocument(await importXlsx(bytes, { documentId: "native-table-move" }), { kind: "workbook" }), session = createDocumentSession(source);
  let cleanup = () => {};
  try {
    const before = source.snapshot(), sheetId = before.metadata.workbook.sheets[0]!.sheetId, targetIndex = item.name === "other-sheet" ? 1 : 0,
      targetSheetId = before.metadata.workbook.sheets[targetIndex]!.sheetId, originalTable = before.metadata.workbook.sheets[0]!.tables![0]!;
    const tx = { protocolVersion: 1 as const, documentId: before.manifest.documentId, baseRevision: before.manifest.revision,
      operations: [{ op: "moveWorkbookRange" as const, sheetId, range: { from: { row: 0, column: 0 }, to: { row: 5, column: 3 } }, targetSheetId, target: area.from }] };
    if (transport === "direct") {
      const candidate = session.prepare(tx); expect(session.snapshot()).toEqual(before); session.commit(candidate);
    } else {
      const channel = new MessageChannel(), server = serveWorkbookSession(channel.port1, session), client = await connectWorkbookSession(channel.port2);
      cleanup = () => { client.dispose(); server.dispose(); };
      const candidate = await client.prepare({ kind: "transaction", transaction: tx }), view = await client.view(candidate);
      expect(view.cells.snapshot.metadata.workbook.sheets[targetIndex]!.tables![0]!.range).toEqual(area); view.dispose();
      expect(session.snapshot()).toEqual(before); await client.discard(candidate); expect(session.snapshot()).toEqual(before);
      await client.commit(await client.prepare({ kind: "transaction", transaction: tx }));
    }
    session.calculateSync();
    const table = session.snapshot().metadata.workbook.sheets[targetIndex]!.tables![0]!;
    expect(table).toMatchObject({ tableId: originalTable.tableId, name: originalTable.name, range: area, style: originalTable.style });
    expect(table.columns.map(column => column.columnId)).toEqual(originalTable.columns.map(column => column.columnId));
    expect(table.columns.map(column => column.name)).toEqual(item.observed.tables[0]!.columns);
    expect(table.columns[3]!.calculatedFormula).toBeDefined(); expect(table.autoFilter!.range).toEqual({ from: area.from, to: { row: 263, column: 10 } });
    for (const [index, expected] of item.observed.sheets.entries()) {
      const sheet = session.snapshot().metadata.workbook.sheets[index]!; expect(sheet.tables?.length ?? 0).toBe(expected.tables);
      for (const formula of expected.formulas) {
        const coordinate = at(formula.address), cell = session.getCell(sheet.sheetId, coordinate.row, coordinate.column)!;
        expect(parseWorkbookFormula(cell.formula!.expression, coordinate)).toEqual(parseWorkbookFormula("=" + formula.expression, coordinate));
        expect(session.getCalculatedCell(sheet.sheetId, coordinate.row, coordinate.column)).toMatchObject({ state: "valid", value: { type: formula.type === "e" ? "error" : "number", value: formula.type === "e" ? formula.value : Number(formula.value) } });
      }
    }
    expect(openDocument(await readDocumentZip(await writeDocumentZip(session.serialize())), { kind: "workbook" }).snapshot()).toEqual(session.snapshot());
    expect(session.history().undoSteps).toBe(1); session.undo(); expect(session.snapshot().metadata).toEqual(before.metadata); expect(session.snapshot().manifest.chunks).toEqual(before.manifest.chunks);
    expect(source.snapshot()).toEqual(before);
    const producer = readFileSync("packages/office/tests/fixtures/workbook-move-definitions-" + item.name + "-calc.xlsx");
    expect(createHash("sha256").update(producer).digest("hex")).toBe(item.producerSha256);
    const reopened = openDocument(await importXlsx(producer, { documentId: "native-producer-move" }), { kind: "workbook" });
    expect(reopened.snapshot().metadata.workbook.sheets[targetIndex]!.tables![0]!.range).toEqual(area);
  } finally { cleanup(); session.dispose(); }
});
