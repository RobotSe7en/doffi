import { createHash } from "node:crypto";
import { execFileSync } from "node:child_process";
import { readFileSync } from "node:fs";
import { expect, test } from "vitest";
import { createDocumentSession, openDocument, parseWorkbookFormula } from "@doffi/core";
import { createWorkbookFillTransaction, createWorkbookRangeCopyTransaction } from "@doffi/editor";
import { importXlsx, readDocumentZip, writeDocumentZip } from "@doffi/office";

type NativeSheet = { merges: string[]; cells: { address: string; type: string; value: string; formula: string | null }[] };
const evidence = JSON.parse(readFileSync("docs/evidence/workbook-copy-merges.json", "utf8")) as {
  sourceSha256: string; cases: { name: string; producerSha256: string; observed: NativeSheet[] }[];
};
const bytes = readFileSync("packages/office/tests/fixtures/workbook-copy-merges.xlsx");
function at(address: string) {
  const match = /^([A-Z]+)([1-9][0-9]*)$/.exec(address)!;
  let column = 0; for (const letter of match[1]!) column = column * 26 + letter.charCodeAt(0) - 64;
  return { row: Number(match[2]) - 1, column: column - 1 };
}
function range(address: string) { const [from, to = from] = address.split(":"); return { from: at(from!), to: at(to!) }; }

test("merged-copy sources and actual producer edits are independent and reproducible", () => {
  expect(createHash("sha256").update(bytes).digest("hex")).toBe(evidence.sourceSha256);
  expect(execFileSync("python3", ["scripts/check-workbook-copy-merges.py"], { encoding: "utf8" })).toContain("six independent");
});
test.each(evidence.cases)("$name preserves native topology/current results, isolated source and one history step", async item => {
  const source = openDocument(await importXlsx(bytes, { documentId: "copy-merges-native" }), { kind: "workbook" }), session = createDocumentSession(source);
  try {
    const before = source.snapshot(), sheetId = before.metadata.workbook.sheets[0]!.sheetId;
    const transaction = () => item.name.startsWith("fill-")
      ? createWorkbookFillTransaction(session, sheetId, range("A256:B257"), range(item.name === "fill-down" ? "A256:B261" : "A252:B257"))
      : createWorkbookRangeCopyTransaction(session, sheetId, range("A256:B258"), at(item.name === "copy-matching" ? "D256" : item.name === "copy-incompatible" ? "H256" : "D1"),
        { targetSheetId: before.metadata.workbook.sheets[item.name === "copy-other" ? 1 : 0]!.sheetId });
    if (item.name === "copy-incompatible") {
      await expect(transaction()).rejects.toMatchObject({ code: "InvalidOperation" });
      expect(session.serialize()).toEqual(source.serialize()); expect(session.history().undoSteps).toBe(0); return;
    }
    const command = await transaction(); expect(session.snapshot()).toEqual(before);
    session.commit(session.prepare(command)); session.calculateSync();
    for (const [index, expected] of item.observed.entries()) {
      const sheet = session.snapshot().metadata.workbook.sheets[index]!;
      const order = (a: typeof sheet.merges[number], b: typeof sheet.merges[number]) => a.from.row - b.from.row || a.from.column - b.from.column;
      expect([...sheet.merges].sort(order)).toEqual(expected.merges.map(range).sort(order));
      for (const cell of expected.cells) {
        const coordinate = at(cell.address), actual = session.getCell(sheet.sheetId, coordinate.row, coordinate.column)!;
        if (cell.formula !== null) {
          expect(parseWorkbookFormula(actual.formula!.expression, coordinate)).toEqual(parseWorkbookFormula("=" + cell.formula, coordinate));
          expect(session.getCalculatedCell(sheet.sheetId, coordinate.row, coordinate.column)).toMatchObject({ state: "valid", value: {
            type: cell.type === "e" ? "error" : "number", value: cell.type === "e" ? cell.value : Number(cell.value),
          } });
        } else expect(actual.value).toEqual({ type: cell.type === "text" ? "text" : "number", value: cell.type === "text" ? cell.value : Number(cell.value) });
      }
    }
    const reopened = openDocument(await readDocumentZip(await writeDocumentZip(session.serialize())), { kind: "workbook" });
    expect(reopened.snapshot()).toEqual(session.snapshot()); expect(session.history().undoSteps).toBe(1);
    session.undo(); expect(session.snapshot().metadata).toEqual(before.metadata); expect(session.snapshot().manifest.chunks).toEqual(before.manifest.chunks);
    expect(source.snapshot()).toEqual(before);
    const producer = readFileSync("packages/office/tests/fixtures/workbook-copy-merges-" + item.name + "-calc.xlsx");
    expect(createHash("sha256").update(producer).digest("hex")).toBe(item.producerSha256);
    const imported = openDocument(await importXlsx(producer, { documentId: "copy-merge-producer" }), { kind: "workbook" });
    expect(imported.snapshot().metadata.workbook.sheets.map(sheet => sheet.merges)).toEqual(item.observed.map(sheet => sheet.merges.map(range)));
  } finally { session.dispose(); }
});
