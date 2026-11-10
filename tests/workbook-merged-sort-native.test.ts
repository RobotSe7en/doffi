import { createHash } from "node:crypto";
import { execFileSync } from "node:child_process";
import { readFileSync } from "node:fs";
import { expect, test } from "vitest";
import { createDocumentSession, openDocument } from "@doffi/core";
import type { WorkbookSort } from "@doffi/core";
import { createWorkbookSortTransaction } from "@doffi/editor";
import { importXlsx, readDocumentZip, writeDocumentZip } from "@doffi/office";

interface Case { name: string; sheet: string; height: number; columns: number[]; descending: boolean[]; reject?: boolean; expectedRecords?: number[];
  outputFile: string; outputSha256: string; nativeCall: { outcome: string; changed: boolean }; observed: { records: string[][]; merges: string[]; outsideFormula: string } }
const input = readFileSync("packages/office/tests/fixtures/workbook-merged-sort.xlsx");
const evidence = JSON.parse(readFileSync("docs/evidence/workbook-merged-sort.json", "utf8")) as { sourceSha256: string; cases: Case[] };
test("independent merged-sort sources and producer observations are reproducible without SDK conversion", () => {
  expect(createHash("sha256").update(input).digest("hex")).toBe(evidence.sourceSha256);
  expect(execFileSync("python3", ["scripts/check-workbook-merged-sort.py"], { encoding: "utf8" })).toContain("four independent");
});
test.each(evidence.cases)("independent merged source $name preserves record blocks, outside addresses and exact undo", async item => {
  const session = createDocumentSession(openDocument(await importXlsx(input, { documentId: "merged-sort-native" }), { kind: "workbook" }));
  try {
    const before = session.snapshot(), sheetId = before.metadata.workbook.sheets.find(sheet => sheet.name === item.sheet)!.sheetId;
    const sort: WorkbookSort = { range: { from: { row: 0, column: 0 }, to: { row: item.height * 3 - 1, column: 3 } }, method: "none", caseSensitive: false,
      keys: item.columns.map((column, index) => ({ column, descending: item.descending[index]! })) };
    const candidate = createWorkbookSortTransaction(session, sheetId, sort, { calculation: session.calculateSync() });
    if (item.reject) { await expect(candidate).rejects.toMatchObject({ code: "InvalidOperation" }); expect(session.snapshot()).toEqual(before); expect(session.history().undoSteps).toBe(0); return; }
    const transaction = await candidate; expect(session.snapshot()).toEqual(before); session.commit(session.prepare(transaction)); session.calculateSync();
    const source = [[3, 30], [1, 20], [1, 25]], expected = item.expectedRecords!.map(index => source[index]!);
    for (let record = 0; record < 3; record++) for (const [index, column] of [0, 2].entries()) {
      expect(session.getCalculatedCell(sheetId, record * item.height, column)).toMatchObject({ state: "valid", value: { type: "number", value: expected[record]![index] } });
    }
    expect(session.getCell(sheetId, 0, 5)?.formula?.expression).toBe("=A1"); expect(session.getCalculatedCell(sheetId, 0, 5)).toMatchObject({ value: { type: "number", value: 1 } });
    const observer = before.metadata.workbook.sheets.find(sheet => sheet.name === "Observer")!;
    const observerColumn = item.sheet === "Uniform" ? 0 : 1;
    expect(session.getCell(observer.sheetId, 0, observerColumn)?.formula?.expression).toBe("=" + item.sheet + "!A1");
    expect(session.getCalculatedCell(observer.sheetId, 0, observerColumn)).toMatchObject({ value: { type: "number", value: 1 } });
    expect(openDocument(readDocumentZip(writeDocumentZip(session.serialize())), { kind: "workbook" }).snapshot()).toEqual(session.snapshot());
    expect(session.history().undoSteps).toBe(1); session.undo(); expect(session.snapshot().metadata).toEqual(before.metadata); expect(session.snapshot().manifest.chunks).toEqual(before.manifest.chunks);
  } finally { session.dispose(); }
});
test.each(evidence.cases)("the actual $name producer is retained with its own ordering and merge geometry", async item => {
  const bytes = readFileSync(item.outputFile); expect(createHash("sha256").update(bytes).digest("hex")).toBe(item.outputSha256);
  const model = openDocument(await importXlsx(bytes, { documentId: "merged-sort-producer" }), { kind: "workbook" });
  const sheet = model.snapshot().metadata.workbook.sheets.find(sheet => sheet.name === item.sheet)!;
  for (const [record, values] of item.observed.records.entries()) for (const [index, column] of [0, 2].entries()) expect(model.getCell(sheet.sheetId, record * item.height, column)?.value).toEqual({ type: "number", value: Number(values[index]) });
  expect(sheet.merges).toHaveLength(item.observed.merges.length); expect(model.getCell(sheet.sheetId, 0, 5)?.formula?.expression).toBe("=" + item.observed.outsideFormula);
  // Returned sort calls can be no-ops. The recorded producer order is never
  // adopted as the SDK expectation or reported as a successful native sort.
  expect(item.nativeCall).toEqual({ outcome: "returned", changed: false });
});
