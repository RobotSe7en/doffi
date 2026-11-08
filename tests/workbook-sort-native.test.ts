import { createHash } from "node:crypto";
import { execFileSync } from "node:child_process";
import { readFileSync } from "node:fs";
import { expect, test } from "vitest";
import { createDocumentSession, openDocument } from "@doffi/core";
import type { WorkbookSort } from "@doffi/core";
import { createWorkbookSortTransaction } from "@doffi/editor";
import { importXlsx, readDocumentZip, writeDocumentZip } from "@doffi/office";

interface Case { name: string; sheet: string; keys: WorkbookSort["keys"]; caseSensitive: boolean; expected: number[];
  observed: { identities: number[]; outsideFormulas: Record<string, string> }; matchesProducer: boolean; outputFile: string; outputSha256: string }
const evidence = JSON.parse(readFileSync("docs/evidence/workbook-sort.json", "utf8")) as { sourceSha256: string; cases: Case[] };
const input = readFileSync("packages/office/tests/fixtures/workbook-sort.xlsx");
test("independent sorting source and unmodified producer files remain tied to their declared hashes", () => {
  expect(createHash("sha256").update(input).digest("hex")).toBe(evidence.sourceSha256);
  for (const item of evidence.cases) {
    expect(createHash("sha256").update(readFileSync(item.outputFile)).digest("hex")).toBe(item.outputSha256);
    expect(item.matchesProducer).toBe(JSON.stringify(item.expected) === JSON.stringify(item.observed.identities));
  }
  expect(execFileSync("python3", ["scripts/check-workbook-sort.py"], { encoding: "utf8" })).toContain("six independent");
});
test.each(evidence.cases)("independent $name sorting follows documented expectations and reopens the same source with exact undo", async item => {
  const model = openDocument(await importXlsx(input, { documentId: "sort-native" }), { kind: "workbook" }), s = createDocumentSession(model);
  try {
    const original = s.snapshot(), sheet = original.metadata.workbook.sheets.find(sheet => sheet.name === item.sheet)!;
    const transaction = await createWorkbookSortTransaction(s, sheet.sheetId, { range: { from: { row: 1, column: 0 }, to: { row: 17, column: 2 } }, keys: item.keys,
      caseSensitive: item.caseSensitive, method: "none" }, { calculation: s.calculateSync(), locale: "en-US" });
    expect(s.snapshot()).toEqual(original); s.commit(s.prepare(transaction));
    expect(Array.from({ length: 17 }, (_, index) => Number(s.getCell(sheet.sheetId, index + 1, 2)!.text.slice(1)))).toEqual(item.expected);
    expect(s.getCell(sheet.sheetId, 0, 5)?.formula?.expression).toBe("=A2");
    const observer = original.metadata.workbook.sheets.find(sheet => sheet.name === "Observer")!;
    expect(s.getCell(observer.sheetId, 0, item.sheet === "Data" ? 0 : 1)?.formula?.expression).toBe("=" + item.sheet + "!A2");
    s.calculateSync(); const leading = s.getCalculatedCell(sheet.sheetId, 1, 0);
    expect(s.getCalculatedCell(sheet.sheetId, 0, 5)).toMatchObject({ state: leading.state, ...(leading.state === "valid" ? { value: leading.value } : {}) });
    const reopened = openDocument(readDocumentZip(writeDocumentZip(s.serialize())), { kind: "workbook" }); expect(reopened.snapshot()).toEqual(s.snapshot()); expect(s.history().undoSteps).toBe(1);
    s.undo(); expect(s.snapshot().metadata).toEqual(original.metadata); expect(s.snapshot().manifest.chunks).toEqual(original.manifest.chunks); s.redo(); expect(s.snapshot().metadata).toEqual(reopened.snapshot().metadata);
  } finally { s.dispose(); }
});
test.each(evidence.cases)("the unmodified $name Calc producer opens with its own recorded ordering", async item => {
  const model = openDocument(await importXlsx(readFileSync(item.outputFile), { documentId: "sort-producer" }), { kind: "workbook" });
  const sheet = model.snapshot().metadata.workbook.sheets.find(sheet => sheet.name === item.sheet)!;
  expect(Array.from({ length: 17 }, (_, index) => Number(model.getCell(sheet.sheetId, index + 1, 2)!.text.slice(1)))).toEqual(item.observed.identities);
  expect(openDocument(readDocumentZip(writeDocumentZip(model.serialize())), { kind: "workbook" }).snapshot()).toEqual(model.snapshot());
});
