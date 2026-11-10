import { readFileSync } from "node:fs";
import { createHash } from "node:crypto";
import { execFileSync } from "node:child_process";
import { expect, test } from "vitest";
import { createWorkbookCalculator, openDocument } from "@doffi/core";
import { importXlsx } from "@doffi/office";
const evidence = JSON.parse(readFileSync(new URL("../docs/evidence/workbook-filter-calculation.json", import.meta.url), "utf8")) as {
  sourceSha256: string; cases: { sheet: string; expected: number[]; observed: number[]; matchesProducer: boolean; basis: string }[];
};
test("current filters and source hidden flags match independent native SUBTOTAL expectations", async () => {
  execFileSync("python3", ["scripts/check-workbook-filter-calculation.py"]);
  const bytes = readFileSync(new URL("../packages/office/tests/fixtures/workbook-filter-calculation.xlsx", import.meta.url));
  expect(createHash("sha256").update(bytes).digest("hex")).toBe(evidence.sourceSha256);
  const model = openDocument(await importXlsx(bytes, { documentId: "filter-calculation" }), { kind: "workbook" }), before = model.serialize();
  const calculator = createWorkbookCalculator(model); expect(calculator.calculateSync().calculated).toBe(16);
  for (const item of evidence.cases) {
    if (item.matchesProducer) expect(item.expected).toEqual(item.observed);
    else { expect(item.expected).not.toEqual(item.observed); expect(item.basis).toContain("manually hidden"); }
    const sheetId = model.snapshot().metadata.workbook.sheets.find(sheet => sheet.name === item.sheet)!.sheetId;
    for (const [row, value] of item.expected.entries()) expect(calculator.getCell({ sheetId, row, column: 3 }), item.sheet + ":" + row).toMatchObject({ state: "valid", value: { type: "number", value } });
  }
  expect(model.serialize()).toEqual(before); calculator.dispose();
});
