import { readFileSync } from "node:fs";
import { createHash } from "node:crypto";
import { execFileSync } from "node:child_process";
import { expect, test } from "vitest";
import { createWorkbookCalculator, openDocument } from "@doffi/core";
import type { WorkbookValue } from "@doffi/core";
import { importXlsx } from "@doffi/office";
const evidence = JSON.parse(readFileSync(new URL("../docs/evidence/workbook-calculation-functions.json", import.meta.url), "utf8")) as {
  sourceSha256: string;
  cases: { address: string; expression: string; expected: WorkbookValue; value: WorkbookValue; matchesProducer: boolean; basis: string }[];
};
test("matches independent native function references while ignoring deliberately wrong source caches", async () => {
  const bytes = readFileSync(new URL("../packages/office/tests/fixtures/workbook-calculation.xlsx", import.meta.url));
  expect(createHash('sha256').update(bytes).digest('hex')).toBe(evidence.sourceSha256);
  execFileSync('python3', ['scripts/check-workbook-calculation.py']);
  const model = openDocument(await importXlsx(bytes, { documentId: "native-calculation" }), { kind: "workbook" }), before = model.serialize();
  const calculator = createWorkbookCalculator(model), result = calculator.calculateSync(), sheetId = model.snapshot().metadata.workbook.activeSheetId;
  expect(result.cells.length).toBeGreaterThanOrEqual(evidence.cases.length);
  const differences: string[] = [];
  for (const [row, item] of evidence.cases.entries()) {
    if (item.matchesProducer) expect(item.expected).toEqual(item.value);
    else { expect(item.expected).not.toEqual(item.value); expect(item.basis).not.toContain('Pinned Calc observation'); }
    const cell = calculator.getCell({ sheetId, row, column: 4 });
    if (cell.state !== "valid" || cell.value.type !== item.expected.type || cell.value.type === "number" && item.expected.type === "number" && Math.abs(cell.value.value - item.expected.value) > Math.max(1e-12, Math.abs(item.expected.value) * 1e-12)
      || cell.state === "valid" && cell.value.type !== "number" && JSON.stringify(cell.value) !== JSON.stringify(item.expected)) differences.push(item.expression + ": " + JSON.stringify({ actual: cell, expected: item.expected }));
  }
  expect(differences).toEqual([]); expect(model.serialize()).toEqual(before); calculator.dispose();
});
