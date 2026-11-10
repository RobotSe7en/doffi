import { expect, test } from "vitest";
import { copyWorkbookFormula, createDocumentSession, createWorkbook, createWorkbookFormulaFiller, openDocument, parseWorkbookFormula } from "../src/index.js";
import type { WorkbookDocumentMetadata, WorkbookTable } from "../src/index.js";

const origin = { row: 10, column: 10 };
function example() {
  const files = new Map(createWorkbook({ documentId: "fill-tables" }));
  const metadata = JSON.parse(new TextDecoder().decode(files.get("metadata.json")!)) as WorkbookDocumentMetadata;
  metadata.workbook.sheets[0]!.tables = [{ tableId: "sales", name: "Sales", range: { from: { row: 0, column: 0 }, to: { row: 4, column: 3 } }, header: false, totals: false, totalsPreviouslyShown: false,
    columns: ["产品", "数量", "单价", "金]额"].map((name, index) => ({ columnId: "column-" + index, name, totalFunction: "none" })),
    rowStripes: true, columnStripes: false, showFirstColumn: false, showLastColumn: false }];
  files.set("metadata.json", new TextEncoder().encode(JSON.stringify(metadata))); return createDocumentSession(openDocument(files, { kind: "workbook" }));
}
function equivalent(actual: string, expected: string, at = origin) { expect(parseWorkbookFormula(actual, at)).toEqual(parseWorkbookFormula(expected, at)); }
test("horizontal fill cycles qualified columns in both directions without changing copy semantics", () => {
  const s = example(), fill = createWorkbookFormulaFiller(s), source = "=SUM(Sales[数量])";
  for (const [delta, column] of [[1, "单价"], [2, "金']额"], [3, "产品"], [4, "数量"], [-1, "产品"], [-2, "金']额"], [17, "单价"]] as const) {
    const at = { row: origin.row, column: origin.column + delta };
    if (at.column < 0) continue;
    equivalent(fill(source, origin, at), "=SUM(Sales[" + column + "])", at);
    equivalent(copyWorkbookFormula(source, origin, at), source, at);
  }
  s.dispose();
});
test("range locks, unqualified columns, item selectors, scalar names and literals retain their distinct identities", () => {
  const s = example(), fill = createWorkbookFormulaFiller(s), at = { row: 12, column: 11 };
  const source = '=SUM(Sales[[数量]:[数量]],Sales[[数量]:[单价]],Sales[[#Headers],[数量]],Sales[@数量],[@数量],Sales,Missing[数量],Sales[Unknown])+A11+$A11+A$1+$A$1+"Sales[数量]"';
  equivalent(fill(source, origin, at), '=SUM(Sales[[数量]:[数量]],Sales[[数量]:[单价]],Sales[[#Headers],[单价]],Sales[@单价],[@数量],Sales,Missing[数量],Sales[Unknown])+B13+$A13+B$1+$A$1+"Sales[数量]"', at);
  equivalent(fill(source, origin, { row: 12, column: 10 }), copyWorkbookFormula(source, origin, { row: 12, column: 10 }), { row: 12, column: 10 });
  expect(fill(source, origin, origin)).toBe(source); s.dispose();
});
test("table lookup belongs to the captured document revision and validates formula and destination bounds", () => {
  const s = example(), fill = createWorkbookFormulaFiller(s), at = { ...origin, column: 11 };
  const table = structuredClone(s.snapshot().metadata.workbook.sheets[0]!.tables![0]!) as WorkbookTable; table.name = "Renamed";
  s.commit(s.prepare({ protocolVersion: 1, documentId: "fill-tables", baseRevision: 0,
    operations: [{ op: "setWorkbookTable", sheetId: "sheet-1", table }] }));
  equivalent(fill("=SUM(Sales[数量])", origin, at), "=SUM(Sales[单价])", at);
  equivalent(createWorkbookFormulaFiller(s)("=SUM(Sales[数量])", origin, at), "=SUM(Sales[数量])", at);
  const other = openDocument(createWorkbook({ documentId: "other" }), { kind: "workbook" });
  equivalent(createWorkbookFormulaFiller(other)("=SUM(Sales[数量])", origin, at), "=SUM(Sales[数量])", at);
  expect(() => fill("=SUM(", origin, at)).toThrow(); expect(() => fill("=A1", origin, { row: 0, column: 16_384 })).toThrow();
  s.dispose();
});
