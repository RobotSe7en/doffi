import { createHash } from "node:crypto";
import { execFileSync } from "node:child_process";
import { readFileSync } from "node:fs";
import { expect, test } from "vitest";
import { createDocumentSession, openDocument } from "@doffi/core";
import type { DocumentOperation, WorkbookSession, WorkbookTable } from "@doffi/core";
import { importXlsx, readDocumentZip, writeDocumentZip } from "@doffi/office";

interface Case {
  name: string; operation: { axis: "row" | "column"; action: "insert" | "delete"; index: number; count: number };
  setHeader?: { column: number; text: string }; producerRejection?: { code: string; reason: string };
  expected: { rejected: boolean; ref: string | null; totals: boolean; columns: string[]; observerValues: string[] };
  observed: { tables: { ref: string; header: string; totals: string; headerCells: { address: string; type: string | null; constantText: string | null }[];
    columns: { name: string; calculated: string | null; totalFormula: string | null }[] }[];
    formulas: Record<string, Record<string, { formula: string; value: string; type: string }>> };
  outputFile: string; outputSha256: string;
}
const evidence = JSON.parse(readFileSync("docs/evidence/workbook-axis-tables.json", "utf8")) as { sourceSha256: string; baseSourceFile: string; baseSourceSha256: string; cases: Case[] };
const bytes = readFileSync("packages/office/tests/fixtures/workbook-axis-tables.xlsx");
function apply(s: WorkbookSession, operations: DocumentOperation[]) { const manifest = s.snapshot().manifest; return s.commit(s.prepare({ protocolVersion: 1, documentId: manifest.documentId, baseRevision: manifest.revision, operations })); }
function point(value: string) { const match = /^([A-Z]+)(\d+)$/.exec(value)!; let column = 0; for (const letter of match[1]!) column = column * 26 + letter.charCodeAt(0) - 64; return { row: Number(match[2]) - 1, column: column - 1 }; }
function range(value: string) { const [from, to = from] = value.split(":"); return { from: point(from!), to: point(to!) }; }
test("independent table source extends the original ZIP/XML fixture without replacing it and pins all producer bytes", () => {
  expect(createHash("sha256").update(bytes).digest("hex")).toBe(evidence.sourceSha256);
  expect(createHash("sha256").update(readFileSync(evidence.baseSourceFile)).digest("hex")).toBe(evidence.baseSourceSha256);
  for (const item of evidence.cases) expect(createHash("sha256").update(readFileSync(item.outputFile)).digest("hex")).toBe(item.outputSha256);
  expect(execFileSync("python3", ["scripts/check-workbook-axis-tables.py"], { encoding: "utf8" })).toContain("eleven independent");
});
test.each(evidence.cases)("native table $name follows independent source expectations, current values and exact package/history closure", async item => {
  const s = createDocumentSession(openDocument(await importXlsx(bytes, { documentId: "axis-table-native" }), { kind: "workbook" }));
  try {
    const before = s.snapshot(), sales = before.metadata.workbook.sheets.find(sheet => sheet.name === "Sales")!, observer = before.metadata.workbook.sheets.find(sheet => sheet.name === "Observer")!;
    const oldCells = before.manifest.chunks.flatMap(chunk => [0, 128].flatMap(offset => s.readWindow({ sheetId: chunk.sheetId, row: chunk.row + offset, column: chunk.column, rows: 128, columns: 32 })));
    const operation: DocumentOperation = { op: "editWorkbookAxis", sheetId: sales.sheetId, ...item.operation, ...(item.operation.action === "insert" ? { formatSource: "none" } : {}) };
    if (item.expected.rejected) { expect(() => apply(s, [operation])).toThrowError(expect.objectContaining({ code: "InvalidOperation" })); expect(s.snapshot()).toEqual(before); }
    else apply(s, [operation]);
    if (item.setHeader) {
      const current = JSON.parse(JSON.stringify(s.snapshot().metadata.workbook.sheets[0]!.tables![0]!)) as WorkbookTable;
      current.columns[item.setHeader.column - current.range.from.column]!.name = item.setHeader.text;
      apply(s, [{ op: "setWorkbookTable", sheetId: sales.sheetId, table: current }]);
    }
    const table = s.snapshot().metadata.workbook.sheets[0]!.tables?.[0];
    if (item.expected.ref === null) expect(table).toBeUndefined();
    else { expect(table?.range).toEqual(range(item.expected.ref)); expect(table?.totals).toBe(item.expected.totals); expect(table?.columns.map(column => column.name)).toEqual(item.expected.columns); }
    if (!item.expected.rejected) for (const old of oldCells) {
      const position = { row: old.row, column: old.column };
      if (old.sheetId === sales.sheetId) {
        const coordinate = position[item.operation.axis];
        if (item.operation.action === "delete" && coordinate >= item.operation.index && coordinate < item.operation.index + item.operation.count) continue;
        if (coordinate >= item.operation.index) position[item.operation.axis] += item.operation.action === "insert" ? item.operation.count : -item.operation.count;
      }
      const current = s.getCell(old.sheetId, position.row, position.column)!;
      expect(current.value).toEqual(old.value); expect(current.html.slice(current.html.indexOf(">") + 1)).toBe(old.html.slice(old.html.indexOf(">") + 1));
    }
    s.calculateSync();
    for (const [column, value] of item.expected.observerValues.entries()) expect(s.getCalculatedCell(observer.sheetId, 0, column)).toMatchObject({ state: "valid", value: value.startsWith("#") ? { type: "error", value } : { type: "number", value: Number(value) } });
    if (!item.expected.rejected && item.operation.axis === "row" && item.operation.action === "insert")
      for (let row = item.operation.index; row < item.operation.index + item.operation.count; row++) expect(s.getCell(sales.sheetId, row, 3)?.formula).toBeDefined();
    expect(openDocument(readDocumentZip(writeDocumentZip(s.serialize())), { kind: "workbook" }).snapshot()).toEqual(s.snapshot());
    const steps = Number(!item.expected.rejected) + Number(Boolean(item.setHeader)); expect(s.history().undoSteps).toBe(steps);
    for (let index = 0; index < steps; index++) s.undo();
    expect(s.snapshot().metadata).toEqual(before.metadata); expect(s.snapshot().styles).toEqual(before.styles); expect(s.snapshot().manifest.chunks).toEqual(before.manifest.chunks);
  } finally { s.dispose(); }
});
test.each(evidence.cases)("unmodified actual producer $name either reopens its own cached source or fails the declared strict boundary", async item => {
  const input = readFileSync(item.outputFile), original = Buffer.from(input);
  if (item.producerRejection) {
    await expect(importXlsx(input, { documentId: "axis-table-producer" })).rejects.toMatchObject({ code: item.producerRejection.code }); expect(input).toEqual(original);
    if (item.producerRejection.reason === "constant-header-mismatch") expect(item.observed.tables.some(table => table.columns.some((column, index) => table.headerCells[index]?.constantText !== column.name))).toBe(true);
    else expect(item.observed.formulas.Observer!["A1"]!.formula).toContain("#NAME?['#REF!]");
    return;
  }
  const model = openDocument(await importXlsx(input, { documentId: "axis-table-producer" }), { kind: "workbook" });
  for (const [name, formulas] of Object.entries(item.observed.formulas)) {
    const sheet = model.snapshot().metadata.workbook.sheets.find(sheet => sheet.name === name)!;
    for (const [address, observed] of Object.entries(formulas)) {
      const at = point(address), cell = model.getCell(sheet.sheetId, at.row, at.column)!;
      expect(cell.formula?.expression).toBe("=" + observed.formula);
      expect(cell.value).toEqual(observed.type === "e" ? { type: "error", value: observed.value } : { type: "number", value: Number(observed.value) });
      expect(cell.formula?.cacheState).toBe("imported");
    }
  }
  expect(openDocument(readDocumentZip(writeDocumentZip(model.serialize())), { kind: "workbook" }).snapshot()).toEqual(model.snapshot());
});
