import { expect, test } from "vitest";
import { captureWorkbookModel, createDocumentSession, createWorkbook, createWorkbookCalculator, openDocument, readWorkbookChangeImpact } from "../src/index.js";
import type { DocumentOperation, WorkbookSession } from "../src/index.js";
import { scalarWorkbook } from "./workbook-scalar-fixture.js";
import { createHash } from "node:crypto";

const sheetId = "sheet-1", number = (value: number) => ({ type: "number" as const, value });
function apply(s: WorkbookSession, operations: DocumentOperation[]) {
  const manifest = s.snapshot().manifest; s.commit(s.prepare({ protocolVersion: 1, documentId: manifest.documentId, baseRevision: manifest.revision, operations }));
}
test("source and changed calculation ranges remain distinct across chunks", () => {
  const s = createDocumentSession(scalarWorkbook([{ row: 0, column: 0, value: number(999), formula: "=B1+1" }, { row: 0, column: 1, value: number(1) },
    { row: 0, column: 2, value: number(999), formula: "=99" }, { row: 1000, column: 0, value: number(999), formula: "=A1+1" },
    { row: 1000, column: 1, value: number(999), formula: "=A1+2" }]));
  const previous = captureWorkbookModel(s), before = s.calculateSync();
  apply(s, [{ op: "setWorkbookCells", sheetId, cells: [{ row: 0, column: 1, source: { kind: "value", value: number(10) } }] }]);
  const after = s.calculateSync(), impact = readWorkbookChangeImpact(previous, s, before, after);
  expect(impact).toMatchObject({ documentId: "scalar-kernel", fromRevision: 0, revision: 1,
    sourceRanges: [{ sheetId, range: { from: { row: 0, column: 1 }, to: { row: 0, column: 1 } } }],
    calculatedRanges: [{ sheetId, range: { from: { row: 0, column: 0 }, to: { row: 0, column: 0 } } }, { sheetId, range: { from: { row: 1000, column: 0 }, to: { row: 1000, column: 1 } } }] });
  expect(s.getCalculatedCell(sheetId, 1000, 0)).toMatchObject({ value: number(12) }); expect(s.getCalculatedCell(sheetId, 1000, 1)).toMatchObject({ value: number(13) });
  expect(s.getCalculatedCell(sheetId, 0, 2)).toMatchObject({ value: number(99) }); expect(Object.isFrozen(impact.calculatedRanges[1]?.range.from)).toBe(true); s.dispose();
});
test("format-only source changes do not invent changed calculation values", () => {
  const s = createDocumentSession(scalarWorkbook([{ row: 0, column: 0, value: number(999), formula: "=1" }])), previous = captureWorkbookModel(s), before = s.calculateSync();
  apply(s, [{ op: "setWorkbookCellFormats", sheetId, cells: [{ row: 0, column: 0, cssClass: "cell-default" }] }]);
  const impact = readWorkbookChangeImpact(previous, s, before, s.calculateSync());
  expect(impact.sourceRanges).toEqual([{ sheetId, range: { from: { row: 0, column: 0 }, to: { row: 0, column: 0 } } }]); expect(impact.calculatedRanges).toEqual([]);
  expect(s.getCell(sheetId, 0, 0)).toMatchObject({ value: number(999), formula: { cacheState: "imported" } }); s.dispose();
});
test("calculation context can change row impact without a new source revision", () => {
  const model = scalarWorkbook([{ row: 0, column: 0, value: number(999), formula: "=TODAY()" }, { row: 0, column: 1, value: number(999), formula: "=1" }]), calculator = createWorkbookCalculator(model);
  const before = calculator.calculateSync({ context: { now: { unixMilliseconds: 0, utcOffsetMinutes: 0 } } });
  const after = calculator.calculateSync({ context: { now: { unixMilliseconds: 86400000, utcOffsetMinutes: 0 } } });
  const impact = readWorkbookChangeImpact(model, model, before, after);
  expect(impact).toMatchObject({ fromRevision: 0, revision: 0, sourceRanges: [], calculatedRanges: [{ sheetId, range: { from: { row: 0, column: 0 }, to: { row: 0, column: 0 } } }] }); calculator.dispose();
});
test("an absent batch stays explicit and never borrows imported caches as completed values", () => {
  const model = scalarWorkbook([{ row: 0, column: 0, value: number(1), formula: "=1" }]), calculator = createWorkbookCalculator(model), calculation = calculator.calculateSync();
  expect(readWorkbookChangeImpact(model, model).calculatedRanges).toEqual([]);
  expect(readWorkbookChangeImpact(model, model, undefined, calculation).calculatedRanges).toHaveLength(1);
  expect(readWorkbookChangeImpact(model, model, calculation).calculatedRanges).toHaveLength(1);
  expect(model.getCell(sheetId, 0, 0)?.formula?.cacheState).toBe("imported"); calculator.dispose();
});
test("batch copies and equal-ID foreign sources cannot replace owned completed batches", () => {
  const a = scalarWorkbook([{ row: 0, column: 0, value: number(999), formula: "=1" }]), b = scalarWorkbook([{ row: 0, column: 0, value: number(999), formula: "=2" }]);
  const calculator = createWorkbookCalculator(a), calculation = calculator.calculateSync();
  expect(() => readWorkbookChangeImpact(a, a, calculation, { ...calculation })).toThrowError(expect.objectContaining({ code: "InvalidOperation" }));
  expect(() => readWorkbookChangeImpact(a, b, calculation, calculation)).toThrowError(expect.objectContaining({ code: "RevisionConflict" }));
  expect(() => readWorkbookChangeImpact(a, openDocument(createWorkbook({ documentId: "foreign" }), { kind: "workbook" }))).toThrowError(expect.objectContaining({ code: "RevisionConflict" })); calculator.dispose();
});

test.each([{ sheetId, row: 256, column: 0 }, { sheetId: "other", row: 256, column: 0 }, { sheetId: "other", row: 0, column: 0 }, { sheetId, row: 0, column: 32 }])("relocated chunks invalidate both original and new coordinates (%j)", destination => {
  const before = scalarWorkbook([{ row: 0, column: 0, value: { type: "text", value: "wrapped content" } }], metadata => {
    metadata.workbook.sheets.push({ ...metadata.workbook.sheets[0]!, sheetId: "other", name: "Other" });
  });
  const files = new Map(before.serialize()), manifest = JSON.parse(new TextDecoder().decode(files.get("manifest.json")!));
  const descriptor = manifest.chunks[0];
  const html = new TextDecoder().decode(files.get(descriptor.path)!).replace('data-sheet-id="sheet-1"', `data-sheet-id="${destination.sheetId}"`)
    .replaceAll('data-row="0"', `data-row="${destination.row}"`).replaceAll('data-column="0"', `data-column="${destination.column}"`);
  const bytes = new TextEncoder().encode(html); files.set(descriptor.path, bytes);
  Object.assign(descriptor, destination); descriptor.sha256 = createHash("sha256").update(bytes).digest("hex"); manifest.revision = 1;
  files.set("manifest.json", new TextEncoder().encode(JSON.stringify(manifest)));
  const after = openDocument(files, { kind: "workbook" });
  expect(readWorkbookChangeImpact(before, after).sourceRanges).toEqual([
    { sheetId, range: { from: { row: 0, column: 0 }, to: { row: 0, column: 0 } } },
    { sheetId: destination.sheetId, range: { from: { row: destination.row, column: destination.column }, to: { row: destination.row, column: destination.column } } },
  ]);
});
