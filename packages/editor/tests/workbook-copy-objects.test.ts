import { afterEach, expect, test } from "vitest";
import { createDocumentSession, createWorkbook, openDocument } from "@doffi/core";
import type { DocumentOperation, WorkbookLayoutMeasurement, WorkbookLayoutRequest, WorkbookSession } from "@doffi/core";
import { createWorkbookRangeCopyTransaction } from "../src/workbook/copy-range.js";

const sessions: WorkbookSession[] = [], sheetId = "sheet-1";
afterEach(() => sessions.splice(0).forEach(s => s.dispose()));
const range = (row: number, column: number, lastRow = row, lastColumn = column) => ({ from: { row, column }, to: { row: lastRow, column: lastColumn } });
const metrics = (request: WorkbookLayoutRequest): WorkbookLayoutMeasurement => ({ sourceDigest: request.sourceDigest, fontEnvironment: "combined-range-object-copy-points", sheets: request.sheetIds.map(sheetId => ({ sheetId, rows: { defaultSizePt: 15, overrides: [] }, columns: { defaultSizePt: 45, overrides: [] } })) });
const apply = (s: WorkbookSession, operations: DocumentOperation[]) => { const m = s.snapshot().manifest; return s.commit(s.prepare({ protocolVersion: 1, documentId: m.documentId, baseRevision: m.revision, operations }, { measureWorkbookLayout: metrics })); };
function example() {
  const s = createDocumentSession(openDocument(createWorkbook({ documentId: "combined-object-copy" }), { kind: "workbook" })); sessions.push(s);
  apply(s, [{ op: "insertWorkbookSheet", name: "Other" }, { op: "setWorkbookTable", sheetId, table: { tableId: "sales", name: "Sales", range: range(0, 0, 3, 1), header: true, totals: false, totalsPreviouslyShown: false,
    columns: [{ columnId: "amount", name: "Amount", totalFunction: "none" }, { columnId: "double", name: "Double", calculatedFormula: "=Sales[@Amount]*2", totalFunction: "none" }], rowStripes: true, columnStripes: false, showFirstColumn: false, showLastColumn: false } },
    { op: "setWorkbookCells", sheetId, cells: [1, 2, 3].map((value, index) => ({ row: index + 1, column: 0, source: { kind: "value", value: { type: "number", value } } })) },
    { op: "setWorkbookNumberFormat", format: { formatId: "percent", builtinId: 9 } }, { op: "setWorkbookCellFormats", sheetId, cells: [{ row: 1, column: 0, formatId: "percent" }] },
    { op: "insertWorkbookObject", sheetId, anchor: { kind: "one-cell", from: { row: 3, column: 3, offsetXPt: 2, offsetYPt: 3 }, widthPt: 80, heightPt: 40 }, object: { kind: "geometry", preset: "rect", html: "<p>图形禰󠄀😀</p>" } },
    { op: "insertWorkbookChart", sheetId, anchor: { kind: "one-cell", from: { row: 4, column: 3, offsetXPt: 2, offsetYPt: 3 }, widthPt: 200, heightPt: 100 }, kind: "line", data: { series: [
      { values: { expression: "=A2:A4", valueType: "number" } }, { values: { expression: "=Sales[Double]", valueType: "number" } },
    ] } }]);
  for (const [format, column] of [["classic", 0], ["modern", 1]] as const) apply(s, [{ op: "addWorkbookComment", format, anchor: { kind: "cells", sheetId, ranges: [range(1, column)] }, html: `<p>${format} 禰󠄀😀</p>` }]);
  s.clearHistory(); return { s, targetId: s.snapshot().metadata.workbook.sheets[1]!.sheetId };
}
test("one region command combines independent table identities, current references, two comment kinds and copied drawings", async () => {
  const { s, targetId } = example(), before = s.snapshot();
  const tx = await createWorkbookRangeCopyTransaction(s, sheetId, range(0, 0, 5, 4), { row: 59, column: 7 }, { targetSheetId: targetId });
  const candidate = s.prepare(tx, { measureWorkbookLayout: metrics }); expect(s.snapshot()).toEqual(before); s.commit(candidate);
  const metadata = s.snapshot().metadata.workbook, table = metadata.sheets[1]!.tables![0]!, copied = metadata.charts![1]!;
  expect(table.name).toBe("Sales_2"); expect(metadata.objects).toHaveLength(4); expect(s.snapshot().annotations).toHaveLength(4); expect(metadata.comments!.map(comment => comment.format)).toEqual(["classic", "modern", "classic", "modern"]);
  expect(copied.series[0]!.values).toMatchObject({ expression: "=H61:H63" }); expect(copied.series[1]!.values).toMatchObject({ expression: "=Sales[Double]", originSheetId: sheetId });
  expect(s.getCell(targetId, 60, 8)?.formula!.expression).toContain("Sales_2"); expect(s.history().undoSteps).toBe(1); expect(openDocument(s.serialize(), { kind: "workbook" }).snapshot()).toEqual(s.snapshot());
  s.undo(); expect(s.snapshot().metadata).toEqual(before.metadata); expect(s.snapshot().content).toEqual(before.content); expect(s.snapshot().manifest.chunks).toEqual(before.manifest.chunks);
});
test.each(["values", "formulas", "formats"] as const)("%s copying preserves drawing and comment identities while omitting table creation", async mode => {
  const { s, targetId } = example(), before = s.snapshot(), calculation = s.calculateSync();
  const tx = await createWorkbookRangeCopyTransaction(s, sheetId, range(0, 0, 5, 4), { row: 59, column: 7 }, { targetSheetId: targetId, mode, calculation }); s.commit(s.prepare(tx, { measureWorkbookLayout: metrics }));
  expect(s.snapshot().metadata.workbook.objects).toEqual(before.metadata.workbook.objects); expect(s.snapshot().metadata.workbook.charts).toEqual(before.metadata.workbook.charts);
  expect(s.snapshot().annotations).toEqual(before.annotations.map(annotation => ({ ...annotation, baseRevision: s.snapshot().manifest.revision })));
  expect(s.snapshot().metadata.workbook.sheets[1]!.tables).toBeUndefined(); expect(s.history().undoSteps).toBe(1);
});
test("skip blanks omits blank-cell drawing anchors but retains populated table and comment footprints", async () => {
  const { s, targetId } = example(), before = s.snapshot();
  const tx = await createWorkbookRangeCopyTransaction(s, sheetId, range(0, 0, 5, 4), { row: 59, column: 7 }, { targetSheetId: targetId, skipBlanks: true }); s.commit(s.prepare(tx, { measureWorkbookLayout: metrics }));
  expect(s.snapshot().metadata.workbook.objects).toEqual(before.metadata.workbook.objects); expect(s.snapshot().metadata.workbook.sheets[1]!.tables).toHaveLength(1); expect(s.snapshot().annotations).toHaveLength(4); expect(s.history().undoSteps).toBe(1);
});
