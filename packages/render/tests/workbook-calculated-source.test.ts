import { expect, test } from "vitest";
import { createDocumentSession, DEFAULT_WORKBOOK_RULE_LIMITS } from "@doffi/core";
import { scalarWorkbook } from "../../core/tests/workbook-scalar-fixture.js";
import { createWorkbookPreviewSource } from "../src/workbook/preview-source.js";
import { checkedWorkbookRules } from "../src/workbook/rule-response.js";
const number = (value: number) => ({ type: "number" as const, value });
test("calculated preview providers preserve source bytes while projecting a complete current-value window", async () => {
  const session = createDocumentSession(scalarWorkbook(Array.from({ length: 150 }, (_, row) => ({ row, column: 0, value: number(999), formula: row ? `=A${row}+1` : "=1" }))));
  const before = session.serialize(), calculation = session.calculateSync(), source = createWorkbookPreviewSource(session, calculation), signal = new AbortController().signal;
  expect(source.rules?.currentValues).toBe(true);
  const request = { ...DEFAULT_WORKBOOK_RULE_LIMITS, documentId: "scalar-kernel", revision: 0, sheetId: "sheet-1", includeCurrent: true, coordinates: [{ row: 149, column: 0 }, { row: 149, column: 1 }] };
  const current = checkedWorkbookRules(await source.rules!.read(request, signal), request, source.cells.snapshot);
  expect(current.cells).toMatchObject([{ current: { value: number(150) }, conditional: [] }, { current: { value: { type: "blank" } }, conditional: [] }]);
  const incomplete = { ...current, cells: current.cells.slice(1) }; expect(() => checkedWorkbookRules(incomplete, request, source.cells.snapshot)).toThrowError(expect.objectContaining({ code: "InvalidDocument" }));
  const raw = await source.cells.read({ documentId: "scalar-kernel", revision: 0, maxCells: 1, maxCodeUnits: 4096, windows: [{ sheetId: "sheet-1", row: 149, column: 0, rows: 1, columns: 1 }] }, signal);
  expect(raw.cells[0]).toMatchObject({ value: number(999), formula: { cacheState: "imported" } }); expect(session.serialize()).toEqual(before);
  session.commit(session.prepare({ protocolVersion: 1, documentId: "scalar-kernel", baseRevision: 0, operations: [{ op: "setWorkbookCells", sheetId: "sheet-1", cells: [{ row: 0, column: 0, source: { kind: "formula", expression: "=10" } }] }] }));
  session.dispose(); expect(await source.rules!.read(request, signal)).toEqual(current);
});
test("an actual previous preview source supplies both source and distant calculation impact", () => {
  const s = createDocumentSession(scalarWorkbook([{ row: 0, column: 0, value: number(999), formula: "=B1" }, { row: 0, column: 1, value: number(1) },
    { row: 1000, column: 0, value: number(999), formula: "=A1+1" }]));
  const before = createWorkbookPreviewSource(s, s.calculateSync());
  s.commit(s.prepare({ protocolVersion: 1, documentId: "scalar-kernel", baseRevision: 0, operations: [{ op: "setWorkbookCells", sheetId: "sheet-1", cells: [
    { row: 0, column: 1, source: { kind: "value", value: number(10) } },
  ] }] }));
  const calculation = s.calculateSync(), current = createWorkbookPreviewSource(s, calculation, before);
  expect(current.impact).toMatchObject({ fromRevision: 0, revision: 1,
    sourceRanges: [{ sheetId: "sheet-1", range: { from: { row: 0, column: 1 }, to: { row: 0, column: 1 } } }],
    calculatedRanges: [{ sheetId: "sheet-1", range: { from: { row: 0, column: 0 }, to: { row: 0, column: 0 } } }, { sheetId: "sheet-1", range: { from: { row: 1000, column: 0 }, to: { row: 1000, column: 0 } } }] });
  expect(createWorkbookPreviewSource(s, calculation, { ...before }).impact).toBeUndefined();
  expect(before.cells.snapshot.manifest.revision).toBe(0); s.dispose();
});
