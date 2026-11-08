import { expect, test } from "vitest";
import { workbookContextContains, workbookContextCovers, workbookWritableContext } from "../src/session/workbook-interaction-scope.js";
import type { Annotation, WorkbookTable } from "../src/index.js";
import type { WorkbookInteractionRange } from "../src/contracts/workbook-interaction.js";
import { scalarWorkbook } from "./workbook-scalar-fixture.js";

const range = (row: number, column: number, lastRow = row, lastColumn = column) => ({ from: { row, column }, to: { row: lastRow, column: lastColumn } });
const located = (row: number, column: number, lastRow = row, lastColumn = column): WorkbookInteractionRange => ({ sheetId: "sheet-1", range: range(row, column, lastRow, lastColumn) });
const noop = () => {};
function annotation(model = scalarWorkbook([]), ranges = [range(0, 0)]): Annotation {
  return { annotationId: "instruction", documentId: model.snapshot().manifest.documentId, baseRevision: 0, status: "open", instruction: "修改选区", anchor: { kind: "cells", sheetId: "sheet-1", ranges } };
}

test("writable coverage compares the entire union, including one-cell holes and another worksheet", () => {
  const target = located(0, 0, 9, 9), bands = [located(0, 0, 4, 9), located(5, 0, 9, 9)];
  expect(workbookContextCovers(bands, target, noop)).toBe(true);
  expect(workbookContextCovers([located(0, 0, 4, 9), located(5, 0, 9, 4), located(5, 6, 9, 9)], target, noop)).toBe(false);
  expect(workbookContextCovers(bands, { ...target, sheetId: "other" }, noop)).toBe(false);
  expect(workbookContextContains(bands, "sheet-1", 9, 9, noop)).toBe(true);
  expect(workbookContextContains(bands, "sheet-1", 10, 0, noop)).toBe(false);
});

test("union coverage agrees with independent small-grid membership across overlapping rectangles", () => {
  let seed = 20261108;
  const random = () => { seed = (Math.imul(seed, 1664525) + 1013904223) >>> 0; return (seed >>> 16) % 8; };
  for (let sample = 0; sample < 400; sample++) {
    const covers = Array.from({ length: random() }, () => { const a = random(), b = random(), c = random(), d = random(); return located(Math.min(a, b), Math.min(c, d), Math.max(a, b), Math.max(c, d)); });
    const a = random(), b = random(), c = random(), d = random(), target = located(Math.min(a, b), Math.min(c, d), Math.max(a, b), Math.max(c, d));
    let expected = true;
    for (let row = target.range.from.row; row <= target.range.to.row; row++) for (let column = target.range.from.column; column <= target.range.to.column; column++)
      if (!covers.some(({ range }) => row >= range.from.row && row <= range.to.row && column >= range.from.column && column <= range.to.column)) expected = false;
    expect(workbookContextCovers(covers, target, noop)).toBe(expected);
  }
});

test("sparse full-sheet coverage remains geometric and fragmented work is bounded", () => {
  let work = 0; const full = located(0, 0, 1_048_575, 16_383);
  expect(workbookContextCovers([full], full, () => work++)).toBe(true); expect(work).toBeLessThan(10);
  const stripes = Array.from({ length: 500 }, (_, column) => located(0, column * 2, 1_048_575, column * 2));
  expect(() => workbookContextCovers(stripes, full, () => { if (++work > 2000) throw Error("work budget"); })).toThrow("work budget");
});

test("default scope preserves all selected anchors while an explicit smaller scope is rejected", () => {
  const model = scalarWorkbook([]), input = annotation(model, [range(0, 0, 2, 1), range(7, 3)]), before = model.serialize();
  expect(workbookWritableContext(model, [input], undefined, noop).ranges).toEqual([located(0, 0, 2, 1), located(7, 3)]);
  expect(() => workbookWritableContext(model, [input], { kind: "ranges", ranges: [located(0, 0, 2, 1)] }, noop)).toThrowError(expect.objectContaining({ code: "AnchorNotFound" }));
  expect(() => workbookWritableContext(model, [input], { kind: "ranges", ranges: [{ ...located(0, 0), sheetId: "foreign" }] }, noop)).toThrowError(expect.objectContaining({ code: "InvalidOperation" }));
  expect(model.serialize()).toEqual(before);
});

test("table scope uses its actual current range and does not authorize an adjacent blank or foreign object", () => {
  const table: WorkbookTable = { tableId: "table", name: "Sales", range: range(0, 0, 3, 1), header: true, totals: false, totalsPreviouslyShown: false,
    columns: [{ columnId: "amount", name: "Amount", totalFunction: "none" }, { columnId: "count", name: "Count", totalFunction: "none" }], rowStripes: true, columnStripes: false, showFirstColumn: false, showLastColumn: false };
  const model = scalarWorkbook([{ row: 0, column: 0, value: { type: "text", value: "Amount" } }, { row: 0, column: 1, value: { type: "text", value: "Count" } }], metadata => { metadata.workbook.sheets[0]!.tables = [table]; });
  expect(workbookWritableContext(model, [annotation(model)], { kind: "tables", tableIds: ["table"] }, noop).ranges).toEqual([located(0, 0, 3, 1)]);
  expect(() => workbookWritableContext(model, [annotation(model, [range(4, 0)])], { kind: "tables", tableIds: ["table"] }, noop)).toThrowError(expect.objectContaining({ code: "AnchorNotFound" }));
  for (const scope of [{ kind: "tables" as const, tableIds: ["table", "table"] }, { kind: "tables" as const, tableIds: ["foreign"] }, { kind: "tables" as const, tableIds: ["table"], objectIds: ["foreign"] }])
    expect(() => workbookWritableContext(model, [annotation(model)], scope, noop)).toThrowError(expect.objectContaining({ code: "InvalidOperation" }));
});
