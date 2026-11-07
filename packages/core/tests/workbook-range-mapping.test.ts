import { expect, test } from "vitest";
import { invertWorkbookMapping, mapWorkbookRanges } from "../src/workbook/range-mapping.js";
import type { WorkbookRangeChange, WorkbookStructureMapping } from "../src/runtime.js";
const range = (first: number, last: number, column = 0, endColumn = column): WorkbookRangeChange => ({ sheetId: "s", range: { from: { row: first, column }, to: { row: last, column: endColumn } } });
function points(ranges: readonly WorkbookRangeChange[]): Set<string> {
  const result = new Set<string>();
  for (const item of ranges) for (let row = item.range.from.row; row <= item.range.to.row; row++) for (let column = item.range.from.column; column <= item.range.to.column; column++) result.add(`${item.sheetId}/${row}/${column}`);
  return result;
}
test("cut ranges keep exact fragments, remove overwritten identities and invert surviving locations", () => {
  const step: WorkbookStructureMapping = { kind: "move", source: range(2, 2), target: range(2, 2, 2) };
  expect(mapWorkbookRanges([range(0, 4)], [step])).toEqual([range(0, 1), range(3, 4), range(2, 2, 2)]);
  expect(mapWorkbookRanges([range(2, 2, 2)], [step])).toEqual([]);
  const moved = mapWorkbookRanges([range(0, 4)], [step]);
  expect(points(mapWorkbookRanges(moved, invertWorkbookMapping([step])))).toEqual(points([range(0, 4)]));
});
test("axis insertions map existing identities without inventing coverage for inserted cells", () => {
  const step: WorkbookStructureMapping = { kind: "axis", sheetId: "s", axis: "row", action: "insert", index: 2, count: 2 };
  expect(mapWorkbookRanges([range(0, 4)], [step])).toEqual([range(0, 1), range(4, 6)]);
  expect(points(mapWorkbookRanges(mapWorkbookRanges([range(0, 4)], [step]), invertWorkbookMapping([step])))).toEqual(points([range(0, 4)]));
});
test("sort mappings preserve full rectangles and precisely fragment partial selections", () => {
  const step: WorkbookStructureMapping = { kind: "sort", sheetId: "s", range: range(0, 5, 0, 2).range, rows: [5, 2, 0, 4, 1, 3] };
  expect(mapWorkbookRanges([range(0, 5, 0, 2)], [step])).toEqual([range(0, 5, 0, 2)]);
  expect(mapWorkbookRanges([range(0, 1)], [step])).toEqual([range(2, 2), range(5, 5)]);
  const output = mapWorkbookRanges([range(0, 3, 0, 3)], [step]);
  expect(points(mapWorkbookRanges(output, invertWorkbookMapping([step])))).toEqual(points([range(0, 3, 0, 3)]));
});
test("rectangle mappings agree with an independent small-grid point projection", () => {
  for (let first = 0; first < 5; first++) for (let last = first; last < 6; last++) for (let column = 0; column < 4; column++) {
    const input = range(first, last, column, column + 1);
    for (const action of ["insert", "delete"] as const) {
      const step: WorkbookStructureMapping = { kind: "axis", sheetId: "s", axis: "row", action, index: 2, count: 2 };
      const expected = new Set<string>();
      for (let row = first; row <= last; row++) for (let col = column; col <= column + 1; col++) {
        if (action === "delete" && row >= 2 && row < 4) continue;
        expected.add(`s/${row < 2 ? row : row + (action === "insert" ? 2 : -2)}/${col}`);
      }
      expect(points(mapWorkbookRanges([input], [step]))).toEqual(expected);
    }
  }
});
test("deletion, ordered steps, cross-sheet cuts and fragment budgets remain explicit", () => {
  const cut: WorkbookStructureMapping = { kind: "move", source: range(1, 2), target: { ...range(3, 4, 1), sheetId: "other" } };
  expect(mapWorkbookRanges([range(0, 3)], [cut, { kind: "deleteSheet", sheetId: "s" }])).toEqual([{ ...range(3, 4, 1), sheetId: "other" }]);
  expect(() => mapWorkbookRanges([range(0, 3)], [cut], 2)).toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded" }));
  expect(() => mapWorkbookRanges([range(0, 2)], [{ kind: "sort", sheetId: "s", range: range(0, 2).range, rows: [1, 1, 2] }])).toThrow();
});
