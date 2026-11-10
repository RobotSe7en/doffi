import { expect, test } from "vitest";
import { createWorkbook, openDocument } from "@doffi/core";
import type { WorkbookSheet } from "@doffi/core";
import { AxisDimensions } from "../src/workbook/dimensions.js";
import { navigateWorkbookSelection } from "../src/workbook/selection-navigation.js";
const point = (row: number, column = 0) => ({ row, column });
const range = (fromRow: number, fromColumn = 0, toRow = fromRow, toColumn = fromColumn) => ({ from: point(fromRow, fromColumn), to: point(toRow, toColumn) });
const sheet = () => JSON.parse(JSON.stringify(openDocument(createWorkbook({ documentId: "navigation" }), { kind: "workbook" }).snapshot().metadata.workbook.sheets[0])) as WorkbookSheet;
test("source navigation skips complete hidden intervals and respects physical RTL direction", () => {
  const source = sheet(), rows = new AxisDimensions(1_048_576, 20, [{ from: 1, to: 1_048_574, sizePt: 0 }]), columns = new AxisDimensions(16_384, 40, [{ from: 1, to: 10, sizePt: 0 }]);
  const dimensions = { rows, columns }, initial = { origin: point(0), focus: point(0) };
  expect(navigateWorkbookSelection(source, dimensions, initial, "down", false).range).toEqual(range(1_048_575));
  expect(navigateWorkbookSelection(source, dimensions, initial, "right", false).range).toEqual(range(0, 11));
  source.rightToLeft = true;
  const right = navigateWorkbookSelection(source, dimensions, { origin: point(0, 11), focus: point(0, 11) }, "right", false);
  expect(right.range).toEqual(range(0)); expect(rows.segmentCount).toBe(3); expect(columns.segmentCount).toBe(3);
});
test("merged navigation leaves the complete merge at its leading source column and expands both anchors", () => {
  const source = sheet(); source.merges = [range(1, 0, 2, 2), range(3, 0, 4, 1)];
  const dimensions = { rows: new AxisDimensions(1_048_576, 20, []), columns: new AxisDimensions(16_384, 40, []) };
  const initial = { origin: point(1), focus: point(1) };
  const down = navigateWorkbookSelection(source, dimensions, initial, "down", false); expect(down.range).toEqual(range(3, 0, 4, 1));
  const extended = navigateWorkbookSelection(source, dimensions, initial, "down", true); expect(extended.range).toEqual(range(1, 0, 4, 2));
  const restored = navigateWorkbookSelection(source, dimensions, extended.navigation, "up", true); expect(restored.range).toEqual(range(1, 0, 2, 2));
});
test("extension crosses its original endpoint and worksheet boundaries retain the current cell", () => {
  const source = sheet(), dimensions = { rows: new AxisDimensions(1_048_576, 20, []), columns: new AxisDimensions(16_384, 40, []) };
  const initial = { origin: point(1, 1), focus: point(1, 1) }, right = navigateWorkbookSelection(source, dimensions, initial, "right", true);
  const back = navigateWorkbookSelection(source, dimensions, right.navigation, "left", true), left = navigateWorkbookSelection(source, dimensions, back.navigation, "left", true);
  expect(left.range).toEqual(range(1, 0, 1, 1)); expect(left.navigation.origin).toEqual(point(1, 1));
  expect(navigateWorkbookSelection(source, dimensions, { origin: point(0), focus: point(0) }, "up", false).range).toEqual(range(0));
});
