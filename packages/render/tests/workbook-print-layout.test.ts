import { describe, expect, it } from "vitest";
import type { WorkbookPrint } from "@doffi/core";
import { AxisDimensions } from "../src/workbook/dimensions.js";
import { planWorkbookPrintSheet, workbookPaper, workbookPrintViewport } from "../src/workbook/print-layout.js";

const band = { left: [], center: [], right: [] };
const print: WorkbookPrint = {
  areas: [], repeatRows: { from: 0, to: 0 }, repeatColumns: { from: 0, to: 0 },
  orientation: "portrait", paper: { widthPt: 100, heightPt: 100 },
  marginsPt: { left: 0, right: 0, top: 0, bottom: 0, header: 0, footer: 0 },
  pageOrder: "downThenOver", horizontalCentered: false, verticalCentered: false, gridLines: false, headings: false, blackAndWhite: false,
  scale: { kind: "percent", percent: 100 }, rowBreaks: [{ before: 3, from: 8, to: 8 }], columnBreaks: [{ before: 5, from: 0, to: 9 }],
  headerFooter: { differentFirst: false, differentOddEven: false, scaleWithDocument: true, alignWithMargins: true, oddHeader: band, oddFooter: band },
};
const dimensions = { rows: new AxisDimensions(1_048_576, 20, []), columns: new AxisDimensions(16_384, 20, []) };
const range = { from: { row: 0, column: 0 }, to: { row: 8, column: 8 } };

describe("workbook print pagination", () => {
  it("applies breaks, repeats titles without duplicating their cells and orders the page grid", () => {
    const { tiles } = planWorkbookPrintSheet(dimensions, [range], print, workbookPaper(print), 20);
    expect(tiles.map(tile => [tile.rows.body.from, tile.columns.body.from])).toEqual([[0, 0], [3, 0], [7, 0], [0, 5], [3, 5], [7, 5]]);
    const layout = workbookPrintViewport(dimensions, tiles[4]!, 100);
    expect(tiles[0]!.rows.parts).toEqual([{ from: 0, to: 2 }]);
    expect(layout.panes.map(pane => [pane.kind, pane.rows.map(row => row.index), pane.columns.map(column => column.index)])).toEqual([
      ["corner", [0], [0]], ["frozenRows", [0], [5, 6, 7, 8]], ["frozenColumns", [3, 4, 5, 6], [0]], ["body", [3, 4, 5, 6], [5, 6, 7, 8]],
    ]);
    const over = planWorkbookPrintSheet(dimensions, [range], { ...print, pageOrder: "overThenDown" }, workbookPaper(print), 20);
    expect(over.tiles.map(tile => [tile.rows.body.from, tile.columns.body.from])).toEqual([[0, 0], [0, 5], [3, 0], [3, 5], [7, 0], [7, 5]]);
  });

  it("fits without enlarging, disregards manual breaks in fit mode and allows an unconstrained axis", () => {
    const fit = { ...print, scale: { kind: "fit", width: 1, height: 1 } } as const;
    const result = planWorkbookPrintSheet(dimensions, [range], fit, workbookPaper(fit), 20);
    expect(result.scale).toBe(0.55);
    expect(result.tiles).toHaveLength(1);
    const tall = { from: range.from, to: { row: 30, column: 2 } };
    const fitWidth = planWorkbookPrintSheet(dimensions, [tall], { ...fit, scale: { kind: "fit", width: 1, height: 0 } }, workbookPaper(fit), 20);
    expect(fitWidth.scale).toBe(1);
    expect(fitWidth.tiles[1]!.rows.body.from).toBe(5);
  });

  it("jumps hidden dimensions and rejects over-budget areas before expanding the grid", () => {
    const rows = new AxisDimensions(1_048_576, 20, [{ from: 1, to: 999_999, sizePt: 0 }]);
    const distant = { from: range.from, to: { row: 1_000_001, column: 0 } };
    const result = planWorkbookPrintSheet({ ...dimensions, rows }, [distant], { ...print, rowBreaks: [] }, workbookPaper(print), 2);
    expect(workbookPrintViewport({ ...dimensions, rows }, result.tiles[0]!, 4).panes[0]!.rows.map(row => row.index)).toEqual([0, 1_000_000, 1_000_001]);
    expect(() => planWorkbookPrintSheet(dimensions, [distant], print, workbookPaper(print), 2)).toThrow(/smaller workbook print area/);
    expect(() => workbookPrintViewport(dimensions, result.tiles[0]!, 1)).toThrow();
  });
});
