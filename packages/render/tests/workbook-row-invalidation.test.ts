import { expect, test } from "vitest";
import { createDocumentSession, readWorkbookChangeImpact } from "@doffi/core";
import type { WorkbookChangeImpact } from "@doffi/core";
import { scalarWorkbook } from "../../core/tests/workbook-scalar-fixture.js";
import { createWorkbookPreviewSource, linkWorkbookPreviewSource, workbookPreviewChangeImpact } from "../src/workbook/preview-source.js";
import { workbookRowInvalidation } from "../src/workbook/row-invalidation.js";

const sheetId = "sheet-1";
const range = (from: number, to = from) => ({ sheetId, range: { from: { row: from, column: 0 }, to: { row: to, column: 0 } } });
test("compatible cell invalidations merge source and distant calculation rows", () => {
  const snapshot = scalarWorkbook([]).snapshot(), impact: WorkbookChangeImpact = { documentId: "scalar-kernel", fromRevision: 0, revision: 0,
    sourceRanges: [range(2, 4), range(0, 1)], calculatedRanges: [range(1000), range(4, 7)] };
  expect(workbookRowInvalidation(snapshot, snapshot, impact, 100)?.get(sheetId)).toEqual([{ from: 0, to: 7 }, { from: 1000, to: 1000 }]);
  expect(workbookRowInvalidation(snapshot, snapshot, { ...impact, fromRevision: 1 }, 100)).toBeUndefined();
  expect(workbookRowInvalidation(snapshot, snapshot, impact, 1)).toBeUndefined();
  expect(workbookRowInvalidation(snapshot, { ...snapshot, styles: [{ path: "styles.css", css: ".changed{color:red}" }] }, impact, 100)).toBeUndefined();
});
test("malformed or foreign rectangles never select a row cache", () => {
  const snapshot = scalarWorkbook([]).snapshot(), base = { documentId: "scalar-kernel", fromRevision: 0, revision: 0, calculatedRanges: [] };
  for (const value of [{ ...range(0), sheetId: "foreign" }, range(2, 1), range(1_048_576)]) {
    expect(() => workbookRowInvalidation(snapshot, snapshot, { ...base, sourceRanges: [value] }, 100)).toThrowError(expect.objectContaining({ code: "InvalidDocument" }));
  }
});
test("equal source identities cannot replace the exact previous provider/calculation pair", () => {
  const s = createDocumentSession(scalarWorkbook([{ row: 0, column: 0, value: { type: "number", value: 999 }, formula: "=TODAY()" }]));
  const a = s.calculateSync({ context: { now: { unixMilliseconds: 0, utcOffsetMinutes: 0 } } }), first = createWorkbookPreviewSource(s, a);
  const b = s.calculateSync({ context: { now: { unixMilliseconds: 86400000, utcOffsetMinutes: 0 } } }), other = createWorkbookPreviewSource(s, b);
  const prior = { ...other, cells: first.cells }, next = { ...first, rules: other.rules! };
  const impact = readWorkbookChangeImpact(s, s, b, a); linkWorkbookPreviewSource(next, prior, impact);
  expect(workbookPreviewChangeImpact(first, next)).toBeUndefined();
  expect(workbookPreviewChangeImpact(prior, next)).toBe(impact); s.dispose();
});
