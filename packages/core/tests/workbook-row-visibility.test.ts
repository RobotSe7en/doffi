import { expect, test } from "vitest";
import { createDocumentSession, DEFAULT_WORKBOOK_ROW_VISIBILITY_LIMITS, readWorkbookRowVisibility } from "../src/index.js";
import { scalarWorkbook } from "./workbook-scalar-fixture.js";

function model() { return scalarWorkbook([3, 7, 5].map((value, index) => ({ row: index + 1, column: 0, value: { type: "number", value } })), metadata => {
  const sheet = metadata.workbook.sheets[0]!; sheet.rowDimensions = [{ from: 1, to: 1, hidden: true }];
  sheet.autoFilter = { range: { from: { row: 0, column: 0 }, to: { row: 3, column: 0 } }, columns: [{ column: 0, hiddenButton: false, showButton: true,
    criterion: { kind: "top", top: true, percent: false, count: 1 } }] };
}); }
const query = { ...DEFAULT_WORKBOOK_ROW_VISIBILITY_LIMITS, documentId: "scalar-kernel", revision: 0, sheetId: "sheet-1", from: 0, to: 3 };
test("manual hiding and current filter predicates remain separate and do not change the package", async () => {
  const s = createDocumentSession(model()); try {
    const before = s.serialize(), calculation = s.calculateSync(), result = await readWorkbookRowVisibility(s, query, { calculation });
    expect(result.rows).toEqual([{ row: 0, hidden: false, filtered: false }, { row: 1, hidden: true, filtered: true }, { row: 2, hidden: false, filtered: false }, { row: 3, hidden: false, filtered: true }]);
    expect(Object.isFrozen(result.rows[0])).toBe(true); expect(s.serialize()).toEqual(before); expect(s.history().undoSteps).toBe(0);
  } finally { s.dispose(); }
});
test("row visibility captures request/context input and rejects cancellation, stale batches and mid-query revision changes", async () => {
  const s = createDocumentSession(model()); try {
    const request = { ...query, to: 600 }, calculation = s.calculateSync(); let count = 0;
    const result = await readWorkbookRowVisibility(s, request, { calculation, yield: async () => { if (!count++) request.to = 1; } });
    expect(result.rows).toHaveLength(601);
    const abort = new AbortController(); await expect(readWorkbookRowVisibility(s, { ...query, to: 600 }, { signal: abort.signal, yield: async () => { abort.abort(); } })).rejects.toThrowError(expect.objectContaining({ code: "OperationCancelled" }));
    const before = s.serialize(); await expect(readWorkbookRowVisibility(s, { ...query, maxRows: 2 })).rejects.toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded" })); expect(s.serialize()).toEqual(before);
    await expect(readWorkbookRowVisibility(s, { ...query, to: 600 }, { calculation, yield: async () => {
      s.commit(s.prepare({ protocolVersion: 1, documentId: "scalar-kernel", baseRevision: 0, operations: [{ op: "setWorkbookSettings", changes: { calculationMode: "manual" } }] }));
    } })).rejects.toThrowError(expect.objectContaining({ code: "RevisionConflict" }));
    await expect(readWorkbookRowVisibility(s, { ...query, revision: 1 }, { calculation })).rejects.toThrowError(expect.objectContaining({ code: "RevisionConflict" }));
  } finally { s.dispose(); }
});
