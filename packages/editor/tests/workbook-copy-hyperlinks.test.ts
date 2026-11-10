import { afterEach, expect, test } from "vitest";
import { createDocumentSession } from "@doffi/core";
import type { DocumentOperation, WorkbookSession } from "@doffi/core";
import { scalarWorkbook } from "../../core/tests/workbook-scalar-fixture.js";
import { createWorkbookRangeCopyTransaction } from "../src/workbook/copy-range.js";
const sessions: WorkbookSession[] = [];
afterEach(() => { for (const session of sessions.splice(0)) session.dispose(); });
const range = (row: number, column: number, lastRow = row, lastColumn = column) => ({ from: { row, column }, to: { row: lastRow, column: lastColumn } });
function fixture() { const s = createDocumentSession(scalarWorkbook([{ row: 0, column: 0, value: { type: "text", value: "原文禰󠄀😀\r\n尾" } }, { row: 2, column: 0, value: { type: "number", value: 3 } }], metadata => {
  metadata.workbook.sheets[0]!.hyperlinks = [{ range: range(0, 0, 1, 0), url: "https://example.org/source", location: "Sheet1!A1", tooltip: "提示\r\n原文", display: "显示\r\n文字" }, { range: range(2, 0), location: "Sheet1!$A$1" }, { range: range(0, 4, 5, 6), url: "https://example.org/old", tooltip: "目标" }];
  })); sessions.push(s);
  const m = s.snapshot().manifest; s.commit(s.prepare({ protocolVersion: 1, documentId: m.documentId, baseRevision: m.revision, operations: [{ op: "setWorkbookCells", sheetId: "sheet-1", cells: [{ row: 0, column: 0, source: { kind: "value", value: { type: "text", value: "原文禰󠄀😀\r\n尾" } } }] }] })); s.clearHistory(); return s; }
const links = (s: WorkbookSession, sheetId = "sheet-1") => s.snapshot().metadata.workbook.sheets.find(sheet => sheet.sheetId === sheetId)!.hyperlinks ?? [];
function apply(s: WorkbookSession, operations: readonly DocumentOperation[]) { const m = s.snapshot().manifest; return s.commit(s.prepare({ protocolVersion: 1, documentId: m.documentId, baseRevision: m.revision, operations: [...operations] })); }
test("content copy snapshots overlapping hyperlink targets and retains exact strings and undo", async () => {
  const s = fixture(), original = links(s), tx = await createWorkbookRangeCopyTransaction(s, "sheet-1", range(0, 0, 2, 0), { row: 1, column: 0 });
  expect(links(s)).toEqual(original); s.commit(s.prepare(tx));
  expect(links(s).find(link => link.range.from.row === 1 && link.range.to.row === 2)).toMatchObject({ url: original[0]!.url, location: original[0]!.location, tooltip: original[0]!.tooltip, display: original[0]!.display });
  expect(links(s).find(link => link.range.from.row === 3)).toMatchObject({ location: "Sheet1!$A$1" }); expect(s.getCell("sheet-1", 1, 0)?.text).toBe("原文禰󠄀😀\r\n尾"); expect(s.history().undoSteps).toBe(1); s.undo(); expect(links(s)).toEqual(original);
});
test("skip blanks retains destination targets while unskipped blanks can copy or clear source link coverage", async () => {
  const s = fixture(), original = links(s);
  let tx = await createWorkbookRangeCopyTransaction(s, "sheet-1", range(0, 0, 3, 0), { row: 1, column: 5 }, { skipBlanks: true }); s.commit(s.prepare(tx));
  expect(links(s).find(link => link.range.from.row === 2 && link.range.from.column === 5)).toMatchObject({ url: "https://example.org/old" });
  expect(links(s).some(link => link.range.from.row <= 4 && link.range.to.row >= 4 && link.range.from.column <= 5 && link.range.to.column >= 5 && link.url === "https://example.org/old")).toBe(true); s.undo();
  tx = await createWorkbookRangeCopyTransaction(s, "sheet-1", range(0, 0, 3, 0), { row: 1, column: 5 }); s.commit(s.prepare(tx));
  expect(links(s).some(link => link.range.from.row <= 2 && link.range.to.row >= 2 && link.range.from.column <= 5 && link.range.to.column >= 5 && link.url === original[0]!.url)).toBe(true);
  expect(links(s).some(link => link.range.from.row <= 4 && link.range.to.row >= 4 && link.range.from.column <= 5 && link.range.to.column >= 5)).toBe(false);
  expect(links(s).filter(link => link.url === "https://example.org/old").every(link => link.tooltip === "目标")).toBe(true); s.undo(); expect(links(s)).toEqual(original);
});
test.each(["values", "formulas", "formats"] as const)("%s copy preserves destination hyperlink metadata", async mode => {
  const s = fixture(), original = links(s); s.commit(s.prepare(await createWorkbookRangeCopyTransaction(s, "sheet-1", range(0, 0, 2, 0), { row: 1, column: 5 }, { mode }))); expect(links(s)).toEqual(original);
});
test("cross-sheet links retain explicit destinations; repeated equivalent and self copies have no net change", async () => {
  const s = fixture(); apply(s, [{ op: "insertWorkbookSheet", name: "Other" }]); const targetSheetId = s.snapshot().metadata.workbook.sheets[1]!.sheetId;
  s.clearHistory(); const original = links(s); const tx = await createWorkbookRangeCopyTransaction(s, "sheet-1", range(0, 0, 2, 0), { row: 255, column: 32 }, { targetSheetId }); s.commit(s.prepare(tx));
  expect(links(s, targetSheetId)[0]).toMatchObject({ range: range(255, 32, 256, 32), url: original[0]!.url, location: original[0]!.location, tooltip: original[0]!.tooltip });
  expect(s.commit(s.prepare(await createWorkbookRangeCopyTransaction(s, "sheet-1", range(0, 0, 2, 0), { row: 255, column: 32 }, { targetSheetId }))).changed).toBe(false);
  expect(s.commit(s.prepare(await createWorkbookRangeCopyTransaction(s, "sheet-1", range(0, 0, 2, 0), { row: 0, column: 0 })))).toMatchObject({ changed: false }); expect(s.history().undoSteps).toBe(1);
});
test("link plan budgets, cancellation and stale copy plans never partially mutate source", async () => {
  const s = fixture(), before = s.serialize(); await expect(createWorkbookRangeCopyTransaction(s, "sheet-1", range(0, 0), { row: 0, column: 5 }, { maxSteps: 1 })).rejects.toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded" })); expect(s.serialize()).toEqual(before);
  const abort = new AbortController(); await expect(createWorkbookRangeCopyTransaction(s, "sheet-1", range(0, 0, 299, 0), { row: 0, column: 5 }, { signal: abort.signal, yield: async () => { abort.abort(); } })).rejects.toThrowError(expect.objectContaining({ code: "OperationCancelled" })); expect(s.serialize()).toEqual(before);
  const tx = await createWorkbookRangeCopyTransaction(s, "sheet-1", range(0, 0), { row: 0, column: 5 }); apply(s, [{ op: "setWorkbookHyperlink", sheetId: "sheet-1", range: range(20, 0), target: { url: "https://example.org/external" } }]); const current = s.serialize(); expect(() => s.prepare(tx)).toThrowError(expect.objectContaining({ code: "RevisionConflict" })); expect(s.serialize()).toEqual(current);
});
