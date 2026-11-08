import { afterEach, expect, test } from "vitest";
import { createDocumentSession, openDocument } from "../src/index.js";
import type { DocumentOperation, WorkbookSession } from "../src/index.js";
import { scalarWorkbook } from "./workbook-scalar-fixture.js";

const sessions: WorkbookSession[] = [];
afterEach(() => { for (const session of sessions.splice(0)) session.dispose(); });
function session() { const value = createDocumentSession(scalarWorkbook([])); sessions.push(value); return value; }
function transaction(s: WorkbookSession, operations: DocumentOperation[]) { const manifest = s.snapshot().manifest; return { protocolVersion: 1 as const, documentId: manifest.documentId, baseRevision: manifest.revision, operations }; }
function apply(s: WorkbookSession, operations: DocumentOperation[]) { return s.commit(s.prepare(transaction(s, operations))); }
function text(s: WorkbookSession, html: string) { apply(s, [{ op: "setWorkbookCells", sheetId: "sheet-1", cells: [{ row: 0, column: 0, source: { kind: "richText", html } }] }]); s.clearHistory(); }
const format = (from: number, to: number, expectedText: string, properties = { color: "#cc2200", "font-weight": "bold" }): Extract<DocumentOperation, { op: "formatWorkbookText" }> => ({ op: "formatWorkbookText", sheetId: "sheet-1", at: { row: 0, column: 0 }, from, to, expectedText, properties });

test("partial fonts preserve exact text, existing runs, CRLF, variation sequences and one-step package/history ownership", () => {
  const s = session(); text(s, '<span class="cell-default">Ab</span><span class="cell-default">禰󠄀😀&#13;<br>尾</span>');
  const before = s.snapshot(), cell = s.getCell("sheet-1", 0, 0)!;
  apply(s, [format(1, 9, "b禰󠄀😀\r\n")]);
  const changed = s.getCell("sheet-1", 0, 0)!;
  expect(changed.text).toBe(cell.text); expect(changed.value).toEqual(cell.value); expect(changed.cssClass).toBe(cell.cssClass); expect(changed.formatId).toBe(cell.formatId);
  expect(changed.html).toContain('<span class="cell-default">A<span class="workbook-run-'); expect(changed.html).toContain("&#13;"); expect(changed.html).toContain("</span>尾</span>");
  expect(s.snapshot().metadata).toEqual(before.metadata); expect(s.history().undoSteps).toBe(1);
  expect(openDocument(s.serialize(), { kind: "workbook" }).snapshot()).toEqual(s.snapshot());
  s.undo(); expect(s.snapshot().styles).toEqual(before.styles); expect(s.snapshot().manifest.chunks).toEqual(before.manifest.chunks); expect(s.getCell("sheet-1", 0, 0)).toEqual(cell);
  s.redo(); expect(s.getCell("sheet-1", 0, 0)).toEqual(changed);
});

test("formatting a plain text constant creates only selected runs and later whole-cell and partial edits share one candidate", () => {
  const s = session(); apply(s, [{ op: "setWorkbookCells", sheetId: "sheet-1", cells: [{ row: 0, column: 0, source: { kind: "value", value: { type: "text", value: "0012😀中文" } } }] }]); s.clearHistory();
  apply(s, [format(0, 2, "00"), { op: "formatWorkbookRanges", sheetId: "sheet-1", ranges: [{ from: { row: 0, column: 0 }, to: { row: 0, column: 0 } }], properties: { "font-style": "italic" } },
    format(6, 8, "中文", { color: "blue", "font-weight": "normal" })]);
  expect(s.getCell("sheet-1", 0, 0)).toMatchObject({ value: { type: "text", value: "0012😀中文" }, text: "0012😀中文" });
  expect(s.history().undoSteps).toBe(1); expect(s.snapshot().styles[0]!.css).toContain("color:blue!important");
  expect(openDocument(s.serialize(), { kind: "workbook" }).snapshot()).toEqual(s.snapshot());
});

test("invalid text/UTF-16 ranges, non-font settings, numbers and formula caches cannot be partially formatted", () => {
  const s = session(); text(s, "A禰󠄀😀Z"); const before = s.snapshot();
  for (const operation of [format(2, 4, "󠄀"), format(4, 5, "\ud83d"), format(0, 1, "wrong"), format(0, 0, ""), format(0, 999, ""),
    { ...format(0, 1, "A"), properties: { "background-color": "red" } }, { ...format(0, 1, "A"), properties: { color: "url(https://private.invalid)" } }]) {
    expect(() => apply(s, [operation])).toThrow(); expect(s.snapshot()).toEqual(before);
  }
  for (const source of [{ kind: "value" as const, value: { type: "number" as const, value: 123 } }, { kind: "formula" as const, expression: "=123" }]) {
    apply(s, [{ op: "setWorkbookCells", sheetId: "sheet-1", cells: [{ row: 0, column: 0, source }] }]); const saved = s.snapshot();
    expect(() => apply(s, [format(0, 1, "1")])).toThrowError(expect.objectContaining({ code: "InvalidOperation" })); expect(s.snapshot()).toEqual(saved);
  }
});

test("empty font patches retain redo, stale candidates and asynchronous cancellation never publish selected fragments", async () => {
  const s = session(); text(s, "ABC"); apply(s, [format(0, 1, "A")]); s.undo(); const before = s.snapshot();
  expect(apply(s, [{ ...format(0, 1, "A"), properties: {} }]).changed).toBe(false); expect(s.history().redoSteps).toBe(1);
  const candidate = s.prepare(transaction(s, [format(1, 2, "B")])); apply(s, [format(2, 3, "C")]);
  expect(() => s.commit(candidate)).toThrowError(expect.objectContaining({ code: "RevisionConflict" })); s.discard(candidate); s.undo();
  const abort = { aborted: false }, saved = s.snapshot();
  await expect(s.prepareAsync(transaction(s, [format(0, 1, "A")]), {}, { signal: abort, yield: async () => { abort.aborted = true; } })).rejects.toMatchObject({ code: "OperationCancelled" });
  expect(s.snapshot()).toEqual(saved); expect(s.snapshot().styles).toEqual(before.styles);
});
