import { expect, test } from "vitest";
import { createDocumentSession, createWorkbook, openDocument, replaceWorkbookCellText } from "../src/index.js";
import type { WorkbookCellSource } from "../src/index.js";
const at = { sheetId: "sheet-1", row: 0, column: 0 };
function seeded(html: string) {
  const source = createDocumentSession(openDocument(createWorkbook({ documentId: "cell-text" }), { kind: "workbook" }));
  source.commit(source.prepare({ protocolVersion: 1, documentId: "cell-text", baseRevision: 0, operations: [{ op: "setWorkbookCells", sheetId: at.sheetId, cells: [{ row: 0, column: 0, source: { kind: "richText", html } }] }] })); source.clearHistory(); return source;
}
function apply(source: ReturnType<typeof seeded>, replacement: WorkbookCellSource): void {
  source.commit(source.prepare({ protocolVersion: 1, documentId: "cell-text", baseRevision: source.snapshot().manifest.revision, operations: [{ op: "setWorkbookCells", sheetId: at.sheetId, cells: [{ row: 0, column: 0, source: replacement }] }] }));
}
test("multiple cell text splices retain untouched font spans, exact UTF-16 and one undo step", () => {
  const source = seeded('<span class="cell-default">old</span><span class="cell-default">中文😀禰󠄀</span><span class="cell-default">old</span>');
  const original = source.getCell(at.sheetId, 0, 0)!, tail = original.text.length - 3, before = source.serialize();
  const replacement = replaceWorkbookCellText(source, at, [{ from: 0, to: 3, expectedText: "old", text: "new" }, { from: tail, to: tail + 3, expectedText: "old", text: "末尾" }]);
  expect(source.serialize()).toEqual(before);
  expect(replacement).toEqual({ kind: "richText", html: '<span class="cell-default">new</span><span class="cell-default">中文😀禰󠄀</span><span class="cell-default">末尾</span>' });
  apply(source, replacement); expect(source.history().undoSteps).toBe(1); source.undo(); expect(source.getCell(at.sheetId, 0, 0)?.html).toBe(original.html);
});
test("cross-span edits and line breaks preserve surrounding runs while inserting literal safe text", () => {
  const source = seeded('<span class="cell-default">abc</span><br><span class="cell-default">def&#13;Z</span>');
  const replacement = replaceWorkbookCellText(source, at, [{ from: 2, to: 5, expectedText: "c\nd", text: '<img onerror="x">\n新' }]);
  expect(replacement).toEqual({ kind: "richText", html: '<span class="cell-default">ab&lt;img onerror="x"&gt;<br>新</span><span class="cell-default">ef&#13;Z</span>' });
  apply(source, replacement); expect(source.getCell(at.sheetId, 0, 0)?.text).toBe('ab<img onerror="x">\n新ef\rZ');
});
test("insertion into an empty rich cell and after a trailing break retains its formatting context", () => {
  const empty = seeded('<span class="cell-default"></span>');
  expect(replaceWorkbookCellText(empty, at, [{ from: 0, to: 0, expectedText: "", text: "首行\n中文" }])).toEqual({ kind: "richText", html: '<span class="cell-default">首行<br>中文</span>' });
  const source = seeded('<span class="cell-default">A<br></span>');
  const replacement = replaceWorkbookCellText(source, at, [{ from: 2, to: 2, expectedText: "", text: "B" }]);
  expect(replacement).toEqual({ kind: "richText", html: '<span class="cell-default">A<br>B</span>' });
});
test("bad source proofs, overlap and detached variation selectors reject before changing source", () => {
  const source = seeded('<span class="cell-default">😀禰󠄀AB</span>'), before = source.serialize();
  for (const edits of [
    [{ from: 0, to: 1, expectedText: "\ud83d", text: "X" }],
    [{ from: 2, to: 3, expectedText: "禰", text: "X" }],
    [{ from: 0, to: 2, expectedText: "wrong", text: "X" }],
    [{ from: 0, to: 5, expectedText: "😀禰󠄀", text: "X" }, { from: 2, to: 5, expectedText: "禰󠄀", text: "Y" }],
    [{ from: 0, to: 2, expectedText: "😀", text: "\ud800" }],
  ]) expect(() => replaceWorkbookCellText(source, at, edits)).toThrowError(expect.objectContaining({ code: "InvalidOperation" }));
  expect(() => replaceWorkbookCellText(source, at, [{ from: 0, to: 0, expectedText: "", text: "X".repeat(32768) }])).toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded" }));
  expect(source.serialize()).toEqual(before); expect(source.history().undoSteps).toBe(0);
});
test("many disjoint replacements across run boundaries match independent string splices", () => {
  const source = seeded(Array.from({ length: 80 }, (_, index) => `<span class="cell-default">${index % 10}xy</span>`).join(""));
  const original = source.getCell(at.sheetId, 0, 0)!.text;
  const edits = Array.from({ length: 40 }, (_, index) => { const from = index * 6 + 2; return { from, to: from + 3, expectedText: original.slice(from, from + 3), text: index % 3 ? "变" : "换\n行" }; });
  let expected = original; for (const edit of [...edits].reverse()) expected = expected.slice(0, edit.from) + edit.text + expected.slice(edit.to);
  apply(source, replaceWorkbookCellText(source, at, edits)); expect(source.getCell(at.sheetId, 0, 0)?.text).toBe(expected);
});
