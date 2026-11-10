import { afterEach, expect, test } from "vitest";
import { createDocumentSession, createWorkbook, openDocument } from "@doffi/core";
import type { WorkbookSession } from "@doffi/core";
import { createWorkbookTextFormatTransaction, readWorkbookInputText } from "../src/index.js";
import { workbookInputTextView } from "../src/workbook/text-projection.js";

const sessions: WorkbookSession[] = [];
afterEach(() => { for (const session of sessions.splice(0)) session.dispose(); });
function session(value: string) {
  const s = createDocumentSession(openDocument(createWorkbook({ documentId: "native-text-font" }), { kind: "workbook" })); sessions.push(s);
  s.commit(s.prepare({ protocolVersion: 1, documentId: "native-text-font", baseRevision: 0, operations: [{ op: "setWorkbookCells", sheetId: "sheet-1", cells: [
    { row: 0, column: 0, source: { kind: "value", value: { type: "text", value } } },
  ] }] })); s.clearHistory(); return s;
}
test("font selection maps native LF and literal escapes to exact source UTF-16 without changing its text or revision", () => {
  for (const value of ["=SUM(A1)", "'quoted", "00123", "A\r\n禰󠄀😀Z"]) {
    const s = session(value), before = s.snapshot(), input = workbookInputTextView(readWorkbookInputText(s.getCell("sheet-1", 0, 0))), properties = { color: "red" };
    const transaction = createWorkbookTextFormatTransaction(s, "sheet-1", { row: 0, column: 0 }, { from: 0, to: input.length, text: input }, properties);
    expect(transaction.operations).toEqual([{ op: "formatWorkbookText", sheetId: "sheet-1", at: { row: 0, column: 0 }, from: 0, to: value.length, expectedText: value, properties }]);
    properties.color = "blue"; expect(s.snapshot()).toEqual(before);
    s.commit(s.prepare(transaction)); expect(s.getCell("sheet-1", 0, 0)?.text).toBe(value); expect(s.snapshot().styles[0]!.css).toContain("color:red!important");
    expect(s.history().undoSteps).toBe(1); s.undo(); expect(s.snapshot().manifest.chunks).toEqual(before.manifest.chunks);
  }
  const s = session("A\r\nB"), result = createWorkbookTextFormatTransaction(s, "sheet-1", { row: 0, column: 0 }, { from: 1, to: 3, text: "A\nB" }, { color: "red" });
  expect(result.operations[0]).toMatchObject({ from: 1, to: 4, expectedText: "\r\nB" });
});
test("font conversion refuses changed input, literal-prefix-only ranges, UTF-16 splits and malformed selection contracts", () => {
  const s = session("=禰󠄀😀"), cell = s.getCell("sheet-1", 0, 0)!, text = readWorkbookInputText(cell), before = s.serialize();
  for (const selected of [{ from: 0, to: 1, text }, { from: 2, to: 3, text }, { from: 5, to: 6, text }, { from: 0, to: text.length + 1, text },
    { from: 0, to: 2, text: "changed" }, { from: 0, to: 2, text, extra: true }])
    expect(() => createWorkbookTextFormatTransaction(s, "sheet-1", { row: 0, column: 0 }, selected, { color: "red" })).toThrow();
  expect(s.serialize()).toEqual(before);
});
