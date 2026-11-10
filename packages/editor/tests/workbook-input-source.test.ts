import { expect, test } from "vitest";
import { createDocumentSession, createWorkbook, openDocument } from "@doffi/core";
import { createWorkbookInputTransaction, readWorkbookInputText } from "../src/workbook/input-source.js";

test("input transaction spelling retains literal text types and idle source bytes", () => {
  const session = createDocumentSession(openDocument(createWorkbook({ documentId: "input-source" }), { kind: "workbook" })), sheetId = "sheet-1", at = { row: 0, column: 0 };
  for (const value of ["", "12", "=bad(", "'quoted", "00123", "=" + "A".repeat(32766)]) {
    session.commit(session.prepare(createWorkbookInputTransaction(session, sheetId, at, value, "text")));
    const before = session.serialize(), revision = session.snapshot().manifest.revision, text = readWorkbookInputText(session.getCell(sheetId, 0, 0));
    const receipt = session.commit(session.prepare(createWorkbookInputTransaction(session, sheetId, at, text)));
    expect(receipt).toMatchObject({ changed: false, revision }); expect(session.serialize()).toEqual(before);
    expect(session.getCell(sheetId, 0, 0)?.value).toEqual({ type: "text", value });
  }
  session.dispose();
});

test("input transactions preserve untouched rich source and allow an explicit whole formula replacement", () => {
  const files = new Map(createWorkbook({ documentId: "input-source" }));
  files.set("styles.css", new TextEncoder().encode(new TextDecoder().decode(files.get("styles.css")!) + '.emphasis{font-style:italic;--office-font-east-asia:"Noto Serif CJK SC"}'));
  const session = createDocumentSession(openDocument(files, { kind: "workbook" })), sheetId = "sheet-1", at = { row: 256, column: 32 };
  session.commit(session.prepare({ protocolVersion: 1, documentId: "input-source", baseRevision: 0, operations: [{ op: "setWorkbookCells", sheetId,
    cells: [{ ...at, source: { kind: "richText", html: '<span class="cell-default">ABC&#13;&#10;</span><span class="emphasis">中文禰󠄀</span>' } }] }] }));
  const cell = session.getCell(sheetId, at.row, at.column)!, original = cell.html;
  session.commit(session.prepare(createWorkbookInputTransaction(session, sheetId, at, readWorkbookInputText(cell) + "尾")));
  expect(session.getCell(sheetId, at.row, at.column)?.html).toContain('<span class="cell-default">ABC&#13;\n</span><span class="emphasis">中文禰󠄀尾</span>');
  session.undo(); expect(session.getCell(sheetId, at.row, at.column)?.html).toBe(original);
  session.commit(session.prepare(createWorkbookInputTransaction(session, sheetId, at, "=1+1")));
  expect(session.getCell(sheetId, at.row, at.column)?.formula?.expression).toBe("=1+1");
  session.dispose();
});

test("invalid input cannot produce a partially applicable command", () => {
  const session = createDocumentSession(openDocument(createWorkbook({ documentId: "input-source" }), { kind: "workbook" })), before = session.serialize();
  for (const text of ["=", "\u0000", "\ud800", "x".repeat(32768)]) expect(() => createWorkbookInputTransaction(session, "sheet-1", { row: 0, column: 0 }, text)).toThrow();
  expect(() => createWorkbookInputTransaction(session, "absent", { row: 0, column: 0 }, "1")).toThrow();
  expect(() => createWorkbookInputTransaction(session, "sheet-1", { row: 1_048_576, column: 0 }, "1")).toThrow();
  expect(session.serialize()).toEqual(before); expect(session.history().undoSteps).toBe(0); session.dispose();
});
