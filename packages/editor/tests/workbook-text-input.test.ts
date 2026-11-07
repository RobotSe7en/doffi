import { expect, test } from "vitest";
import { createDocumentSession, createWorkbook, openDocument } from "@doffi/core";
import { createWorkbookPasteTransaction, parseWorkbookTextInput, parseWorkbookTsv } from "../src/workbook/text-input.js";
const origin = { row: 0, column: 0 };
const source = (text: string) => parseWorkbookTextInput(text, origin);
test("text input conservatively recognizes constants and formulas while retaining IDs and dates", () => {
  for (const text of ["00123", "1234567890123456", "2024-01-02", "1/2", " 12 "]) expect(source(text)).toEqual({ kind: "value", value: { type: "text", value: text } });
  for (const text of ["12", "-1.25", ".5", "1e20", "+2"]) expect(source(text)).toEqual({ kind: "value", value: { type: "number", value: Number(text) } });
  expect(source("false")).toEqual({ kind: "value", value: { type: "boolean", value: false } });
  expect(source("=A1+1")).toEqual({ kind: "formula", expression: "=A1+1" });
  expect(source("'=A1")).toEqual({ kind: "value", value: { type: "text", value: "=A1" } });
  expect(parseWorkbookTextInput("00123", origin, "number")).toEqual({ kind: "value", value: { type: "number", value: 123 } });
  expect(parseWorkbookTextInput("=A1", origin, "text")).toEqual({ kind: "value", value: { type: "text", value: "=A1" } });
  expect(source("")).toEqual({ kind: "value", value: { type: "blank" } }); expect(parseWorkbookTextInput("", origin, "text")).toEqual({ kind: "value", value: { type: "text", value: "" } });
  expect(source("中文😀禰󠄀\n混排")).toEqual({ kind: "value", value: { type: "text", value: "中文😀禰󠄀\n混排" } });
});
test("quoted TSV preserves tabs, multiline Unicode, quotes and trailing empty fields", () => {
  expect(parseWorkbookTsv('"甲\t乙"\t"中文😀\r\n禰󠄀"\r\n"a""b"\t\r\n')).toEqual([["甲\t乙", "中文😀\n禰󠄀"], ['a"b', ""]]);
  expect(parseWorkbookTsv("a\n\n")).toEqual([["a"], [""]]); expect(parseWorkbookTsv("")).toEqual([[""]]);
  expect(Object.isFrozen(parseWorkbookTsv("a\tb")[0])).toBe(true);
});
test("clipboard and literal failures reject before any source publication", () => {
  for (const value of ['"unclosed', '"a"junk', "a\tb\nc", "\ud800", "\u0000"]) expect(() => parseWorkbookTsv(value)).toThrow();
  expect(() => parseWorkbookTsv("a\tb", { maxCells: 1 })).toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded" }));
  expect(() => parseWorkbookTsv("abc", { maxCodeUnits: 2 })).toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded" }));
  expect(() => source("=")).toThrow(); expect(() => source("x".repeat(32768))).toThrow();
  expect(() => parseWorkbookTextInput("Infinity", origin, "number")).toThrow();
});
test("paste is one isolated transaction and undo step across block boundaries", () => {
  const session = createDocumentSession(openDocument(createWorkbook({ documentId: "input" }), { kind: "workbook" })), before = session.serialize();
  const transaction = createWorkbookPasteTransaction(session.snapshot(), "sheet-1", { row: 255, column: 31 }, "00123\t=AF256+1\n中文😀禰󠄀\tTRUE");
  const candidate = session.prepare(transaction); expect(session.serialize()).toEqual(before); session.commit(candidate);
  expect(session.getCell("sheet-1", 255, 31)?.value).toEqual({ type: "text", value: "00123" });
  expect(session.getCell("sheet-1", 255, 32)?.formula?.expression).toBe("=AF256+1");
  expect(session.getCell("sheet-1", 256, 31)?.value).toEqual({ type: "text", value: "中文😀禰󠄀" }); expect(session.history().undoSteps).toBe(1);
  session.undo(); expect(session.snapshot().manifest.chunks).toEqual([]);
  expect(() => createWorkbookPasteTransaction(session.snapshot(), "sheet-1", { row: 1_048_575, column: 16_383 }, "a\tb")).toThrow();
  expect(() => session.prepare(transaction)).toThrowError(expect.objectContaining({ code: "RevisionConflict" }));
});
