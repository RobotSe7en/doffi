import { expect, test } from "vitest";
import { createDocumentSession, openDocument } from "../src/index.js";
import type { DocumentOperation, WorkbookSession } from "../src/index.js";
import { scalarWorkbook } from "./workbook-scalar-fixture.js";
const range = (row: number, column: number, lastRow = row, lastColumn = column) => ({ from: { row, column }, to: { row: lastRow, column: lastColumn } });
function apply(session: WorkbookSession, operations: DocumentOperation[]) {
  const manifest = session.snapshot().manifest;
  return session.commit(session.prepare({ protocolVersion: 1, documentId: manifest.documentId, baseRevision: manifest.revision, operations }));
}
test("format-only edits retain imported formula caches, rich UTF-16 source and unrelated chunk bytes in one undo", () => {
  const files = new Map(scalarWorkbook([{ row: 255, column: 31, value: { type: "number", value: 44 }, formula: '=IF(A1>0,1,2)' },
    { row: 600, column: 60, value: { type: "text", value: "unchanged" } }]).serialize());
  files.set("styles.css", new TextEncoder().encode(new TextDecoder().decode(files.get("styles.css")!) + ".strong{font-weight:700}"));
  const session = createDocumentSession(openDocument(files, { kind: "workbook" }));
  apply(session, [{ op: "setWorkbookCells", sheetId: "sheet-1", cells: [{ row: 256, column: 32, source: { kind: "richText", html: '<span class="strong">中文&#13;禰󠄀</span><br><span class="cell-default">😀</span>' } }] }]); session.clearHistory();
  const before = session.serialize(), rich = session.getCell("sheet-1", 256, 32)!, formula = session.getCell("sheet-1", 255, 31)!;
  const receipt = apply(session, [{ op: "setWorkbookNumberFormat", format: { formatId: "fixed", code: '"值 "0.00' } },
    { op: "setWorkbookCellFormats", sheetId: "sheet-1", cells: [{ row: 255, column: 31, cssClass: "strong", formatId: "fixed" }, { row: 256, column: 32, formatId: "fixed" }] }]);
  expect(receipt.changed).toBe(true); expect(session.history().undoSteps).toBe(1);
  expect(session.getCell("sheet-1", 255, 31)).toMatchObject({ value: formula.value, formula: formula.formula, text: formula.text, formatId: "fixed" });
  expect(session.getCell("sheet-1", 256, 32)).toMatchObject({ value: rich.value, text: rich.text });
  expect(session.getCell("sheet-1", 256, 32)?.html.slice(session.getCell("sheet-1", 256, 32)!.html.indexOf(">") + 1)).toBe(rich.html.slice(rich.html.indexOf(">") + 1));
  const untouched = session.snapshot().manifest.chunks.find(chunk => chunk.row === 512)!;
  expect(session.serialize().get(untouched.path)).toEqual(before.get(untouched.path));
  session.undo(); expect(session.getCell("sheet-1", 255, 31)?.html).toBe(formula.html); expect(session.getCell("sheet-1", 256, 32)?.html).toBe(rich.html);
  session.redo(); expect(session.getCell("sheet-1", 255, 31)?.formula?.cacheState).toBe("imported");
});
test("format no-ops, restoration and sparse blanks do not erase formula or rich content", () => {
  const source = scalarWorkbook([{ row: 0, column: 0, value: { type: "number", value: 99 }, formula: "=1" }]);
  const session = createDocumentSession(source), before = session.serialize();
  expect(apply(session, [{ op: "setWorkbookCellFormats", sheetId: "sheet-1", cells: [{ row: 0, column: 0, cssClass: "cell-default" }] },
    { op: "setWorkbookCellFormats", sheetId: "sheet-1", cells: [{ row: 0, column: 0, cssClass: null }] }])).toMatchObject({ changed: false, revision: 0 });
  expect(session.serialize()).toEqual(before); expect(session.history().undoSteps).toBe(0);
  apply(session, [{ op: "setWorkbookCellFormats", sheetId: "sheet-1", cells: [{ row: 1_048_575, column: 16383, cssClass: "cell-default" }] }]);
  expect(session.getCell("sheet-1", 1_048_575, 16383)?.value).toEqual({ type: "blank" }); expect(session.snapshot().manifest.chunks).toHaveLength(2);
  apply(session, [{ op: "setWorkbookCellFormats", sheetId: "sheet-1", cells: [{ row: 1_048_575, column: 16383, cssClass: null }] }]);
  expect(session.getCell("sheet-1", 1_048_575, 16383)).toBeUndefined(); expect(session.snapshot().manifest.chunks).toHaveLength(1);
  expect(session.getCell("sheet-1", 0, 0)?.formula?.cacheState).toBe("imported");
});
test("number-format definitions update shared displays without changing cell sources, and a restored definition adds no history", () => {
  const session = createDocumentSession(scalarWorkbook([{ row: 0, column: 0, value: { type: "number", value: .25 } }], metadata => {
    metadata.workbook.numberFormats.push({ formatId: "shared", code: "0.00" }); metadata.workbook.sheets[0]!.defaultNumberFormatId = "shared";
  })), original = session.getCell("sheet-1", 0, 0)!.html;
  const before = session.serialize();
  expect(apply(session, [{ op: "setWorkbookNumberFormat", format: { formatId: "shared", code: "0%" } },
    { op: "setWorkbookNumberFormat", format: { formatId: "shared", code: "0.00" } }])).toMatchObject({ changed: false });
  expect(session.serialize()).toEqual(before);
  const result = apply(session, [{ op: "setWorkbookNumberFormat", format: { formatId: "shared", code: "0%" } }]);
  expect(result.changedRanges).toEqual([]); expect(result.affectedSheetIds).toEqual(["sheet-1"]);
  expect(session.getCell("sheet-1", 0, 0)?.html).toBe(original); expect(session.snapshot().metadata.workbook.numberFormats.find(format => format.formatId === "shared")).toEqual({ formatId: "shared", code: "0%" });
  session.undo(); expect(session.snapshot().metadata.workbook.numberFormats.find(format => format.formatId === "shared")).toEqual({ formatId: "shared", code: "0.00" });
});
test("format-only changes preserve an imported stop-validation violation without bypassing later content validation", () => {
  const session = createDocumentSession(scalarWorkbook([{ row: 0, column: 0, value: { type: "number", value: 5 } }], metadata => {
    metadata.workbook.sheets[0]!.dataValidations = [{ ruleId: "existing", kind: "whole", operator: "between", origin: { row: 0, column: 0 }, ranges: [range(0, 0)], formulas: ["=1", "=3"], allowBlank: true,
      showInputMessage: true, showErrorMessage: true, errorStyle: "stop", errorTitle: "", error: "", promptTitle: "", prompt: "" }];
  }));
  apply(session, [{ op: "setWorkbookCellFormats", sheetId: "sheet-1", cells: [{ row: 0, column: 0, cssClass: "cell-default" }] }]);
  expect(session.getCell("sheet-1", 0, 0)?.value).toEqual({ type: "number", value: 5 }); expect(session.history().undoSteps).toBe(1);
  expect(() => apply(session, [{ op: "setWorkbookCells", sheetId: "sheet-1", cells: [{ row: 0, column: 0, source: { kind: "value", value: { type: "number", value: 6 } } }] }])).toThrowError(expect.objectContaining({ code: "InvalidOperation" }));
  expect(session.getCell("sheet-1", 0, 0)?.value).toEqual({ type: "number", value: 5 }); expect(session.history().undoSteps).toBe(1);
});
test("invalid formats, duplicate or covered coordinates and unsafe links fail atomically", () => {
  const session = createDocumentSession(scalarWorkbook([], metadata => { metadata.workbook.sheets[0]!.merges = [range(0, 0, 1, 1)]; })), before = session.serialize();
  const edits: DocumentOperation[][] = [
    [{ op: "setWorkbookCellFormats", sheetId: "sheet-1", cells: [{ row: 0, column: 0, formatId: "missing" }] }],
    [{ op: "setWorkbookCellFormats", sheetId: "sheet-1", cells: [{ row: 0, column: 1, cssClass: "cell-default" }] }],
    [{ op: "setWorkbookCellFormats", sheetId: "sheet-1", cells: [{ row: 0, column: 0 }, { row: 0, column: 0 }] }],
    [{ op: "setWorkbookNumberFormat", format: { formatId: "bad", code: '"unclosed' } }],
    [{ op: "setWorkbookHyperlink", sheetId: "sheet-1", range: range(0, 0), target: { url: "javascript:alert(1)" } }],
  ];
  for (const operations of edits) { expect(() => apply(session, operations)).toThrow(); expect(session.serialize()).toEqual(before); expect(session.history().undoSteps).toBe(0); }
});
test("hyperlink replacement and partial removal keep exact outside rectangles, with no implicit cell nodes", () => {
  const session = createDocumentSession(scalarWorkbook([], metadata => { metadata.workbook.sheets[0]!.hyperlinks = [{ range: range(0, 0, 9, 9), url: "https://example.org/old", tooltip: "原始" }]; }));
  const chunks = session.snapshot().manifest.chunks, before = session.serialize();
  apply(session, [{ op: "setWorkbookHyperlink", sheetId: "sheet-1", range: range(2, 2, 4, 4), target: { location: "Sheet1!A1", tooltip: "内部" } }]);
  const links = session.snapshot().metadata.workbook.sheets[0]!.hyperlinks!;
  expect(links.filter(link => link.url).map(link => link.range)).toEqual([range(0, 0, 1, 9), range(5, 0, 9, 9), range(2, 0, 4, 1), range(2, 5, 4, 9)]);
  expect(links.find(link => link.location)?.range).toEqual(range(2, 2, 4, 4)); expect(session.snapshot().manifest.chunks).toEqual(chunks); expect(session.history().undoSteps).toBe(1);
  apply(session, [{ op: "setWorkbookHyperlink", sheetId: "sheet-1", range: range(3, 3), target: null }]);
  expect(session.snapshot().metadata.workbook.sheets[0]!.hyperlinks!.some(link => link.range.from.row <= 3 && link.range.to.row >= 3 && link.range.from.column <= 3 && link.range.to.column >= 3)).toBe(false);
  session.undo(); session.undo(); expect(session.snapshot().metadata.workbook.sheets[0]!.hyperlinks).toEqual([{ range: range(0, 0, 9, 9), url: "https://example.org/old", tooltip: "原始" }]);
  expect([...before].filter(([path]) => path.includes("sheet-1/"))).toEqual([...session.serialize()].filter(([path]) => path.includes("sheet-1/")));
});
test("equivalent hyperlink fragmentation and absent/empty lists restore the original source without history", () => {
  for (const original of [undefined, [], [{ range: range(0, 0, 9, 9), url: "https://example.org/old" }]]) {
    const session = createDocumentSession(scalarWorkbook([], metadata => { if (original) metadata.workbook.sheets[0]!.hyperlinks = original; })), before = session.serialize();
    const receipt = apply(session, [
      { op: "setWorkbookHyperlink", sheetId: "sheet-1", range: range(2, 2, 4, 4), target: { url: "https://example.org/new" } },
      { op: "setWorkbookHyperlink", sheetId: "sheet-1", range: range(2, 2, 4, 4), target: original?.length ? { url: "https://example.org/old" } : null },
    ]);
    expect(receipt).toMatchObject({ changed: false, revision: 0 }); expect(session.serialize()).toEqual(before); expect(session.history().undoSteps).toBe(0);
  }
});
