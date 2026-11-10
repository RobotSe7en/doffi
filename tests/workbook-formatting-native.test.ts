import { readFileSync } from "node:fs";
import { expect, test } from "vitest";
import { createDocumentSession, openDocument, readWorkbookCellSource } from "@doffi/core";
import { importXlsx } from "@doffi/office";

test.each(["01-values-text", "02-styles-dates", "04-formulas-names"])("native %s retains source types, rich fonts, style bindings and formula caches through formatting and package reopen", async name => {
  const model = openDocument(await importXlsx(readFileSync("examples/xlsx/" + name + ".xlsx"), { documentId: "format-native" }), { kind: "workbook" });
  const session = createDocumentSession(model), original = session.snapshot(), sheetId = original.metadata.workbook.sheets[0]!.sheetId;
  const cells = original.manifest.chunks.filter(chunk => chunk.sheetId === sheetId).flatMap(chunk => [0, 128].flatMap(offset => session.readWindow({ sheetId, row: chunk.row + offset, column: chunk.column, rows: 128, columns: 32 })));
  const inputs = cells.map(cell => ({ cell, source: readWorkbookCellSource(session, { sheetId, row: cell.row, column: cell.column }) }));
  const at = { row: cells[0]!.row, column: cells[0]!.column };
  session.commit(session.prepare({ protocolVersion: 1, documentId: "format-native", baseRevision: 0, operations: [
    { op: "setWorkbookNumberFormat", format: { formatId: "format-test-general", builtinId: 0 } },
    { op: "setWorkbookCellFormats", sheetId, cells: cells.map(cell => ({ row: cell.row, column: cell.column, formatId: "format-test-general" })) },
    { op: "setWorkbookHyperlink", sheetId, range: { from: at, to: at }, target: { url: "https://example.org/format", tooltip: "明确目标" } },
  ] }));
  const reopened = openDocument(session.serialize(), { kind: "workbook" });
  for (const { cell, source } of inputs) {
    const location = { sheetId, row: cell.row, column: cell.column }, current = reopened.getCell(sheetId, cell.row, cell.column)!;
    expect(readWorkbookCellSource(reopened, location)).toEqual(source);
    expect(current.value).toEqual(cell.value); expect(current.formula).toEqual(cell.formula); expect(current.text).toBe(cell.text); expect(current.cssClass).toBe(cell.cssClass);
    expect(current.formatId).toBe("format-test-general");
  }
  expect(reopened.snapshot().metadata.workbook.cellStyles).toEqual(original.metadata.workbook.cellStyles);
  expect(session.history().undoSteps).toBe(1); session.undo();
  expect(session.snapshot().metadata).toEqual(original.metadata); expect(session.snapshot().manifest.chunks).toEqual(original.manifest.chunks);
  for (const { cell } of inputs) expect(session.getCell(sheetId, cell.row, cell.column)?.html).toBe(cell.html);
  session.dispose();
});
