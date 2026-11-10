import { readFileSync } from "node:fs";
import { expect, test } from "vitest";
import { createDocumentSession, openDocument, readWorkbookCellSource, resolveWorkbookCellFormat } from "@doffi/core";
import { importXlsx, readDocumentZip, writeDocumentZip } from "@doffi/office";
import { createWorkbookRangeCopyTransaction } from "@doffi/editor";

test.each(["01-values-text", "02-styles-dates", "04-formulas-names"])("native %s copies across worksheets and chunks without changing its source or named style definitions", async name => {
  const model = openDocument(await importXlsx(readFileSync("examples/xlsx/" + name + ".xlsx"), { documentId: "copy-native" }), { kind: "workbook" }), session = createDocumentSession(model);
  try {
    session.commit(session.prepare({ protocolVersion: 1, documentId: "copy-native", baseRevision: 0, operations: [{ op: "insertWorkbookSheet", name: "Destination" }] }));
    session.clearHistory(); const original = session.snapshot(), source = original.metadata.workbook.sheets[0]!, destination = original.metadata.workbook.sheets.find(sheet => sheet.name === "Destination")!;
    const cells = original.manifest.chunks.filter(chunk => chunk.sheetId === source.sheetId).flatMap(chunk => [0, 128].flatMap(offset => session.readWindow({ sheetId: source.sheetId, row: chunk.row + offset, column: chunk.column, rows: 128, columns: 32 })));
    expect(cells.length).toBeGreaterThan(0);
    const range = { from: { row: Math.min(...cells.map(cell => cell.row)), column: Math.min(...cells.map(cell => cell.column)) },
      to: { row: Math.max(...cells.map(cell => cell.row)), column: Math.max(...cells.map(cell => cell.column)) } }, target = { row: 255, column: 31 };
    const before = session.serialize();
    const transaction = await createWorkbookRangeCopyTransaction(session, source.sheetId, range, target, { targetSheetId: destination.sheetId });
    expect(session.serialize()).toEqual(before); session.commit(session.prepare(transaction));
    const reopened = openDocument(readDocumentZip(writeDocumentZip(session.serialize())), { kind: "workbook" });
    expect(reopened.snapshot().metadata.workbook.cellStyles).toEqual(original.metadata.workbook.cellStyles);
    for (const cell of cells) {
      const at = { sheetId: destination.sheetId, row: target.row + cell.row - range.from.row, column: target.column + cell.column - range.from.column }, copied = reopened.getCell(at.sheetId, at.row, at.column)!;
      expect(copied).toBeDefined(); expect(session.getCell(source.sheetId, cell.row, cell.column)?.html).toBe(cell.html);
      const format = resolveWorkbookCellFormat(source, cell, cell);
      expect(copied.cssClass).toBe(format.cssClass); if (format.numberFormatId) expect(copied.formatId).toBe(format.numberFormatId);
      if (cell.formula) expect(copied.formula).toMatchObject({ cacheState: "stale" });
      else expect(readWorkbookCellSource(reopened, at)).toEqual(readWorkbookCellSource(model, { sheetId: source.sheetId, row: cell.row, column: cell.column }));
    }
    expect(session.history().undoSteps).toBe(1); session.undo();
    expect(session.snapshot().manifest.chunks).toEqual(original.manifest.chunks); expect(session.snapshot().metadata).toEqual(original.metadata);
  } finally { session.dispose(); }
});
