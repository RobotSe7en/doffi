import type { WorkbookCell, WorkbookModel } from "../runtime.js";
import { DocumentError } from "../errors.js";
import { CHUNK_COLUMNS, CHUNK_ROWS, decodeWorkbookCell } from "./chunk-index.js";
import { captureWorkbookModel, workbookSource } from "./source.js";

/** Read actual cells in worksheet order. Only one row of cells and one band
 * of compact source offsets are visited at a time. Empty grid space is free. */
export function* iterateWorkbookRows(model: WorkbookModel, sheetId: string): Generator<readonly WorkbookCell[]> {
  model = captureWorkbookModel(model);
  const source = workbookSource(model), snapshot = model.snapshot();
  if (!source || !snapshot.metadata.workbook.sheets.some(sheet => sheet.sheetId === sheetId))
    throw new DocumentError("InvalidOperation", "Reading worksheet rows requires a validated worksheet.");
  const chunks = snapshot.manifest.chunks.filter(chunk => chunk.sheetId === sheetId)
    .sort((left, right) => left.row - right.row || left.column - right.column);
  for (let start = 0; start < chunks.length;) {
    const band = chunks[start]!.row;
    let end = start + 1;
    while (chunks[end]?.row === band) end++;
    const columns = chunks.slice(start, end).map(chunk => ({
      cells: source.chunks.get(chunk.path)!.readWindow({ ...chunk, rows: CHUNK_ROWS, columns: CHUNK_COLUMNS }),
      next: 0,
    }));
    for (let row = band; row < band + CHUNK_ROWS; row++) {
      const cells: WorkbookCell[] = [];
      for (const column of columns) {
        while (column.cells[column.next]?.row === row) {
          const { column: coordinate, html } = column.cells[column.next++]!;
          cells.push({ sheetId, row, column: coordinate, html, ...decodeWorkbookCell(html) });
        }
      }
      if (cells.length) yield cells;
    }
    start = end;
  }
}
