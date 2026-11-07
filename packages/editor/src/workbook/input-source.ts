import { captureWorkbookModel, DocumentError, replaceWorkbookCellText, WorkbookCoordinateSchema } from "@doffi/core";
import type { DocumentTransaction, WorkbookCell, WorkbookCoordinate, WorkbookModel } from "@doffi/core";
import { parseWorkbookTextInput } from "./text-input.js";
import type { WorkbookTextInputMode } from "./text-input.js";
import { workbookTextSplice } from "./text-projection.js";

/** Reversible input spelling, separate from display formatting. Literal text
 * retains its type even when it resembles a number, formula or empty cell. */
export function readWorkbookInputText(cell: WorkbookCell | undefined): string {
  let text = cell?.formula?.expression ?? (!cell || cell.value.type === "blank" ? "" : cell.value.type === "boolean" ? cell.value.value ? "TRUE" : "FALSE" : String(cell.value.value));
  if (!cell?.formula && cell?.value.type === "text") {
    let literal = false;
    try { const parsed = parseWorkbookTextInput(text, { row: cell.row, column: cell.column }); literal = parsed.kind === "value" && parsed.value.type === "text" && parsed.value.value === text; }
    catch { /* Formula-like text has an explicit literal spelling. */ }
    if (!literal) text = "'" + text;
  }
  return text;
}

/** Both direct and Worker input generate the same typed transaction against
 * an immutable source. The caller still owns preparation and application. */
export function createWorkbookInputTransaction(input: WorkbookModel, sheetId: string, origin: WorkbookCoordinate, text: string, mode: WorkbookTextInputMode = "auto"): DocumentTransaction {
  const coordinate = WorkbookCoordinateSchema.safeParse(origin);
  if (input?.kind !== "workbook" || !coordinate.success || typeof text !== "string" || !["auto", "text", "number"].includes(mode)) throw new DocumentError("InvalidOperation", "Workbook input requires a worksheet coordinate, text and supported mode.");
  const model = captureWorkbookModel(input), snapshot = model.snapshot(), at = coordinate.data;
  if (!snapshot.metadata.workbook.sheets.some(sheet => sheet.sheetId === sheetId)) throw new DocumentError("InvalidOperation", "Workbook input requires an existing worksheet.");
  const cell = model.getCell(sheetId, at.row, at.column);
  let source = mode === "auto" && text === readWorkbookInputText(cell) ? undefined : parseWorkbookTextInput(text, at, mode);
  if (source?.kind === "value" && source.value.type === "text" && cell?.value.type === "text" && !cell.formula) {
    const edit = workbookTextSplice(cell.text, source.value.value);
    source = edit ? replaceWorkbookCellText(model, { sheetId, ...at }, [edit]) : undefined;
  }
  return { protocolVersion: 1, documentId: snapshot.manifest.documentId, baseRevision: snapshot.manifest.revision,
    operations: source ? [{ op: "setWorkbookCells", sheetId, cells: [{ ...at, source }] }] : [] };
}
