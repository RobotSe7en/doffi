import { WorkbookCoordinateSchema } from "../contracts/workbook.js";
import type { WorkbookCellSource } from "../contracts/workbook-operations.js";
import type { WorkbookModel } from "../runtime.js";
import { DocumentError } from "../errors.js";
import { serializeHtmlTree } from "../document/html-parser.js";
import { parseWorkbookCellElement } from "./chunk-index.js";
import type { WorkbookReferenceLocation } from "./reference.js";
import { captureWorkbookModel } from "./source.js";

/** Read an authorized immutable cell as typed transaction input. Source spans
 * keep their font/style references; formula caches never become constants.
 * Preparing the eventual transaction still validates its complete candidate. */
export function readWorkbookCellSource(model: WorkbookModel, location: WorkbookReferenceLocation): WorkbookCellSource {
  model = captureWorkbookModel(model);
  if (!location || typeof location !== "object" || Object.keys(location).some(key => !["sheetId", "row", "column"].includes(key))
    || !WorkbookCoordinateSchema.safeParse({ row: location.row, column: location.column }).success
    || !model.snapshot().metadata.workbook.sheets.some(sheet => sheet.sheetId === location.sheetId)) throw new DocumentError("InvalidOperation", "Reading workbook input requires an existing worksheet coordinate.");
  const cell = model.getCell(location.sheetId, location.row, location.column);
  if (cell && (cell.sheetId !== location.sheetId || cell.row !== location.row || cell.column !== location.column)) throw new DocumentError("InvalidDocument", "The workbook provider returned another cell's source.");
  if (cell?.formula) return Object.freeze({ kind: "formula", expression: cell.formula.expression });
  if (cell?.value.type === "text" && /<span(?:[\t\n\f\r />])/i.test(cell.html)) {
    const html = serializeHtmlTree(parseWorkbookCellElement(cell.html));
    if (html.length > 2 * 1024 * 1024) throw new DocumentError("ResourceLimitExceeded", "Rich workbook input exceeds its source budget.");
    return Object.freeze({ kind: "richText", html });
  }
  return Object.freeze({ kind: "value", value: Object.freeze(cell ? { ...cell.value } : { type: "blank" }) });
}
