import { captureWorkbookModel, DocumentError, isDocumentTextBoundary, WorkbookCoordinateSchema, WorkbookOperationSchema } from "@doffi/core";
import type { DocumentTransaction, WorkbookCoordinate, WorkbookModel } from "@doffi/core";
import { readWorkbookInputText } from "./input-source.js";
import { workbookInputTextSourceRange, workbookInputTextView } from "./text-projection.js";

export interface WorkbookTextFontSelection { readonly from: number; readonly to: number; readonly text: string }
/** Native LF offsets and a literal escape are view spelling. Resolve them
 * against this authorized source before entering the raw UTF-16 transaction. */
export function createWorkbookTextFormatTransaction(input: WorkbookModel, sheetId: string, at: WorkbookCoordinate,
  selection: WorkbookTextFontSelection, properties: Readonly<Record<string, string>>): DocumentTransaction {
  const coordinate = WorkbookCoordinateSchema.safeParse(at);
  if (!coordinate.success || !selection || typeof selection !== "object" || Array.isArray(selection) || Object.keys(selection).some(key => !["from", "to", "text"].includes(key)) || typeof selection.text !== "string")
    throw new DocumentError("InvalidOperation", "Invalid workbook text font selection.");
  const model = captureWorkbookModel(input), snapshot = model.snapshot(), cell = model.getCell(sheetId, coordinate.data.row, coordinate.data.column);
  if (!cell || cell.formula || cell.value.type !== "text") throw new DocumentError("InvalidOperation", "只有文本常量支持选中文字的字体设置。");
  const raw = readWorkbookInputText(cell);
  if (workbookInputTextView(raw) !== selection.text) throw new DocumentError("RevisionConflict", "选中文字的来源已变化，请重新选择。");
  const mapped = workbookInputTextSourceRange(raw, selection.from, selection.to), prefix = raw === "'" + cell.text ? 1 : 0;
  const from = Math.max(0, mapped.from - prefix), to = Math.max(0, mapped.to - prefix);
  if (!isDocumentTextBoundary(cell.text, from) || !isDocumentTextBoundary(cell.text, to)) throw new DocumentError("InvalidOperation", "字体选区不能拆分代理对或变体序列。");
  const parsed = WorkbookOperationSchema.safeParse({ op: "formatWorkbookText", sheetId, at: coordinate.data, from, to, expectedText: cell.text.slice(from, to), properties });
  if (!parsed.success || from >= to) throw new DocumentError("InvalidOperation", "请选择来源文字并提供有效字体设置。");
  return { protocolVersion: 1, documentId: snapshot.manifest.documentId, baseRevision: snapshot.manifest.revision, operations: [parsed.data] };
}
