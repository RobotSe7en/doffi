import { captureWorkbookModel, DocumentError, replaceWorkbookCellText, WorkbookCoordinateSchema, WorkbookValueSchema } from "@doffi/core";
import type { DocumentTransaction, WorkbookCellInput, WorkbookCellSource, WorkbookModel } from "@doffi/core";
import { invalidWorkbookFind, scanWorkbookText, workbookFindUnicode } from "./find.js";
import type { WorkbookFindMatch, WorkbookFindQuery, WorkbookReplaceOptions } from "./find-contracts.js";
import { parseWorkbookTextInput } from "./text-input.js";

/** Replacement is literal, including $& and backslashes. Rich text constants
 * retain untouched spans. The caller prepares/commits one complete batch. */
export async function createWorkbookReplaceTransaction(input: WorkbookModel, query: WorkbookFindQuery, replacement: string, options: WorkbookReplaceOptions = {}): Promise<DocumentTransaction> {
  workbookFindUnicode(replacement);
  if (replacement.length > 32767) throw new DocumentError("ResourceLimitExceeded", "Workbook replacement exceeds its text budget.");
  if (!options || typeof options !== "object" || Array.isArray(options)) invalidWorkbookFind("Invalid workbook replacement options.");
  const { valueMode = "preserve", only: suppliedOnly, ...scanOptions } = options;
  let only: WorkbookReplaceOptions["only"];
  if (!["preserve", "auto", "text"].includes(valueMode) || query?.field === "display") invalidWorkbookFind("Workbook replacement requires editable source or formula text and a valid type mode.");
  const model = captureWorkbookModel(input), snapshot = model.snapshot(), updates = new Map<string, WorkbookCellInput[]>(); let selected = false;
  if (suppliedOnly !== undefined) {
    if (!suppliedOnly || typeof suppliedOnly !== "object" || Array.isArray(suppliedOnly) || Object.keys(suppliedOnly).some(key => !["documentId", "baseRevision", "match"].includes(key))) invalidWorkbookFind("Invalid single replacement proof.");
    if (suppliedOnly.documentId !== snapshot.manifest.documentId || suppliedOnly.baseRevision !== snapshot.manifest.revision) throw new DocumentError("RevisionConflict", "Workbook replacement match belongs to another source revision.");
    const match = suppliedOnly.match;
    if (!match || typeof match !== "object" || Array.isArray(match) || Object.keys(match).some(key => !["sheetId", "row", "column", "from", "to", "text"].includes(key))
      || !WorkbookCoordinateSchema.safeParse({ row: match.row, column: match.column }).success || typeof match.sheetId !== "string"
      || ![match.from, match.to].every(Number.isSafeInteger) || match.from < 0 || match.to <= match.from || typeof match.text !== "string" || !match.text) invalidWorkbookFind("Invalid single replacement source proof.");
    only = { documentId: suppliedOnly.documentId, baseRevision: suppliedOnly.baseRevision, match: { ...match } };
  }
  const same = (a: WorkbookFindMatch, b: WorkbookFindMatch): boolean => a.sheetId === b.sheetId && a.row === b.row && a.column === b.column && a.from === b.from && a.to === b.to && a.text === b.text;
  await scanWorkbookText(input, query, scanOptions, ({ cell, text, matches }, account) => {
    const chosen = only ? matches.filter(match => same(match, only.match)) : matches;
    if (!chosen.length) return; selected = true;
    const edits = chosen.map(match => ({ from: match.from, to: match.to, expectedText: match.text, text: replacement }));
    const pieces: string[] = []; let cursor = 0;
    for (const edit of edits) { pieces.push(text.slice(cursor, edit.from), replacement); cursor = edit.to; } pieces.push(text.slice(cursor));
    const length = text.length + edits.reduce((sum, edit) => sum + replacement.length - (edit.to - edit.from), 0);
    if (length > (cell.formula ? 8192 : 32767)) throw new DocumentError("ResourceLimitExceeded", "A replaced workbook cell exceeds its source text budget.");
    const result = pieces.join(""); if (result === text) return;
    const coordinate = { row: cell.row, column: cell.column };
    let source: WorkbookCellSource;
    if (cell.formula) {
      source = parseWorkbookTextInput(result, coordinate);
      if (source.kind !== "formula") invalidWorkbookFind("A formula replacement must retain a valid formula expression.");
    } else if (cell.value.type === "text" && valueMode === "preserve") source = replaceWorkbookCellText(model, { sheetId: cell.sheetId, row: cell.row, column: cell.column }, edits);
    else if (valueMode === "text") source = parseWorkbookTextInput(result, coordinate, "text");
    else {
      if (valueMode === "preserve" && cell.value.type === "error") {
        const error = WorkbookValueSchema.safeParse({ type: "error", value: result });
        if (!error.success) invalidWorkbookFind("Replacement would change an error constant's type; select an explicit conversion mode.");
        source = { kind: "value", value: error.data };
      } else {
        source = parseWorkbookTextInput(result, coordinate, valueMode === "preserve" && cell.value.type === "number" ? "number" : "auto");
        if (valueMode === "preserve" && (source.kind !== "value" || source.value.type !== cell.value.type)) invalidWorkbookFind("Replacement would change a constant's type; select an explicit conversion mode.");
      }
    }
    const update = { row: cell.row, column: cell.column, source }; account(JSON.stringify(update).length);
    const cells = updates.get(cell.sheetId) ?? []; cells.push(update); updates.set(cell.sheetId, cells);
  }, only?.match);
  if (only && !selected) invalidWorkbookFind("The replacement source proof no longer matches the query.");
  if (captureWorkbookModel(input) !== model) throw new DocumentError("RevisionConflict", "Workbook replacement source changed before completion.");
  return { protocolVersion: 1, documentId: snapshot.manifest.documentId, baseRevision: snapshot.manifest.revision,
    operations: [...updates].map(([sheetId, cells]) => ({ op: "setWorkbookCells", sheetId, cells })) };
}
