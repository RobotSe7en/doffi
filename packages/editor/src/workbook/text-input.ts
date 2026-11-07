import { DocumentError, parseWorkbookFormula, WorkbookCoordinateSchema } from "@doffi/core";
import type { DocumentTransaction, WorkbookCellSource, WorkbookCoordinate, WorkbookSnapshot } from "@doffi/core";

export type WorkbookTextInputMode = "auto" | "text" | "number";
export interface WorkbookPasteOptions { readonly mode?: WorkbookTextInputMode; readonly maxCells?: number; readonly maxCodeUnits?: number }
function invalid(message: string): never { throw new DocumentError("InvalidOperation", message); }
function unicode(text: string): void {
  for (const character of text) {
    const value = character.codePointAt(0)!;
    if (value === 0 || value >= 0xd800 && value <= 0xdfff) invalid("Workbook input requires well-formed Unicode without null characters.");
  }
}
const numeric = /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/;
/** Conservative user input; formatting and ambiguous dates never change the
 * source type. The caller still prepares and commits the shared transaction. */
export function parseWorkbookTextInput(text: string, origin: WorkbookCoordinate, mode: WorkbookTextInputMode = "auto"): WorkbookCellSource {
  if (typeof text !== "string" || !["auto", "text", "number"].includes(mode) || !WorkbookCoordinateSchema.safeParse(origin).success) invalid("Invalid workbook text input or coordinate.");
  if (text.length > 32767 && !(mode === "auto" && text.startsWith("'") && text.length === 32768)) throw new DocumentError("ResourceLimitExceeded", "Workbook cell input exceeds its text budget.");
  unicode(text);
  const value = (item: Extract<WorkbookCellSource, { kind: "value" }>["value"]): WorkbookCellSource => Object.freeze({ kind: "value", value: Object.freeze(item) });
  if (mode === "text") return value({ type: "text", value: text });
  if (mode === "number") {
    const source = text.trim(); if (!numeric.test(source) || !Number.isFinite(Number(source))) invalid("Explicit numeric input requires a finite numeric literal.");
    return value({ type: "number", value: Number(source) });
  }
  if (text.startsWith("'")) return value({ type: "text", value: text.slice(1) });
  if (text.startsWith("=")) {
    try { parseWorkbookFormula(text, origin); }
    catch (error) { if (error instanceof DocumentError && error.code === "InvalidDocument") throw new DocumentError("InvalidOperation", error.message, error.details); throw error; }
    return Object.freeze({ kind: "formula", expression: text });
  }
  if (!text) return value({ type: "blank" });
  if (/^(?:TRUE|FALSE)$/i.test(text)) return value({ type: "boolean", value: text.toUpperCase() === "TRUE" });
  if (numeric.test(text)) {
    const mantissa = text.replace(/^[+-]/, "").split(/[eE]/)[0]!, whole = mantissa.split(".")[0]!;
    const digits = mantissa.replace(".", "").replace(/^0+/, "");
    // Leading zeros and more than fifteen significant digits are often IDs.
    if (!(whole.length > 1 && whole.startsWith("0")) && digits.length <= 15 && Number.isFinite(Number(text))) return value({ type: "number", value: Number(text) });
  }
  return value({ type: "text", value: text });
}

/** Excel-style quoted TSV, bounded before materialization. A final line
 * separator terminates its row; an additional empty row remains meaningful. */
export function parseWorkbookTsv(text: string, options: Pick<WorkbookPasteOptions, "maxCells" | "maxCodeUnits"> = {}): readonly (readonly string[])[] {
  const limits = { maxCells: 25_000, maxCodeUnits: 2 * 1024 * 1024, ...options };
  if (typeof text !== "string" || !options || typeof options !== "object" || Array.isArray(options) || Object.keys(options).some(key => !["maxCells", "maxCodeUnits"].includes(key))
    || !Object.values(limits).every(value => Number.isSafeInteger(value) && value > 0)) invalid("Invalid workbook clipboard options.");
  if (text.length > limits.maxCodeUnits) throw new DocumentError("ResourceLimitExceeded", "Clipboard text exceeds its input budget.");
  unicode(text);
  const rows: string[][] = []; let row: string[] = [], field = "", state: "bare" | "quoted" | "closed" = "bare", total = 0, terminated = false;
  const cell = (): void => { if (++total > limits.maxCells) throw new DocumentError("ResourceLimitExceeded", "Clipboard cells exceed their input budget."); row.push(field); field = ""; state = "bare"; };
  for (let index = 0; index < text.length; index++) {
    const character = text[index]!; terminated = false;
    if (state === "quoted") {
      if (character === '"') { if (text[index + 1] === '"') { field += '"'; index++; } else state = "closed"; }
      else if (character === "\r" && text[index + 1] === "\n") { field += "\n"; index++; }
      else field += character;
    } else if (character === "\t" || character === "\n" || character === "\r") {
      cell(); if (character !== "\t") { if (character === "\r" && text[index + 1] === "\n") index++; rows.push(row); row = []; terminated = true; }
    } else {
      if (state === "closed") invalid("Quoted clipboard fields must end at a tab or line separator.");
      if (character === '"' && !field) state = "quoted"; else field += character;
    }
    if (field.length > 32767) throw new DocumentError("ResourceLimitExceeded", "A clipboard cell exceeds its text budget.");
  }
  if (state === "quoted") invalid("Clipboard text contains an unclosed quoted field.");
  if (!terminated) { cell(); rows.push(row); }
  const columns = rows[0]!.length;
  if (rows.some(row => row.length !== columns)) invalid("Clipboard cells must form a complete rectangle.");
  return Object.freeze(rows.map(row => Object.freeze(row)));
}

export function createWorkbookPasteTransaction(snapshot: WorkbookSnapshot, sheetId: string, origin: WorkbookCoordinate, text: string, options: WorkbookPasteOptions = {}): DocumentTransaction {
  if (!options || typeof options !== "object" || Array.isArray(options) || Object.keys(options).some(key => !["mode", "maxCells", "maxCodeUnits"].includes(key))) invalid("Invalid workbook paste options.");
  if (!WorkbookCoordinateSchema.safeParse(origin).success || !snapshot.metadata.workbook.sheets.some(sheet => sheet.sheetId === sheetId)) invalid("Paste requires a coordinate on an existing worksheet.");
  const rows = parseWorkbookTsv(text, { ...(options.maxCells !== undefined ? { maxCells: options.maxCells } : {}), ...(options.maxCodeUnits !== undefined ? { maxCodeUnits: options.maxCodeUnits } : {}) });
  if (origin.row + rows.length > 1_048_576 || origin.column + rows[0]!.length > 16_384) invalid("Clipboard cells exceed their destination worksheet.");
  const cells = rows.flatMap((values, row) => values.map((value, column) => {
    const position = { row: origin.row + row, column: origin.column + column };
    return { ...position, source: parseWorkbookTextInput(value, position, options.mode) };
  }));
  return { protocolVersion: 1, documentId: snapshot.manifest.documentId, baseRevision: snapshot.manifest.revision, operations: [{ op: "setWorkbookCells", sheetId, cells }] };
}
