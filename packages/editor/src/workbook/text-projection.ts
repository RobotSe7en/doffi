import { isDocumentTextBoundary } from "@doffi/core";
import type { WorkbookCellTextEdit } from "@doffi/core";

/** Native textareas expose LF line endings. This is a transient projection,
 * never a request to normalize the document's untouched CR/CRLF source. */
export function workbookInputTextView(raw: string): string { return raw.replace(/\r\n?/g, "\n"); }
export function workbookTextSplice(before: string, after: string): WorkbookCellTextEdit | undefined {
  if (before === after) return;
  let from = 0;
  while (from < before.length && from < after.length && before[from] === after[from]) from++;
  let oldTo = before.length, newTo = after.length;
  while (oldTo > from && newTo > from && before[oldTo - 1] === after[newTo - 1]) { oldTo--; newTo--; }
  while (from > 0 && (!isDocumentTextBoundary(before, from) || !isDocumentTextBoundary(after, from))) from--;
  while (oldTo < before.length && newTo < after.length && (!isDocumentTextBoundary(before, oldTo) || !isDocumentTextBoundary(after, newTo))) { oldTo++; newTo++; }
  return { from, to: oldTo, expectedText: before.slice(from, oldTo), text: after.slice(from, newTo) };
}
export function reconcileWorkbookInputText(raw: string, field: string): string {
  if (!raw.includes("\r")) return field;
  const offsets = [0]; let before = "";
  for (let index = 0; index < raw.length; index++) {
    if (raw[index] === "\r") { if (raw[index + 1] === "\n") index++; before += "\n"; }
    else before += raw[index]!;
    offsets.push(index + 1);
  }
  const edit = workbookTextSplice(before, field); if (!edit) return raw;
  return raw.slice(0, offsets[edit.from]!) + edit.text + raw.slice(offsets[edit.to]!);
}
