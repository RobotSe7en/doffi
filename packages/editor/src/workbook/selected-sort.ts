import { DocumentError } from "@doffi/core";
import type { DocumentSelection, WorkbookSnapshot, WorkbookSort } from "@doffi/core";

export function selectedWorkbookSort(snapshot: WorkbookSnapshot, selection: DocumentSelection | undefined,
  settings: { columns: string; header: boolean; caseSensitive: boolean; method: string }, descending: boolean): { sheetId: string; sort: WorkbookSort } {
  const anchor = selection?.anchor;
  if (!selection || anchor?.kind !== "cells" || anchor.ranges.length !== 1) throw new DocumentError("InvalidOperation", "请选择一个排序区域。");
  if (selection.documentId !== snapshot.manifest.documentId || selection.baseRevision !== snapshot.manifest.revision) throw new DocumentError("RevisionConflict", "Sort selection belongs to a previous source.");
  const sheet = snapshot.metadata.workbook.sheets.find(sheet => sheet.sheetId === anchor.sheetId);
  if (!sheet) throw new DocumentError("InvalidOperation", "排序工作表不存在。");
  const selected = anchor.ranges[0]!, range = { from: { ...selected.from }, to: { ...selected.to } };
  const table = sheet.tables?.find(table => range.from.row >= table.range.from.row && range.to.row <= table.range.to.row && range.from.column >= table.range.from.column && range.to.column <= table.range.to.column);
  if (table) { range.from = { ...table.range.from, row: table.range.from.row + Number(table.header) }; range.to = { ...table.range.to, row: table.range.to.row - Number(table.totals) }; }
  else if (settings.header || sheet.autoFilter && range.from.row === sheet.autoFilter.range.from.row && range.from.column <= sheet.autoFilter.range.to.column && range.to.column >= sheet.autoFilter.range.from.column) range.from.row++;
  if (range.from.row > range.to.row) throw new DocumentError("InvalidOperation", "排序区域没有数据行。");
  const names = settings.columns.trim() ? settings.columns.toUpperCase().split(/[,，]/).map(value => value.trim()) : [];
  const columns = names.length ? names.map(name => {
    if (!/^[A-Z]{1,3}$/.test(name)) throw new DocumentError("InvalidOperation", "排序列请输入列字母，例如 A,C。");
    let column = 0; for (const letter of name) column = column * 26 + letter.charCodeAt(0) - 64;
    return column - 1;
  }) : [selected.from.column];
  if (new Set(columns).size !== columns.length || columns.some(column => column < range.from.column || column > range.to.column)) throw new DocumentError("InvalidOperation", "排序列必须在区域内且不能重复。");
  return { sheetId: sheet.sheetId, sort: { range, keys: columns.map(column => ({ column, descending })), caseSensitive: settings.caseSensitive, method: settings.method as WorkbookSort["method"] } };
}
