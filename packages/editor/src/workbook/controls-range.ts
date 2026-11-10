import { DocumentError, parseWorkbookFormula } from "@doffi/core";
import type { DeepReadonly, WorkbookRange } from "@doffi/core";

export function workbookControlRange(text: string): WorkbookRange {
  const parsed = parseWorkbookFormula("=" + text.trim(), { row: 0, column: 0 });
  const from = parsed.kind === "cell" ? parsed : parsed.kind === "range" ? parsed.from : undefined;
  const to = parsed.kind === "cell" ? parsed : parsed.kind === "range" ? parsed.to : undefined;
  if (!from || !to || from.sheet !== undefined || to.sheet !== undefined || from.row > to.row || from.column > to.column)
    throw new DocumentError("InvalidOperation", "请输入当前工作表内的单元格区域，例如 A1:C10。");
  return { from: { row: from.row, column: from.column }, to: { row: to.row, column: to.column } };
}
export function workbookControlAddress(point: DeepReadonly<WorkbookRange>["from"]): string {
  let letters = ""; for (let value = point.column + 1; value; value = Math.floor((value - 1) / 26)) letters = String.fromCharCode(65 + (value - 1) % 26) + letters;
  return letters + (point.row + 1);
}
export function workbookControlRangeText(range: DeepReadonly<WorkbookRange>): string { return workbookControlAddress(range.from) + ":" + workbookControlAddress(range.to); }
