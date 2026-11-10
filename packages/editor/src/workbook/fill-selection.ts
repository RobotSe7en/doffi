import { DocumentError } from "@doffi/core";
import type { DeepReadonly, WorkbookRange } from "@doffi/core";

export function workbookSeriesAxis(range: DeepReadonly<WorkbookRange>, merges: readonly DeepReadonly<WorkbookRange>[]): "row" | "column" {
  const first = merges.find(merge => merge.from.row === range.from.row && merge.from.column === range.from.column);
  return first && first.to.column === range.to.column || range.from.column === range.to.column ? "row" : "column";
}

/** Numeric series use one logical cell per row/column, which may be a complete
 * merged rectangle. Other merged shapes cannot define an arithmetic seed. */
export function workbookSeriesUnit(range: DeepReadonly<WorkbookRange>, merges: readonly DeepReadonly<WorkbookRange>[], axis: "row" | "column"): number {
  const other = axis === "row" ? "column" : "row", first = merges.find(merge => merge.from.row <= range.from.row && merge.to.row >= range.from.row
    && merge.from.column <= range.from.column && merge.to.column >= range.from.column);
  if (!first) {
    if (range.from[other] !== range.to[other]) throw new DocumentError("InvalidOperation", "数列填充请选择一列或一行逻辑单元格。");
    return 1;
  }
  if (first.from.row !== range.from.row || first.from.column !== range.from.column || first.to[other] !== range.to[other] || first.to[axis] > range.to[axis])
    throw new DocumentError("InvalidOperation", "数列填充的种子必须是完整且对齐的逻辑单元格。");
  return first.to[axis] - first.from[axis] + 1;
}

/** A first row/column seed grows to include every merged region crossing that
 * band. Sorting intervals lets connected bands close in one bounded pass. */
export function workbookFillSeedCount(range: DeepReadonly<WorkbookRange>, merges: readonly DeepReadonly<WorkbookRange>[], axis: "row" | "column"): number {
  const other = axis === "row" ? "column" : "row";
  const intersected = merges.filter(merge => merge.to[axis] >= range.from[axis] && merge.from[axis] <= range.to[axis]
    && merge.to[other] >= range.from[other] && merge.from[other] <= range.to[other]).sort((a, b) => a.from[axis] - b.from[axis]);
  let end = range.from[axis];
  for (const merge of intersected) {
    if (merge.from[axis] > end) break;
    if (merge.from[axis] < range.from[axis] || merge.to[axis] > range.to[axis] || merge.from[other] < range.from[other] || merge.to[other] > range.to[other])
      throw new DocumentError("InvalidOperation", "填充选区不能截断合并单元格。");
    end = Math.max(end, merge.to[axis]);
  }
  return end - range.from[axis] + 1;
}
