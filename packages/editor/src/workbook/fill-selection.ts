import { DocumentError } from "@doffi/core";
import type { DeepReadonly, WorkbookRange } from "@doffi/core";

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
