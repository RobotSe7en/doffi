import type { DeepReadonly, WorkbookCoordinate, WorkbookRange, WorkbookSheet } from "@doffi/core";
import type { AxisDimensions } from "./dimensions.js";
import { WorkbookMergeIndex } from "./projection.js";

export interface WorkbookSelectionNavigation {
  readonly origin: DeepReadonly<WorkbookCoordinate>;
  readonly focus: DeepReadonly<WorkbookCoordinate>;
}
/** Reuse displayed sparse dimensions: hidden/zero-size runs are skipped as
 * intervals, and merged cells contribute their complete source geometry. */
export function navigateWorkbookSelection(sheet: DeepReadonly<WorkbookSheet>, dimensions: { readonly rows: AxisDimensions; readonly columns: AxisDimensions },
  state: WorkbookSelectionNavigation, direction: "up" | "down" | "left" | "right", extend: boolean): { readonly range: DeepReadonly<WorkbookRange>; readonly navigation: WorkbookSelectionNavigation } {
  const merges = new WorkbookMergeIndex(sheet.merges), current = merges.at(state.focus.row, state.focus.column) ?? { from: state.focus, to: state.focus };
  const vertical = direction === "up" || direction === "down", axis = vertical ? "row" : "column";
  const step = vertical ? direction === "up" ? -1 : 1 : (direction === "left") !== sheet.rightToLeft ? -1 : 1;
  const start = step === -1 ? current.from[axis] : current.to[axis], dimension = vertical ? dimensions.rows : dimensions.columns;
  const next = dimension.neighbor(start, step), focus = { ...current.from, [axis]: next ?? state.focus[axis] };
  const target = merges.at(focus.row, focus.column) ?? { from: focus, to: focus }, origin = extend ? state.origin : target.from;
  const anchor = merges.at(origin.row, origin.column) ?? { from: origin, to: origin };
  const range = extend ? { from: { row: Math.min(anchor.from.row, target.from.row), column: Math.min(anchor.from.column, target.from.column) },
    to: { row: Math.max(anchor.to.row, target.to.row), column: Math.max(anchor.to.column, target.to.column) } } : target;
  return { range, navigation: { origin, focus: target.from } };
}
