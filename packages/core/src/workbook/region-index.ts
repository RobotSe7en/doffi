import type { DeepReadonly } from "../runtime.js";
import type { WorkbookRange } from "../contracts/workbook.js";
type Range = DeepReadonly<WorkbookRange>;
interface Branch { readonly row: number; readonly crossing: readonly Range[]; readonly left?: Branch; readonly right?: Branch }

/** The core already validates nonoverlapping worksheet regions. This derived interval tree
 * indexes their row spans without expanding any covered source cells. */
export class WorkbookRegionIndex {
  private readonly root: Branch | undefined;
  constructor(ranges: readonly Range[]) {
    const sorted = [...ranges].sort((a, b) => a.from.row - b.from.row || a.from.column - b.from.column);
    function build(values: readonly Range[]): Branch | undefined {
      if (!values.length) return undefined;
      const row = values[values.length >>> 1]!.from.row, before: Range[] = [], after: Range[] = [], crossing: Range[] = [];
      for (const range of values) {
        if (range.to.row < row) before.push(range);
        else if (range.from.row > row) after.push(range);
        else crossing.push(range);
      }
      // Every crossing range contains this row, so validated nonoverlap makes
      // its column interval disjoint. A point needs one binary column lookup,
      // even if thousands of long row intervals have shuffled column order.
      crossing.sort((a, b) => a.from.column - b.from.column);
      const left = build(before), right = build(after);
      return { row, crossing, ...(left ? { left } : {}), ...(right ? { right } : {}) };
    }
    this.root = build(sorted);
  }
  at(row: number, column: number): Range | undefined {
    let branch = this.root;
    while (branch) {
      let low = 0, high = branch.crossing.length;
      while (low < high) { const middle = (low + high) >>> 1; if (branch.crossing[middle]!.to.column < column) low = middle + 1; else high = middle; }
      const range = branch.crossing[low];
      if (range && range.from.column <= column && range.from.row <= row && range.to.row >= row) return range;
      branch = row < branch.row ? branch.left : row > branch.row ? branch.right : undefined;
    }
    return undefined;
  }
}
