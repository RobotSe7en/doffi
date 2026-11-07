import type { DeepReadonly, WorkbookRangeChange, WorkbookStructureMapping } from "../runtime.js";
import type { WorkbookRange } from "../contracts/workbook.js";
import { checkLimit, freeze, invalid } from "../document/validation.js";
import { MAX_SHEET_COLUMNS, MAX_SHEET_ROWS } from "./chunk-index.js";
import { checkRange } from "./ranges.js";
type Range = DeepReadonly<WorkbookRange>;
export function intersectWorkbookRange(a: Range, b: Range): WorkbookRange | undefined {
  const from = { row: Math.max(a.from.row, b.from.row), column: Math.max(a.from.column, b.from.column) };
  const to = { row: Math.min(a.to.row, b.to.row), column: Math.min(a.to.column, b.to.column) };
  return from.row <= to.row && from.column <= to.column ? { from, to } : undefined;
}
/** At most four exact rectangles, with no enumeration of their cells. */
export function subtractWorkbookRange(a: Range, b: Range): WorkbookRange[] {
  const intersection = intersectWorkbookRange(a, b); if (!intersection) return [{ from: { ...a.from }, to: { ...a.to } }];
  const result: WorkbookRange[] = [];
  if (a.from.row < intersection.from.row) result.push({ from: { ...a.from }, to: { row: intersection.from.row - 1, column: a.to.column } });
  if (intersection.to.row < a.to.row) result.push({ from: { row: intersection.to.row + 1, column: a.from.column }, to: { ...a.to } });
  if (a.from.column < intersection.from.column) result.push({ from: { row: intersection.from.row, column: a.from.column }, to: { row: intersection.to.row, column: intersection.from.column - 1 } });
  if (intersection.to.column < a.to.column) result.push({ from: { row: intersection.from.row, column: intersection.to.column + 1 }, to: { row: intersection.to.row, column: a.to.column } });
  return result;
}
function shifted(range: Range, row: number, column: number): WorkbookRange {
  return { from: { row: range.from.row + row, column: range.from.column + column }, to: { row: range.to.row + row, column: range.to.column + column } };
}
/** Maps identities of existing cells. Inserted cells never enlarge an anchor;
 * fragmented selections stay fragmented instead of becoming bounding boxes. */
export function mapWorkbookRanges(input: readonly WorkbookRangeChange[], steps: readonly WorkbookStructureMapping[], maxFragments = 4096): readonly WorkbookRangeChange[] {
  if (!Number.isSafeInteger(maxFragments) || maxFragments < 1) invalid("Invalid workbook mapping fragment budget.");
  checkLimit(input.length, maxFragments);
  let result = input.map(item => ({ sheetId: item.sheetId, range: { from: { ...item.range.from }, to: { ...item.range.to } } }));
  for (const item of result) checkRange(item.range);
  for (const step of steps) {
    const next: WorkbookRangeChange[] = [];
    const add = (item: WorkbookRangeChange): void => { checkLimit(next.length + 1, maxFragments); next.push(item); };
    for (const item of result) {
      const { sheetId, range } = item;
      if (step.kind === "restoreSheet") { add(item); continue; }
      if (step.kind === "deleteSheet") { if (sheetId !== step.sheetId) add(item); continue; }
      if (step.kind === "axis") {
        if (sheetId !== step.sheetId) { add(item); continue; }
        const axis = step.axis, bound = axis === "row" ? MAX_SHEET_ROWS : MAX_SHEET_COLUMNS;
        if (!Number.isSafeInteger(step.index) || !Number.isSafeInteger(step.count) || step.index < 0 || step.count < 1 || step.index + step.count > bound) invalid("Invalid workbook axis mapping.");
        if (range.from[axis] < step.index) add({ sheetId, range: { from: range.from, to: { ...range.to, [axis]: Math.min(range.to[axis], step.index - 1) } } });
        const start = Math.max(range.from[axis], step.index + (step.action === "delete" ? step.count : 0));
        if (start <= range.to[axis]) {
          const delta = step.action === "insert" ? step.count : -step.count, stop = Math.min(range.to[axis], bound - 1 - delta);
          if (start <= stop) add({ sheetId, range: { from: { ...range.from, [axis]: start + delta }, to: { ...range.to, [axis]: stop + delta } } });
        }
      } else if (step.kind === "move") {
        checkRange(step.source.range); checkRange(step.target.range);
        const source = step.source, target = step.target;
        if (source.range.to.row - source.range.from.row !== target.range.to.row - target.range.from.row || source.range.to.column - source.range.from.column !== target.range.to.column - target.range.from.column) invalid("A move must retain the exact range dimensions.");
        const moved = sheetId === source.sheetId ? intersectWorkbookRange(range, source.range) : undefined;
        const kept = sheetId === source.sheetId ? subtractWorkbookRange(range, source.range) : [range];
        for (const part of kept) for (const remaining of sheetId === target.sheetId ? subtractWorkbookRange(part, target.range) : [part]) add({ sheetId, range: remaining });
        if (moved) add({ sheetId: target.sheetId, range: shifted(moved, target.range.from.row - source.range.from.row, target.range.from.column - source.range.from.column) });
      } else {
        if (sheetId !== step.sheetId) { add(item); continue; }
        const count = step.range.to.row - step.range.from.row + 1;
        if (step.rows.length !== count || new Set(step.rows).size !== count || step.rows.some(row => !Number.isSafeInteger(row) || row < step.range.from.row || row > step.range.to.row)) invalid("A sort mapping requires a complete row permutation.");
        for (const part of subtractWorkbookRange(range, step.range)) add({ sheetId, range: part });
        const intersection = intersectWorkbookRange(range, step.range);
        if (intersection) {
          const mapped = step.rows.slice(intersection.from.row - step.range.from.row, intersection.to.row - step.range.from.row + 1).sort((a, b) => a - b);
          for (let index = 0; index < mapped.length;) {
            const first = mapped[index]!; let last = first; index++;
            while (index < mapped.length && mapped[index] === last + 1) last = mapped[index++]!;
            add({ sheetId, range: { from: { row: first, column: intersection.from.column }, to: { row: last, column: intersection.to.column } } });
          }
        }
      }
    }
    result = next.map(item => ({ sheetId: item.sheetId, range: { from: { ...item.range.from }, to: { ...item.range.to } } }));
  }
  return freeze(result);
}
export function invertWorkbookMapping(steps: readonly WorkbookStructureMapping[]): readonly WorkbookStructureMapping[] {
  return freeze([...steps].reverse().map((step): WorkbookStructureMapping => {
    if (step.kind === "axis") return { ...step, action: step.action === "insert" ? "delete" : "insert" };
    if (step.kind === "move") return { kind: "move", source: step.target, target: step.source };
    if (step.kind === "deleteSheet" || step.kind === "restoreSheet") return { ...step, kind: step.kind === "deleteSheet" ? "restoreSheet" : "deleteSheet" };
    const rows: number[] = []; step.rows.forEach((target, index) => { rows[target - step.range.from.row] = step.range.from.row + index; }); return { ...step, rows };
  }));
}
