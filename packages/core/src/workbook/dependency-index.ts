import type { DeepReadonly } from "../runtime.js";
import type { WorkbookReferenceArea } from "./reference.js";
export interface WorkbookDependencyEntry { readonly area: DeepReadonly<WorkbookReferenceArea>; readonly owner: string }

/** Rectangles remain single index entries; whole-column dependencies never
 * become a million point edges. Tree traversal reports actual intersections. */
export class WorkbookDependencyIndex {
  private readonly sheets = new Map<string, { entries: WorkbookDependencyEntry[]; maximum: Int32Array }>();
  constructor(entries: readonly WorkbookDependencyEntry[], step: () => void) {
    for (const entry of entries) {
      step(); let sheet = this.sheets.get(entry.area.sheetId); if (!sheet) { sheet = { entries: [], maximum: new Int32Array(0) }; this.sheets.set(entry.area.sheetId, sheet); } sheet.entries.push(entry);
    }
    for (const sheet of this.sheets.values()) {
      sheet.entries.sort((a, b) => { step(); return a.area.from.row - b.area.from.row; }); sheet.maximum = new Int32Array(sheet.entries.length * 4);
      function build(node: number, first: number, last: number): number {
        step(); if (first === last) return sheet.maximum[node] = sheet.entries[first]!.area.to.row;
        const middle = (first + last) >>> 1; return sheet.maximum[node] = Math.max(build(node * 2, first, middle), build(node * 2 + 1, middle + 1, last));
      }
      if (sheet.entries.length) build(1, 0, sheet.entries.length - 1);
    }
  }
  *query(area: DeepReadonly<WorkbookReferenceArea>, step: () => void): Generator<string> {
    const sheet = this.sheets.get(area.sheetId); if (!sheet?.entries.length) return;
    function* visit(node: number, first: number, last: number): Generator<string> {
      step(); if (sheet!.maximum[node]! < area.from.row || sheet!.entries[first]!.area.from.row > area.to.row) return;
      if (first === last) {
        const entry = sheet!.entries[first]!;
        if (entry.area.from.column <= area.to.column && entry.area.to.column >= area.from.column) yield entry.owner;
        return;
      }
      const middle = (first + last) >>> 1; yield* visit(node * 2, first, middle); yield* visit(node * 2 + 1, middle + 1, last);
    }
    yield* visit(1, 0, sheet.entries.length - 1);
  }
}
