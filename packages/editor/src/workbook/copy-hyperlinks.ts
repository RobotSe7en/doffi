import { WorkbookRegionIndex } from "@doffi/core";
import type { DeepReadonly, WorkbookHyperlink, WorkbookRange } from "@doffi/core";

type Target = Omit<WorkbookHyperlink, "range">;
interface SourceTarget { readonly key: string; readonly target: Target }
/** Frozen source indexes and row runs keep skip-blank and overlapping copies
 * precise. Repeated targets coalesce vertically without expanding a region. */
export function createWorkbookHyperlinkCopy(source: readonly DeepReadonly<WorkbookHyperlink>[], target: readonly DeepReadonly<WorkbookHyperlink>[], sourceRange: DeepReadonly<WorkbookRange>, targetRange: DeepReadonly<WorkbookRange>, step: () => void): {
  cell(fromRow: number, fromColumn: number, row: number, column: number, copied: boolean): void;
  finish(): WorkbookHyperlink[] | undefined;
} {
  const targets = new Map<string, SourceTarget>();
  function index(links: readonly DeepReadonly<WorkbookHyperlink>[], scope: DeepReadonly<WorkbookRange>) {
    const ranges: DeepReadonly<WorkbookRange>[] = [];
    const values = new Map<DeepReadonly<WorkbookRange>, SourceTarget>();
    for (const link of links) {
      step(); if (link.range.to.row < scope.from.row || link.range.from.row > scope.to.row || link.range.to.column < scope.from.column || link.range.from.column > scope.to.column) continue;
      ranges.push(link.range); const target: Target = { ...(link.url !== undefined ? { url: link.url } : {}), ...(link.location !== undefined ? { location: link.location } : {}), ...(link.tooltip !== undefined ? { tooltip: link.tooltip } : {}), ...(link.display !== undefined ? { display: link.display } : {}) };
      const key = JSON.stringify(target); let value = targets.get(key); if (!value) { value = { key, target }; targets.set(key, value); } values.set(link.range, value);
    }
    return { regions: new WorkbookRegionIndex(ranges), values };
  }
  const from = index(source, sourceRange), destination = index(target, targetRange), links: WorkbookHyperlink[] = [];
  let changed = false, rowNumber = -1, previous = new Map<string, WorkbookHyperlink>(), current = new Map<string, WorkbookHyperlink>();
  let run: { value: SourceTarget; from: number; to: number } | undefined;
  function flush(): void {
    if (!run) return;
    const key = JSON.stringify([run.value.key, run.from, run.to]), prior = previous.get(key);
    if (prior && prior.range.to.row === rowNumber - 1) { prior.range.to.row = rowNumber; current.set(key, prior); }
    else { const link = { ...run.value.target, range: { from: { row: rowNumber, column: run.from }, to: { row: rowNumber, column: run.to } } }; links.push(link); current.set(key, link); }
    run = undefined;
  }
  return { cell(fromRow, fromColumn, row, column, copied) {
    step();
    if (rowNumber !== row) { flush(); previous = current; current = new Map(); rowNumber = row; }
    const sourceRange = copied ? from.regions.at(fromRow, fromColumn) : undefined, targetRange = destination.regions.at(row, column);
    const old = targetRange ? destination.values.get(targetRange) : undefined, next = copied ? sourceRange ? from.values.get(sourceRange) : undefined : old;
    changed ||= next?.key !== old?.key;
    if (run && (run.value.key !== next?.key || run.to !== column - 1)) flush();
    if (next) { if (run) run.to = column; else run = { value: next, from: column, to: column }; }
  }, finish() { flush(); return changed ? links : undefined; } };
}
