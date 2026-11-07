import type { WorkbookOperation } from "../contracts/workbook-operations.js";
import type { WorkbookHyperlink, WorkbookSheet } from "../contracts/workbook.js";
import type { DeepReadonly } from "../runtime.js";
import { checkLimit } from "../document/validation.js";
import { validateWorkbookHyperlink } from "../workbook/hyperlinks.js";
import { intersectWorkbookRange, subtractWorkbookRange } from "../workbook/range-mapping.js";
import { checkRange, sameWorkbookRangeCoverage } from "../workbook/ranges.js";

/** Replace only the requested coordinates. A partial removal retains exact
 * outside fragments and never expands an implicit region into cell nodes. */
export function editWorkbookHyperlink(sheet: WorkbookSheet, operation: Extract<WorkbookOperation, { op: "setWorkbookHyperlink" }>, maxItems: number): void {
  checkRange(operation.range); if (operation.target) validateWorkbookHyperlink(operation.target);
  const previous = sheet.hyperlinks ?? [], next: WorkbookHyperlink[] = [];
  for (const link of previous) {
    if (!intersectWorkbookRange(link.range, operation.range)) next.push(link);
    else for (const range of subtractWorkbookRange(link.range, operation.range)) next.push({ ...link, range });
    checkLimit(next.length, maxItems);
  }
  if (operation.target) next.push({ ...operation.target, range: operation.range });
  checkLimit(next.length, maxItems);
  if (!operation.target && next.length === previous.length && next.every((link, index) => link === previous[index])) return;
  sheet.hyperlinks = next;
}

export function sameWorkbookHyperlinks(left: readonly DeepReadonly<WorkbookHyperlink>[], right: readonly DeepReadonly<WorkbookHyperlink>[]): boolean {
  const key = ({ url, location, tooltip, display }: DeepReadonly<WorkbookHyperlink>): string => JSON.stringify({ url, location, tooltip, display });
  const group = (links: readonly DeepReadonly<WorkbookHyperlink>[]) => {
    const groups = new Map<string, WorkbookHyperlink["range"][]>();
    for (const link of links) { const identity = key(link), ranges = groups.get(identity) ?? []; ranges.push(link.range); groups.set(identity, ranges); }
    return groups;
  };
  const before = group(left), after = group(right);
  return before.size === after.size && [...before].every(([key, ranges]) => after.has(key) && sameWorkbookRangeCoverage(ranges, after.get(key)!));
}
