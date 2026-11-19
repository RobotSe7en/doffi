import type { WrapBounds } from "./word-wrap-geometry.js";
import type { DeepReadonly, DocumentSection } from "@doffi/core";
import { viewNodeRole } from "./view-nodes.js";
import { wordTableRepeatHeader } from "./word-table-layout.js";

export interface ColumnRegion extends WrapBounds {
  readonly ordinal: number;
  readonly fragment?: WrapBounds;
  readonly fragmented?: boolean;
}

type Locate = (node: HTMLElement, rect: DOMRect, owner?: HTMLElement) => ColumnRegion | undefined;
const regions = new WeakMap<HTMLElement, Locate>();
type TableOffset = (width: number, pageIndex: number) => number;
export const floatingTableColumnOffset: WeakMap<HTMLElement, TableOffset> = new WeakMap<
  HTMLElement,
  TableOffset
>();

export function floatingTableFragments(table: HTMLTableElement): ReadonlyMap<number, WrapBounds> {
  const bounds = new Map<number, WrapBounds>();
  const repeats = table.rows[0] && wordTableRepeatHeader(table.rows[0]);
  for (const row of table.rows) {
    if (viewNodeRole(row) === "remove" || !row.getClientRects().length) continue;
    const rect = row.getBoundingClientRect(),
      column = columnRegion(row, rect);
    if (!column) continue;
    const current = bounds.get(column.ordinal);
    bounds.set(column.ordinal, {
      left: Math.min(current?.left ?? Infinity, rect.left),
      right: Math.max(current?.right ?? -Infinity, rect.right),
      top: current?.top ?? (repeats && bounds.size ? column.top : rect.top),
      bottom: rect.bottom,
    });
  }
  return bounds;
}

export function setColumnRegions(flow: HTMLElement, locate: Locate): void {
  regions.set(flow, locate);
}

/** Coordinates in the containing section's projected flow. A physical page
 * can contain several bands with independent column geometry. */
export function columnRegion(
  node: HTMLElement,
  rect: DOMRect = node.getBoundingClientRect(),
  owner?: HTMLElement,
): ColumnRegion | undefined {
  for (let parent: HTMLElement | null = node; parent; parent = parent.parentElement) {
    const locate = regions.get(parent);
    if (locate) return locate(node, rect, owner);
  }
  return undefined;
}

export function sameColumnGeometry(
  a: DeepReadonly<DocumentSection>["columns"],
  b: DeepReadonly<DocumentSection>["columns"],
  width: number,
): boolean {
  const count = a?.count ?? 1;
  if (count !== (b?.count ?? 1)) return false;
  const gap = (columns: typeof a, index: number) => columns?.spaces?.[index] ?? columns?.gap ?? 0;
  const span = (columns: typeof a, index: number) =>
    columns?.widths?.[index] ?? (width - (count - 1) * (columns?.gap ?? 0)) / count;
  for (let index = 0; index < count; index++)
    if (
      Math.abs(span(a, index) - span(b, index)) > 0.001 ||
      (index < count - 1 && Math.abs(gap(a, index) - gap(b, index)) > 0.001)
    )
      return false;
  return true;
}
