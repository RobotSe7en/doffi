import { DocumentError } from "@doffi/core";
import type { DeepReadonly, WordDrawing, WordDrawingLayout } from "@doffi/core";

export interface WrapPoint {
  x: number;
  y: number;
}
export interface WrapBounds {
  left: number;
  right: number;
  top: number;
  bottom: number;
}
export interface WrapInterval {
  left: number;
  right: number;
}

export function wordDrawingCollisionBounds(
  position: { left: number; top: number },
  geometry: DeepReadonly<WordDrawing["geometry"]>,
): WrapBounds {
  const points = wordWrapPoints(
    [
      { x: 0, y: 0 },
      { x: 21600, y: 0 },
      { x: 21600, y: 21600 },
      { x: 0, y: 21600 },
    ],
    { ...position, width: geometry.width / 0.75, height: geometry.height / 0.75 },
    geometry,
  );
  return {
    left: Math.min(...points.map(({ x }) => x)),
    right: Math.max(...points.map(({ x }) => x)),
    top: Math.min(...points.map(({ y }) => y)),
    bottom: Math.max(...points.map(({ y }) => y)),
  };
}

/** Word uses a 21600-square coordinate space. Apply the same center rotation
 * and reflections as the drawing, before taking any exclusion bounds. */
export function wordWrapPoints(
  points: readonly DeepReadonly<WrapPoint>[],
  box: { left: number; top: number; width: number; height: number },
  orientation: { rotation: number; flipH: boolean; flipV: boolean },
  reserve: () => void = () => {},
): WrapPoint[] {
  const angle = (orientation.rotation * Math.PI) / 180,
    c = Math.cos(angle),
    s = Math.sin(angle);
  return points.map((p) => {
    reserve();
    const x = (p.x / 21600 - 0.5) * box.width * (orientation.flipH ? -1 : 1),
      y = (p.y / 21600 - 0.5) * box.height * (orientation.flipV ? -1 : 1);
    const point = {
      x: box.left + box.width / 2 + x * c - y * s,
      y: box.top + box.height / 2 + x * s + y * c,
    };
    if (!Number.isFinite(point.x) || !Number.isFinite(point.y))
      throw new DocumentError(
        "ResourceLimitExceeded",
        "Word wrap coordinates exceed the render geometry budget.",
      );
    return point;
  });
}

function merge(intervals: WrapInterval[]): WrapInterval[] {
  intervals.sort((a, b) => a.left - b.left);
  const result: WrapInterval[] = [];
  for (const span of intervals) {
    const last = result.at(-1);
    if (last && span.left <= last.right + 0.01) last.right = Math.max(last.right, span.right);
    else result.push({ ...span });
  }
  return result;
}

/** Project the filled polygon over the entire glyph-height band, rather than
 * sampling a line's center. Through keeps concave openings; tight closes them.
 * The caller accounts for edge work against its render measurement budget. */
export function wordPolygonIntervals(
  points: readonly WrapPoint[],
  top: number,
  bottom: number,
  through: boolean,
  reserve: () => void = () => {},
): WrapInterval[] {
  if (bottom <= top || points.length < 3) return [];
  const ys = [
      ...new Set([top, bottom, ...points.map((p) => p.y).filter((y) => y > top && y < bottom)]),
    ].sort((a, b) => a - b),
    spans: WrapInterval[] = [];
  for (let index = 1; index < ys.length; index++) {
    const lo = ys[index - 1]!,
      hi = ys[index]!,
      mid = (lo + hi) / 2,
      edges: { x: number; lo: number; hi: number }[] = [];
    for (let i = 0; i < points.length; i++) {
      reserve();
      const a = points[i]!,
        b = points[(i + 1) % points.length]!;
      if (a.y > mid === b.y > mid) continue;
      const x = (y: number) => a.x + ((b.x - a.x) * (y - a.y)) / (b.y - a.y);
      edges.push({ x: x(mid), lo: x(lo), hi: x(hi) });
    }
    edges.sort((a, b) => a.x - b.x);
    for (let i = 1; i < edges.length; i += 2) {
      const a = edges[i - 1]!,
        b = edges[i]!;
      spans.push({
        left: Math.min(a.lo, a.hi, b.lo, b.hi),
        right: Math.max(a.lo, a.hi, b.lo, b.hi),
      });
    }
  }
  const result = merge(spans);
  return through || !result.length
    ? result
    : [{ left: result[0]!.left, right: result.at(-1)!.right }];
}

export function wordWrapIntervals(
  bounds: WrapBounds,
  points: readonly WrapPoint[] | undefined,
  layout: DeepReadonly<WordDrawingLayout>,
  line: WrapBounds,
  paragraph: WrapInterval,
  rtl: boolean,
  reserve: () => void = () => {},
): WrapInterval[] {
  const d = layout.distances ?? { top: 0, right: 0, bottom: 0, left: 0 },
    top = line.top - d.bottom / 0.75,
    bottom = line.bottom + d.top / 0.75;
  if (bottom <= bounds.top + 0.25 || top >= bounds.bottom - 0.25) return [];
  const wrap = layout.wrap ?? "square";
  if (wrap === "none" || layout.behindText) return [];
  if (wrap === "topAndBottom") return [{ left: paragraph.left, right: paragraph.right }];
  let spans =
    points && (wrap === "tight" || wrap === "through")
      ? wordPolygonIntervals(points, top, bottom, wrap === "through", reserve)
      : [{ left: bounds.left, right: bounds.right }];
  spans = merge(
    spans.map((span) => ({ left: span.left - d.left / 0.75, right: span.right + d.right / 0.75 })),
  );
  if (!spans.length) return spans;
  const first = spans[0]!,
    last = spans.at(-1)!;
  let side = layout.wrapSide ?? "bothSides";
  if (side === "largest") {
    const left = Math.max(0, first.left - paragraph.left),
      right = Math.max(0, paragraph.right - last.right);
    side =
      Math.abs(left - right) < 0.25 ? (rtl ? "right" : "left") : left > right ? "left" : "right";
  }
  if (side === "left") spans = [{ left: first.left, right: Math.max(last.right, paragraph.right) }];
  if (side === "right") spans = [{ left: Math.min(first.left, paragraph.left), right: last.right }];
  return spans;
}

/** Resolve collisions in source order. Earlier objects keep their positions;
 * a later object moves downward if either participant disallows overlap. */
export function avoidWordDrawingOverlap(
  bounds: WrapBounds,
  allowOverlap: boolean,
  placed: readonly (WrapBounds & { allowOverlap: boolean })[],
  reserve: () => void = () => {},
): number {
  let offset = 0;
  for (let pass = 0; pass <= placed.length; pass++) {
    let next = offset;
    for (const other of placed) {
      reserve();
      if (
        (allowOverlap && other.allowOverlap) ||
        bounds.right <= other.left + 0.01 ||
        bounds.left >= other.right - 0.01 ||
        bounds.bottom + offset <= other.top + 0.01 ||
        bounds.top + offset >= other.bottom - 0.01
      )
        continue;
      next = Math.max(next, other.bottom - bounds.top);
    }
    if (next <= offset) break;
    offset = next;
  }
  return offset;
}
