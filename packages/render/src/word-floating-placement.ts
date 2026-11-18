import { DocumentError } from "@doffi/core";
import type { RenderedPage, RenderLimits } from "./contracts.js";
import { avoidWordDrawingOverlap } from "./word-wrap-geometry.js";
import type { WrapBounds } from "./word-wrap-geometry.js";

interface Placement {
  anchor: HTMLElement;
  page: RenderedPage;
  scope: HTMLElement;
  bounds: WrapBounds;
  allowOverlap: boolean;
  apply?(offset: number): void;
}

export function wordFloatingPlacement(limits: Readonly<RenderLimits>) {
  const placements: Placement[] = [];
  return {
    add(placement: Placement): void {
      placements.push(placement);
    },
    place(): void {
      const pages = new Map<
        RenderedPage,
        Map<HTMLElement, (WrapBounds & { allowOverlap: boolean; anchor: HTMLElement })[]>
      >();
      let work = 0;
      const reserve = () => {
        if (++work > limits.maxNodes * 32)
          throw new DocumentError(
            "ResourceLimitExceeded",
            "Floating object collisions exceed the render measurement budget.",
          );
      };
      // Rows on continuation columns already occupy measured flow. Place
      // adjustable objects against those spans and then against one another.
      placements.sort(
        (a, b) =>
          Number(Boolean(a.apply)) - Number(Boolean(b.apply)) ||
          (a.anchor === b.anchor ? 0 : a.anchor.compareDocumentPosition(b.anchor) & 4 ? -1 : 1),
      );
      const offsets = placements.map(({ page, scope, bounds, allowOverlap, anchor, apply }) => {
        let scopes = pages.get(page);
        if (!scopes) pages.set(page, (scopes = new Map()));
        let seen = scopes.get(scope);
        if (!seen) scopes.set(scope, (seen = []));
        const others = apply
          ? seen.filter((other) => {
              reserve();
              return other.anchor !== anchor;
            })
          : [];
        const delta = avoidWordDrawingOverlap(bounds, allowOverlap, others, reserve);
        seen.push({
          ...bounds,
          top: bounds.top + delta,
          bottom: bounds.bottom + delta,
          allowOverlap,
          anchor,
        });
        return delta;
      });
      for (const [index, placement] of placements.entries()) placement.apply?.(offsets[index]!);
      placements.length = 0;
    },
  };
}

export type WordFloatingPlacement = ReturnType<typeof wordFloatingPlacement>;

const storyPaint = new WeakMap<HTMLElement, HTMLElement>();

export function registerFloatingStoryPaint(node: HTMLElement, story: HTMLElement): void {
  storyPaint.set(node, story);
}

export function floatingStory(node?: HTMLElement | null): HTMLElement | undefined {
  for (let parent = node; parent; parent = parent.parentElement) {
    if (parent.classList.contains("sdk-page-story")) return parent;
    const story = storyPaint.get(parent);
    if (story) return story;
  }
  return undefined;
}
