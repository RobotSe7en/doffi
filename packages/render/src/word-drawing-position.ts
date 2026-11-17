import type { DocumentSection, DeepReadonly, WordDrawing } from "@doffi/core";
import type { RenderedPage } from "./contracts.js";

/** Coordinates are measured after pagination and column projection. The
 * paragraph's CSS box can span several columns; its anchor identifies the
 * actual column and line containing this object. All results use CSS pixels. */
export function wordDrawingPosition(
  drawing: DeepReadonly<WordDrawing>,
  page: RenderedPage,
  surface: DOMRect,
  paragraph: HTMLElement,
  anchor: HTMLElement,
  section?: DeepReadonly<DocumentSection>,
): { left: number; top: number; offsetX: number; offsetY: number } {
  const g = drawing.geometry,
    l = drawing.layout,
    p = paragraph.getBoundingClientRect(),
    a = anchor.getBoundingClientRect(),
    s = page.setup,
    m = s.margins;
  const left = surface.left + page.rect.x / 0.75,
    top = surface.top + page.rect.y / 0.75,
    width = s.width / 0.75,
    height = s.height / 0.75;
  const odd = page.index % 2 === 0;
  const horizontal = l.horizontal ?? "column",
    vertical = l.vertical ?? "paragraph";
  let h = { start: left + m.left / 0.75, size: (s.width - m.left - m.right) / 0.75 },
    v = { start: p.top, size: p.height };
  if (horizontal === "page") h = { start: left, size: width };
  else if (horizontal === "character") h = { start: a.left, size: 0 };
  else if (["leftMargin", "rightMargin", "insideMargin", "outsideMargin"].includes(horizontal)) {
    const leading =
      horizontal === "leftMargin" ||
      (horizontal === "insideMargin" && odd) ||
      (horizontal === "outsideMargin" && !odd);
    h = leading
      ? { start: left, size: m.left / 0.75 }
      : { start: left + width - m.right / 0.75, size: m.right / 0.75 };
  } else if (horizontal === "column") {
    const columns = section?.columns,
      count = columns?.count ?? 1,
      gap = (columns?.gap ?? 0) / 0.75,
      widths =
        columns?.widths?.map((value) => value / 0.75) ??
        Array.from({ length: count }, () => (h.size - gap * (count - 1)) / count);
    let start = h.start;
    for (let index = 0; index < count; index++) {
      const size = widths[index]!,
        space = (columns?.spaces?.[index] ?? columns?.gap ?? 0) / 0.75;
      if (a.left < start + size + space / 2 || index === count - 1) {
        h = { start, size };
        break;
      }
      start += size + space;
    }
  }
  if (vertical === "page") v = { start: top, size: height };
  else if (vertical === "margin")
    v = { start: top + m.top / 0.75, size: (s.height - m.top - m.bottom) / 0.75 };
  else if (vertical === "topMargin") v = { start: top, size: m.top / 0.75 };
  else if (vertical === "bottomMargin")
    v = { start: top + height - m.bottom / 0.75, size: m.bottom / 0.75 };
  else if (vertical === "insideMargin" || vertical === "outsideMargin") {
    // Side-bound margin strips extend over the page height. A top-bound
    // document instead has top/bottom inside and outside margin strips.
    v =
      section?.gutterPosition === "top"
        ? vertical === "insideMargin"
          ? { start: top, size: m.top / 0.75 }
          : { start: top + height - m.bottom / 0.75, size: m.bottom / 0.75 }
        : { start: top, size: height };
  } else if (vertical === "line") v = { start: a.top, size: a.height };
  if (l.layoutInCell !== false) {
    const cell = paragraph.closest<HTMLElement>("td,th");
    if (cell) {
      const rect = cell.getBoundingClientRect();
      const css = cell.ownerDocument.defaultView!.getComputedStyle(cell),
        leftInset = (parseFloat(css.paddingLeft) || 0) + (parseFloat(css.borderLeftWidth) || 0),
        rightInset = (parseFloat(css.paddingRight) || 0) + (parseFloat(css.borderRightWidth) || 0),
        topInset =
          (parseFloat(css.paddingTop) || 0) +
          (parseFloat(css.borderTopWidth) || 0) +
          (l.distances?.top ?? 0) / 0.75,
        bottomInset =
          (parseFloat(css.paddingBottom) || 0) +
          (parseFloat(css.borderBottomWidth) || 0) +
          (l.distances?.bottom ?? 0) / 0.75;
      if (horizontal !== "character")
        h = { start: rect.left + leftInset, size: rect.width - leftInset - rightInset };
      if (vertical !== "line" && vertical !== "paragraph")
        v = {
          start: rect.top + topInset,
          size: Math.max(g.height / 0.75, rect.height - topInset - bottomInset),
        };
    }
  }
  const aligned = (
    kind: string | undefined,
    size: number,
    object: number,
    offset: number,
    horizontal: boolean,
  ): number => {
    if (kind === "center") return (size - object) / 2;
    if (kind === "inside") return odd ? 0 : size - object;
    if (kind === "outside") return odd ? size - object : 0;
    if (kind === (horizontal ? "right" : "bottom")) return size - object;
    if (kind === (horizontal ? "left" : "top")) return 0;
    return offset / 0.75;
  };
  const offsetX = aligned(l.alignH, h.size, g.width / 0.75, g.x, true),
    offsetY = aligned(l.alignV, v.size, g.height / 0.75, g.y, false);
  let x = h.start + offsetX,
    y = v.start + offsetY;
  const cell = l.layoutInCell !== false ? paragraph.closest<HTMLElement>("td,th") : null;
  if (cell) {
    const bounds = cell.getBoundingClientRect(),
      css = cell.ownerDocument.defaultView!.getComputedStyle(cell),
      left =
        bounds.left + (parseFloat(css.paddingLeft) || 0) + (parseFloat(css.borderLeftWidth) || 0),
      right =
        bounds.right -
        (parseFloat(css.paddingRight) || 0) -
        (parseFloat(css.borderRightWidth) || 0),
      top = bounds.top + (parseFloat(css.paddingTop) || 0) + (parseFloat(css.borderTopWidth) || 0);
    x = Math.max(left, Math.min(x, Math.max(left, right - g.width / 0.75)));
    y = Math.max(top, y);
  }
  return {
    left: x,
    top: y,
    offsetX: (x - h.start) * 0.75,
    offsetY: (y - v.start) * 0.75,
  };
}
