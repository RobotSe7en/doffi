import { floatingStory } from "./word-floating-placement.js";
import type { DocumentSection, DeepReadonly, WordDrawing, WordDrawingLayout } from "@doffi/core";
import type { RenderedPage } from "./contracts.js";
import { isDocumentTextBoundary } from "@doffi/core";
import { isMeasuredViewContent, viewNodeRole } from "./view-nodes.js";
import { columnRegion } from "./column-regions.js";

/** Word 2013+ constrains cell-anchored drawings to the cell even when the
 * saved legacy flag is false. Preserve that flag for native export. */
export function wordDrawingLayoutInCell(layout: DeepReadonly<WordDrawingLayout>, compatibilityMode = 12): boolean {
  return compatibilityMode >= 15 || layout.layoutInCell !== false;
}

/** A page gap may sit after a leading zero-width anchor. Find its first real
 * following glyph so the floating object follows the paragraph onto that page. */
export function wordDrawingLeadingTextRect(
  paragraph: HTMLElement,
  anchor: HTMLElement,
): DOMRect | undefined {
  const document = paragraph.ownerDocument,
    walker = document.createTreeWalker(paragraph, 4);
  for (let node = walker.nextNode(); node; node = walker.nextNode()) {
    const text = node as Text;
    if (
      !text.length ||
      text.parentElement?.closest("[data-sdk-word-object],[data-word-native-object]")
    )
      continue;
    let ignored = false;
    for (let p: Node | null = text.parentNode; p && p !== paragraph; p = p.parentNode)
      if (viewNodeRole(p) === "remove" && !isMeasuredViewContent(p)) {
        ignored = true;
        break;
      }
    if (
      ignored ||
      document.defaultView!.getComputedStyle(text.parentElement!).visibility !== "visible"
    )
      continue;
    if (!(anchor.compareDocumentPosition(text) & 4)) return;
    let end = 1;
    while (end < text.length && !isDocumentTextBoundary(text.data, end)) end++;
    const range = document.createRange();
    range.setStart(text, 0);
    range.setEnd(text, end);
    const rect = range.getBoundingClientRect();
    if (rect.width > 0 && rect.height > 0) return rect;
  }
  return;
}

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
  anchorRect?: DOMRect,
  compatibilityMode = 12,
): { left: number; top: number; offsetX: number; offsetY: number } {
  const g = drawing.geometry,
    l = drawing.layout,
    p = paragraph.getBoundingClientRect(),
    a = anchorRect ?? anchor.getBoundingClientRect(),
    s = page.setup,
    m = s.margins;
  const left = surface.left + page.rect.x / 0.75,
    top = surface.top + page.rect.y / 0.75,
    width = s.width / 0.75,
    height = s.height / 0.75;
  const odd = page.index % 2 === 0;
  const horizontal = l.horizontal ?? "column",
    vertical = l.vertical ?? "paragraph",
    story = floatingStory(anchor),
    cell = wordDrawingLayoutInCell(l, compatibilityMode) ? paragraph.closest<HTMLElement>("td,th") : null,
    column = !story ? columnRegion(anchor, a, cell ?? paragraph) : undefined,
    paragraphFragment = !cell ? column?.fragment : undefined;
  let h = { start: left + m.left / 0.75, size: (s.width - m.left - m.right) / 0.75 },
    v = {
      start: paragraphFragment?.top ?? (anchorRect ? Math.max(p.top, top + m.top / 0.75) : p.top),
      size: anchorRect
        ? Math.max(a.height, p.bottom - Math.max(p.top, top + m.top / 0.75))
        : paragraphFragment
          ? paragraphFragment.bottom - paragraphFragment.top
          : p.height,
    };
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
    const region = story?.getBoundingClientRect() ?? column;
    if (region) h = { start: region.left, size: region.right - region.left };
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
  if (cell) {
    const rect = column?.fragment ?? cell.getBoundingClientRect();
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
      h = { start: rect.left + leftInset, size: rect.right - rect.left - leftInset - rightInset };
    if (vertical !== "line" && vertical !== "paragraph")
      v = {
        start: rect.top + topInset,
        size: Math.max(g.height / 0.75, rect.bottom - rect.top - topInset - bottomInset),
      };
  } else if (vertical === "paragraph") {
    // A legacy object outside its cell uses the cell's row origin, including
    // when its anchor is in a later paragraph. Cell-contained objects keep
    // that paragraph's own origin instead.
    const owner = paragraph.closest<HTMLElement>("td,th");
    if (owner) v.start = owner.getBoundingClientRect().top;
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
  if (cell) {
    const bounds = column?.fragment ?? cell.getBoundingClientRect(),
      css = cell.ownerDocument.defaultView!.getComputedStyle(cell),
      top = bounds.top + (parseFloat(css.paddingTop) || 0) + (parseFloat(css.borderTopWidth) || 0);
    // Word constrains a fitting object to the cell grid, including its padding
    // area. Negative offsets may enter that padding; the text inset is not a
    // clipping boundary. Oversized objects retain their native position.
    if (g.width / .75 <= bounds.right - bounds.left)
      x = Math.max(bounds.left, Math.min(x, bounds.right - g.width / 0.75));
    y = Math.max(top, y);
  }
  return { left: x, top: y, offsetX: (x - h.start) * 0.75, offsetY: (y - v.start) * 0.75 };
}
