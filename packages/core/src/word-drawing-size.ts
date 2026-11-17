import { DocumentError } from "./errors.js";
import type { DeepReadonly } from "./runtime.js";
import type { WordDrawing } from "./contracts/word-drawing.js";
import type { PageSetup, DocumentSection } from "./contracts/document.js";

/** The page setup already contains physical mirrored/gutter margins. Relative
 * sizing refers to the page, independently of the object's position reference
 * and the cell in which its anchor happens to live. Percent zero retains the
 * native disabled declaration and uses the authored absolute extent. */
export function wordDrawingDimensions(
  drawing: DeepReadonly<WordDrawing>,
  page: { index: number; setup: DeepReadonly<PageSetup> },
  section?: DeepReadonly<DocumentSection>,
): WordDrawing["geometry"] {
  const { setup: s, index } = page,
    m = s.margins,
    odd = index % 2 === 0;
  const horizontal = (from: string): number =>
    from === "page"
      ? s.width
      : from === "margin"
        ? s.width - m.left - m.right
        : from === "leftMargin"
          ? m.left
          : from === "rightMargin"
            ? m.right
            : (from === "insideMargin") === odd
              ? m.left
              : m.right;
  const vertical = (from: string): number =>
    from === "page"
      ? s.height
      : from === "margin"
        ? s.height - m.top - m.bottom
        : from === "topMargin"
          ? m.top
          : from === "bottomMargin"
            ? m.bottom
            : section?.gutterPosition === "top"
              ? from === "insideMargin"
                ? m.top
                : m.bottom
              : s.height;
  const w = drawing.layout.relativeWidth,
    h = drawing.layout.relativeHeight,
    width =
      w && w.percent > 0 ? (horizontal(w.relativeFrom) * w.percent) / 100 : drawing.geometry.width,
    height =
      h && h.percent > 0 ? (vertical(h.relativeFrom) * h.percent) / 100 : drawing.geometry.height;
  if (![width, height].every(Number.isFinite) || width < 0 || height < 0)
    throw new DocumentError(
      "ResourceLimitExceeded",
      "Relative drawing dimensions exceed the render geometry budget.",
    );
  return { ...drawing.geometry, width, height };
}
