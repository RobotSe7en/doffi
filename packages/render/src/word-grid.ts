import type { DeepReadonly, DocumentSection, PageSetup, WordMetadata } from "@doffi/core";
import type { RenderLimits } from "./contracts.js";
import { ownViewStyle, viewNodeRole, isMeasuredViewContent } from "./view-nodes.js";
import { wordCharacterGrid } from "./word-character-grid.js";

/** Native margins exclude the binding gutter. Keep the effective box derived. */
export function wordPageSetup(
  section: DeepReadonly<DocumentSection>,
  fallback: DeepReadonly<PageSetup>,
): DeepReadonly<PageSetup> {
  const setup = section.pageSetup ?? fallback;
  if (!section.gutter) return setup;
  const side = section.gutterPosition ?? "left";
  return { ...setup, margins: { ...setup.margins, [side]: setup.margins[side] + section.gutter } };
}
/** Facing-page parity is physical and independent of printed page labels. */
export function facingWordPageSetup(
  setup: DeepReadonly<PageSetup>,
  index: number,
  mirror: boolean,
): DeepReadonly<PageSetup> {
  return mirror && index % 2 === 1
    ? {
        ...setup,
        margins: { ...setup.margins, left: setup.margins.right, right: setup.margins.left },
      }
    : setup;
}
/** A grid changes measured text, never authored paragraph/run formatting. */
export function applyWordGrid(
  document: Document,
  root: HTMLElement,
  section: DeepReadonly<DocumentSection>,
  limits: Readonly<RenderLimits>,
  word?: DeepReadonly<WordMetadata>,
  diagnostic: (message: string) => void = () => {},
): { measure(): void; dispose(): void } | undefined {
  const grid = section.grid;
  if (!grid || grid.type === "default") return;
  const view = document.defaultView!,
    lines = new Map<HTMLElement, number>(),
    paragraphs = [...root.querySelectorAll<HTMLElement>("p,h1,h2,h3,h4,h5,h6,li")].filter(
      (paragraph) => !paragraph.closest("[data-sdk-word-endnotes]"),
    );
  const natural = (element: HTMLElement): number => {
    const css = view.getComputedStyle(element),
      font = parseFloat(css.fontSize) || 0;
    return Math.max(font, parseFloat(css.lineHeight) || font * 1.2);
  };
  // Capture all run metrics before changing the paragraph strut. A large run
  // reserves a multiple of the pitch on its own line, leaving other lines small.
  for (const paragraph of paragraphs) {
    const css = view.getComputedStyle(paragraph);
    if (
      !grid.linePitch ||
      css.getPropertyValue("--word-paragraph-snap-to-grid").trim() === "0" ||
      (!word?.adjustLineHeightInTable && paragraph.closest("td,th")) ||
      (paragraph.dataset.wordLineRule ?? css.getPropertyValue("--word-line-rule").trim()) ===
        "exact"
    )
      continue;
    for (const node of [
      paragraph,
      ...paragraph.querySelectorAll<HTMLElement>("span,strong,em,u,s,sub,sup,ins,del"),
    ]) {
      if (
        node.closest("[data-word-drawing],[data-word-native-object],[data-word-equation]") ||
        (viewNodeRole(node) === "remove" && !isMeasuredViewContent(node)) ||
        parseFloat(view.getComputedStyle(node).fontSize) < 0.01
      )
        continue;
      lines.set(node, natural(node));
    }
  }
  let characters: ReturnType<typeof wordCharacterGrid> | undefined;
  if (grid.type === "snapToChars" || grid.type === "linesAndChars") {
    const probe = document.createElement("span"),
      normal =
        word?.styles.find((style) => style.kind === "paragraph" && style.styleId === "Normal") ??
        word?.styles.find((style) => style.kind === "paragraph" && style.default),
      chain: DeepReadonly<WordMetadata>["styles"][number][] = [];
    let style = normal;
    const seen = new Set<string>();
    while (style && !seen.has(style.styleId)) {
      seen.add(style.styleId);
      chain.unshift(style);
      style = word?.styles.find((item) => item.styleId === style!.basedOn);
    }
    probe.className = [
      word?.defaultRunCssClass,
      ...chain.flatMap((style) => [style.cssClass, style.runCssClass]),
    ]
      .filter(Boolean)
      .join(" ");
    probe.style.cssText = "position:absolute;visibility:hidden;pointer-events:none";
    document.body.append(probe);
    const spacing = (((grid.characterSpacing ?? 0) / 4096) * 4) / 3,
      pitch = parseFloat(view.getComputedStyle(probe).fontSize) + spacing;
    probe.remove();
    if (grid.type === "linesAndChars" || pitch > 0)
      characters = wordCharacterGrid(document, root, pitch, grid.type, spacing, limits, diagnostic);
    else
      diagnostic(
        "A nonpositive character pitch is preserved and displayed without fixed character snapping.",
      );
  }
  return {
    measure() {
      if (grid.linePitch)
        for (const [node, line] of lines)
          if (node.isConnected) {
            ownViewStyle(node, "line-height");
            node.style.lineHeight =
              Math.max(1, Math.ceil((line - 0.01) / ((grid.linePitch * 4) / 3))) *
                ((grid.linePitch * 4) / 3) +
              "px";
          }
      characters?.measure();
    },
    dispose() {
      characters?.dispose();
    },
  };
}
