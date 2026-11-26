import type { DeepReadonly, DocumentSection, PageSetup, WordMetadata } from "@doffi/core";
import type { RenderLimits } from "./contracts.js";
import { ownViewStyle, viewNodeRole, isMeasuredViewContent } from "./view-nodes.js";
import { wordCharacterGrid } from "./word-character-grid.js";
import { wordAutoLineMultipliers } from "./word-line-spacing.js";

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
  includeCharacters = true,
): { measure(): void; dispose(): void } | undefined {
  const grid = section.grid;
  if (!grid || grid.type === "default") return;
  const view = document.defaultView!,
    pitch = ((grid.linePitch ?? 0) * 4) / 3,
    lines = new Map<HTMLElement, number>(),
    paragraphs = [...root.querySelectorAll<HTMLElement>("p,h1,h2,h3,h4,h5,h6,li")].filter(
      (paragraph) => !paragraph.closest("[data-sdk-word-endnotes]"),
    );
  const normalHeights = new Map<string, number>();
  const probe = document.createElement("span");
  probe.style.cssText =
    "position:absolute;visibility:hidden;display:block;line-height:normal;white-space:pre";
  const singleLine = (css: CSSStyleDeclaration, eastAsian: boolean): number => {
    const { fontFamily, fontSize, fontWeight, fontStyle } = css,
      sample = eastAsian ? "汉" : "M",
      key = JSON.stringify([fontFamily, fontSize, fontWeight, fontStyle, sample]),
      known = normalHeights.get(key);
    if (known !== undefined) return known;
    if (!probe.isConnected) document.body.append(probe);
    Object.assign(probe.style, { fontFamily, fontSize, fontWeight, fontStyle });
    probe.textContent = sample;
    const height = probe.getBoundingClientRect().height;
    normalHeights.set(key, height);
    return height;
  };
  // Capture all run metrics before changing the paragraph strut. A large run
  // reserves a multiple of the pitch on its own line, leaving other lines small.
  for (const paragraph of paragraphs) {
    const css = view.getComputedStyle(paragraph),
      rule = paragraph.dataset.wordLineRule ?? css.getPropertyValue("--word-line-rule").trim();
    if (
      !pitch ||
      css.getPropertyValue("--word-paragraph-snap-to-grid").trim() === "0" ||
      (!word?.adjustLineHeightInTable && root.contains(paragraph.closest("td,th"))) ||
      rule === "exact" ||
      rule === "atLeast"
    )
      continue;
    const savedMultiple = Number(paragraph.dataset.sdkWordAutoLineMultiple),
      multiple = wordAutoLineMultipliers.get(paragraph) ?? (savedMultiple > 0 ? savedMultiple : undefined);
    for (const node of [
      paragraph,
      ...paragraph.querySelectorAll<HTMLElement>("span,strong,em,u,s,sub,sup,ins,del"),
    ]) {
      const run = view.getComputedStyle(node),
        font = parseFloat(run.fontSize);
      if (
        root.contains(node.closest("[data-word-drawing],[data-word-native-object],[data-word-equation]")) ||
        (viewNodeRole(node) === "remove" && !isMeasuredViewContent(node)) ||
        font < 0.01
      )
        continue;
      const height = parseFloat(run.lineHeight),
        single = multiple
          ? Math.max(font, singleLine(run, node.closest('[data-word-script="eastAsia"]') !== null))
          : Math.max(font, height || font * 1.2);
      // Measure script roles separately so fallback glyphs from another script
      // do not expand the strut. Large runs can still require more grid rows.
      const snapped = Math.ceil((single - 0.01) / pitch);
      lines.set(node, Math.max(1, multiple ?? 1, snapped) * pitch);
    }
  }
  probe.remove();
  let characters: ReturnType<typeof wordCharacterGrid> | undefined;
  if (includeCharacters && (grid.type === "snapToChars" || grid.type === "linesAndChars")) {
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
      for (const [node, line] of lines)
        if (node.isConnected) {
          ownViewStyle(node, "line-height");
          node.style.lineHeight = line + "px";
        }
      characters?.measure();
    },
    dispose() {
      characters?.dispose();
    },
  };
}
