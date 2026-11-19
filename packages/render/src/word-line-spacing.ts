import { ownViewStyle } from "./view-nodes.js";

/** Preserve the native multiplier for section grids after resolving it to pixels. */
export const wordAutoLineMultipliers: WeakMap<HTMLElement, number> = new WeakMap();

function autoLineMultiplier(node: HTMLElement, view: Window): number {
  const typed = node as HTMLElement & {
    computedStyleMap?: () => { get: (name: string) => { unit?: string; value?: number } | undefined };
  };
  const line = typed.computedStyleMap?.().get("line-height") as
    | { unit?: string; value?: number }
    | undefined;
  if (line) return line.unit === "number" ? line.value ?? 0 : 0;
  const inline = Number(node.style.lineHeight);
  if (inline > 0) return inline;

  // Distinguish a unitless inherited rule from an absolute line height on
  // engines that only expose the resolved pixel value.
  const css = view.getComputedStyle(node),
    size = parseFloat(css.fontSize),
    height = parseFloat(css.lineHeight);
  if (!(size > 0) || !Number.isFinite(height)) return 0;
  const value = node.style.getPropertyValue("font-size"),
    priority = node.style.getPropertyPriority("font-size");
  node.style.setProperty("font-size", size * 2 + "px", "important");
  const doubled = parseFloat(view.getComputedStyle(node).lineHeight);
  if (value) node.style.setProperty("font-size", value, priority);
  else node.style.removeProperty("font-size");
  return Math.abs(doubled - height * 2) < 0.2 ? height / size : 0;
}

function applyParagraphMarkFont(paragraph: HTMLElement, mark: HTMLElement, view: Window): void {
  const block = view.getComputedStyle(paragraph),
    marker = view.getComputedStyle(mark),
    properties = ["font-family", "font-size", "font-style", "font-weight"],
    overrides = properties.flatMap((name) => {
      const value = marker.getPropertyValue(name);
      return value && value !== block.getPropertyValue(name) ? [{ name, value }] : [];
    });
  if (!overrides.length) return;

  // The mark defines the line strut. Pin any changed inheritance on its
  // siblings so a mark's font does not alter the authored text runs.
  const siblings = [...paragraph.children].filter(
    (child): child is HTMLElement =>
      child !== mark && child.namespaceURI === paragraph.namespaceURI && Boolean(child.textContent),
  ).map((child) => {
    const css = view.getComputedStyle(child);
    return { child, values: overrides.map(({ name }) => css.getPropertyValue(name)) };
  });
  for (const { name, value } of overrides) {
    ownViewStyle(paragraph, name);
    paragraph.style.setProperty(name, value);
  }
  for (const { child, values } of siblings) {
    const css = view.getComputedStyle(child);
    for (const [index, { name }] of overrides.entries()) {
      const value = values[index]!;
      if (!value || value === css.getPropertyValue(name)) continue;
      ownViewStyle(child, name);
      child.style.setProperty(name, value);
    }
  }
}

export function composeWordAutoLineSpacing(document: Document, scope: ParentNode): void {
  const paragraphs = scope.querySelectorAll<HTMLElement>('[data-word-line-rule="auto"]');
  if (!paragraphs.length) return;

  const view = document.defaultView!,
    context = document.createElement("canvas").getContext("2d"),
    probe = document.createElement("span"),
    heights = new Map<string, number>();
  probe.style.cssText =
    "position:absolute;visibility:hidden;display:block;line-height:normal;white-space:pre";
  probe.textContent = "M";
  const single = (css: CSSStyleDeclaration, eastAsian: boolean): number => {
    const { fontFamily, fontSize, fontWeight, fontStyle } = css,
      size = parseFloat(fontSize),
      key = JSON.stringify([fontFamily, fontSize, fontWeight, fontStyle, eastAsian]),
      known = heights.get(key);
    if (known !== undefined) return known;

    let height: number;
    if (eastAsian) {
      // Word's East Asian single spacing uses the selected font's bounding
      // height with 30% leading. Measure at a larger size to avoid pixel
      // rounding before scaling to the run's authored point size.
      const sampleSize = 256;
      if (context) context.font = `${fontStyle} ${fontWeight} ${sampleSize}px ${fontFamily}`;
      const metrics = context?.measureText("汉"),
        measured = (metrics?.fontBoundingBoxAscent ?? 0) + (metrics?.fontBoundingBoxDescent ?? 0);
      height = (measured > 0 ? measured / sampleSize : 1) * size * 1.3;
    } else {
      if (!probe.isConnected) document.body.append(probe);
      Object.assign(probe.style, { fontFamily, fontSize, fontWeight, fontStyle });
      height = Math.max(probe.getBoundingClientRect().height, size * 1.15);
    }
    heights.set(key, height);
    return height;
  };

  for (const paragraph of paragraphs) {
    const multiple = autoLineMultiplier(paragraph, view);
    if (!Number.isFinite(multiple) || multiple <= 0) continue;
    wordAutoLineMultipliers.set(paragraph, multiple);
    const mark = paragraph.querySelector<HTMLElement>(':scope > [data-word-paragraph-mark="true"]');
    if (mark) applyParagraphMarkFont(paragraph, mark, view);
    const nodes = [paragraph, ...paragraph.querySelectorAll<HTMLElement>("[data-word-script]")];
    for (const node of nodes) {
      const height = single(view.getComputedStyle(node), node.dataset.wordScript === "eastAsia");
      ownViewStyle(node, "line-height");
      node.style.lineHeight = height * multiple + "px";
    }
    // A run containing only line breaks contributes no glyph box. The
    // paragraph strut still supplies the height of a genuinely blank line.
    for (const run of paragraph.querySelectorAll<HTMLElement>("[data-word-run]")) {
      if (
        run.textContent || !run.children.length ||
        ![...run.children].every((child) => child.matches("br:not([data-word-break])"))
      )
        continue;
      ownViewStyle(run, "line-height");
      run.style.lineHeight = "0";
    }
  }
  probe.remove();
}
