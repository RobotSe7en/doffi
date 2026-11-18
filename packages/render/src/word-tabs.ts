import type { DeepReadonly, WordMetadata } from "@doffi/core";
import { markViewNode, ownViewStyle } from "./view-nodes.js";

const tabSelector = "[data-word-tab]";
const paragraphSelector = "p,h1,h2,h3,h4,h5,h6,li";
const pixelsPerPoint = 4 / 3;
const leaderText = { dot: ".", hyphen: "-", middleDot: "·", underscore: "_", heavy: "_" };
// Leave less than a visible pixel for rounding at a shared inline edge.
const roundoff = 1 / 32;

export interface WordTabLayout {
  measure(root: ParentNode): void;
}

function prepareTab(tab: HTMLElement): void {
  ownViewStyle(
    tab, "display", "width", "inline-size", "white-space", "overflow",
    "vertical-align", "text-indent", "text-align",
  );
  tab.textContent = "";
  tab.style.display = "inline-block";
  tab.style.width = "";
  tab.style.inlineSize = "0px";
  tab.style.whiteSpace = "nowrap";
  tab.style.overflow = "hidden";
  tab.style.verticalAlign = "bottom";
  tab.style.textIndent = "0";
  tab.style.textAlign = "start";
}

/** Measure the following content with the same inline styles, shaping and
 * font fallback as the document. The probe never participates in line breaking. */
function createTabProbe(
  paragraph: HTMLElement, tab: HTMLElement, next: HTMLElement | undefined,
  css: CSSStyleDeclaration, range: Range,
): HTMLElement | undefined {
  range.setStartAfter(tab);
  if (next) range.setEndBefore(next);
  else range.setEnd(paragraph, paragraph.childNodes.length);
  if (range.collapsed) return;
  const probe = markViewNode(paragraph.cloneNode(false) as HTMLElement);
  probe.removeAttribute("id");
  const whitespace = /^(normal|nowrap|pre-line)$/.test(css.whiteSpace) ? "nowrap" : "pre";
  probe.style.cssText +=
    ";position:absolute!important;visibility:hidden!important;pointer-events:none;" +
    "inline-size:max-content!important;min-inline-size:0!important;max-inline-size:none!important;" +
    "margin:0!important;padding:0!important;border:0!important;text-indent:0!important;" +
    "text-align:start!important;text-align-last:start!important;white-space:" + whitespace + "!important";
  probe.append(range.cloneContents());
  paragraph.after(probe);
  return probe;
}

function decimalOffset(
  probe: HTMLElement, range: Range, start: (rect: DOMRect) => number, content: DOMRect,
): number | undefined {
  const walker = probe.ownerDocument.createTreeWalker(probe, 4);
  for (let node = walker.nextNode(); node; node = walker.nextNode()) {
    const offset = node.textContent!.search(/[.,]/);
    if (offset < 0) continue;
    let owner = node.parentElement;
    for (; owner && owner !== probe; owner = owner.parentElement)
      if (owner.hidden || owner.style.visibility === "hidden") break;
    if (owner !== probe) continue;
    range.setStart(node, offset);
    range.setEnd(node, offset + 1);
    if (range.getClientRects().length)
      return start(range.getBoundingClientRect()) - start(content);
  }
  return undefined;
}

export function wordTabLayout(document: Document, word: DeepReadonly<WordMetadata>): WordTabLayout {
  const view = document.defaultView!,
    definitions = new Map(word.tabs.map((definition) => [definition.paragraphNodeId, definition])),
    range = document.createRange();

  function measure(root: ParentNode): void {
    const paragraphs = new Set(
      Array.from(root.querySelectorAll(tabSelector), (tab) => tab.closest<HTMLElement>(paragraphSelector)),
    );
    const layouts = [...paragraphs].flatMap((paragraph) => {
      if (!paragraph || !paragraph.getClientRects().length) return [];
      const definition = definitions.get(paragraph.dataset.nodeId ?? "");
      if (!definition) return [];
      const css = view.getComputedStyle(paragraph),
        vertical = /^(vertical|sideways)/.test(css.writingMode),
        reversed = css.direction === "rtl",
        edge = vertical ? (reversed ? "bottom" : "top") : reversed ? "right" : "left",
        start = (rect: DOMRect) => (reversed ? -1 : 1) * rect[edge],
        extent = (rect: DOMRect) => vertical ? rect.height : rect.width,
        stops = definition.stops.filter((stop) => stop.alignment !== "clear")
          .sort((left, right) => left.position - right.position),
        boundaries = [...paragraph.querySelectorAll<HTMLElement>(tabSelector + ",br")].filter(
          (node) => node.closest(paragraphSelector) === paragraph &&
            node.getClientRects().length && view.getComputedStyle(node).visibility === "visible",
        ),
        tabs = boundaries.flatMap((tab, index) =>
          tab.matches(tabSelector) ? [{ tab, next: boundaries[index + 1] }] : [],
        );
      return tabs.length ? [{ paragraph, css, start, extent, definition, stops, tabs }] : [];
    });
    for (const layout of layouts)
      for (const { tab } of layout.tabs) prepareTab(tab);

    const count = layouts.reduce((max, layout) => Math.max(max, layout.tabs.length), 0);
    for (let index = 0; index < count; index++) {
      // The next tab in a paragraph depends on the previous one. Batch the
      // reads and writes of independent paragraphs to avoid serial reflows.
      const placements = layouts.flatMap((layout) => {
        const item = layout.tabs[index];
        if (!item) return [];
        const { paragraph, css, start, extent, stops, definition } = layout,
          box = paragraph.getBoundingClientRect(),
          // The paragraph's indent moves the text, but does not move its tab stops.
          origin = start(box) - (parseFloat(css.marginInlineStart) || 0),
          cursor = start(item.tab.getBoundingClientRect()) - origin,
          end = start(box) + extent(box) - origin -
            (parseFloat(css.paddingInlineEnd) || 0) - (parseFloat(css.borderInlineEndWidth) || 0),
          stop = stops.find((entry) => entry.position * pixelsPerPoint > cursor),
          interval = definition.defaultInterval * pixelsPerPoint,
          target = stop ? stop.position * pixelsPerPoint : (Math.floor(cursor / interval) + 1) * interval,
          alignedEnd = Math.abs(target - end) < roundoff ? Math.min(target, end) : target;
        return [{ ...item, layout, available: alignedEnd - cursor, stop }];
      });
      for (let offset = 0; offset < placements.length; offset += 64) {
        const batch = placements.slice(offset, offset + 64);
        const probes: (HTMLElement | undefined)[] = [];
        let widths: number[];
        try {
          for (const { tab, next, layout, stop } of batch) {
            probes.push(stop && ["right", "center", "decimal"].includes(stop.alignment)
              ? createTabProbe(layout.paragraph, tab, next, layout.css, range)
              : undefined);
          }
          widths = batch.map(({ layout, stop, available }, item) => {
            const probe = probes[item];
            let adjustment = 0;
            if (probe) {
              const content = probe.getBoundingClientRect();
              adjustment = layout.extent(content);
              if (stop?.alignment === "center") adjustment /= 2;
              if (stop?.alignment === "decimal")
                adjustment = decimalOffset(probe, range, layout.start, content) ?? adjustment;
            }
            return Math.max(0, available - adjustment - roundoff);
          });
        } finally {
          for (const probe of probes) probe?.remove();
        }
        batch.forEach(({ tab, stop }, item) => {
          const width = widths[item]!;
          tab.style.inlineSize = width + "px";
          const leader = stop?.leader;
          if (leader && leader !== "none")
            tab.textContent = leaderText[leader].repeat(Math.min(2000, Math.ceil(width / pixelsPerPoint)));
        });
      }
    }
  }
  return { measure };
}
