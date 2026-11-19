import { DocumentError } from "@doffi/core";
import type { DeepReadonly, WordMetadata } from "@doffi/core";
import { resolvePreviewFonts } from "./fonts.js";
import type { HostedWordFonts } from "./fonts.js";
import type { FontResolution } from "./fonts.js";
import { markViewNode, ownViewStyle } from "./view-nodes.js";

const pixelsPerPoint = 4 / 3;

function numeral(value: number, format: string): string {
  if (format === "none") return "";
  if (format === "upperLetter" || format === "lowerLetter") {
    let result = "";
    for (let n = value; n > 0; n = Math.floor((n - 1) / 26))
      result = String.fromCharCode(65 + ((n - 1) % 26)) + result;
    return format === "lowerLetter" ? result.toLowerCase() : result;
  }
  if (format === "upperRoman" || format === "lowerRoman") {
    if (value > 1_000_000)
      throw new DocumentError(
        "ResourceLimitExceeded",
        "A derived Roman numbering label exceeds the preview budget.",
      );
    let result = "",
      n = value;
    for (const [amount, letters] of [
      [1000, "M"],
      [900, "CM"],
      [500, "D"],
      [400, "CD"],
      [100, "C"],
      [90, "XC"],
      [50, "L"],
      [40, "XL"],
      [10, "X"],
      [9, "IX"],
      [5, "V"],
      [4, "IV"],
      [1, "I"],
    ] as const) {
      while (n >= amount) {
        result += letters;
        n -= amount;
      }
    }
    return format === "lowerRoman" ? result.toLowerCase() : result;
  }
  return String(value);
}

export function composeWordNumbering(
  document: Document,
  word: DeepReadonly<WordMetadata> | undefined,
  hostedFonts: HostedWordFonts | undefined,
): readonly FontResolution[] {
  if (!word?.numbering.length) return [];
  const view = document.defaultView!,
    definitions = new Map(word.numbering.map((item) => [item.numberingId, item])),
    tabs = new Map(word.tabs.map((item) => [item.paragraphNodeId, item])),
    entries: { paragraph: HTMLElement; marker: HTMLElement; suffix: "tab" | "space" | "nothing" }[] = [];
  for (const story of word.stories) {
    const root = document.querySelector('[data-node-id="' + CSS.escape(story.rootNodeId) + '"]');
    if (!root) continue;
    const counters = new Map<string, number[]>();
    for (const paragraph of root.querySelectorAll<HTMLElement>("[data-word-numbering]")) {
      const definition = definitions.get(paragraph.dataset.wordNumbering ?? "");
      const level = Number(paragraph.dataset.wordLevel);
      const rule = definition?.levels.find((item) => item.level === level);
      if (!definition || !rule) continue;
      const values = counters.get(definition.numberingId) ?? [];
      values[level] = (values[level] ?? rule.start - 1) + 1;
      for (const deeper of definition.levels)
        if (
          deeper.level > level &&
          deeper.restart !== 0 &&
          (deeper.restart === undefined || deeper.restart - 1 >= level)
        )
          values[deeper.level] = deeper.start - 1;
      counters.set(definition.numberingId, values);
      const marker = markViewNode(document.createElement("span"));
      marker.className = "sdk-number-marker " + rule.cssClass;
      marker.setAttribute("aria-hidden", "true");
      marker.textContent = rule.text.replace(/%([1-9])/g, (_match, token: string) => {
        const index = Number(token) - 1;
        const format = definition.levels.find((item) => item.level === index);
        return numeral(values[index] ?? format?.start ?? 1, format?.format ?? "decimal");
      });
      if (rule.suffix === "space") marker.textContent += " ";
      ownViewStyle(paragraph, "position", "text-indent");
      paragraph.style.position = "relative";
      marker.style.cssText =
        "position:absolute;width:max-content;inline-size:max-content;margin:0;padding:0;text-indent:0;white-space:pre";
      paragraph.prepend(marker);
      entries.push({ paragraph, marker, suffix: rule.suffix });
    }
  }
  const fonts = resolvePreviewFonts(document, ".sdk-number-marker", hostedFonts);
  const offsets = entries.map(({ paragraph, marker, suffix }) => {
    const css = view.getComputedStyle(paragraph),
      start = parseFloat(css.textIndent) || 0,
      margin = parseFloat(css.marginInlineStart) || 0,
      rect = marker.getBoundingClientRect(),
      width = css.writingMode === "horizontal-tb" ? rect.width : rect.height,
      end = margin + start + width;
    if (suffix !== "tab" || !marker.textContent)
      return { start, textIndent: start + width };
    const tab = tabs.get(paragraph.dataset.nodeId ?? ""),
      interval = (tab?.defaultInterval ?? word.tabs[0]?.defaultInterval ?? 36) * pixelsPerPoint,
      explicit = tab?.stops.find(
        (stop) => stop.alignment !== "clear" && stop.position * pixelsPerPoint > end,
      ),
      next = explicit
        ? explicit.position * pixelsPerPoint
        : start < 0 && end <= margin
          ? margin
          : (Math.floor(end / interval) + 1) * interval;
    return { start, textIndent: next - margin };
  });
  entries.forEach(({ paragraph, marker }, index) => {
    const { start, textIndent } = offsets[index]!;
    marker.style.insetInlineStart = start + "px";
    paragraph.style.textIndent = textIndent + "px";
  });
  return fonts;
}
