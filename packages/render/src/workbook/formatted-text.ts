import { DocumentError } from "@doffi/core";
import type { SourceStyleResolver, WorkbookDisplayFormat, WorkbookNumberFormatter } from "@doffi/core";
import { resolvePreviewFontsInScope } from "../fonts.js";
import type { FontProbeCache } from "../font-probe-cache.js";
import type { WorkbookSourceCell } from "./cell-source.js";
import { createWorkbookTextView } from "./text-view.js";
import type { WorkbookTextProjection, WorkbookTextView } from "./text-view.js";
import { createWorkbookNumberText } from "./number-text.js";
import { layoutWorkbookTexts } from "./text-layout.js";
import type { WorkbookTextLayoutInput, WorkbookTextLayouts } from "./text-layout.js";
import { workbookTextAlignment, workbookTextGutters, workbookReadingOrder } from "./text-alignment.js";
import { WorkbookTextDirection } from "./text-direction.js";
import { fitWorkbookFill } from "./number-layout.js";
import { createSourceStyleResolver } from "../source-style-probe.js";

export interface WorkbookFormattedTextInput {
  readonly source: WorkbookSourceCell;
  readonly value: string;
  readonly format: WorkbookDisplayFormat;
  readonly widthPt: number;
  readonly heightPt: number;
  readonly bordered?: boolean;
  /** Formula text is derived; @ fields cannot claim source text offsets. */
  readonly derived?: boolean;
}

/** A text format derives affixes/spacing/repetition around independently
 * projected @ fields. It neither flattens rich runs nor assigns fake source
 * offsets to literals. Repeated fields retain repeated source ranges. */
export function layoutWorkbookFormattedTexts(scope: HTMLElement, inputs: readonly WorkbookFormattedTextInput[], formatter: WorkbookNumberFormatter,
  normalSpaceWidthPx: number, maxNodes = 50_000, maxTextUnits: number = 8 * 1024 * 1024, probes?: FontProbeCache, suppliedSourceStyles?: () => SourceStyleResolver): WorkbookTextLayouts {
  const document = scope.ownerDocument, view = document.defaultView;
  if (!scope.isConnected || !view || document.fonts.status === "loading" || !Number.isFinite(normalSpaceWidthPx) || normalSpaceWidthPx < 0
    || ![maxNodes, maxTextUnits].every(n => Number.isSafeInteger(n) && n > 0)) throw new DocumentError("InvalidOperation", "Text-format layout requires a connected settled surface and positive budgets.");
  const root = document.createElement("div"); root.dataset.officeView = "workbook-format-measurement"; root.setAttribute("aria-hidden", "true");
  root.style.cssText = "position:absolute;visibility:hidden;pointer-events:none;contain:layout;left:0;top:0;width:0;height:0;overflow:hidden;direction:ltr";
  const direction = new WorkbookTextDirection(scope);
  let resolver: SourceStyleResolver | undefined;
  const sourceStyles = suppliedSourceStyles ?? (() => resolver ??= createSourceStyleResolver(document));
  let nodes = 1, units = 0;
  const remaining = (): number => { if (nodes >= maxNodes) throw new DocumentError("ResourceLimitExceeded", "Formatted text exceeds its cumulative node budget."); return maxNodes - nodes; };
  const accountText = (amount: number): void => { units += amount; if (units > maxTextUnits) throw new DocumentError("ResourceLimitExceeded", "Formatted text exceeds its cumulative text budget."); };
  try {
    const prepared = inputs.map(input => {
      if (input.source.element.ownerDocument !== document || !input.source.element.isConnected || !Number.isFinite(input.widthPt / .75) || input.widthPt <= 0
        || !Number.isFinite(input.heightPt / .75) || input.heightPt <= 0) throw new DocumentError("InvalidOperation", "Text formats require a matching connected source and positive finite cell geometry.");
      const display = formatter.format(input.value, input.format);
      if (display.kind !== "parts") throw new DocumentError("InvalidDocument", "A text value cannot use a numeric overflow display.");
      const base = createWorkbookNumberText(input.source, [], display.color, remaining()); nodes += base.nodes;
      const element = base.element, projections: WorkbookTextProjection[] = [], spaces: HTMLElement[] = [];
      let count = base.nodes, fixedUnits = 0, fill: { element: HTMLElement; character: string; nodes: number } | undefined, lastFill = -1;
      for (const [index, part] of display.parts.entries()) if (part.kind === "fill") lastFill = index;
      for (const [index, part] of display.parts.entries()) {
        if (part.kind === "fill" && index !== lastFill) continue;
        let text: WorkbookTextView;
        if (part.kind === "text" && part.source && !input.derived) {
          if (part.source.from !== 0 || part.source.to !== input.value.length || part.text !== input.value) throw new DocumentError("InvalidDocument", "A text-format field does not match its source value.");
          text = createWorkbookTextView(input.source, remaining(), sourceStyles);
          if (text.projections.map(projection => projection.node.data).join("") !== input.value) throw new DocumentError("InvalidDocument", "A text-format source differs from its stored text value.");
          if (display.color) for (const node of [text.element, ...text.element.querySelectorAll<HTMLElement>("[data-office-text-style],[data-office-script]")]) node.style.color = element.style.color;
          projections.push(...text.projections);
        } else text = createWorkbookNumberText(input.source, [{ kind: "text", text: part.kind === "text" ? part.text : part.character }], display.color, remaining());
        nodes += text.nodes; count += text.nodes; element.append(text.element);
        if (text.element.dataset.officeAccountingText) element.dataset.officeAccountingText = "true";
        if (text.element.dataset.officeUnderlinedText) element.dataset.officeUnderlinedText = "true";
        const length = text.element.textContent?.length ?? 0; accountText(length);
        if (part.kind !== "fill") fixedUnits += length;
        if (part.kind === "space") spaces.push(text.element);
        if (part.kind === "fill") fill = { element: text.element, character: part.character, nodes: text.nodes };
      }
      element.style.cssText += ';display:inline-block;position:absolute;width:max-content;white-space:pre;font-kerning:none;font-feature-settings:"kern" 0;direction:ltr;unicode-bidi:isolate';
      root.append(element);
      return { input, element, projections, spaces, fill, count, fixedUnits };
    });
    scope.append(root);
    const fonts = resolvePreviewFontsInScope(document, root, "[data-office-text-style],[data-office-script]", probes);
    const finals: WorkbookTextLayoutInput[] = [];
    for (const item of prepared) {
      for (const space of item.spaces) {
        const width = space.getBoundingClientRect().width;
        if (!Number.isFinite(width) || width < 0) throw new DocumentError("UnsupportedFeature", "This browser cannot measure text-format spacers.");
        const removed = space.querySelectorAll("*").length + countTextNodes(space); space.replaceChildren(); nodes -= removed; item.count -= removed;
        space.dataset.officeNumberSpace = "true"; space.setAttribute("aria-hidden", "true"); space.style.cssText += `;display:inline-block;width:${width}px`;
      }
      if (item.fill) {
        const { fill, input } = item;
        // The innermost text node may sit inside native underline/strike spans.
        const walker = document.createTreeWalker(fill.element, view.NodeFilter.SHOW_TEXT), node = walker.nextNode() as Text | null;
        if (!node) throw new DocumentError("InvalidDocument", "A text-format fill has no measurable text.");
        node.data = "";
        const style = view.getComputedStyle(input.source.element), order = direction.read(input.value, workbookReadingOrder(style));
        const alignment = workbookTextAlignment(style, "text", order), gutters = workbookTextGutters(alignment, normalSpaceWidthPx);
        const available = Math.max(0, input.widthPt / .75 - gutters.left - gutters.right - item.element.getBoundingClientRect().width);
        const result = fitWorkbookFill(fill.character, available, item.fixedUnits, value => {
          node.data = value; return fill.element.getBoundingClientRect().width;
        });
        node.data = result.text; accountText(Math.max(0, result.text.length - fill.character.length));
      }
      item.element.remove();
      const text = Object.freeze({ element: item.element, projections: Object.freeze(item.projections), nodes: item.count });
      finals.push({ source: item.input.source, text, valueType: "text", widthPt: item.input.widthPt, heightPt: item.input.heightPt,
        ...(item.input.bordered === undefined ? {} : { bordered: item.input.bordered }) });
    }
    const layout = layoutWorkbookTexts(scope, finals, normalSpaceWidthPx, maxNodes, maxTextUnits, probes);
    return Object.freeze({ ...layout, fonts });
  } finally { root.remove(); root.replaceChildren(); direction.dispose(); }
}

function countTextNodes(root: HTMLElement): number {
  const walker = root.ownerDocument.createTreeWalker(root, root.ownerDocument.defaultView!.NodeFilter.SHOW_TEXT); let count = 0;
  while (walker.nextNode()) count++; return count;
}
