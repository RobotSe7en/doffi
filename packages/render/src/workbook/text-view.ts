import { DocumentError } from "@doffi/core";
import type { SourceStyleResolver } from "@doffi/core";
import { markViewNode } from "../view-nodes.js";
import type { WorkbookSourceCell } from "./cell-source.js";
import { workbookTextFonts } from "./text-fonts.js";
import { copyWorkbookTextStyle, decorateWorkbookText } from "./text-style.js";
import { createSourceStyleResolver, sourceStyleProjection } from "../source-style-probe.js";

export interface WorkbookTextProjection { readonly node: Text; readonly from: number; readonly to: number }
export interface WorkbookTextView {
  readonly element: HTMLElement;
  /** Empty for derived scalar displays: those are anchored to a cell, not to
   * the UTF-16 offsets of an imported display cache. */
  readonly projections: readonly WorkbookTextProjection[];
  readonly nodes: number;
}

/** Build display-only rich text from computed source styles, before any virtual
 * grid ancestry can change their meaning. Keep source span nesting for CSS
 * decoration propagation; native font decorations belong to individual runs so
 * a rich-text run can explicitly turn off its cell's underline/strike. */
export function createWorkbookTextView(source: WorkbookSourceCell, maxNodes = 50_000, sourceStyles?: () => SourceStyleResolver): WorkbookTextView {
  const { element } = source, document = element.ownerDocument, view = document.defaultView;
  if (!view || !element.isConnected || !Number.isSafeInteger(maxNodes) || maxNodes < 1) throw new DocumentError("InvalidOperation", "Workbook text display requires a connected source and a positive node budget.");
  const runs = workbookTextFonts(source, maxNodes), root = markViewNode(document.createElement("span")), projections: WorkbookTextProjection[] = [];
  const resets = [element, ...element.querySelectorAll<HTMLElement>("span")].some(node => {
    const value = view.getComputedStyle(node).getPropertyValue("--office-font-decoration-reset").trim(); return !!value && value !== "none";
  });
  const resolver = resets ? sourceStyles?.() ?? createSourceStyleResolver(document) : undefined;
  let nodes = 1, offset = 0, runIndex = 0;
  const account = (amount: number): void => { nodes += amount; if (nodes > maxNodes) throw new DocumentError("ResourceLimitExceeded", "Workbook text display exceeds its node budget."); };
  const append = (original: HTMLElement, target: HTMLElement, depth: number): void => {
    if (depth > 16) throw new DocumentError("ResourceLimitExceeded", "Workbook text display exceeds its nesting budget.");
    const computed = copyWorkbookTextStyle(original, target, original !== element);
    const paint = resolver?.get(sourceStyleProjection(original, true));
    if (resolver) target.style.textDecorationLine = "none";
    for (const child of original.childNodes) {
      if (child.nodeType === 3) {
        const text = child.textContent ?? "", end = offset + text.length;
        while (offset < end) {
          const run = runs[runIndex++];
          if (!run || run.from !== offset || run.to > end) throw new DocumentError("InvalidDocument", "Worksheet font runs do not match their source text.");
          account(2); const span = document.createElement("span"), node = document.createTextNode(run.text);
          span.dataset.officeScript = run.role; span.style.fontFamily = run.font.family;
          const carrier = decorateWorkbookText(computed, span, account);
          if (paint) {
            const wrapper = document.createElement("span"); account(1);
            for (const property of ["text-decoration-line", "text-decoration-style", "text-decoration-color"]) wrapper.style.setProperty(property, paint[property]!);
            wrapper.append(node); carrier.append(wrapper);
          } else carrier.append(node);
          target.append(span);
          if (span.dataset.officeAccountingUnderline) root.dataset.officeAccountingText = "true";
          if (span.dataset.officeNativeUnderline) root.dataset.officeUnderlinedText = "true";
          projections.push(Object.freeze({ node, from: run.from, to: run.to })); offset = run.to;
        }
      } else if (child.nodeType === 1 && (child as Element).localName === "br") {
        const run = runs[runIndex++];
        if (!run || run.from !== offset || run.text !== "\n") throw new DocumentError("InvalidDocument", "Worksheet line breaks do not match their source text.");
        // A literal newline keeps one UTF-16 source unit and can participate in
        // DOM ranges. White-space handling belongs to the cell layout owner.
        account(1); const node = document.createTextNode("\n"); target.append(node);
        projections.push(Object.freeze({ node, from: offset, to: ++offset }));
      } else if (child.nodeType === 1 && (child as Element).localName === "span") {
        account(1); const span = document.createElement("span"); append(child as HTMLElement, span, depth + 1); target.append(span);
      } else throw new DocumentError("InvalidDocument", "Worksheet display text contains unsupported source content.");
    }
  };
  append(element, root, 0); root.style.whiteSpace = "pre";
  if (runIndex !== runs.length) throw new DocumentError("InvalidDocument", "Worksheet display did not consume all font runs.");
  return Object.freeze({ element: root, projections: Object.freeze(projections), nodes });
}
