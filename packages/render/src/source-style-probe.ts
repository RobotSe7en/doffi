import { DEFAULT_DOCUMENT_READ_LIMITS, DocumentError, SourceStyleResolver } from "@doffi/core";
import type { SourceStyleNode } from "@doffi/core";

const decorationKeys = ["text-decoration", "text-decoration-line", "text-decoration-style", "text-decoration-color", "--office-font-decoration-reset"] as const;
export function createSourceStyleResolver(document: Document): SourceStyleResolver {
  const sheets = [...document.head.querySelectorAll<HTMLStyleElement>("style[data-office-source-styles],style[data-sdk-source-styles]")], css = sheets.map(sheet => sheet.textContent ?? "");
  if (!sheets.length) throw new DocumentError("InvalidOperation", "Text styles require their owning source styles.");
  // These styles have already passed their package/surface admission budgets.
  // Bound this second, mask-only parse by those captured input strings, rather
  // than imposing a smaller package limit on a host's admitted source.
  const units = css.reduce((total, text) => total + text.length, 0), resolver = new SourceStyleResolver(css, {
    ...DEFAULT_DOCUMENT_READ_LIMITS, maxFiles: Math.max(DEFAULT_DOCUMENT_READ_LIMITS.maxFiles, sheets.length),
    maxCssNodes: Math.max(DEFAULT_DOCUMENT_READ_LIMITS.maxCssNodes, units * 2), maxDepth: Math.max(DEFAULT_DOCUMENT_READ_LIMITS.maxDepth, units),
    maxFileBytes: Math.max(DEFAULT_DOCUMENT_READ_LIMITS.maxFileBytes, units * 4), maxPackageBytes: Math.max(DEFAULT_DOCUMENT_READ_LIMITS.maxPackageBytes, units * 4),
  });
  return resolver;
}
export function sourceStyleProjection(element: HTMLElement, includeFontValues = false): SourceStyleNode {
  let parent: SourceStyleNode | undefined; const ancestors: HTMLElement[] = [];
  for (let current: HTMLElement | null = element; current; current = current.parentElement) ancestors.push(current);
  for (const current of ancestors.reverse()) {
    const attrs = Object.fromEntries([...current.attributes].filter(attribute => attribute.name !== "style").map(attribute => [attribute.name, attribute.value]));
    // Exclude private layout positioning. Role queries inspect decoration
    // presence; full text queries also retain inline color, size and variables.
    const keys: readonly string[] = includeFontValues ? [...decorationKeys, "color", "font-size", ...[...current.style].filter(name => name.startsWith("--"))] : decorationKeys;
    attrs.style = keys.flatMap(name => { const value = current.style.getPropertyValue(name); return value ? [name + ":" + value + (current.style.getPropertyPriority(name) ? "!important" : "")] : []; }).join(";");
    parent = { tag: current.localName, attrs, parent };
  }
  return parent!;
}
