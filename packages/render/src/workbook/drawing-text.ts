import { DocumentError } from "@doffi/core";
import type { DeepReadonly, DrawingTextProperties, WorkbookChart, WorkbookObject } from "@doffi/core";
import { markViewNode } from "../view-nodes.js";
import { workbookRichTextFonts } from "./text-fonts.js";
import { createSourceStyleResolver, sourceStyleProjection } from "../source-style-probe.js";

export interface WorkbookDrawingTextProjection {
  readonly node: Text;
  readonly sourceNodeId: string;
  /** UTF-16 offsets in the whole text body; paragraph boundaries add no units. */
  readonly from: number;
  readonly to: number;
}
export interface WorkbookDrawingTextView {
  readonly element: HTMLElement;
  readonly projections: readonly WorkbookDrawingTextProjection[];
  readonly nodes: number;
}

const inline = new Set(["span", "strong", "em", "u", "s", "sub", "sup", "br"]);
const typography = ["font-family", "font-size", "font-weight", "font-style", "font-variant", "line-height", "letter-spacing", "word-spacing", "color", "text-decoration-line", "text-decoration-style", "text-decoration-color", "text-transform", "vertical-align", "direction", "unicode-bidi", "tab-size"];
const paragraphStyle = ["text-align", "text-indent", "margin-left", "margin-right", "margin-top", "margin-bottom"];
const listAttributes = ["data-office-bullet", "data-office-numbering", "data-office-numbering-start", "data-office-level", "data-office-list-none"];

/** Materialize typography while the source still has its validated ancestry.
 * The detached result owns no source IDs/classes and carries no executable or
 * resource attributes. Layout, lists, local-font diagnostics and pane clipping
 * belong to its compositor; no source node is moved or changed here. */
export function createWorkbookDrawingText(source: HTMLElement, object: DeepReadonly<WorkbookObject>, maxNodes = 50_000): WorkbookDrawingTextView {
  if (object.kind !== "text" && object.kind !== "geometry" || source.dataset.nodeId !== object.nodeId || source.dataset.officeKind !== object.kind) throw new DocumentError("InvalidDocument", "Drawing text requires its own shape or text-box source.");
  return createTextBody(source, object.nodeId, maxNodes);
}

/** Rich chart titles retain their own text-body identity and exact UTF-16
 * projection. Cell-linked titles instead use the chart reference data adapter. */
export type WorkbookChartTextRole = "title" | "category-axis-title" | "value-axis-title";
export function createWorkbookChartText(chartSource: HTMLElement, chart: DeepReadonly<WorkbookChart>, role: WorkbookChartTextRole, maxNodes = 50_000): WorkbookDrawingTextView {
  const { source } = workbookChartTextSource(chartSource, chart, role);
  return createTextBody(source, source.dataset.nodeId!, maxNodes);
}

export function workbookChartTextSource(chartSource: HTMLElement, chart: DeepReadonly<WorkbookChart>, role: WorkbookChartTextRole): { readonly source: HTMLElement; readonly properties: DeepReadonly<DrawingTextProperties> | undefined } {
  if (chartSource.dataset.nodeId !== chart.nodeId || chartSource.dataset.officeKind !== "chart") throw new DocumentError("InvalidDocument", "Chart text requires its own source chart.");
  const definition = role === "title" ? { nodeId: chart.titleNodeId, text: chart.titleText } : role === "category-axis-title" ? chart.categoryAxisTitle : role === "value-axis-title" ? chart.valueAxisTitle : undefined;
  const nodeId = definition?.nodeId;
  if (!nodeId) throw new DocumentError("InvalidOperation", "This chart title has no rich-text source body.");
  const source = [...chartSource.children].find(child => (child as HTMLElement).dataset.nodeId === nodeId) as HTMLElement | undefined;
  if (!source) throw new DocumentError("InvalidDocument", "A rich chart title must belong directly to its source chart.");
  return { source, properties: definition?.text };
}

/** A private derived source can share typography without claiming that its
 * display characters are canonical document text. */
export function createWorkbookDerivedDrawingText(source: HTMLElement, sourceNodeId: string, maxNodes = 50_000): WorkbookDrawingTextView {
  return createTextBody(source, sourceNodeId, maxNodes, false);
}

function createTextBody(source: HTMLElement, nodeId: string, maxNodes: number, project = true): WorkbookDrawingTextView {
  const document = source.ownerDocument, view = document.defaultView;
  if (!source.isConnected || !view || !Number.isSafeInteger(maxNodes) || maxNodes < 1) throw new DocumentError("InvalidOperation", "Drawing text requires a connected source and a positive node budget.");
  const root = markViewNode(document.createElement("div")), projections: WorkbookDrawingTextProjection[] = [];
  // Explicit Office run resets require per-run paint. Browser CSS cannot
  // cancel an ancestor's decoration with a child's `none`; keep the original
  // compositor for documents without the source override.
  const reset = "--office-font-decoration-reset", elements = [source, ...source.querySelectorAll<HTMLElement>("*")];
  if (elements.length > maxNodes) throw new DocumentError("ResourceLimitExceeded", "Drawing text exceeds its source node budget.");
  const explicit = elements.some(element => { const value = view.getComputedStyle(element).getPropertyValue(reset).trim().toLowerCase(); return value !== "" && value !== "none"; });
  const resolver = explicit ? createSourceStyleResolver(document) : undefined;
  const decoration = new WeakMap<HTMLElement, Readonly<Record<string, string>>>();
  const effective = (original: HTMLElement): Readonly<Record<string, string>> => {
    let value = decoration.get(original);
    if (!value) { value = resolver!.get(sourceStyleProjection(original, true)); decoration.set(original, value); }
    return value;
  };
  root.dataset.officeView = "workbook-drawing-text"; root.dataset.sourceNodeId = nodeId;
  let nodes = 1, offset = 0;
  const account = (amount: number): void => { nodes += amount; if (nodes > maxNodes) throw new DocumentError("ResourceLimitExceeded", "Drawing text exceeds its node budget."); };
  const style = (original: HTMLElement, target: HTMLElement): CSSStyleDeclaration => {
    const computed = view.getComputedStyle(original); target.dataset.officeTextStyle = "true";
    for (const name of typography) target.style.setProperty(name, computed.getPropertyValue(name));
    if (explicit) target.style.textDecorationLine = "none";
    if (original !== source) target.style.backgroundColor = computed.backgroundColor;
    if (original.lang) target.lang = original.lang;
    if (original.dataset.nodeId) target.dataset.sourceNodeId = original.dataset.nodeId;
    return computed;
  };
  style(source, root); root.style.whiteSpace = "pre-wrap";
  for (const child of source.childNodes) {
    if (child.nodeType === 3 && !child.textContent?.trim()) { offset += child.textContent?.length ?? 0; continue; }
    if (child.nodeType !== 1 || (child as Element).localName !== "p") throw new DocumentError("InvalidDocument", "Drawing text must consist of source paragraphs.");
    account(1);
    const paragraph = child as HTMLElement, runs = workbookRichTextFonts(paragraph, Math.max(1, maxNodes - nodes)), baseOffset = offset;
    let runIndex = 0, localOffset = 0, lineHasText = false, insertionStyle = paragraph;
    const target = document.createElement("p"), computed = style(paragraph, target);
    target.style.margin = "0"; target.style.minHeight = "1em"; target.style.flex = "none";
    for (const name of paragraphStyle) target.style.setProperty(name, computed.getPropertyValue(name));
    for (const name of listAttributes) { const value = paragraph.getAttribute(name); if (value !== null) target.setAttribute(name, value); }
    for (const name of computed) if (name.startsWith("--office-list-") || name.startsWith("--office-bullet-")) target.style.setProperty(name, computed.getPropertyValue(name));
    const emptyLine = (parent: HTMLElement, end: boolean): void => {
      // Empty lines carry the preceding break's insertion style, overridden
      // by endParaRPr only at the paragraph end. This view-only character never
      // becomes a source UTF-16 unit. The temporary inheritance probe is private
      // to the already projected source and is removed on every exit path.
      account(1); const placeholder = markViewNode(document.createElement("span"));
      placeholder.dataset.officeView = "workbook-drawing-empty-line"; placeholder.setAttribute("aria-hidden", "true");
      account(1); const probe = document.createElement("span"), base = style(insertionStyle, probe);
      for (const name of base) if (name.startsWith("--office-font-") || name.startsWith("--office-major-") || name.startsWith("--office-minor-")) probe.style.setProperty(name, base.getPropertyValue(name));
      placeholder.className = end ? paragraph.dataset.officeEndStyle ?? "" : ""; probe.append(placeholder); paragraph.append(probe);
      try {
        const computed = style(placeholder, placeholder);
        placeholder.style.fontFamily = computed.getPropertyValue("--office-font-ascii").trim() || computed.getPropertyValue("--office-minor-ascii").trim() || computed.fontFamily;
      } finally { probe.remove(); nodes--; }
      placeholder.removeAttribute("class"); account(1); placeholder.append(document.createTextNode("\u200b")); parent.append(placeholder);
    };
    const appendText = (text: string, original: HTMLElement, parent: HTMLElement): void => {
      const end = localOffset + text.length;
      const computed = view.getComputedStyle(original);
      while (localOffset < end) {
        const run = runs[runIndex++];
        if (!run || run.from !== localOffset || run.to > end) throw new DocumentError("InvalidDocument", "Drawing font runs do not match the source text.");
        account(2); const span = document.createElement("span"), node = document.createTextNode(run.text);
        span.dataset.officeScript = run.role; span.style.fontFamily = run.font.family; span.style.fontSize = run.font.sizePx + "px"; span.style.lineHeight = computed.lineHeight; span.append(node); parent.append(span);
        if (explicit) {
          const paint = effective(original);
          for (const name of ["text-decoration-line", "text-decoration-style", "text-decoration-color"]) span.style.setProperty(name, paint[name]!);
        }
        if (project) projections.push(Object.freeze({ node, sourceNodeId: original.dataset.nodeId!, from: baseOffset + run.from, to: baseOffset + run.to })); localOffset = run.to;
      }
      if (text.length) { lineHasText = text.includes("\n") ? !text.endsWith("\n") : true; insertionStyle = original; }
    };
    const append = (original: HTMLElement, parent: HTMLElement, depth: number): void => {
      if (depth > 64) throw new DocumentError("ResourceLimitExceeded", "Drawing text exceeds its nesting budget.");
      for (const child of original.childNodes) {
        if (child.nodeType === 3) appendText(child.textContent ?? "", original, parent);
        else if (child.nodeType === 1 && inline.has((child as Element).localName)) {
          const element = child as HTMLElement;
          if (element.localName === "br" && !lineHasText) { if (!localOffset) insertionStyle = element; emptyLine(parent, false); }
          account(1); const span = document.createElement("span"); style(element, span); parent.append(span);
          if (element.localName === "br") { span.dataset.officeTextBreak = "true"; appendText("\n", element, span); }
          else append(element, span, depth + 1);
        } else throw new DocumentError("InvalidDocument", "Drawing text contains unsupported inline content.");
      }
    };
    append(paragraph, target, 0);
    if (runIndex !== runs.length) throw new DocumentError("InvalidDocument", "Drawing text did not consume all source font runs.");
    if (!lineHasText) emptyLine(target, true);
    offset += localOffset; root.append(target);
  }
  return Object.freeze({ element: root, projections: Object.freeze(projections), nodes });
}
