import { DocumentError, evaluatePresentationPreset } from "@doffi/core";
import { measuredFontLineHeight } from "../font-line-height.js";
import { applyDrawingTextOverflow } from "./drawing-text-overflow.js";
import type { DeepReadonly, DrawingTextProperties, PresentationPresetGeometry, WorkbookChart, WorkbookObject } from "@doffi/core";
import type { FontResolution } from "../contracts.js";
import { resolvePreviewFontsInScope } from "../fonts.js";
import { markViewNode } from "../view-nodes.js";
import type { FontProbeCache } from "../font-probe-cache.js";
import { property } from "../presentation/appearance.js";
import { decorateDrawingLists } from "../presentation/text.js";
import { workbookDrawingOrientation } from "./drawing-view.js";
import { createWorkbookChartText, createWorkbookDerivedDrawingText, workbookChartTextSource } from "./drawing-text.js";
import type { WorkbookChartTextRole, WorkbookDrawingTextView } from "./drawing-text.js";
import type { WorkbookRect } from "./projection.js";

export interface WorkbookDrawingTextLayoutInput {
  readonly source: HTMLElement; readonly object: DeepReadonly<WorkbookObject>; readonly text: WorkbookDrawingTextView;
  readonly widthPt: number; readonly heightPt: number;
}
export interface WorkbookDrawingTextLayout {
  readonly element: HTMLElement; readonly text: WorkbookDrawingTextView; readonly widthPt: number; readonly heightPt: number;
  readonly geometry: PresentationPresetGeometry; readonly scale: number; readonly overflow: boolean;
  /** Measured local text bounds, including source reflection and overflow. */
  readonly paintBounds: WorkbookRect;
  readonly nodes: number;
}
export interface WorkbookDrawingTextLimits { readonly maxNodes?: number; readonly maxPasses?: number; readonly maxExtentPt?: number }
export interface WorkbookDrawingTextBatch {
  readonly objects: readonly WorkbookDrawingTextLayout[]; readonly fonts: readonly FontResolution[]; readonly usedFallbackLineMetrics: boolean;
}
export interface WorkbookChartTextLayoutInput {
  readonly source: HTMLElement; readonly chart: DeepReadonly<WorkbookChart>; readonly role: WorkbookChartTextRole; readonly text: WorkbookDrawingTextView;
  readonly widthPt: number; readonly heightPt: number;
}
export interface WorkbookChartTitleContentInput {
  readonly source: HTMLElement; readonly chart: DeepReadonly<WorkbookChart>; readonly role: WorkbookChartTextRole;
  readonly widthPt: number; readonly heightPt: number;
  /** Current formatted scalar for a reference title; absent for rich text.
   * Imported chart cache text is never a fallback. */
  readonly value?: string;
}
interface TextBodyLayoutInput {
  readonly source: HTMLElement; readonly text: WorkbookDrawingTextView; readonly widthPt: number; readonly heightPt: number;
  readonly properties: DeepReadonly<DrawingTextProperties> | undefined;
  readonly geometry: (width: number, height: number) => PresentationPresetGeometry;
  readonly flipH: boolean; readonly flipV: boolean;
}

/** Resolve local fonts and measure complete, unrotated object rectangles before
 * their pane projections are created. Text counters its object's reflections;
 * the outer object still owns rotation, flips, fill/effects and clipping. */
export function layoutWorkbookDrawingTexts(scope: HTMLElement, inputs: readonly WorkbookDrawingTextLayoutInput[], limits: WorkbookDrawingTextLimits = {}, probes?: FontProbeCache): WorkbookDrawingTextBatch {
  return layoutTextBodies(scope, inputs, limits, input => {
    const { object, source } = input;
    if (object.kind !== "text" && object.kind !== "geometry" || source.dataset.nodeId !== object.nodeId || source.dataset.officeKind !== object.kind) throw new DocumentError("InvalidOperation", "Drawing layout requires its own shape or text-box source.");
    const orientation = workbookDrawingOrientation(source), adjustments: Record<string, number> = {};
    for (const attribute of source.attributes) if (attribute.name.startsWith("data-office-adjust-")) adjustments[attribute.name.slice(19)] = Number(attribute.value);
    const preset = object.kind === "geometry" ? object.preset : "rect";
    return { source, text: input.text, widthPt: input.widthPt, heightPt: input.heightPt, properties: object.text,
      geometry: (width: number, height: number) => evaluatePresentationPreset(preset, width, height, adjustments), flipH: orientation.flipH, flipV: orientation.flipV };
  }, probes);
}

/** Measure each rich title in a compositor-supplied local rectangle. The chart
 * compositor owns title placement, axis-title rotation and surrounding graphics;
 * the title keeps its real source identity and DrawingML text-body properties. */
export function layoutWorkbookChartTexts(scope: HTMLElement, inputs: readonly WorkbookChartTextLayoutInput[], limits: WorkbookDrawingTextLimits = {}): WorkbookDrawingTextBatch {
  return layoutTextBodies(scope, inputs, limits, input => {
    const { source, properties } = workbookChartTextSource(input.source, input.chart, input.role);
    return { source, properties, text: input.text, widthPt: input.widthPt, heightPt: input.heightPt,
      geometry: (width: number, height: number) => evaluatePresentationPreset("rect", width, height), flipH: false, flipV: false };
  });
}

/** Compose rich and current-reference titles in one font/layout batch. Private
 * style probes have no canonical node IDs and are released on every exit.
 * A reference display owns no source UTF-16 offsets, even if its characters
 * happen to equal an imported title cache. */
export function layoutWorkbookChartTitleContents(scope: HTMLElement, inputs: readonly WorkbookChartTitleContentInput[],
  limits: WorkbookDrawingTextLimits & { readonly maxCodeUnits?: number } = {}, fontProbes?: FontProbeCache): WorkbookDrawingTextBatch {
  const document = scope.ownerDocument, view = document.defaultView, maxNodes = limits.maxNodes ?? 50_000, maxUnits = limits.maxCodeUnits ?? 8 * 1024 * 1024;
  if (!scope.isConnected || !view || !Number.isSafeInteger(maxNodes) || maxNodes < 1 || !Number.isSafeInteger(maxUnits) || maxUnits < 1)
    throw new DocumentError("InvalidOperation", "Chart title contents require a connected surface and positive budgets.");
  let nodes = 1, units = 0, transientNodes = 0;
  const account = (count: number): void => { nodes += count; if (nodes > maxNodes) throw new DocumentError("ResourceLimitExceeded", "Chart title contents exceed their cumulative node budget."); };
  const probes: HTMLElement[] = [], bodies: TextBodyLayoutInput[] = [];
  try {
    for (const input of inputs) {
      const { source, chart, role } = input;
      if (source.ownerDocument !== document || !source.isConnected || source.dataset.nodeId !== chart.nodeId || source.dataset.officeKind !== "chart")
        throw new DocumentError("InvalidOperation", "Chart title contents require their own connected chart source.");
      const definition = role === "title" ? { nodeId: chart.titleNodeId, reference: chart.titleReference, cssClass: chart.titleCssClass, text: chart.titleText }
        : role === "category-axis-title" ? chart.categoryAxisTitle : role === "value-axis-title" ? chart.valueAxisTitle : undefined;
      if (!definition || !!definition.nodeId === !!definition.reference || (definition.reference ? input.value === undefined : input.value !== undefined))
        throw new DocumentError("InvalidOperation", "Chart title contents require exactly their declared rich or current-reference source.");
      let bodySource: HTMLElement, text: WorkbookDrawingTextView, properties = definition.text;
      if (definition.nodeId) {
        bodySource = workbookChartTextSource(source, chart, role).source;
        units += (bodySource.textContent?.length ?? 0) + bodySource.querySelectorAll("br").length;
        if (units > maxUnits) throw new DocumentError("ResourceLimitExceeded", "Chart title contents exceed their text budget.");
        text = createWorkbookChartText(source, chart, role, Math.max(1, maxNodes - nodes));
      } else {
        units += input.value!.length;
        if (units > maxUnits) throw new DocumentError("ResourceLimitExceeded", "Chart title contents exceed their text budget.");
        account(3); transientNodes += 3;
        const probe = markViewNode(document.createElement("div")), paragraph = document.createElement("p");
        probe.dataset.officeView = "workbook-reference-title-source"; probe.dataset.sourceNodeId = chart.nodeId;
        probe.className = [chart.textCssClass, definition.cssClass].filter(Boolean).join(" ");
        probe.style.cssText = "position:absolute;visibility:hidden";
        paragraph.textContent = input.value!; probe.append(paragraph); source.append(probe); probes.push(probe);
        const computed = view.getComputedStyle(probe);
        // txPr paragraph properties are stored on the title style; move their
        // resolved values to the private paragraph without changing the source.
        for (const key of ["text-align", "text-indent", "margin-left", "margin-right", "margin-top", "margin-bottom"]) paragraph.style.setProperty(key, computed.getPropertyValue(key));
        const wrap = property(computed, "--office-text-wrap"), autoFit = property(computed, "--office-text-auto-fit");
        const numeric = (name: string, minimum: number, exclusive: boolean): number | undefined => {
          const raw = property(computed, name); if (!raw) return undefined;
          const value = Number(raw);
          if (!Number.isFinite(value) || (exclusive ? value <= minimum : value < minimum) || value > 1) throw new DocumentError("InvalidDocument", "Chart title text properties are invalid.");
          return value;
        };
        if (wrap && !["true", "false"].includes(wrap) || autoFit && !["none", "normal", "shape"].includes(autoFit)) throw new DocumentError("InvalidDocument", "Chart title text properties are invalid.");
        const fontScale = numeric("--office-text-font-scale", 0, true), lineSpacingReduction = numeric("--office-text-line-spacing-reduction", 0, false);
        properties = { ...(wrap ? { wrap: wrap === "true" } : {}), ...(autoFit ? { autoFit: autoFit as "none" | "normal" | "shape" } : {}),
          ...(fontScale === undefined ? {} : { fontScale }), ...(lineSpacingReduction === undefined ? {} : { lineSpacingReduction }) };
        bodySource = probe; text = createWorkbookDerivedDrawingText(probe, chart.nodeId, Math.max(1, maxNodes - nodes));
      }
      account(text.nodes + 1);
      units += text.element.textContent?.length ?? 0;
      if (units > maxUnits) throw new DocumentError("ResourceLimitExceeded", "Chart title contents exceed their source and display text budget.");
      text.element.dataset.officeChartTextRole = role;
      bodies.push({ source: bodySource, text, properties, widthPt: input.widthPt, heightPt: input.heightPt,
        geometry: (width, height) => evaluatePresentationPreset("rect", width, height), flipH: false, flipV: false });
    }
    // The measurement pass's own accounting includes all display nodes and its
    // root; keep simultaneous source probes inside the same peak budget.
    return layoutTextBodies(scope, bodies, { ...limits, maxNodes: Math.max(1, maxNodes - transientNodes) }, input => input, fontProbes);
  } finally {
    for (const probe of probes) probe.remove();
  }
}

function layoutTextBodies<T>(scope: HTMLElement, inputs: readonly T[], limits: WorkbookDrawingTextLimits, adapt: (input: T) => TextBodyLayoutInput, probes?: FontProbeCache): WorkbookDrawingTextBatch {
  const document = scope.ownerDocument, view = document.defaultView, maxNodes = limits.maxNodes ?? 50_000, maxPasses = limits.maxPasses ?? 8, maxExtent = limits.maxExtentPt ?? 1_000_000;
  if (!scope.isConnected || !view || document.fonts.status === "loading" || !Number.isSafeInteger(maxNodes) || maxNodes < 1 || !Number.isSafeInteger(maxPasses) || maxPasses < 1 || maxPasses > 128 || !Number.isFinite(maxExtent) || maxExtent <= 0) throw new DocumentError("InvalidOperation", "Drawing text layout requires a connected surface, settled fonts and finite layout budgets.");
  let nodes = 1;
  const account = (amount: number): void => { nodes += amount; if (nodes > maxNodes) throw new DocumentError("ResourceLimitExceeded", "Drawing text layout exceeds its node budget."); };
  const seen = new Set<HTMLElement>(), bodies: TextBodyLayoutInput[] = [];
  for (const original of inputs) {
    const input = adapt(original);
    if (input.source.ownerDocument !== document || !input.source.isConnected || input.text.element.ownerDocument !== document || input.text.element.parentNode || seen.has(input.text.element)
      || input.text.element.dataset.sourceNodeId !== (input.source.dataset.nodeId ?? input.source.dataset.sourceNodeId)
      || ![input.widthPt, input.heightPt].every(n => Number.isFinite(n) && n >= 0)) throw new DocumentError("InvalidOperation", "Drawing layout requires its own connected source, distinct detached text and finite object extents.");
    if (Math.max(input.widthPt, input.heightPt) > maxExtent) throw new DocumentError("ResourceLimitExceeded", "Drawing text layout exceeds its extent budget.");
    seen.add(input.text.element); account(input.text.nodes + 1); bodies.push(input);
  }
  const root = markViewNode(document.createElement("div")); root.dataset.officeView = "workbook-drawing-text-measurement"; root.setAttribute("aria-hidden", "true");
  root.style.cssText = "position:absolute;visibility:hidden;pointer-events:none;contain:layout;left:0;top:0;width:0;height:0;overflow:hidden;direction:ltr";
  let complete = false;
  try {
    const prepared = bodies.map(input => {
      const computed = view.getComputedStyle(input.source), text = input.text.element, element = markViewNode(document.createElement("div"));
      const anchor = property(computed, "--office-text-anchor", "t"), centered = property(computed, "--office-text-anchor-center", "false");
      const horizontal = property(computed, "--office-text-horizontal-overflow", "overflow"), vertical = property(computed, "--office-text-vertical-overflow", "overflow");
      if (!["t", "ctr", "b", "just", "dist"].includes(anchor) || !["false", "true"].includes(centered)) throw new DocumentError("InvalidDocument", "A drawing text anchor is invalid.");
      if (!["overflow", "clip"].includes(horizontal) || !["overflow", "clip", "ellipsis"].includes(vertical)) throw new DocumentError("InvalidDocument", "A drawing text overflow mode is invalid.");
      element.style.cssText = "position:relative;direction:ltr";
      Object.assign(text.style, { position: "absolute", display: "flex", flexDirection: "column", boxSizing: "border-box", margin: "0", overflow: "visible", justifyContent: "flex-start",
        whiteSpace: input.properties?.wrap === false ? "pre" : "pre-wrap", overflowWrap: input.properties?.wrap === false ? "normal" : "anywhere", transformOrigin: "50% 50%", transform: "none" });
      for (const side of ["left", "right", "top", "bottom"]) text.style.setProperty("padding-" + side, computed.getPropertyValue("padding-" + side));
      element.append(text); root.append(element);
      const state = { input, element, text, anchor, horizontal: horizontal as "overflow" | "clip", vertical: vertical as "overflow" | "clip" | "ellipsis", centered: centered === "true", nodes: input.text.nodes + 1, width: input.widthPt, height: input.heightPt, geometry: input.geometry(input.widthPt, input.heightPt),
        resize(width: number, height: number) {
          if (![width, height].every(n => Number.isFinite(n) && n >= 0 && n <= maxExtent)) throw new DocumentError("ResourceLimitExceeded", "Drawing text auto-fit exceeds its extent budget.");
          this.width = width; this.height = height; this.geometry = input.geometry(width, height);
          element.style.width = width + "pt"; element.style.height = height + "pt";
          const rect = this.geometry.textRect;
          Object.assign(text.style, rect ? { display: "flex", left: rect.x + "pt", top: rect.y + "pt",
            // anchorCtr centers a minimal text block, independently of the
            // paragraphs' alignments. Intrinsic width keeps mixed alignment
            // inside that block; max-width preserves the native wrap boundary.
            width: centered === "true" ? "max-content" : rect.width + "pt", maxWidth: centered === "true" ? rect.width + "pt" : "none", height: rect.height + "pt" } : { display: "none" });
        } };
      state.resize(state.width, state.height); return state;
    });
    scope.append(root);
    for (const item of prepared) decorateDrawingLists(item.text, amount => { account(amount); item.nodes += amount; });
    const fonts = [...resolvePreviewFontsInScope(document, root, "[data-office-text-style],[data-office-script],.sdk-ppt-marker", probes)];
    const normalLines = new Map<string, number>(), lineContext = document.createElement("canvas").getContext("2d");
    let usedFallbackLineMetrics = false;
    const lineHeight = (computed: CSSStyleDeclaration): number => {
      const line = parseFloat(computed.lineHeight); if (Number.isFinite(line)) return line;
      if (parseFloat(computed.fontSize) === 0) return 0;
      const properties = ["font-family", "font-size", "font-weight", "font-style", "font-stretch", "font-variant", "letter-spacing"], key = JSON.stringify(properties.map(name => computed.getPropertyValue(name)));
      const cached = normalLines.get(key); if (cached !== undefined) return cached;
      // Measure scalable font bounds at a fixed large size: browser pixel
      // rounding at the actual 14–18pt size otherwise adds a full pixel to
      // every line in some engines. No large canvas bitmap is allocated.
      const height = measuredFontLineHeight(lineContext, computed);
      if (height !== undefined) { normalLines.set(key, height); return height; }
      // Legacy engines do not expose font bounds. Keep a measured DOM fallback
      // and report its use; do not silently claim native normal-leading parity.
      usedFallbackLineMetrics = true;
      account(2); const probe = document.createElement("span");
      probe.style.cssText = "position:absolute;display:block;width:max-content;white-space:pre;line-height:normal;margin:0;padding:0;border:0";
      for (const name of properties) probe.style.setProperty(name, computed.getPropertyValue(name));
      probe.textContent = "M"; root.append(probe);
      try {
        const height = probe.getBoundingClientRect().height;
        if (!Number.isFinite(height) || height < 0 || height === 0 && parseFloat(computed.fontSize) > 0) throw new DocumentError("UnsupportedFeature", "This browser cannot measure the resolved drawing font's normal leading.");
        normalLines.set(key, height); return height;
      } finally { probe.remove(); nodes -= 2; }
    };
    const objects = prepared.map(item => {
      const { text } = item, properties = item.input.properties, styles = [text, ...text.querySelectorAll<HTMLElement>("[data-office-text-style],[data-office-script],.sdk-ppt-marker")].map(node => {
        const computed = view.getComputedStyle(node); return { node, size: parseFloat(computed.fontSize), line: lineHeight(computed) };
      });
      const apply = (scale: number): void => { for (const { node, size, line } of styles) {
        // DrawingML's paragraph default is an inheritance/insertion style,
        // not an extra large glyph on every line. Actual run/empty-line boxes
        // carry explicit sizes and leading; suppress only the paragraph strut.
        if (node.localName === "p" || node.closest('[data-office-text-break="true"]')) { node.style.fontSize = "0px"; node.style.lineHeight = "0px"; node.style.minHeight = "0px"; }
        else { node.style.fontSize = size * scale + "px"; if (Number.isFinite(line)) node.style.lineHeight = line * scale * (1 - (properties?.lineSpacingReduction ?? 0)) + "px"; }
      } };
      let scale = properties?.fontScale ?? 1;
      const measured = () => ({ width: Math.max(text.scrollWidth, text.clientWidth), height: Math.max(text.scrollHeight, text.clientHeight),
        availableWidth: (item.geometry.textRect?.width ?? 0) / .75, availableHeight: (item.geometry.textRect?.height ?? 0) / .75 });
      // A CSS border box cannot shrink below its padding. Compare against the
      // declared text rectangle so that oversized insets do not appear to fit.
      const overflow = (): boolean => { const size = measured(); return !!item.geometry.textRect && (size.height > size.availableHeight + 1 || size.width > size.availableWidth + 1); };
      apply(scale);
      if (properties?.autoFit === "normal") for (let pass = 0; pass < maxPasses && overflow(); pass++) {
        const size = measured(), next = Math.max(.01, scale * Math.min(.95, Math.sqrt(size.availableHeight / Math.max(1, size.height)), size.availableWidth / Math.max(1, size.width)));
        if (next >= scale) break; scale = next; apply(scale);
      }
      if (properties?.autoFit === "shape") for (let pass = 0; pass < maxPasses && overflow(); pass++) {
        const size = measured(); item.resize(item.width + Math.max(0, size.width - size.availableWidth) * .75, item.height + Math.max(0, size.height - size.availableHeight) * .75);
      }
      const exceeded = overflow();
      text.style.justifyContent = item.anchor === "ctr" ? "center" : item.anchor === "b" ? "flex-end" : item.anchor === "just" ? "space-between" : item.anchor === "dist" ? "space-around" : "flex-start";
      if (item.centered && item.geometry.textRect) {
        const rect = item.geometry.textRect, width = text.getBoundingClientRect().width * .75;
        text.style.left = rect.x + (rect.width - width) / 2 + "pt";
      }
      if (item.geometry.textRect && (item.horizontal !== "overflow" || item.vertical !== "overflow")) {
        const rect = item.geometry.textRect, left = (rect.x - parseFloat(text.style.left)) / .75, top = (rect.y - parseFloat(text.style.top)) / .75;
        applyDrawingTextOverflow(text, { left, top, right: left + rect.width / .75, bottom: top + rect.height / .75 }, item.horizontal, item.vertical,
          amount => { account(amount); item.nodes += amount; });
        fonts.push(...resolvePreviewFontsInScope(document, text, '[data-office-view="workbook-drawing-ellipsis"]', probes));
      }
      text.style.transform = `scale(${item.input.flipH ? -1 : 1},${item.input.flipV ? -1 : 1})`;
      const origin = item.element.getBoundingClientRect();
      let left = 0, top = 0, right = item.width, bottom = item.height;
      // Element rectangles include paragraph/run backgrounds and reflected
      // line boxes. Text ranges also include visible unbroken text overflowing
      // its containing block, including negative-side/right-aligned overflow.
      const include = (rect: DOMRect): void => {
        left = Math.min(left, (rect.left - origin.left) * .75); top = Math.min(top, (rect.top - origin.top) * .75);
        right = Math.max(right, (rect.right - origin.left) * .75); bottom = Math.max(bottom, (rect.bottom - origin.top) * .75);
      };
      if (text.style.display !== "none") {
        const walker = document.createTreeWalker(text, view.NodeFilter.SHOW_ELEMENT | view.NodeFilter.SHOW_TEXT), range = document.createRange();
        include(text.getBoundingClientRect());
        for (let node = walker.nextNode(); node; node = walker.nextNode()) {
          if (node.nodeType === 1) for (const rect of (node as Element).getClientRects()) include(rect);
          else { range.selectNodeContents(node); for (const rect of range.getClientRects()) include(rect); }
        }
      }
      if (![left, top, right, bottom].every(Number.isFinite) || Math.max(Math.abs(left), Math.abs(top), Math.abs(right), Math.abs(bottom)) > maxExtent)
        throw new DocumentError("ResourceLimitExceeded", "Drawing text overflow exceeds its measurable extent budget.");
      if (item.geometry.textRect) {
        const rect = item.geometry.textRect;
        if (item.horizontal === "clip") { left = Math.max(left, Math.min(0, rect.x)); right = Math.min(right, Math.max(item.width, rect.x + rect.width)); }
        if (item.vertical !== "overflow") { top = Math.max(top, Math.min(0, rect.y)); bottom = Math.min(bottom, Math.max(item.height, rect.y + rect.height)); }
      }
      const paintBounds = Object.freeze({ x: left, y: top, width: right - left, height: bottom - top });
      const displayedText = Object.freeze({ ...item.input.text, nodes: item.nodes - 1 });
      item.element.remove(); return Object.freeze({ element: item.element, text: displayedText, widthPt: item.width, heightPt: item.height, geometry: item.geometry, scale, overflow: exceeded, paintBounds, nodes: item.nodes });
    });
    complete = true;
    return Object.freeze({ objects: Object.freeze(objects), fonts, usedFallbackLineMetrics });
  } finally {
    if (!complete) for (const input of bodies) input.text.element.remove();
    root.remove(); root.replaceChildren();
  }
}
