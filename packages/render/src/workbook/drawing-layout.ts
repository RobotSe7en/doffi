import { DocumentError, evaluatePresentationPreset } from "@doffi/core";
import type { DeepReadonly, WorkbookNumberFormatter, WorkbookObject, WorkbookSnapshot } from "@doffi/core";
import type { FontResolution } from "../contracts.js";
import { createDrawingGeometryView } from "../drawing-geometry.js";
import { strokeAttributes } from "../presentation/appearance.js";
import { markViewNode } from "../view-nodes.js";
import { prepareWorkbookAssets } from "./assets.js";
import type { WorkbookAssetLimits, WorkbookAssetSource, WorkbookPreparedAssets } from "./assets.js";
import { DEFAULT_CHART_READ_LIMITS } from "./chart-source.js";
import type { WorkbookChartSource } from "./chart-source.js";
import { bindWorkbookChartData } from "./chart-values.js";
import type { WorkbookChartBinding, WorkbookChartDataDiagnostic } from "./chart-values.js";
import { layoutWorkbookChart } from "./chart-layout.js";
import type { WorkbookChartViewDiagnostic } from "./chart-view.js";
import type { WorkbookContentSources } from "./content-source.js";
import type { AxisDimensions } from "./dimensions.js";
import { projectWorkbookDrawing, workbookDrawingRect } from "./drawing.js";
import type { WorkbookDrawingProjection } from "./drawing.js";
import { workbookDrawingShadow } from "./drawing-shadow.js";
import type { WorkbookDrawingShadow } from "./drawing-shadow.js";
import { createWorkbookDrawingFragment, workbookDrawingOrientation } from "./drawing-view.js";
import type { WorkbookDrawingOrientation } from "./drawing-view.js";
import { createWorkbookDrawingText } from "./drawing-text.js";
import { layoutWorkbookDrawingTexts } from "./drawing-text-layout.js";
import type { WorkbookDrawingTextLayout } from "./drawing-text-layout.js";
import { createWorkbookImageContent, decodeWorkbookImages } from "./image-view.js";
import type { WorkbookRect } from "./projection.js";
import type { WorkbookSurface } from "./surface.js";
import type { WorkbookWindowFrame } from "./window-loader.js";
import { checkWindowCancellation, checkWindowIdentity } from "./window-source.js";
import { layoutWorkbookNoteBoxes } from "./note-boxes.js";
import type { WorkbookNoteProjection } from "./note-boxes.js";

export interface WorkbookDrawingLayoutLimits {
  readonly maxObjects?: number; readonly maxNodes?: number; readonly maxSourceNodes?: number;
  readonly maxCodeUnits?: number; readonly maxExtentPt?: number;
  readonly maxShadowPixels?: number; readonly maxObjectShadowPixels?: number;
  readonly assets?: Partial<WorkbookAssetLimits>;
}
export type WorkbookDrawingDiagnostic = WorkbookChartDataDiagnostic | WorkbookChartViewDiagnostic
  | { readonly code: "drawing-text-overflow" | "approximate-shadow" | "chart-title-overflow"; readonly nodeId: string };
export interface WorkbookDrawingLayout {
  readonly element: HTMLElement; readonly nodes: number; readonly shadowPixels: number;
  readonly fonts: readonly FontResolution[]; readonly usedFallbackLineMetrics: boolean;
  readonly diagnostics: readonly WorkbookDrawingDiagnostic[];
  readonly objects: readonly { readonly nodeId: string; readonly projection: WorkbookDrawingProjection; readonly text?: WorkbookDrawingTextLayout }[];
  readonly notes: readonly WorkbookNoteProjection[];
  readonly assets: Readonly<{ count: number; byteLength: number; pixels: number }>;
  /** Release this candidate only. The caller owns replacement of an active window. */
  dispose(): void;
}
interface PreparedDrawing {
  readonly object: DeepReadonly<WorkbookObject>; readonly source: HTMLElement; readonly style: CSSStyleDeclaration;
  readonly rect: WorkbookRect; readonly orientation: WorkbookDrawingOrientation; readonly parts: readonly WorkbookDrawingProjection[];
  readonly shadow: WorkbookDrawingShadow; readonly text?: WorkbookDrawingTextLayout; readonly imagePath?: string;
}

/** Build one detached object layer against the same dimensions and pane frame
 * as the cells. Source order controls stacking. All panes share one candidate
 * resource lease, while each owns fresh view nodes, projections and SVG IDs. */
export async function layoutWorkbookDrawings(surface: Pick<WorkbookSurface, "source" | "view" | "fontProbes">, snapshot: WorkbookSnapshot,
  frame: WorkbookWindowFrame, dimensions: { readonly rows: AxisDimensions; readonly columns: AxisDimensions },
  sources: WorkbookContentSources, assetSource: WorkbookAssetSource, chartSource: WorkbookChartSource,
  formatter: WorkbookNumberFormatter, signal: AbortSignal, input: WorkbookDrawingLayoutLimits = {}): Promise<WorkbookDrawingLayout> {
  const document = surface.source.ownerDocument, view = document.defaultView;
  const limits = { maxObjects: 128, maxNodes: 50_000, maxSourceNodes: 50_000, maxCodeUnits: 8 * 1024 * 1024,
    maxExtentPt: 1_000_000, maxShadowPixels: 64 * 1024 * 1024, maxObjectShadowPixels: 16 * 1024 * 1024, ...input };
  if (!surface.source.isConnected || !surface.view.isConnected || surface.view.ownerDocument !== document || !view
    || Object.keys(input).some(key => !["maxObjects", "maxNodes", "maxSourceNodes", "maxCodeUnits", "maxExtentPt", "maxShadowPixels", "maxObjectShadowPixels", "assets"].includes(key))
    || ![limits.maxObjects, limits.maxNodes, limits.maxSourceNodes, limits.maxCodeUnits, limits.maxShadowPixels, limits.maxObjectShadowPixels].every(value => Number.isSafeInteger(value) && value > 0)
    || !Number.isFinite(limits.maxExtentPt) || limits.maxExtentPt <= 0 || limits.maxExtentPt > 1_000_000)
    throw new DocumentError("InvalidOperation", "Drawing composition requires one connected surface and supported finite budgets.");
  checkWindowCancellation(signal); checkWindowIdentity(frame, snapshot);
  checkWindowIdentity(sources.identity, snapshot);
  checkWindowIdentity(assetSource.snapshot.manifest, snapshot); checkWindowIdentity(chartSource.snapshot.manifest, snapshot);
  const sheet = snapshot.metadata.workbook.sheets.find(sheet => sheet.sheetId === frame.sheetId);
  if (!sheet) throw new DocumentError("InvalidOperation", "Drawing composition requires its owning worksheet.");
  const objects = (snapshot.metadata.workbook.objects ?? []).filter(object => object.sheetId === frame.sheetId);
  const projection = sources.project(surface.source, objects.map(object => object.nodeId), limits.maxSourceNodes);
  const element = markViewNode(document.createElement("div")), images: HTMLImageElement[] = [], prepared: PreparedDrawing[] = [];
  element.dataset.officeView = "workbook-drawings"; element.style.cssText = `position:absolute;left:0;top:0;width:${frame.layout.widthPt}pt;height:${frame.layout.heightPt}pt;pointer-events:none`;
  let nodes = 1, retained = 0, units = 0, shadowPixels = 0, lease: WorkbookPreparedAssets | undefined, usedFallbackLineMetrics = false, disposed = false;
  const fonts: FontResolution[] = [], diagnostics: WorkbookDrawingDiagnostic[] = [], displayed: { nodeId: string; projection: WorkbookDrawingProjection; text?: WorkbookDrawingTextLayout }[] = [];
  const remaining = (): number => { if (nodes + retained >= limits.maxNodes) throw new DocumentError("ResourceLimitExceeded", "Drawing composition exceeds its cumulative display-node budget."); return limits.maxNodes - nodes - retained; };
  const account = (count: number): void => { nodes += count; if (nodes + retained > limits.maxNodes) throw new DocumentError("ResourceLimitExceeded", "Drawing composition exceeds its cumulative display-node budget."); };
  const textUnits = (text: string): void => { units += text.length; if (units > limits.maxCodeUnits) throw new DocumentError("ResourceLimitExceeded", "Drawing composition exceeds its cumulative text budget."); };
  const countNodes = (root: Node): number => { const walker = document.createTreeWalker(root, 0xffff_ffff); let count = 1; while (walker.nextNode()) count++; return count; };
  const dispose = (): void => { if (disposed) return; disposed = true; for (const image of images) image.removeAttribute("src"); element.remove(); element.replaceChildren(); lease?.dispose(); };
  let abort = (): void => {};
  const cancelled = new Promise<never>((_, reject) => { abort = () => reject(new DocumentError("OperationCancelled", "Workbook drawing composition was cancelled.")); });
  // Cancellation may arrive while yielding between CPU batches, outside a
  // provider race. Keep that rejection handled; checkpoints still throw it.
  void cancelled.catch(() => {});
  signal.addEventListener("abort", abort, { once: true });
  const checkpoint = async (index: number): Promise<void> => { if (index % 16 === 0) await new Promise<void>(resolve => setTimeout(resolve, 0)); checkWindowCancellation(signal); };
  try {
    const ordered = objects.map(object => ({ object, source: projection.get(object.nodeId)! }))
      .sort((a, b) => a.source === b.source ? 0 : a.source.compareDocumentPosition(b.source) & view.Node.DOCUMENT_POSITION_FOLLOWING ? -1 : 1);
    const imagePaths = new Map(projection.assets.map(asset => [asset.nodeId, asset.path]));
    for (const [index, { object, source }] of ordered.entries()) {
      await checkpoint(index);
      const orientation = workbookDrawingOrientation(source); if (orientation.hidden) continue;
      let rect = workbookDrawingRect(object.anchor, dimensions), text: WorkbookDrawingTextLayout | undefined;
      if (Math.max(rect.width, rect.height) > limits.maxExtentPt) throw new DocumentError("ResourceLimitExceeded", "A worksheet object exceeds its measurable extent budget.");
      textUnits(source.textContent ?? "");
      if (object.kind === "geometry" || object.kind === "text") {
        const batch = layoutWorkbookDrawingTexts(surface.view, [{ source, object, text: createWorkbookDrawingText(source, object, remaining()), widthPt: rect.width, heightPt: rect.height }], { maxNodes: remaining(), maxExtentPt: limits.maxExtentPt }, surface.fontProbes);
        text = batch.objects[0]!; rect = { ...rect, width: text.widthPt, height: text.heightPt };
        fonts.push(...batch.fonts); usedFallbackLineMetrics ||= batch.usedFallbackLineMetrics;
      }
      const style = view.getComputedStyle(source), stroke = strokeAttributes(document, style, "--office-line", object.kind === "geometry" && object.preset === "line" ? "#000000" : "none");
      const width = Number(stroke["stroke-width"]), miter = Number(stroke["stroke-miterlimit"]);
      if (![width, miter].every(value => Number.isFinite(value) && value >= 0)) throw new DocumentError("InvalidDocument", "A worksheet outline has invalid bounds.");
      const outline = stroke.stroke === "none" ? 0 : width * Math.max(6, miter);
      const painted = text?.paintBounds ?? { x: 0, y: 0, width: rect.width, height: rect.height };
      const outset = Math.max(-painted.x, -painted.y, painted.x + painted.width - rect.width, painted.y + painted.height - rect.height, outline);
      if (!Number.isFinite(outset) || outset > limits.maxExtentPt) throw new DocumentError("ResourceLimitExceeded", "Worksheet painted bounds exceed their extent budget.");
      const paintWidth = Math.max(rect.width + outline, painted.x + painted.width) - Math.min(-outline, painted.x);
      const paintHeight = Math.max(rect.height + outline, painted.y + painted.height) - Math.min(-outline, painted.y);
      const shadow = workbookDrawingShadow(source, orientation, paintWidth, paintHeight, { maxRasterPixels: Number.MAX_SAFE_INTEGER });
      const parts = Object.freeze(projectWorkbookDrawing(rect, orientation.rotationDegrees, frame.layout, outset + shadow.outsetPt)
        .map(part => Object.freeze({ ...part, rect: Object.freeze(part.rect), clip: Object.freeze(part.clip) })));
      if (!parts.length) continue;
      if (prepared.length >= limits.maxObjects) throw new DocumentError("ResourceLimitExceeded", "The worksheet window exceeds its visible-object budget.");
      shadowPixels += shadow.rasterPixels * parts.length;
      if (shadow.rasterPixels > limits.maxObjectShadowPixels || !Number.isSafeInteger(shadowPixels) || shadowPixels > limits.maxShadowPixels) throw new DocumentError("ResourceLimitExceeded", "Worksheet object shadows exceed the window raster budget.");
      if (text) { retained += text.nodes; remaining(); if (text.overflow) diagnostics.push({ code: "drawing-text-overflow", nodeId: object.nodeId }); }
      if (shadow.approximateBlur) diagnostics.push({ code: "approximate-shadow", nodeId: object.nodeId });
      const imagePath = object.kind === "image" ? imagePaths.get((source.firstElementChild as HTMLElement).dataset.nodeId!) : undefined;
      if (object.kind === "image" && !imagePath) throw new DocumentError("InvalidDocument", "A worksheet image has no declared source resource.");
      prepared.push({ object, source, style, rect, orientation, parts, shadow, ...(text ? { text } : {}), ...(imagePath ? { imagePath } : {}) });
    }
    const chartIds = prepared.flatMap(item => item.object.kind === "chart" ? [item.object.chartId] : []);
    let bindings: readonly WorkbookChartBinding[] = [];
    if (chartIds.length) {
      const request = { ...DEFAULT_CHART_READ_LIMITS, maxCharts: limits.maxObjects, maxCodeUnits: Math.max(1, limits.maxCodeUnits - units), documentId: frame.documentId, revision: frame.revision, chartIds };
      const response = await Promise.race([chartSource.read(request, signal), cancelled]); checkWindowCancellation(signal);
      if (response.charts.length !== chartIds.length || response.charts.some(chart => !chartIds.includes(chart.chartId)))
        throw new DocumentError("InvalidDocument", "The visible chart response differs from its requested objects.");
      bindings = bindWorkbookChartData(snapshot, response, formatter, { maxCharts: limits.maxObjects, maxCodeUnits: request.maxCodeUnits });
    }
    lease = await prepareWorkbookAssets(assetSource, prepared.flatMap(item => item.imagePath ? [item.imagePath] : []), signal, limits.assets);
    const charts = new Map((snapshot.metadata.workbook.charts ?? []).map(chart => [chart.chartId, chart])), bound = new Map(bindings.map(binding => [binding.chartId, binding]));
    for (const [index, item] of prepared.entries()) {
      await checkpoint(index);
      const { object, source, style, rect, orientation, shadow } = item;
      const binding = object.kind === "chart" ? bound.get(object.chartId) : undefined;
      if (object.kind === "chart" && !binding) throw new DocumentError("InvalidDocument", "A visible chart has no current binding response.");
      if (binding) diagnostics.push(...binding.diagnostics);
      for (const [partIndex, part] of item.parts.entries()) {
        account(3); // Content group and the fragment's two wrappers.
        const content = markViewNode(document.createElement("div")); content.style.cssText = `position:absolute;left:0;top:0;width:${rect.width}pt;height:${rect.height}pt;filter:${shadow.filter}`;
        const preset = object.kind === "geometry" ? object.preset : "rect", geometry = item.text?.geometry ?? evaluatePresentationPreset(preset, rect.width, rect.height);
        const id = `workbook-object-${index}-${partIndex}`;
        const framePaint = (paint: "fill" | "stroke"): void => {
          const background = style.backgroundColor, hasBackground = background !== "transparent" && !/rgba\([^)]*,\s*0\)$/.test(background);
          const svg = createDrawingGeometryView(document, style, { preset, geometry, widthPt: rect.width, heightPt: rect.height, id: id + "-" + paint, paint, maxNodes: remaining(), orientation: item.orientation,
            fillFallback: preset === "line" ? "none" : hasBackground ? background : object.kind === "geometry" ? "#4472c4" : "none", strokeFallback: preset === "line" ? "#000000" : "none" });
          account(countNodes(svg)); content.append(svg);
        };
        framePaint("fill");
        let text: WorkbookDrawingTextLayout | undefined;
        if (item.imagePath) {
          const image = createWorkbookImageContent(source, item.imagePath, lease, { widthPt: rect.width, heightPt: rect.height, orientation: item.orientation }); images.push(image.image); account(image.nodes);
          if (rect.width * image.scaleX > limits.maxExtentPt || rect.height * image.scaleY > limits.maxExtentPt)
            throw new DocumentError("ResourceLimitExceeded", "A cropped worksheet image exceeds its measurable extent budget.");
          content.append(image.element);
        }
        if (item.text) {
          if (!partIndex) { text = item.text; retained -= text.nodes; } else text = cloneDrawingText(item.text, remaining());
          account(text.nodes); textUnits(text.element.textContent ?? ""); content.append(text.element);
        }
        if (object.kind === "chart") {
          if (binding!.kind === "unavailable") {
            account(2); const message = markViewNode(document.createElement("div")); message.textContent = "Chart data unavailable"; textUnits(message.textContent);
            message.dataset.officeView = "workbook-chart-unavailable"; message.style.cssText = "position:absolute;inset:0;display:flex;align-items:center;justify-content:center"; content.append(message);
          } else {
            const chart = layoutWorkbookChart(surface.view, source, charts.get(object.chartId)!, binding as Extract<WorkbookChartBinding, { kind: "available" }>, rect.width, rect.height, id + "-chart", formatter,
              { maxNodes: remaining(), maxCodeUnits: Math.max(1, limits.maxCodeUnits - units) }, surface.fontProbes);
            account(chart.nodes); textUnits(chart.element.textContent ?? ""); chart.element.style.overflow = "hidden"; content.append(chart.element);
            if (!partIndex) { fonts.push(...chart.fonts); usedFallbackLineMetrics ||= chart.usedFallbackLineMetrics;
              diagnostics.push(...chart.diagnostics.filter(diagnostic => diagnostic.code === "negative-pie-magnitude" || diagnostic.code === "chart-label-overflow"));
              if (chart.overflowingTitles.length) diagnostics.push({ code: "chart-title-overflow", nodeId: object.nodeId }); }
          }
        }
        framePaint("stroke");
        const fragment = createWorkbookDrawingFragment(document, object.nodeId, part, frame.layout.widthPt, orientation, sheet.rightToLeft, content);
        fragment.style.pointerEvents = "auto";
        if (partIndex) fragment.setAttribute("aria-hidden", "true");
        else { fragment.setAttribute("role", "group"); const label = source.getAttribute("data-office-alt") || object.name; if (label) { textUnits(label); fragment.setAttribute("aria-label", label); } }
        element.append(fragment); displayed.push(Object.freeze({ nodeId: object.nodeId, projection: part, ...(text ? { text } : {}) }));
      }
    }
    const nativeNotes = snapshot.metadata.workbook.comments?.some(comment => comment.format === "classic" && comment.box?.visible) ? await layoutWorkbookNoteBoxes(surface, snapshot, frame, dimensions, sources, signal, {
      maxNodes: remaining(), maxSourceNodes: limits.maxSourceNodes, maxCodeUnits: Math.max(1, limits.maxCodeUnits - units), maxExtentPt: limits.maxExtentPt, maxObjects: Math.max(0, limits.maxObjects - prepared.length),
    }) : undefined;
    if (nativeNotes) { account(nativeNotes.nodes); units += nativeNotes.codeUnits; fonts.push(...nativeNotes.fonts); element.append(nativeNotes.element); }
    await decodeWorkbookImages(images, signal); checkWindowCancellation(signal);
    return Object.freeze({ element, nodes, shadowPixels, fonts: Object.freeze(fonts), usedFallbackLineMetrics,
      diagnostics: Object.freeze(diagnostics.map(diagnostic => Object.freeze(diagnostic))), objects: Object.freeze(displayed), notes: nativeNotes?.notes ?? Object.freeze([]),
      assets: Object.freeze({ count: lease.count, byteLength: lease.byteLength, pixels: lease.pixels }), dispose });
  } catch (error) { dispose(); throw error; }
  finally { signal.removeEventListener("abort", abort); projection.dispose(); }
}

function cloneDrawingText(source: WorkbookDrawingTextLayout, maxNodes: number): WorkbookDrawingTextLayout {
  if (source.nodes > maxNodes) throw new DocumentError("ResourceLimitExceeded", "Repeated drawing text exceeds its node budget.");
  const element = source.element.cloneNode(true) as HTMLElement, document = element.ownerDocument;
  const original = document.createTreeWalker(source.element, 0xffff_ffff), cloned = document.createTreeWalker(element, 0xffff_ffff), map = new Map<Node, Node>([[source.element, element]]);
  while (original.nextNode()) { if (!cloned.nextNode()) throw new DocumentError("InvalidOperation", "Drawing text clone does not match its source view."); map.set(original.currentNode, cloned.currentNode); }
  const text = Object.freeze({ ...source.text, element: map.get(source.text.element) as HTMLElement,
    projections: Object.freeze(source.text.projections.map(projection => Object.freeze({ ...projection, node: map.get(projection.node) as Text }))) });
  return Object.freeze({ ...source, element, text });
}
