import { DocumentError } from "@doffi/core";
import type { DeepReadonly, WorkbookChart } from "@doffi/core";
import { createChartStyleProbe } from "../chart-style-probe.js";
import { renderBasicChart } from "../chart.js";
import type { BasicChartData, BasicChartLayout, BasicChartOptions } from "../chart.js";
import type { BasicChartTextRun } from "../chart.js";
import { markViewNode } from "../view-nodes.js";
import { resolvePreviewFontsInScope } from "../fonts.js";
import type { FontProbeCache } from "../font-probe-cache.js";
import type { FontResolution } from "../contracts.js";
import { workbookDrawingOrientation } from "./drawing-view.js";
import { prepareChartTextDecorations } from "../chart-text-decoration.js";
import { prepareChartLabelFrames } from "../chart-label-frame.js";
import { layoutBasicChartLabels } from "../chart-label-layout.js";
import { layoutChartNumberFields } from "./chart-number-layout.js";

export interface WorkbookChartViewOptions {
  /** Required: the workbook compositor owns number formats, locale and dates.
   * This adapter must never inherit the presentation formatter's policy. */
  readonly formatNumber: BasicChartOptions["formatNumber"];
  readonly formatCategory?: NonNullable<BasicChartOptions["formatCategory"]>;
  /** A fresh, already materialized title view; never a source/active view node. */
  readonly title?: HTMLElement;
  readonly layout?: BasicChartLayout;
  readonly maxNodes?: number; readonly maxPoints?: number; readonly maxCodeUnits?: number;
  readonly probes?: FontProbeCache;
}
export type WorkbookChartViewDiagnostic = { readonly code: "negative-pie-magnitude"; readonly chartId: string; readonly negativePoints: number }
  | { readonly code: "chart-label-overflow"; readonly chartId: string; readonly labels: number };
export interface WorkbookChartView {
  readonly element: HTMLElement; readonly nodes: number; readonly peakNodes: number;
  readonly fonts: readonly FontResolution[];
  readonly diagnostics: readonly WorkbookChartViewDiagnostic[];
}

/** Materialize graphics from already bound and formatted plot data. Resolve
 * CSS against the connected source and remove all temporary probes on failure.
 * The detached view carries no source classes or source DOM IDs. Frame effects,
 * title materialization, source projections and pane placement belong to the compositor. */
export function createWorkbookChartView(source: HTMLElement, chart: DeepReadonly<WorkbookChart>, data: BasicChartData, widthPt: number, heightPt: number, id: string, options: WorkbookChartViewOptions): WorkbookChartView {
  const document = source.ownerDocument, view = document.defaultView;
  const maxNodes = options.maxNodes ?? 50_000, maxPoints = options.maxPoints ?? 50_000, maxCodeUnits = options.maxCodeUnits ?? 8 * 1024 * 1024;
  if (!source.isConnected || !view || ![maxNodes, maxPoints, maxCodeUnits].every(value => Number.isSafeInteger(value) && value > 0)
    || ![widthPt, heightPt].every(value => Number.isFinite(value) && value > 0 && value <= 1_000_000) || !/^[A-Za-z_][A-Za-z0-9_.-]*$/.test(id)) throw new DocumentError("InvalidOperation", "Chart graphics require a connected source, finite positive dimensions, a local SVG identity and positive budgets.");
  if (source.dataset.nodeId !== chart.nodeId || source.dataset.officeKind !== "chart") throw new DocumentError("InvalidDocument", "Chart graphics require their own source object.");
  if (options.layout && [options.layout.plot, ...(options.layout.title ? [options.layout.title] : []), ...(options.layout.legend ? [options.layout.legend] : [])].some(rect => ![rect.x, rect.y, rect.width, rect.height].every(Number.isFinite)
    || rect.x < 0 || rect.y < 0 || rect.width <= 0 || rect.height <= 0 || rect.x + rect.width > widthPt || rect.y + rect.height > heightPt))
    throw new DocumentError("InvalidOperation", "Chart layout rectangles must fit inside the finite object canvas.");
  const title = options.title;
  if (title && (title.ownerDocument !== document || title.isConnected || title.parentNode)) throw new DocumentError("InvalidOperation", "A chart title must be a fresh detached view in its own document.");
  if (chart.titleVisible !== false && (chart.titleNodeId || chart.titleReference) && !title) throw new DocumentError("InvalidOperation", "A visible chart title requires a materialized title view.");
  if (data.series.length !== chart.series.length || data.series.some((series, index) => series.seriesId !== chart.series[index]!.seriesId || series.values.length !== data.categories.length
    || series.values.some(value => value !== null && !Number.isFinite(value)) || series.hidden !== undefined && typeof series.hidden !== "boolean"
    || series.missing !== undefined && (series.missing.length !== data.categories.length || series.missing.some((mode, index) => mode !== undefined && (mode !== "gap" && mode !== "span" || series.values[index] !== null))))
    || data.sourceIndexes !== undefined && (data.sourceIndexes.length !== data.categories.length || data.sourceIndexes.some((index, position) => !Number.isSafeInteger(index) || index < 0 || index > 0xffff_ffff || position > 0 && index <= data.sourceIndexes![position - 1]!))) throw new DocumentError("InvalidOperation", "Derived chart values must match their source series and category positions.");
  let units = 0, points = data.categories.length;
  const accountText = (amount: number): void => { units += amount; if (units > maxCodeUnits) throw new DocumentError("ResourceLimitExceeded", "Chart text exceeds its source and display budget."); };
  for (const name of data.categories) accountText(name.length);
  for (const series of data.series) { accountText(series.name.length); points += series.values.length; }
  if (points > maxPoints) throw new DocumentError("ResourceLimitExceeded", "Derived chart data exceeds its point budget.");
  const negativePoints = chart.kind === "pie" ? data.series.reduce((count, series) => count + (series.hidden ? 0 : series.values.filter(value => value !== null && value < 0).length), 0) : 0;
  const diagnostics: WorkbookChartViewDiagnostic[] = negativePoints
    ? [Object.freeze({ code: "negative-pie-magnitude" as const, chartId: chart.chartId, negativePoints })] : [];
  const element = markViewNode(document.createElement("div")), probes: HTMLElement[] = [];
  const textRuns: BasicChartTextRun[] = []; let fontScope: HTMLElement | undefined, fontNodes = 0;
  element.dataset.officeView = "workbook-chart"; element.dataset.sourceNodeId = chart.nodeId;
  element.style.cssText = `position:absolute;width:${widthPt}pt;height:${heightPt}pt`;
  let nodes = 1, peakNodes = 1;
  const account = (amount: number): void => {
    nodes += amount; peakNodes = Math.max(peakNodes, nodes);
    if (peakNodes > maxNodes) throw new DocumentError("ResourceLimitExceeded", "Chart graphics exceed their node budget.");
  };
  try {
    let styleScope = source;
    if (chart.textCssClass) {
      account(1); styleScope = markViewNode(document.createElement("span")); styleScope.className = chart.textCssClass;
      styleScope.style.cssText = "position:absolute;visibility:hidden"; source.append(styleScope); probes.push(styleScope);
    }
    const styleProbe = createChartStyleProbe(styleScope, source, account, probes);
    if (title) {
      const walker = document.createTreeWalker(title, 0xffff_ffff);
      do { account(1); if (walker.currentNode.nodeType === 3) accountText(walker.currentNode.textContent!.length); } while (walker.nextNode());
      element.append(title);
    }
    const scene = renderBasicChart(element, chart, data, widthPt, heightPt, id, account, { formatNumber: options.formatNumber, ...(options.formatCategory ? { formatCategory: options.formatCategory } : {}), pieValues: "magnitude", negativeColumns: "white", showLabelsOverMaximum: chart.showLabelsOverMaximum ?? false, orientation: workbookDrawingOrientation(source),
      // MS-OI29500 §2.1.1571: specified multiple series disable by-point
      // variation, even when visibility later leaves only one rendered series.
      varyColors: (chart.varyColors ?? chart.kind === "pie") && chart.series.length === 1, accountText, onTextRun: run => { textRuns.push(run); }, ...(options.layout ? { layout: options.layout } : {}), ...(title ? { title } : {}), style: styleProbe });
    let fonts: readonly FontResolution[] = Object.freeze([]);
    if (textRuns.length) {
      account(1); fontNodes++; fontScope = document.createElement("span");
      fontScope.dataset.officeView = "workbook-chart-font-measurement"; fontScope.style.cssText = "position:absolute;display:none";
      const pairs = textRuns.map(run => {
        account(2); fontNodes += 2; accountText(run.text.length);
        const probe = document.createElement("span"); probe.dataset.officeScript = run.role; probe.textContent = run.text;
        probe.style.fontFamily = run.family;
        for (const name of ["font-size", "font-weight", "font-style", "font-variant"]) probe.style.setProperty(name, run.style.getPropertyValue(name));
        fontScope!.append(probe); return { run, probe };
      });
      source.append(fontScope);
      fonts = resolvePreviewFontsInScope(document, fontScope, "[data-office-script]", options.probes);
      for (const { run, probe } of pairs) run.node.setAttribute("font-family", view.getComputedStyle(probe).fontFamily);
    }
    // A private shadow measurement keeps source-wide SVG/text selectors out
    // while the final resolved fonts determine real label bounds.
    account(2); const measurement = document.createElement("span");
    measurement.style.cssText = "all:initial!important;position:absolute!important;visibility:hidden!important";
    measurement.dataset.officeView = "workbook-chart-label-measurement";
    measurement.attachShadow({ mode: "closed" }).append(element); source.append(measurement);
    try { layoutChartNumberFields(scene, accountText); prepareChartTextDecorations(scene, account, accountText); prepareChartLabelFrames(scene, account, workbookDrawingOrientation(source)); const labels = layoutBasicChartLabels(scene, account); if (labels) diagnostics.push(Object.freeze({ code: "chart-label-overflow", chartId: chart.chartId, labels })); }
    finally { element.remove(); measurement.remove(); account(-2); }
    return Object.freeze({ element, nodes: nodes - probes.length - fontNodes, peakNodes, diagnostics: Object.freeze(diagnostics), fonts });
  } catch (error) { element.replaceChildren(); throw error; }
  finally { fontScope?.remove(); for (const probe of probes) probe.remove(); }
}
