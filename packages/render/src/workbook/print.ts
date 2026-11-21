import { captureWorkbookModel, collectWorkbookFilteredRows, copyWorkbookScalarContext, createWorkbookCalculatedEvaluator, createWorkbookCalculator, createWorkbookNumberFormatter,
  DocumentError, iterateWorkbookRows, WorkbookRangeSchema } from "@doffi/core";
import type { DeepReadonly, WorkbookCalculationResult, WorkbookModel, WorkbookRange, WorkbookSheet } from "@doffi/core";
import { renderLimits } from "../contracts.js";
import type { FontResolution, LayoutDiagnostic, RenderLimits, WorkbookPrintedPage, WorkbookPrintReport, WorkbookPrintSettings } from "../contracts.js";
import { collectDocumentFontNames, loadWordWebFonts } from "../fonts.js";
import type { WordWebFont } from "../fonts.js";
import { AxisDimensions } from "./dimensions.js";
import type { AxisOverride } from "./dimensions.js";
import { workbookDrawingRect } from "./drawing.js";
import { workbookColumnLabel } from "./preview-chrome.js";
import { createWorkbookPreviewSource } from "./preview-source.js";
import { paintWorkbookPrintBands, workbookPrintFields } from "./print-bands.js";
import { planWorkbookPrintSheet, PRINT_HEADING_HEIGHT, PRINT_HEADING_WIDTH, workbookPaper, workbookPrintLimit, workbookPrintViewport } from "./print-layout.js";
import type { PrintDimensions, WorkbookPrintTile } from "./print-layout.js";
import { waitForWorkbookRead } from "./read-cancellation.js";
import { modelChartSource } from "./chart-source.js";
import { modelRuleSource } from "./rule-source.js";
import { createWorkbookSurface } from "./surface.js";
import { WorkbookWindowController } from "./window-controller.js";
import type { WorkbookRenderedWindow } from "./window-controller.js";
import { checkWindowCancellation } from "./window-source.js";

const MAX_WORKBOOK_PRINT_PAPER_PIXELS = 256 * 1024 * 1024;
const MAX_PAGE_CELLS = 5000;
const cachedBadge = { zh: "未重算 · 使用公式缓存", en: "Not recalculated · cached formula values" };
let nextCanvas = 0;

export interface WorkbookPrintOptions extends WorkbookPrintSettings {
  readonly calculation?: WorkbookCalculationResult;
  readonly signal?: AbortSignal;
  readonly limits?: Partial<RenderLimits>;
  readonly webFonts?: readonly WordWebFont[];
}
export interface WorkbookPrintDocument {
  readonly element: HTMLElement;
  readonly pages: readonly HTMLElement[];
  readonly report: WorkbookPrintReport;
  dispose(): void;
}

type Sheet = DeepReadonly<WorkbookSheet>;
type Plan = { readonly page: WorkbookPrintedPage; readonly sheet: Sheet; readonly sheetIndex: number; readonly dimensions: PrintDimensions; readonly layout: ReturnType<typeof workbookPrintViewport> };
const nextTask = (): Promise<void> => new Promise(resolve => setTimeout(resolve, 0));

function selectSheets(sheets: readonly Sheet[], settings: WorkbookPrintSettings, maxPages: number) {
  if (settings.sheets === undefined) return sheets.filter(sheet => sheet.visibility === "visible").map(sheet => ({ sheet, areas: sheet.print?.areas }));
  if (!Array.isArray(settings.sheets) || settings.sheets.length === 0) throw new DocumentError("InvalidOperation", "Select at least one worksheet to print.");
  if (settings.sheets.length > maxPages) workbookPrintLimit("Too many worksheets were selected.");
  const seen = new Set<string>();
  return settings.sheets.map(selection => {
    if (!selection || typeof selection !== "object" || Object.keys(selection).some(key => key !== "sheetId" && key !== "areas"))
      throw new DocumentError("InvalidOperation", "Invalid workbook print selection.");
    const sheet = sheets.find(candidate => candidate.sheetId === selection.sheetId);
    if (!sheet || sheet.visibility !== "visible" || seen.has(sheet.sheetId))
      throw new DocumentError("InvalidOperation", "Print selection must contain distinct visible worksheets.");
    seen.add(sheet.sheetId);
    if (selection.areas !== undefined && (!Array.isArray(selection.areas) || !selection.areas.length))
      throw new DocumentError("InvalidOperation", "A print area override must contain at least one range.");
    if (selection.areas && selection.areas.length > maxPages) workbookPrintLimit("Too many print areas were selected.");
    const areas = selection.areas?.map((area: DeepReadonly<WorkbookRange>) => {
      const checked = WorkbookRangeSchema.safeParse(area);
      if (!checked.success || checked.data.from.row > checked.data.to.row || checked.data.from.column > checked.data.to.column)
        throw new DocumentError("InvalidOperation", "Invalid workbook print range.");
      return checked.data;
    });
    return { sheet, areas: areas ?? sheet.print?.areas };
  });
}

async function usedArea(model: WorkbookModel, sheet: Sheet, dimensions: PrintDimensions, signal: AbortSignal): Promise<WorkbookRange> {
  const from = { row: 1_048_576, column: 16_384 }, to = { row: -1, column: -1 };
  const include = (r: number, c: number, b: number, e: number): void => {
    from.row = Math.min(from.row, r); from.column = Math.min(from.column, c);
    to.row = Math.max(to.row, b); to.column = Math.max(to.column, e);
  };
  let rows = 0;
  for (const cells of iterateWorkbookRows(model, sheet.sheetId)) {
    if (++rows % 128 === 0) { await nextTask(); checkWindowCancellation(signal); }
    for (const cell of cells) if (dimensions.rows.size(cell.row) && dimensions.columns.size(cell.column))
      include(cell.row, cell.column, cell.row, cell.column);
  }
  for (const { from: a, to: b } of sheet.merges) if (dimensions.rows.position(b.row + 1) > dimensions.rows.position(a.row)
    && dimensions.columns.position(b.column + 1) > dimensions.columns.position(a.column)) include(a.row, a.column, b.row, b.column);
  const index = (axis: PrintDimensions["rows"], point: number): number | undefined =>
    axis.hit(Math.min(axis.extentPt - Math.max(1e-7, axis.extentPt * Number.EPSILON), Math.max(0, point)));
  for (const object of model.snapshot().metadata.workbook.objects ?? []) {
    if (object.sheetId !== sheet.sheetId) continue;
    const { x, y, width, height } = workbookDrawingRect(object.anchor, dimensions);
    if (x >= dimensions.columns.extentPt || y >= dimensions.rows.extentPt || x + width <= 0 || y + height <= 0) continue;
    const r = index(dimensions.rows, y), c = index(dimensions.columns, x);
    const b = index(dimensions.rows, y + height - 1e-7), e = index(dimensions.columns, x + width - 1e-7);
    if (r !== undefined && c !== undefined && b !== undefined && e !== undefined) include(r, c, b, e);
  }
  return to.row < 0 ? { from: { row: 0, column: 0 }, to: { row: 0, column: 0 } } : { from, to };
}

function applyFilteredRows(dimensions: PrintDimensions, filtered?: ReadonlySet<number>): PrintDimensions {
  if (!filtered?.size) return dimensions;
  const sorted = [...filtered].sort((a, b) => a - b), overrides: AxisOverride[] = [];
  let index = 0;
  for (const run of dimensions.rows.runs()) {
    let start = run.from;
    while (index < sorted.length && sorted[index]! <= run.to) {
      const row = sorted[index++]!;
      if (!run.sizePt) continue;
      if (start < row) overrides.push({ from: start, to: row - 1, sizePt: run.sizePt });
      overrides.push({ from: row, to: row, sizePt: 0 });
      start = row + 1;
    }
    if (start <= run.to) overrides.push({ from: start, to: run.to, sizePt: run.sizePt });
  }
  return { rows: new AxisDimensions(dimensions.rows.count, 0, overrides), columns: dimensions.columns };
}

function printedPage(setup: ReturnType<typeof workbookPaper>, sheet: Sheet, tile: WorkbookPrintTile, scale: number, index: number, y: number, number: number): WorkbookPrintedPage {
  return Object.freeze({
    pageId: "workbook-page-" + index, index, sectionId: sheet.sheetId, sheetId: sheet.sheetId, areaIndex: tile.areaIndex,
    number, label: String(number), rect: Object.freeze({ x: 0, y, width: setup.width, height: setup.height }),
    setup: Object.freeze({ ...setup, margins: Object.freeze(setup.margins) }), scale,
    blank: !tile.rows.extent || !tile.columns.extent,
    range: Object.freeze({ from: Object.freeze({ row: tile.rows.body.from, column: tile.columns.body.from }),
      to: Object.freeze({ row: tile.rows.body.to, column: tile.columns.body.to }) }),
  });
}

function paintHeadings(target: HTMLElement, layout: Plan["layout"], rtl: boolean, account: (count: number) => void): void {
  const document = target.ownerDocument, columnAxis = new Map<number, { start: number; size: number }>(), rowAxis = new Map<number, { start: number; size: number }>();
  for (const pane of layout.panes) {
    for (const col of pane.columns) columnAxis.set(col.index, { start: col.startPt - pane.sourceLeftPt + pane.inlineStartPt, size: col.sizePt });
    for (const row of pane.rows) rowAxis.set(row.index, { start: row.startPt - pane.sourceTopPt + pane.blockStartPt, size: row.sizePt });
  }
  const append = (text: string, left: number, top: number, width: number, height: number): void => {
    account(text ? 2 : 1);
    const box = document.createElement("span");
    box.textContent = text;
    box.style.cssText = `position:absolute;left:${left}pt;top:${top}pt;width:${width}pt;height:${height}pt;box-sizing:border-box;border:.5pt solid #aaa;overflow:hidden;white-space:nowrap;text-align:center;color:black;background:white;font:9pt/${height}pt Arial,sans-serif`;
    target.append(box);
  };
  const rowLeft = rtl ? layout.widthPt : 0, gridLeft = rtl ? 0 : PRINT_HEADING_WIDTH;
  append("", rowLeft, 0, PRINT_HEADING_WIDTH, PRINT_HEADING_HEIGHT);
  for (const [index, col] of columnAxis) {
    const size = Math.min(col.size, layout.widthPt - col.start);
    append(workbookColumnLabel(index), gridLeft + (rtl ? layout.widthPt - col.start - size : col.start), 0, size, PRINT_HEADING_HEIGHT);
  }
  for (const [index, row] of rowAxis)
    append(String(index + 1), rowLeft, PRINT_HEADING_HEIGHT + row.start, PRINT_HEADING_WIDTH, Math.min(row.size, layout.heightPt - row.start));
}

function paperElement(document: Document, { page, sheet, layout }: Plan, name: string, grid: HTMLElement, account: (count: number) => void) {
  const { setup, scale } = page, { width, height, margins } = setup, { print, rightToLeft } = sheet;
  const paper = document.createElement("div");
  paper.dataset.workbookPrintPage = String(page.index);
  paper.style.cssText = `all:initial;display:block;position:relative;overflow:hidden;box-sizing:border-box;background:white;width:${width}pt;height:${height}pt;page:${name};break-after:page;print-color-adjust:exact;-webkit-print-color-adjust:exact`;
  const target = paper.attachShadow({ mode: "open" });
  const ink = document.createElement("div");
  ink.style.cssText = `position:absolute;inset:0;${print?.blackAndWhite ? "filter:grayscale(1)" : ""}`;
  target.append(ink);
  const headings = print?.headings && !page.blank;
  const gridWidth = layout.widthPt + (headings ? PRINT_HEADING_WIDTH : 0), gridHeight = layout.heightPt + (headings ? PRINT_HEADING_HEIGHT : 0);
  const printableWidth = width - margins.left - margins.right, printableHeight = height - margins.top - margins.bottom;
  const dx = Math.max(0, printableWidth - gridWidth * scale), dy = Math.max(0, printableHeight - gridHeight * scale);
  const left = margins.left + (print?.horizontalCentered ? dx / 2 : rightToLeft ? dx : 0);
  const top = margins.top + (print?.verticalCentered ? dy / 2 : 0);
  const body = document.createElement("div");
  body.style.cssText = `position:absolute;left:${left}pt;top:${top}pt;width:${gridWidth}pt;height:${gridHeight}pt;overflow:hidden;transform:scale(${scale});transform-origin:top left`;
  grid.style.position = "absolute";
  grid.style.left = (headings && !rightToLeft ? PRINT_HEADING_WIDTH : 0) + "pt";
  grid.style.top = (headings ? PRINT_HEADING_HEIGHT : 0) + "pt";
  body.append(grid);
  if (headings) paintHeadings(body, layout, rightToLeft, account);
  ink.append(body);
  account(4);
  return { paper, target, ink };
}

function renderedFormulaStates(window: WorkbookRenderedWindow, cached: boolean, seen: Set<string>, counts: Record<"valid" | "imported" | "stale" | "unavailable", number>): void {
  const { frame } = window, current = new Map(window.rules?.cells.map(cell => [cell.row * 16_384 + cell.column, cell.current.kind]));
  for (const cell of window.cells) {
    const { row, column } = cell.projection, value = frame.getCell(row, column);
    if (!value?.formula) continue;
    const identity = `${frame.sheetId}:${row}:${column}`;
    if (seen.has(identity)) continue;
    seen.add(identity);
    if (cached) counts[value.formula.cacheState]++;
    else if (current.get(row * 16_384 + column) === "value") counts.valid++;
    else throw new DocumentError("UnsupportedFeature", "A printed worksheet formula has no current result. Select a different area or request the formula-cache snapshot.");
  }
  if (!cached && window.diagnostics.some(item => item.code === "unavailable-cell" || item.code === "stale-cell" || item.code === "imported-cell"))
    throw new DocumentError("UnsupportedFeature", "A printed chart has no current formula result. Select a different area or request the formula-cache snapshot.");
}

/** Captures one revision and composes bounded printable pages with the grid's
 * normal cell, text, border, rule and drawing painters. The caller owns the
 * returned view and its image/font leases until PDF or image capture finishes. */
export async function renderWorkbookPrint(container: HTMLElement, input: WorkbookModel, options: WorkbookPrintOptions = {}): Promise<WorkbookPrintDocument> {
  if (!container.isConnected || !container.ownerDocument.defaultView || input.kind !== "workbook")
    throw new DocumentError("InvalidOperation", "Workbook printing requires a connected owner and workbook.");
  const document = container.ownerDocument, model = captureWorkbookModel(input), snapshot = model.snapshot(), workbook = snapshot.metadata.workbook;
  const limits = renderLimits(options.limits), maxPaperPixels = options.maxPaperPixels ?? MAX_WORKBOOK_PRINT_PAPER_PIXELS;
  const mode = options.formulaValues ?? "current";
  if (!["current", "cached"].includes(mode) || mode === "cached" && options.calculation !== undefined || !Number.isSafeInteger(maxPaperPixels) || maxPaperPixels < 1
    || maxPaperPixels > MAX_WORKBOOK_PRINT_PAPER_PIXELS || [options.fileName, options.filePath].some(text => text !== undefined && (typeof text !== "string" || text.length > 4096)))
    throw new DocumentError("InvalidOperation", "Invalid workbook print settings.");
  const context = copyWorkbookScalarContext({ now: { unixMilliseconds: Date.now(), utcOffsetMinutes: 0 }, ...options.context });
  const settings = { fileName: options.fileName ?? "", filePath: options.filePath ?? "", context };
  const formatter = createWorkbookNumberFormatter({ dateSystem: workbook.dateSystem, ...(context.locale ? { hostLocale: context.locale } : {}),
    ...(snapshot.metadata.locale ? { documentLocale: snapshot.metadata.locale } : {}) });
  const locale = context.locale ?? snapshot.metadata.locale ?? "zh-CN";
  const selected = selectSheets(workbook.sheets, options, limits.maxPages);
  if (selected.length > limits.maxPages) workbookPrintLimit("Too many worksheets were selected.");
  const signal = options.signal ?? new AbortController().signal;
  const check = (): void => {
    checkWindowCancellation(signal);
    if (captureWorkbookModel(input) !== model) throw new DocumentError("RevisionConflict", "The workbook changed during print rendering.");
  };
  check();
  const releases: (() => void)[] = [], faces: FontFace[] = [];
  const root = document.createElement("div"), pages: HTMLElement[] = [];
  root.style.cssText = "all:initial;display:block;margin:0;padding:0;direction:ltr;position:relative";
  let surface: Awaited<ReturnType<typeof createWorkbookSurface>> | undefined, controller: WorkbookWindowController | undefined;
  const dispose = (): void => {
    root.remove();
    for (const release of releases.splice(0)) release();
    for (const face of faces.splice(0)) document.fonts.delete(face);
    root.replaceChildren();
    controller?.dispose();
    surface?.dispose();
  };
  try {
    let calculation = options.calculation;
    if (calculation !== undefined) createWorkbookCalculatedEvaluator(model, calculation).dispose();
    if (mode === "current" && !calculation) {
      const calculator = createWorkbookCalculator(model);
      try { calculation = await calculator.calculate({ signal, context, yield: nextTask }); }
      finally { calculator.dispose(); }
    }
    check();
    surface = await createWorkbookSurface(container, snapshot, { widthPt: 960, heightPt: 675, signal });
    if (options.webFonts?.length) {
      const existing = new Set(surface.document.fonts);
      const requested = collectDocumentFontNames({ html: snapshot.content?.html ?? "", styles: snapshot.styles, metadata: snapshot.metadata });
      await waitForWorkbookRead(loadWordWebFonts(surface.document, options.webFonts, requested), signal);
      for (const face of surface.document.fonts) if (!existing.has(face)) {
        document.fonts.add(face);
        faces.push(face);
      }
      await waitForWorkbookRead(Promise.all([document.fonts.ready, surface.document.fonts.ready]), signal);
    }
    const filtered = new Map<string, ReadonlySet<number>>(), diagnostics: LayoutDiagnostic[] = [];
    for (const { sheet } of selected) {
      const result = await collectWorkbookFilteredRows(model, sheet.sheetId, { signal, yield: nextTask,
        ...(calculation ? { calculation } : { context, cachedValues: true }) });
      if (result.rows) filtered.set(sheet.sheetId, result.rows);
      if (result.unresolved) {
        if (mode === "current") throw new DocumentError("UnsupportedFeature", "A worksheet filter has no current formula result.");
        diagnostics.push({ code: "NativeFormatting", message: `Worksheet ${sheet.name}: some filtered rows have no saved result.` });
      }
    }
    check();
    const base = createWorkbookPreviewSource(model, calculation), cached = mode === "cached" ? context : undefined;
    const providers = { ...base, ...(cached ? { rules: modelRuleSource(model, undefined, cached) } : {}),
      charts: modelChartSource(model, calculation, cached, filtered) };
    const cellBudget = Math.min(MAX_PAGE_CELLS, Math.floor(limits.maxPrintNodes / 2) || 1);
    controller = new WorkbookWindowController(surface, providers, formatter,
      { maxNodes: Math.min(limits.maxNodes, limits.maxPrintNodes), readAheadRows: 0, read: { maxCells: cellBudget } }, undefined, { deferPublication: true });
    const cancel = (): void => controller?.cancelPending();
    signal.addEventListener("abort", cancel, { once: true });
    try {
      const plans: Plan[] = [], fonts = new Map<string, FontResolution>();
      const collectFonts = (values: readonly FontResolution[]): void => { for (const item of values) fonts.set(JSON.stringify(item), item); };
      let y = 0, number = 1, pixels = 0, generated = 0;
      for (const { sheet, areas } of selected) {
        check();
        const metrics = await controller.measureSheet(sheet.sheetId, signal);
        const dimensions = applyFilteredRows(metrics.dimensions, filtered.get(sheet.sheetId));
        const setup = workbookPaper(sheet.print), sourceAreas = areas?.length ? areas : [await usedArea(model, sheet, dimensions, signal)];
        const { scale, tiles } = planWorkbookPrintSheet(dimensions, sourceAreas, sheet.print, setup, limits.maxPages - plans.length);
        number = sheet.print?.firstPage ?? number;
        pixels += tiles.length * Math.ceil(setup.width * 4 / 3) * Math.ceil(setup.height * 4 / 3);
        if (pixels > maxPaperPixels) workbookPrintLimit("Workbook print layout exceeds its paper pixel budget.");
        for (const [sheetIndex, tile] of tiles.entries()) {
          const page = printedPage(setup, sheet, tile, scale, plans.length, y, number++);
          const layout = workbookPrintViewport(dimensions, tile, cellBudget);
          if (tile.rows.clipped || tile.columns.clipped) diagnostics.push({ code: "ContentOverflow", message: `Page ${page.index + 1} clips an oversized worksheet dimension.` });
          generated += layout.cellCount * 2 + 4;
          if (generated > limits.maxPrintNodes) workbookPrintLimit("Workbook print layout exceeds its node budget.");
          plans.push({ page, sheet, sheetIndex, dimensions, layout });
          y += setup.height;
        }
      }
      const canvas = ++nextCanvas, style = document.createElement("style");
      style.textContent = plans.map(({ page }) => `@page doffi-workbook-${canvas}-${page.index}{size:${page.setup.width}pt ${page.setup.height}pt;margin:0}`).join("\n");
      root.append(style);
      let nodes = 3, imageBytes = 0, imagePixels = 0;
      const account = (count: number): void => { nodes += count; if (nodes > limits.maxPrintNodes) workbookPrintLimit("Workbook print projection exceeds its node budget."); };
      const complete = { valid: 0, imported: 0, stale: 0, unavailable: 0 }, seen = new Set<string>();
      for (const plan of plans) {
        check();
        const { page, sheet, sheetIndex, dimensions, layout } = plan;
        let grid = document.createElement("div");
        if (!page.blank) {
          if (!await controller.render(sheet.sheetId, layout, { layout, dimensions, gridLines: sheet.print?.gridLines ?? false }))
            throw new DocumentError("OperationCancelled", "Workbook print rendering was cancelled.");
          const lease = controller.takePrintWindow();
          releases.push(lease.dispose);
          renderedFormulaStates(lease.state, mode === "cached", seen, complete);
          collectFonts(lease.state.fonts);
          for (const issue of lease.state.diagnostics) {
            if ("chartId" in issue || !("sheetId" in issue) || issue.code === "conditional-rule-unavailable" || issue.code === "conditional-formula-error")
              diagnostics.push({ code: "NativeFormatting", ...("nodeId" in issue ? { nodeId: issue.nodeId } : {}), message: `Worksheet ${sheet.name}, page ${page.index + 1}: ${issue.code}.` });
          }
          imageBytes += lease.assets.byteLength; imagePixels += lease.assets.pixels;
          if (imageBytes > 64 * 1024 * 1024 || imagePixels > MAX_WORKBOOK_PRINT_PAPER_PIXELS) workbookPrintLimit("Workbook print projection exceeds its image budget.");
          account(lease.state.nodes);
          grid = document.adoptNode(lease.state.element) as HTMLDivElement;
        } else account(1);
        const { paper, target, ink } = paperElement(document, plan, `doffi-workbook-${canvas}-${page.index}`, grid, account);
        if (sheet.print) {
          const fields = workbookPrintFields(sheet, page, plans.length, settings, locale);
          collectFonts(paintWorkbookPrintBands(ink, surface, page, sheetIndex, sheet.print, workbook.themeCssClass, fields, account));
        }
        if (mode === "cached") {
          account(2);
          const badge = document.createElement("span");
          badge.dataset.workbookCalculation = "cached";
          badge.style.cssText = "position:absolute;right:3pt;bottom:2pt;background:white;color:#563600;font:bold 9pt Arial,sans-serif";
          badge.textContent = /^zh\b/i.test(locale) ? cachedBadge.zh : cachedBadge.en;
          target.append(badge);
        }
        pages.push(paper);
        root.append(paper);
      }
      check();
      pages.at(-1)?.style.setProperty("break-after", "auto");
      container.append(root);
      const report: WorkbookPrintReport = Object.freeze({ kind: "workbook", layout: "pages", documentId: snapshot.manifest.documentId, revision: snapshot.manifest.revision,
        pages: Object.freeze(plans.map(plan => plan.page)), nodePages: Object.freeze({}), fields: Object.freeze([]),
        diagnostics: Object.freeze(diagnostics), fonts: Object.freeze([...fonts.values()]), formulaValues: mode, formulaCells: Object.freeze(complete) });
      return Object.freeze({ element: root, pages: Object.freeze(pages), report, dispose });
    } finally { signal.removeEventListener("abort", cancel); }
  } catch (error) {
    dispose();
    if (error instanceof DocumentError && error.code === "ResourceLimitExceeded" && !error.message.includes("smaller workbook print area"))
      workbookPrintLimit(error.message);
    throw error;
  }
}
