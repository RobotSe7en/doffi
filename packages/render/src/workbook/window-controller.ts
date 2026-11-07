import { DEFAULT_WORKBOOK_RULE_LIMITS, DocumentError } from "@doffi/core";
import type { WorkbookNumberFormatter, WorkbookRuleReadLimits, WorkbookRuleReadResponse, WorkbookSnapshot } from "@doffi/core";
import type { FontResolution } from "../contracts.js";
import type { WorkbookAssetSource } from "./assets.js";
import { projectWorkbookAnnotations } from "./annotation-projection.js";
import type { WorkbookAnnotationProjectionLimits, WorkbookAnnotationWindow } from "./annotation-projection.js";
import type { WorkbookChartSource } from "./chart-source.js";
import { measureWorkbookColumnFont } from "./column-font.js";
import type { WorkbookColumnFont } from "./column-font.js";
import { WorkbookContentSources } from "./content-source.js";
import { worksheetDimensions } from "./dimensions.js";
import { layoutWorkbookDrawings } from "./drawing-layout.js";
import type { WorkbookDrawingLayout, WorkbookDrawingLayoutLimits } from "./drawing-layout.js";
import type { WorkbookSurface } from "./surface.js";
import type { WorkbookViewport } from "./viewport.js";
import { layoutWorkbookWindow } from "./window-layout.js";
import type { WorkbookWindowLayout, WorkbookWindowLayoutOptions } from "./window-layout.js";
import { WorkbookWindowLoader } from "./window-loader.js";
import type { WorkbookViewportReadLimits, WorkbookWindowFrame } from "./window-loader.js";
import { checkWindowIdentity } from "./window-source.js";
import type { WorkbookWindowSource } from "./window-source.js";
import { checkedWorkbookRules } from "./rule-response.js";
import { workbookRuleDiagnostics } from "./rule-source.js";
import type { WorkbookRuleSource, WorkbookRuleDiagnostic } from "./rule-source.js";
import { maximumWorkbookSkew } from "./skew-sources.js";
import { measureWorkbookAutoRows } from "./auto-rows.js";
import type { WorkbookAutoRowLimits, WorkbookAutoRows } from "./auto-rows.js";
import type { WorkbookRowInterval } from "./row-invalidation.js";
import { WorkbookCellTextCache } from "./cell-text-cache.js";
import { WorkbookReadAhead, workbookIdle } from "./read-ahead.js";

export interface WorkbookWindowProviders {
  readonly cells: WorkbookWindowSource;
  readonly assets: WorkbookAssetSource;
  readonly charts: WorkbookChartSource;
  readonly rules?: WorkbookRuleSource;
}
export interface WorkbookWindowControllerOptions {
  /** Optional idle warming for sources that opt in. Zero disables it. */
  readonly readAheadRows?: number;
  /** Combined final cell/object nodes, independently of each phase's budget. */
  readonly maxNodes?: number;
  readonly read?: Partial<WorkbookViewportReadLimits>;
  readonly cells?: Omit<WorkbookWindowLayoutOptions, "normalSpaceWidthPx">;
  readonly drawings?: WorkbookDrawingLayoutLimits;
  readonly annotations?: WorkbookAnnotationProjectionLimits;
  readonly rules?: Partial<WorkbookRuleReadLimits>;
  readonly autoRows?: Partial<WorkbookAutoRowLimits>;
}
export interface WorkbookRenderedWindow {
  readonly frame: WorkbookWindowFrame;
  readonly element: HTMLElement;
  readonly nodes: number;
  readonly dimensions: ReturnType<typeof worksheetDimensions>;
  readonly cells: WorkbookWindowLayout["cells"];
  readonly objects: WorkbookDrawingLayout["objects"];
  readonly notes: WorkbookDrawingLayout["notes"];
  readonly annotations: WorkbookAnnotationWindow;
  readonly fonts: readonly FontResolution[];
  readonly diagnostics: readonly (WorkbookDrawingLayout["diagnostics"][number] | WorkbookRuleDiagnostic)[];
  readonly rules?: WorkbookRuleReadResponse;
  readonly usedFallbackLineMetrics: boolean;
}
export interface WorkbookWindowDecoration {
  /** May change as an owned panel/selection changes, within its assigned cap. */
  readonly nodes: number;
  readonly fonts?: readonly FontResolution[];
  dispose(): void;
}
export type WorkbookWindowDecorator = (candidate: WorkbookRenderedWindow, content: WorkbookContentSources, maxNodes: number) => WorkbookWindowDecoration;
interface WorkbookSheetMetrics extends WorkbookAutoRows { readonly normal: WorkbookColumnFont; readonly fontEpoch?: number; readonly fontSet?: object }
export interface WorkbookMeasurementSeed { readonly previous: WorkbookSheetMetrics; readonly rows: readonly WorkbookRowInterval[] }
export interface WorkbookWindowOwnership { readonly measurements?: ReadonlyMap<string, WorkbookMeasurementSeed>; readonly deferPublication?: boolean }

/** Internal lifecycle for one immutable workbook. The owner supplies a ready,
 * dedicated surface and authorized providers, and owns their disposal. This
 * controller owns only its window candidates and auxiliary source index.
 * A slow, failed or cancelled request never replaces the published window. */
export class WorkbookWindowController {
  private readonly snapshot: WorkbookSnapshot;
  private readonly content: WorkbookContentSources;
  private readonly options: WorkbookWindowControllerOptions;
  private readonly maxNodes: number;
  private readonly ruleLimits: WorkbookRuleReadLimits;
  private readonly skewCoefficient: number;
  private readonly classStyles: boolean;
  private readonly textCache: WorkbookCellTextCache;
  private readonly readAhead: WorkbookReadAhead;
  private readonly chunkRows = new Map<string, Set<number>>();
  private readonly interruptReadAhead = (event: Event): void => {
    if (event.type !== "wheel") this.clearPreparedGrid();
    this.readAhead.interrupt();
  };
  private readonly sheets = new Map<string, WorkbookSheetMetrics>();
  private readonly measurements: Map<string, WorkbookMeasurementSeed>;
  private deferred: boolean;
  private pending: { cancel(): void } | undefined;
  private active: { readonly state: WorkbookRenderedWindow; readonly drawings: WorkbookDrawingLayout; readonly decoration?: WorkbookWindowDecoration } | undefined;
  private disposed = false;
  private preparedGrid: { readonly frame: WorkbookWindowFrame; readonly cells: WorkbookWindowLayout; readonly fontRevision: number | undefined } | undefined;

  constructor(private readonly surface: WorkbookSurface, private readonly providers: WorkbookWindowProviders,
    private readonly formatter: WorkbookNumberFormatter, input: WorkbookWindowControllerOptions = {}, private readonly decorate: WorkbookWindowDecorator | undefined = undefined, ownership: WorkbookWindowOwnership = {}) {
    this.measurements = new Map(ownership.measurements); this.deferred = ownership.deferPublication ?? false;
    this.maxNodes = input.maxNodes ?? 50_000;
    this.ruleLimits = Object.freeze({ ...DEFAULT_WORKBOOK_RULE_LIMITS, ...input.rules });
    if (Object.keys(this.ruleLimits).some(key => !Object.hasOwn(DEFAULT_WORKBOOK_RULE_LIMITS, key)) || !Object.values(this.ruleLimits).every(value => Number.isSafeInteger(value) && value > 0))
      throw new DocumentError("InvalidOperation", "Workbook rules require supported positive safe-integer budgets.");
    if (!Number.isSafeInteger(this.maxNodes) || this.maxNodes < 1 || !Number.isSafeInteger(input.readAheadRows ?? 8) || (input.readAheadRows ?? 8) < 0 || (input.readAheadRows ?? 8) > 64
      || providers.cells.allowReadAhead !== undefined && typeof providers.cells.allowReadAhead !== "boolean"
      || Object.keys(input).some(key => !["maxNodes", "read", "cells", "drawings", "annotations", "rules", "autoRows", "readAheadRows"].includes(key))
      || !surface.source.isConnected || !surface.view.isConnected || surface.source.ownerDocument !== surface.document || surface.view.ownerDocument !== surface.document
      || surface.view.childNodes.length) throw new DocumentError("InvalidOperation", "A workbook controller requires its own empty connected surface and supported node budget.");
    this.snapshot = providers.cells.snapshot;
    this.classStyles = this.snapshot.styles.every(style => /^(?:\s*\.[a-zA-Z_][-\w]*\s*\{[^{}]*\}\s*)*$/.test(style.css));
    this.textCache = new WorkbookCellTextCache(input.read?.maxCells ?? 5000, Math.min(this.maxNodes, input.cells?.maxNodes ?? this.maxNodes), input.cells?.maxTextCodeUnits ?? 8 * 1024 * 1024);
    this.skewCoefficient = maximumWorkbookSkew(surface.document);
    checkWindowIdentity(providers.assets.snapshot.manifest, this.snapshot); checkWindowIdentity(providers.charts.snapshot.manifest, this.snapshot);
    if (providers.rules) checkWindowIdentity(providers.rules.snapshot.manifest, this.snapshot);
    this.options = Object.freeze({ ...input, read: Object.freeze({ ...input.read }), cells: Object.freeze({ ...input.cells }), annotations: Object.freeze({ ...input.annotations }), autoRows: Object.freeze({ ...input.autoRows }),
      drawings: Object.freeze({ ...input.drawings, ...(input.drawings?.assets ? { assets: Object.freeze({ ...input.drawings.assets }) } : {}) }) });
    // Validate read limits before retaining auxiliary content or starting work.
    const probe = new WorkbookWindowLoader(providers.cells, this.options.read, true); probe.dispose();
    this.content = new WorkbookContentSources(surface.document, this.snapshot);
    this.readAhead = new WorkbookReadAhead(surface.document.defaultView!);
    for (const chunk of this.snapshot.manifest.chunks) {
      let rows = this.chunkRows.get(chunk.sheetId); if (!rows) { rows = new Set(); this.chunkRows.set(chunk.sheetId, rows); } rows.add(chunk.row);
    }
    for (const name of ["pointerdown", "keydown", "wheel"]) surface.document.addEventListener(name, this.interruptReadAhead, { capture: true, passive: true });
  }

  get current(): WorkbookRenderedWindow | undefined { return this.active?.state; }

  /** Cancel only unfinished work while retaining the accepted window. */
  cancelPending(): void { this.readAhead.stop(); this.pending?.cancel(); this.pending = undefined; }
  cancelForReplacement(): void { this.cancelPending(); this.clearPreparedGrid(); }
  measurementSeeds(rows: ReadonlyMap<string, readonly WorkbookRowInterval[]>): ReadonlyMap<string, WorkbookMeasurementSeed> {
    const state = this.surface.fontState?.(), seeds = new Map<string, WorkbookMeasurementSeed>();
    if (!state?.ready) return seeds;
    for (const [sheetId, metrics] of this.sheets) if (metrics.reusable && metrics.fontEpoch === state.epoch && metrics.fontSet === state.identity) {
      seeds.set(sheetId, { previous: metrics, rows: rows.get(sheetId) ?? [] });
    }
    return seeds;
  }
  unpublish(): void { this.active?.state.element.remove(); }
  publish(): void {
    if (this.disposed || !this.active) throw new DocumentError("InvalidOperation", "A workbook window must be fully prepared before publication.");
    const state = this.active.state, viewport = state.frame.layout;
    this.surface.activate?.(); this.surface.view.replaceChildren(state.element);
    this.surface.frame.style.setProperty("width", viewport.widthPt + "pt", "important"); this.surface.frame.style.setProperty("height", viewport.heightPt + "pt", "important");
    this.surface.frame.style.setProperty("visibility", "visible", "important");
    if (this.deferred) { this.deferred = false; this.warm(state, undefined); }
  }

  /** False means superseded/disposed; genuine failures reject. A newer request
   * also settles an older request whose provider ignores its AbortSignal. */
  async render(sheetId: string, input: WorkbookViewport): Promise<boolean> {
    if (this.disposed) throw new DocumentError("InvalidOperation", "This workbook window controller has been disposed.");
    const viewport = Object.freeze({ widthPt: input.widthPt, heightPt: input.heightPt, scrollLeftPt: input.scrollLeftPt, scrollTopPt: input.scrollTopPt });
    if (![viewport.widthPt, viewport.heightPt, viewport.scrollLeftPt, viewport.scrollTopPt].every(Number.isFinite)
      || Math.min(viewport.widthPt, viewport.heightPt) <= 0 || Math.max(viewport.widthPt, viewport.heightPt) > 1_000_000)
      throw new DocumentError("InvalidOperation", "A workbook window requires finite positive viewport extents.");
    const sheet = this.snapshot.metadata.workbook.sheets.find(sheet => sheet.sheetId === sheetId);
    if (!sheet) throw new DocumentError("InvalidOperation", "The requested worksheet is not part of this workbook.");
    this.readAhead.stop(); this.pending?.cancel();
    const abort = new AbortController(), loader = new WorkbookWindowLoader(this.providers.cells, this.options.read, true, this.skewCoefficient);
    let cancel!: () => void;
    const cancelled = new Promise<false>(resolve => { cancel = () => { abort.abort(); loader.dispose(); resolve(false); }; });
    const request = { cancel }; this.pending = request;
    const current = (): boolean => !this.disposed && this.pending === request && !abort.signal.aborted;
    let candidateCells: WorkbookWindowLayout | undefined;
    let drawings: WorkbookDrawingLayout | undefined;
    let decoration: WorkbookWindowDecoration | undefined;
    try {
      let metrics = this.sheets.get(sheetId);
      if (!metrics) {
        const seed = this.measurements.get(sheetId); this.measurements.delete(sheetId);
        const normal = measureWorkbookColumnFont(this.surface.source, this.snapshot.metadata.workbook, sheetId);
        const state = this.surface.fontState?.();
        const reuse = seed && state?.ready && state.epoch === seed.previous.fontEpoch && state.identity === seed.previous.fontSet
          && normal.maximumDigitWidthPx === seed.previous.normal.maximumDigitWidthPx && normal.spaceWidthPx === seed.previous.normal.spaceWidthPx
          ? { previous: seed.previous, rows: seed.rows, guard: () => {
            const current = this.surface.fontState?.(); return !!current?.ready && current.epoch === state.epoch && current.identity === state.identity;
          } } : undefined;
        const measured = await Promise.race([measureWorkbookAutoRows(this.surface, this.providers.cells, sheetId, normal, this.formatter, abort.signal, this.options.autoRows,
          this.providers.rules ? { source: this.providers.rules, limits: this.ruleLimits } : undefined, reuse), cancelled]);
        if (measured === false || !current()) return false;
        metrics = { normal, ...measured, ...(state ? { fontEpoch: state.epoch, fontSet: state.identity } : {}) }; this.sheets.set(sheetId, metrics);
      }
      const { normal, dimensions } = metrics;
      if (!await Promise.race([loader.load(sheetId, dimensions, viewport), cancelled]) || !current()) return false;
      const frame = loader.frame!;
      let rules: WorkbookRuleReadResponse | undefined;
      if (sheet.conditionalRules?.length || sheet.dataValidations?.length || this.providers.rules?.currentValues) {
        if (!this.providers.rules) throw new DocumentError("InvalidOperation", "This worksheet requires an authorized current-value rule provider.");
        const request = Object.freeze({ documentId: frame.documentId, revision: frame.revision, sheetId, ...this.ruleLimits,
          coordinates: Object.freeze(frame.sourceCoordinates.map(point => Object.freeze({ row: point.row, column: point.column }))), ...(this.providers.rules.currentValues ? { includeCurrent: true } : {}) });
        const response = await Promise.race([this.providers.rules.read(request, abort.signal), cancelled]);
        if (response === false || !current()) return false;
        rules = checkedWorkbookRules(response, request, this.snapshot);
      }
      const annotations = projectWorkbookAnnotations(this.snapshot, frame, dimensions, this.options.annotations);
      const prepared = this.takePreparedGrid(frame);
      const cells = prepared ?? layoutWorkbookWindow(this.surface, this.snapshot, frame, dimensions, this.formatter,
        { ...this.options.cells, normalSpaceWidthPx: normal.spaceWidthPx, maxNodes: Math.min(this.maxNodes, this.options.cells?.maxNodes ?? this.maxNodes),
          maxCellFragments: Math.min(this.options.read?.maxCells ?? 5000, this.options.cells?.maxCellFragments ?? this.options.read?.maxCells ?? 5000) }, rules,
        this.classStyles && !sheet.tables?.length && !sheet.conditionalRules?.length ? this.textCache : undefined, this.textCache.probes);
      candidateCells = cells;
      const walker = this.surface.document.createTreeWalker(cells.element, 0xffff_ffff); let nodes = 1;
      while (walker.nextNode()) if (++nodes >= this.maxNodes) throw new DocumentError("ResourceLimitExceeded", "The combined workbook window exceeds its node budget.");
      if (nodes >= this.maxNodes) throw new DocumentError("ResourceLimitExceeded", "The combined workbook window has no remaining object-layer budget.");
      drawings = await layoutWorkbookDrawings(this.surface, this.snapshot, frame, dimensions, this.content, this.providers.assets, this.providers.charts,
        this.formatter, abort.signal, { ...this.options.drawings, maxNodes: Math.min(this.maxNodes - nodes, this.options.drawings?.maxNodes ?? this.maxNodes) });
      if (!current()) return false;
      nodes += drawings.nodes; cells.element.append(drawings.element);
      const base: WorkbookRenderedWindow = Object.freeze({ frame, element: cells.element, nodes, dimensions, cells: cells.cells, objects: drawings.objects, notes: drawings.notes, annotations,
        fonts: Object.freeze([...normal.fonts, ...metrics.fonts, ...cells.fonts, ...drawings.fonts]), diagnostics: Object.freeze([...drawings.diagnostics, ...rules ? workbookRuleDiagnostics(rules) : []]), ...(rules ? { rules } : {}),
        usedFallbackLineMetrics: metrics.usedFallbackLineMetrics || cells.usedFallbackLineMetrics || drawings.usedFallbackLineMetrics });
      decoration = this.decorate?.(base, this.content, this.maxNodes - nodes);
      if (decoration && (!Number.isSafeInteger(decoration.nodes) || decoration.nodes < 0 || nodes + decoration.nodes > this.maxNodes))
        throw new DocumentError("ResourceLimitExceeded", "Workbook controls exceed their assigned window-node budget.");
      if (!current()) return false;
      const extra = decoration;
      const state: WorkbookRenderedWindow = extra ? Object.freeze({ ...base, get nodes() { return base.nodes + extra.nodes; },
        get fonts() { return extra.fonts?.length ? Object.freeze([...base.fonts, ...extra.fonts]) : base.fonts; } }) : base;
      const previous = this.active;
      if (!this.deferred && cells.element.parentNode === this.surface.view) {
        for (const child of [...this.surface.view.childNodes]) if (child !== cells.element) child.remove();
        cells.element.style.removeProperty("visibility"); cells.element.style.removeProperty("pointer-events"); cells.element.removeAttribute("aria-hidden");
      } else if (!this.deferred) this.surface.view.replaceChildren(cells.element);
      candidateCells = undefined;
      this.surface.frame.style.setProperty("width", viewport.widthPt + "pt", "important");
      this.surface.frame.style.setProperty("height", viewport.heightPt + "pt", "important");
      if (!this.deferred) { this.surface.activate?.(); this.surface.frame.style.setProperty("visibility", "visible", "important"); }
      this.active = { state, drawings, ...(decoration ? { decoration } : {}) }; drawings = undefined; decoration = undefined;
      previous?.decoration?.dispose(); previous?.drawings.dispose(); previous?.state.element.replaceChildren();
      if (!this.deferred) this.warm(state, previous?.state);
      return true;
    } catch (error) {
      if (!current()) return false;
      throw error;
    } finally {
      decoration?.dispose(); drawings?.dispose(); candidateCells?.element.remove(); candidateCells?.element.replaceChildren(); loader.dispose();
      if (this.pending === request) this.pending = undefined;
    }
  }

  private warm(state: WorkbookRenderedWindow, previous: WorkbookRenderedWindow | undefined): void {
    try {
      const sheetId = state.frame.sheetId, sheet = this.snapshot.metadata.workbook.sheets.find(sheet => sheet.sheetId === sheetId)!;
      if (!this.providers.cells.allowReadAhead || !this.classStyles || !(this.options.readAheadRows ?? 8)
        || sheet.tables?.length || sheet.conditionalRules?.length || sheet.dataValidations?.length) return;
      const dimensions = state.dimensions, metrics = this.sheets.get(sheetId)!, layout = state.frame.layout;
      if (this.forecastGrid(state, previous)) return;
      const direction = previous?.frame.sheetId === sheetId && layout.scrollTopPt < previous.frame.layout.scrollTopPt ? -1 : 1;
      let first = Infinity, last = -1;
      for (const pane of layout.panes) if (pane.kind === "body" || pane.kind === "frozenColumns") {
        if (pane.rows.length > 128) return;
        for (const row of pane.rows) { first = Math.min(first, row.index); last = Math.max(last, row.index); }
      }
      if (last < 0) return;
      const rows: number[] = []; let edge = direction === 1 ? last : first;
      for (let index = 0; index < (this.options.readAheadRows ?? 8); index++) {
        const next = dimensions.rows.neighbor(edge, direction); if (next === undefined || next < sheet.frozen.rows) break;
        if (!this.chunkRows.get(sheetId)?.has(Math.floor(next / 256) * 256)) break;
        rows.push(next); edge = next;
      }
      const frozen = dimensions.rows.position(sheet.frozen.rows); let index = 0;
      if (!rows.length) return;
      this.readAhead.set(async signal => {
        if (signal.aborted || this.active?.state !== state) return true;
        const end = Math.min(index + 4, rows.length), row = Math.min(rows[index]!, rows[end - 1]!), lastRow = Math.max(rows[index]!, rows[end - 1]!);
        const loader = new WorkbookWindowLoader(this.providers.cells, {
          maxCells: Math.min(this.options.read?.maxCells ?? 5000, 128), maxSourceCells: Math.min(this.options.read?.maxSourceCells ?? 25000, 256),
          maxRowReadCells: Math.min(this.options.read?.maxRowReadCells ?? 250000, 4096), maxCodeUnits: Math.min(this.options.read?.maxCodeUnits ?? 8 * 1024 * 1024, 32768),
        }, true, this.skewCoefficient);
        const cancel = (): void => loader.dispose(); signal.addEventListener("abort", cancel, { once: true });
        try {
          if (!await loader.load(sheetId, dimensions, { widthPt: layout.widthPt, heightPt: frozen + dimensions.rows.position(lastRow + 1) - dimensions.rows.position(row),
            scrollLeftPt: layout.scrollLeftPt, scrollTopPt: dimensions.rows.position(row) - frozen }) || signal.aborted) return true;
          if (!await workbookIdle(this.surface.document.defaultView!, signal) || signal.aborted || this.active?.state !== state) return true;
          const started = performance.now();
          const cells = layoutWorkbookWindow(this.surface, this.snapshot, loader.frame!, dimensions, this.formatter, {
            normalSpaceWidthPx: metrics.normal.spaceWidthPx, maxNodes: Math.min(this.maxNodes, this.options.cells?.maxNodes ?? 50000, 2048),
            maxSourceNodes: Math.min(this.options.cells?.maxSourceNodes ?? 50000, 2048), maxTextCodeUnits: Math.min(this.options.cells?.maxTextCodeUnits ?? 8 * 1024 * 1024, 32768),
            maxCellFragments: Math.min(this.options.cells?.maxCellFragments ?? 5000, 128),
          }, undefined, this.textCache, this.textCache.probes);
          cells.element.replaceChildren(); index = end;
          return index < rows.length && performance.now() - started <= 8;
        } finally { signal.removeEventListener("abort", cancel); loader.dispose(); }
      });
    } catch { this.readAhead.stop(); }
  }

  private clearPreparedGrid(): void {
    this.preparedGrid?.cells.element.remove(); this.preparedGrid?.cells.element.replaceChildren(); this.preparedGrid = undefined;
  }

  private takePreparedGrid(frame: WorkbookWindowFrame): WorkbookWindowLayout | undefined {
    const prepared = this.preparedGrid; this.preparedGrid = undefined;
    if (!prepared) return;
    const before = prepared.frame, a = before.layout, b = frame.layout;
    const same = before.sheetId === frame.sheetId && before.documentId === frame.documentId && before.revision === frame.revision
      && a.widthPt === b.widthPt && a.heightPt === b.heightPt && a.scrollLeftPt === b.scrollLeftPt && a.scrollTopPt === b.scrollTopPt
      && this.textCache.settled(this.surface.document.fonts) && this.textCache.fontRevision === prepared.fontRevision
      && before.completeRows.length === frame.completeRows.length && before.completeRows.every((row, index) => row === frame.completeRows[index])
      && before.sourceCoordinates.length === frame.sourceCoordinates.length && before.sourceCoordinates.every((point, index) => {
        const coordinate = frame.sourceCoordinates[index]!;
        if (point.row !== coordinate.row || point.column !== coordinate.column) return false;
        const old = before.getCell(point.row, point.column), cell = frame.getCell(point.row, point.column);
        return old === cell || !!old && !!cell && old.html === cell.html && old.text === cell.text && old.cssClass === cell.cssClass && old.formatId === cell.formatId
          && old.value.type === cell.value.type && (old.value.type === "blank" || "value" in cell.value && Object.is(old.value.value, cell.value.value))
          && old.formula?.expression === cell.formula?.expression && old.formula?.cacheState === cell.formula?.cacheState;
      });
    if (same) return prepared.cells;
    prepared.cells.element.remove(); prepared.cells.element.replaceChildren(); return undefined;
  }

  /** One hidden, bounded grid only; fresh foreground reads establish identity
   * and every source value before its already measured DOM may be accepted. */
  private forecastGrid(state: WorkbookRenderedWindow, previous: WorkbookRenderedWindow | undefined): boolean {
    this.clearPreparedGrid();
    const layout = state.frame.layout, old = previous?.frame.layout;
    if (!old || previous.frame.sheetId !== state.frame.sheetId || old.widthPt !== layout.widthPt || old.heightPt !== layout.heightPt
      || old.scrollLeftPt !== layout.scrollLeftPt || this.snapshot.annotations.length || this.snapshot.metadata.workbook.objects?.length) return false;
    const delta = layout.scrollTopPt - old.scrollTopPt;
    if (!delta || !Number.isFinite(delta)) return false;
    const dimensions = state.dimensions, target = Math.max(0, layout.scrollTopPt + delta);
    const from = dimensions.rows.hit(layout.scrollTopPt), to = dimensions.rows.hit(target);
    if (from === undefined || to === undefined || Math.abs(to - from) > (this.options.readAheadRows ?? 8)) return false;
    const viewport = { widthPt: layout.widthPt, heightPt: layout.heightPt, scrollLeftPt: layout.scrollLeftPt, scrollTopPt: target };
    const metrics = this.sheets.get(state.frame.sheetId)!;
    this.readAhead.set(async signal => {
      const loader = new WorkbookWindowLoader(this.providers.cells, this.options.read, true, this.skewCoefficient);
      const cancel = (): void => loader.dispose(); signal.addEventListener("abort", cancel, { once: true });
      let cells: WorkbookWindowLayout | undefined;
      try {
        if (!await loader.load(state.frame.sheetId, dimensions, viewport) || signal.aborted || this.active?.state !== state) return false;
        if (!this.textCache.settled(this.surface.document.fonts)) return false;
        const fontRevision = this.textCache.fontRevision;
        cells = layoutWorkbookWindow(this.surface, this.snapshot, loader.frame!, dimensions, this.formatter,
          { ...this.options.cells, normalSpaceWidthPx: metrics.normal.spaceWidthPx, maxNodes: Math.min(this.maxNodes - state.nodes, this.options.cells?.maxNodes ?? this.maxNodes),
            maxCellFragments: Math.min(this.options.read?.maxCells ?? 5000, this.options.cells?.maxCellFragments ?? this.options.read?.maxCells ?? 5000) - state.cells.length },
          undefined, this.textCache, this.textCache.probes);
        if (signal.aborted || this.active?.state !== state || !this.textCache.settled(this.surface.document.fonts) || fontRevision !== this.textCache.fontRevision) return false;
        // Keep one inaccessible candidate in this fixed viewport so the
        // browser can retain its layout across acceptance, within live caps.
        cells.element.style.position = "absolute"; cells.element.style.left = "0"; cells.element.style.top = "0";
        cells.element.style.visibility = "hidden"; cells.element.style.pointerEvents = "none"; cells.element.setAttribute("aria-hidden", "true");
        this.surface.view.append(cells.element); cells.element.getBoundingClientRect();
        this.preparedGrid = { frame: loader.frame!, cells, fontRevision }; cells = undefined;
        return false;
      } finally { cells?.element.remove(); cells?.element.replaceChildren(); signal.removeEventListener("abort", cancel); loader.dispose(); }
    });
    return true;
  }

  dispose(): void {
    if (this.disposed) return;
    this.disposed = true; this.readAhead.dispose(); this.pending?.cancel(); this.pending = undefined;
    this.clearPreparedGrid();
    for (const name of ["pointerdown", "keydown", "wheel"]) this.surface.document.removeEventListener(name, this.interruptReadAhead, true);
    this.active?.decoration?.dispose(); this.active?.drawings.dispose(); this.active?.state.element.remove(); this.active?.state.element.replaceChildren(); this.active = undefined;
    this.content.dispose(); this.sheets.clear(); this.measurements.clear(); this.chunkRows.clear(); this.textCache.clear();
  }
}
