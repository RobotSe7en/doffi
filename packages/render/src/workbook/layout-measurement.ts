import { createWorkbookCalculator, createWorkbookNumberFormatter, DEFAULT_WORKBOOK_RULE_LIMITS, DocumentError } from "@doffi/core";
import type { WorkbookCalculationResult, WorkbookLayoutMeasurement, WorkbookLayoutMeasurer, WorkbookLayoutRequest } from "@doffi/core";
import { FontProbeCache } from "../font-probe-cache.js";
import { measureWorkbookAutoRows } from "./auto-rows.js";
import type { WorkbookAutoRows } from "./auto-rows.js";
import { measureWorkbookColumnFont } from "./column-font.js";
import { createWorkbookPreviewSource, workbookPreviewChangeImpact } from "./preview-source.js";
import { workbookRowInvalidation } from "./row-invalidation.js";
import type { WorkbookPreviewOptions, WorkbookPreviewSource } from "./preview-contracts.js";
import { createWorkbookSurface } from "./surface.js";
import type { WorkbookViewport } from "./viewport.js";

export interface WorkbookLayoutSourceRequest extends Omit<WorkbookLayoutRequest, "model" | "signal"> {
  readonly source: WorkbookPreviewSource;
  readonly signal?: AbortSignal;
}
export interface WorkbookLayoutService {
  readonly measure: WorkbookLayoutMeasurer;
  /** Worker adapters supply a scoped read endpoint for the private candidate. */
  measureSource(request: WorkbookLayoutSourceRequest, viewport?: Pick<WorkbookViewport, "widthPt" | "heightPt">): Promise<WorkbookLayoutMeasurement>;
  dispose(): void;
}
let nextFontEnvironment = 0;

/** The same scriptless source cascade, authorized local fonts, current rule
 * appearances and automatic row measurement as the grid. No rendered pixels
 * or persisted row-height copy become document source. */
export function createWorkbookLayoutService(container: HTMLElement, options: Pick<WorkbookPreviewOptions, "widthPt" | "heightPt" | "locale" | "limits"> = {}): WorkbookLayoutService {
  if (!container.isConnected || !container.ownerDocument.defaultView || !options || Object.keys(options).some(key => !["widthPt", "heightPt", "locale", "limits"].includes(key)))
    throw new DocumentError("InvalidOperation", "Workbook layout measurement requires a connected owner and supported options.");
  const widthPt = options.widthPt ?? 600, heightPt = options.heightPt ?? 360, locale = options.locale;
  const limits = { autoRows: { ...options.limits?.autoRows }, rules: { ...DEFAULT_WORKBOOK_RULE_LIMITS, ...options.limits?.rules } };
  const fonts = new FontProbeCache(), fontOwner = ++nextFontEnvironment, document = container.ownerDocument;
  let disposed = false, pending: AbortController | undefined;
  let cached: { surface: Awaited<ReturnType<typeof createWorkbookSurface>>; source: WorkbookPreviewSource; rows: ReadonlyMap<string, WorkbookAutoRows>; epoch: number; surfaceEpoch: number; size: string; locale: string } | undefined;
  let directCalculation: WorkbookCalculationResult | undefined;
  function check(signal: { readonly aborted: boolean }): void {
    if (disposed || signal.aborted) throw new DocumentError("OperationCancelled", "Workbook layout measurement was cancelled.");
  }
  async function run<Value>(signal: WorkbookLayoutRequest["signal"], work: (abort: AbortController) => Promise<Value>): Promise<Value> {
    if (disposed) throw new DocumentError("InvalidOperation", "This workbook layout service has been disposed.");
    if (pending) throw new DocumentError("ResourceLimitExceeded", "Another workbook layout measurement is pending.");
    const abort = new AbortController(), native = signal as AbortSignal | undefined, stop = (): void => abort.abort();
    pending = abort;
    if (signal?.aborted) abort.abort();
    else if (typeof native?.addEventListener === "function") native.addEventListener("abort", stop, { once: true });
    try { check(signal ?? abort.signal); const result = await work(abort); check(signal ?? abort.signal); return result; }
    catch (error) { cached?.surface.dispose(); cached = undefined; directCalculation = undefined; throw error; }
    finally { if (typeof native?.removeEventListener === "function") native.removeEventListener("abort", stop); if (pending === abort) pending = undefined; }
  }
  async function measureSource(request: WorkbookLayoutSourceRequest, abort: AbortController, viewport = { widthPt, heightPt }): Promise<WorkbookLayoutMeasurement> {
    check(abort.signal);
    const size = { widthPt: viewport.widthPt, heightPt: viewport.heightPt };
    const before = fonts.watch(document.fonts);
    if (!before.ready) throw new DocumentError("InvalidOperation", "Workbook layout measurement requires settled local fonts.");
    const snapshot = request.source.cells.snapshot, ids = [...request.sheetIds];
    if (!ids.length || new Set(ids).size !== ids.length || ids.some(id => !snapshot.metadata.workbook.sheets.some(sheet => sheet.sheetId === id)))
      throw new DocumentError("InvalidOperation", "Workbook layout measurement requires existing, distinct worksheets.");
    const sizeKey = JSON.stringify(size), localeKey = request.calculationContext?.locale ?? locale ?? snapshot.metadata.locale ?? "zh-CN";
    const surfaceFonts = cached?.surface.fontState?.();
    const previous = cached && surfaceFonts?.ready && surfaceFonts.epoch === cached.surfaceEpoch && cached.epoch === before.epoch && cached.size === sizeKey && cached.locale === localeKey ? cached : undefined;
    const invalidation = previous ? workbookRowInvalidation(previous.source.cells.snapshot, snapshot, workbookPreviewChangeImpact(previous.source, request.source), limits.autoRows.maxCells ?? 2_500_000) : undefined;
    if (!invalidation) { cached?.surface.dispose(); cached = undefined; }
    const surface = cached?.surface ?? await createWorkbookSurface(container, snapshot, { ...size, signal: abort.signal });
    const localFonts = surface.fontState?.();
    const sameFonts = (): boolean => {
      const main = fonts.watch(document.fonts), local = surface.fontState?.();
      return main.ready && main.epoch === before.epoch && local?.ready === true && local.epoch === localFonts?.epoch;
    };
    const measuredRows = new Map<string, WorkbookAutoRows>(); let accepted = false;
    const sheets: WorkbookLayoutMeasurement["sheets"][number][] = [];
    try {
      const formatter = createWorkbookNumberFormatter({ dateSystem: snapshot.metadata.workbook.dateSystem,
        ...(request.calculationContext?.locale !== undefined ? { hostLocale: request.calculationContext.locale } : locale !== undefined ? { hostLocale: locale } : {}),
        ...(snapshot.metadata.locale ? { documentLocale: snapshot.metadata.locale } : {}) });
      try {
        for (const sheetId of ids) {
          check(request.signal ?? abort.signal);
          const prior = previous?.rows.get(sheetId);
          const changedRows = invalidation?.get(sheetId) ?? [];
          const rows = invalidation && prior?.reusable && !changedRows.length ? prior : await measureWorkbookAutoRows(surface, request.source.cells, sheetId,
            measureWorkbookColumnFont(surface.source, snapshot.metadata.workbook, sheetId), formatter, abort.signal, limits.autoRows,
            request.source.rules ? { source: request.source.rules, limits: limits.rules } : undefined,
            invalidation && prior ? { previous: prior, rows: changedRows, guard: () => !abort.signal.aborted && sameFonts() } : undefined);
          measuredRows.set(sheetId, rows);
          sheets.push({ sheetId, rows: { defaultSizePt: 0, overrides: rows.dimensions.rows.runs() }, columns: { defaultSizePt: 0, overrides: rows.dimensions.columns.runs() } });
        }
        const after = fonts.watch(document.fonts);
        if (!after.ready || before.epoch !== after.epoch || !sameFonts()) throw new DocumentError("InvalidOperation", "Local workbook fonts changed during measurement.");
        check(request.signal ?? abort.signal);
        cached = { surface, source: request.source, rows: measuredRows, epoch: before.epoch, surfaceEpoch: localFonts!.epoch, size: sizeKey, locale: localeKey }; accepted = true;
        return Object.freeze({ sourceDigest: request.sourceDigest, fontEnvironment: `local-fonts-${fontOwner}-${before.epoch}-${size.widthPt}-${size.heightPt}`, sheets: Object.freeze(sheets) });
      } finally { formatter.dispose(); }
    } finally { if (!accepted) { surface.dispose(); cached = undefined; } }
  }
  return Object.freeze({
    measure: (request: WorkbookLayoutRequest) => run(request.signal, async abort => {
      const calculator = request.model.snapshot().metadata.workbook.calculationMode === "auto" ? createWorkbookCalculator(request.model, {}, directCalculation) : undefined;
      try {
        const calculation = calculator ? await calculator.calculate({ signal: abort.signal, ...(request.calculationContext ? { context: request.calculationContext } : {}),
          yield: async () => { await new Promise<void>(resolve => setTimeout(resolve, 0)); check(request.signal ?? abort.signal); } }) : undefined;
        check(request.signal ?? abort.signal);
        const source = createWorkbookPreviewSource(request.model, calculation, cached?.source);
        const result = await measureSource({ source, sourceDigest: request.sourceDigest, sheetIds: request.sheetIds,
          ...(request.calculationContext ? { calculationContext: request.calculationContext } : {}), signal: abort.signal }, abort);
        directCalculation = calculation; return result;
      } finally { calculator?.dispose(); }
    }),
    measureSource: (request: WorkbookLayoutSourceRequest, viewport?: Pick<WorkbookViewport, "widthPt" | "heightPt">) => run(request.signal, abort => measureSource(request, abort, viewport)),
    dispose() { if (disposed) return; disposed = true; pending?.abort(); cached?.surface.dispose(); cached = undefined; directCalculation = undefined; fonts.clear(); },
  });
}
