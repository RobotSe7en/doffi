import { createWorkbookNumberFormatter, DocumentError, WorkbookRangeSchema } from "@doffi/core";
import type { DeepReadonly, DocumentSelection, WorkbookCoordinate, WorkbookRange, WorkbookSnapshot } from "@doffi/core";
import { createWorkbookPreviewChrome, workbookColumnLabel } from "./preview-chrome.js";
import type { WorkbookPreview, WorkbookPreviewInput, WorkbookPreviewOptions, WorkbookRenderReport } from "./preview-contracts.js";
import { decorateWorkbookPreview } from "./preview-decoration.js";
import type { WorkbookPreviewDecoration } from "./preview-decoration.js";
import { workbookPreviewChangeImpact, workbookPreviewSource } from "./preview-source.js";
import { hitWorkbookCell, WorkbookMergeIndex } from "./projection.js";
import { createWorkbookSurface } from "./surface.js";
import type { WorkbookSurface } from "./surface.js";
import { WorkbookWindowController } from "./window-controller.js";
import type { WorkbookRenderedWindow } from "./window-controller.js";
import type { WorkbookViewport } from "./viewport.js";
import { navigateWorkbookSelection } from "./selection-navigation.js";
import type { WorkbookSelectionNavigation } from "./selection-navigation.js";
import { workbookRowInvalidation } from "./row-invalidation.js";
import { DEFAULT_AUTO_ROW_LIMITS } from "./auto-rows.js";
import { waitForWorkbookRead } from "./read-cancellation.js";

interface PreviewWindow { readonly controls: WorkbookPreviewDecoration; commitChrome(): void }
interface PreviewContext {
  readonly snapshot: WorkbookSnapshot;
  readonly providers: ReturnType<typeof workbookPreviewSource>;
  readonly viewportDependentStyles: boolean;
  readonly surface: WorkbookSurface;
  readonly controller: WorkbookWindowController;
  readonly windows: WeakMap<HTMLElement, PreviewWindow>;
  readonly slot: HTMLElement;
  dispose(): void;
}
function viewportOf(window: WorkbookRenderedWindow): WorkbookViewport {
  const { widthPt, heightPt, scrollLeftPt, scrollTopPt } = window.frame.layout;
  return Object.freeze({ widthPt, heightPt, scrollLeftPt, scrollTopPt });
}
function usesViewportStyles(snapshot: WorkbookSnapshot): boolean {
  // Inline HTML attributes may encode units with character references. Any
  // inline style or escaped CSS conservatively takes the independent path.
  if (/\bstyle\s*=/i.test(snapshot.content?.html ?? "")) return true;
  return snapshot.styles.some(({ css }) => css.includes("\\")
    || /[.\d](?:[sld]?v(?:w|h|i|b|min|max)|cq(?:w|h|i|b|min|max))\b|@(?:media|container)\b|\benv\s*\(/i.test(css));
}
function reportOf(window: WorkbookRenderedWindow): WorkbookRenderReport {
  const counts = { imported: 0, valid: 0, stale: 0, unavailable: 0 }, seen = new Set<number>();
  for (const cell of window.frame.projections) {
    const { row, column } = cell.range.from, key = row * 16_384 + column;
    if (seen.has(key)) continue; seen.add(key);
    const formula = window.frame.getCell(row, column)?.formula;
    if (formula) counts[formula.cacheState]++;
  }
  const { documentId, revision, sheetId, storedCellCount, layout } = window.frame;
  const ruleFormulaValues = { current: 0, unavailable: 0 };
  for (const cell of window.rules?.cells ?? []) if (seen.has(cell.row * 16_384 + cell.column) && window.frame.getCell(cell.row, cell.column)?.formula)
    ruleFormulaValues[cell.current.kind === "value" ? "current" : "unavailable"]++;
  return Object.freeze({ kind: "workbook", layout: "grid", documentId, revision, sheetId, viewport: viewportOf(window),
    maxScrollLeftPt: layout.maxScrollLeftPt, maxScrollTopPt: layout.maxScrollTopPt, visibleCellCount: seen.size, cellFragmentCount: window.cells.length, storedCellCount,
    objectFragmentCount: window.objects.length, formulaCaches: Object.freeze(counts), ruleFormulaValues: Object.freeze(ruleFormulaValues), diagnostics: window.diagnostics, fonts: window.fonts,
    usedFallbackLineMetrics: window.usedFallbackLineMetrics });
}

/** One preview owns its surfaces, candidates and format caches. Authorized
 * providers (including their Worker) remain owned by the host. */
export function mountWorkbookPreview(container: HTMLElement, input: WorkbookPreviewInput, options: WorkbookPreviewOptions = {}): WorkbookPreview {
  if (Object.keys(options).some(key => !["title", "widthPt", "heightPt", "sheetId", "locale", "retainSelection", "limits", "onError", "onSelectionChange"].includes(key))
    || options.retainSelection !== undefined && typeof options.retainSelection !== "boolean")
    throw new DocumentError("InvalidOperation", "Workbook preview received an unsupported option.");
  const config = { ...options, limits: { ...options.limits, read: { ...options.limits?.read }, cells: { ...options.limits?.cells },
    annotations: { ...options.limits?.annotations }, rules: { ...options.limits?.rules }, autoRows: { ...options.limits?.autoRows }, drawings: { ...options.limits?.drawings, assets: { ...options.limits?.drawings?.assets } } } };
  let desired: WorkbookViewport = Object.freeze({ widthPt: config.widthPt ?? 600, heightPt: config.heightPt ?? 360, scrollLeftPt: 0, scrollTopPt: 0 });
  let active: PreviewContext | undefined, pending: { readonly kind: "update" | "resize"; cancel(): void } | undefined, disposed = false, generation = 0;
  let selection: DocumentSelection | undefined, latest: Promise<void> = Promise.resolve();
  let navigation: WorkbookSelectionNavigation | undefined;
  let resolveReady!: () => void, rejectReady!: (error: unknown) => void;
  const ready = new Promise<void>((resolve, reject) => { resolveReady = resolve; rejectReady = reject; });
  // Hosts may observe ready after mounting other UI. Do not create an unhandled
  // rejection in the meantime; the original promise still rejects to callers.
  void ready.catch(() => {});
  const check = (): void => { if (disposed) throw new DocumentError("InvalidOperation", "This workbook preview has been disposed."); };
  const current = (): { context: PreviewContext; window: WorkbookRenderedWindow; ui: PreviewWindow } => {
    check(); const window = active?.controller.current;
    if (!active || !window) throw new DocumentError("InvalidOperation", "The workbook preview has no accepted window yet.");
    return { context: active, window, ui: active.windows.get(window.element)! };
  };
  const errorOf = (error: unknown): DocumentError => error instanceof DocumentError ? error : new DocumentError("InvalidOperation", "The workbook preview could not complete its authorized read.");
  const failed = (error: unknown): void => {
    if (disposed) return;
    const problem = errorOf(error); chrome.status("预览未更新：" + problem.code, true);
    try { config.onError?.(problem.toJSON()); } catch { /* Host callbacks do not own the renderer lifecycle. */ }
  };
  const attempt = (action: () => void): void => { if (disposed || pending) return; try { action(); } catch (error) { failed(error); } };
  const chrome = createWorkbookPreviewChrome(container, config.title ?? "工作簿只读预览", {
    sheet: id => { void showSheet(id).catch(() => {}); },
    scroll: (left, top) => { if (!pending) void setViewport({ ...desired, scrollLeftPt: left, scrollTopPt: top }).catch(() => {}); },
    comments: () => attempt(() => {
      const { context, window, ui } = current(), sheets = new Set(context.snapshot.metadata.workbook.sheets.map(sheet => sheet.sheetId));
      ui.controls.showComments(context.snapshot.annotations.filter(annotation => annotation.anchor.kind === "cells"
        && (annotation.anchor.sheetId === window.frame.sheetId || annotation.status === "needs_reanchor" && !sheets.has(annotation.anchor.sheetId))).map(annotation => annotation.annotationId));
    }),
  });
  const clearSelection = (): void => {
    chrome.cell("", ""); chrome.validation(""); navigation = undefined;
    if (!selection) return; selection = undefined;
    const ticket = generation;
    queueMicrotask(() => { if (!disposed && ticket === generation && !selection) {
      try { config.onSelectionChange?.(undefined); } catch { /* Host-owned callback. */ }
    } });
  };
  const accepted = (context: PreviewContext): void => {
    const window = context.controller.current!;
    const retained = config.retainSelection && selection?.documentId === context.snapshot.manifest.documentId && selection.baseRevision === context.snapshot.manifest.revision
      && selection.anchor.kind === "cells" && selection.anchor.sheetId === window.frame.sheetId ? selection.anchor.ranges[0] : undefined;
    context.windows.get(window.element)!.commitChrome(); desired = viewportOf(window);
    if (retained) selectSourceRange(retained, true, false); else clearSelection();
    const report = reportOf(window), counts = report.formulaCaches;
    chrome.status(report.diagnostics.some(item => ["rule-current-value-unavailable", "conditional-rule-unavailable", "validation-unavailable"].includes(item.code))
      ? "部分条件格式或数据验证无法按当前值求出。"
      : report.diagnostics.some(item => item.code === "validation-invalid") ? "当前窗口含不符合数据验证的来源值，选中单元格可查看提示。"
      : report.ruleFormulaValues.current === Object.values(counts).reduce((a, b) => a + b, 0) && report.ruleFormulaValues.current > 0 ? "当前窗口的公式已按当前值求出。"
      : counts.stale || counts.unavailable ? "当前窗口含过期或不可用的公式结果。"
      : counts.imported ? "当前窗口含导入的公式缓存。" : "只读预览");
    resolveReady();
  };
  const observe = (work: Promise<void>, ticket: number): void => {
    latest = work;
    void work.catch(error => {
      if (disposed || ticket !== generation) return;
      if (active?.controller.current) desired = viewportOf(active.controller.current);
      failed(error); if (!active && ticket === generation) rejectReady(errorOf(error));
    });
  };
  async function settled(): Promise<WorkbookRenderReport> {
    check();
    for (;;) { const work = latest; try { await work; } catch (error) { if (work === latest) throw errorOf(error); } if (work === latest) return reportOf(current().window); }
  }
  function setViewport(viewport: WorkbookViewport): Promise<WorkbookRenderReport> {
    check();
    if (pending?.kind === "update") return Promise.reject(new DocumentError("InvalidOperation", "A replacement workbook is still loading."));
    const { context, window } = current(); return render(context, window.frame.sheetId, viewport);
  }
  function showSheet(sheetId: string): Promise<WorkbookRenderReport> {
    check();
    if (pending?.kind === "update") return Promise.reject(new DocumentError("InvalidOperation", "A replacement workbook is still loading."));
    return render(current().context, sheetId, { ...desired, scrollLeftPt: 0, scrollTopPt: 0 });
  }
  function selectSourceRange(range: DeepReadonly<WorkbookRange>, showRange = false, notify = true): DocumentSelection {
    const { context, window, ui } = current(), sheet = context.snapshot.metadata.workbook.sheets.find(sheet => sheet.sheetId === window.frame.sheetId)!;
    const cell = range.from, value = window.frame.getCell(cell.row, cell.column);
    ui.controls.selectRange(range);
    const anchor = Object.freeze({ kind: "cells" as const, sheetId: sheet.sheetId, ranges: Object.freeze([Object.freeze({ from: Object.freeze({ ...range.from }), to: Object.freeze({ ...range.to }) })]) });
    selection = Object.freeze({ documentId: context.snapshot.manifest.documentId, baseRevision: context.snapshot.manifest.revision, anchor });
    if (notify) navigation = { origin: anchor.ranges[0]!.from, focus: anchor.ranges[0]!.from };
    const address = (at: WorkbookRange["from"]) => workbookColumnLabel(at.column) + (at.row + 1);
    chrome.cell(address(cell) + (!showRange || cell.row === range.to.row && cell.column === range.to.column ? "" : ":" + address(range.to)), value?.formula?.expression ?? (value?.value.type === "blank" || !value ? "" : String(value.value.value)));
    const validation = window.rules?.cells.find(item => item.row === cell.row && item.column === cell.column)?.validation;
    const definition = validation && sheet.dataValidations?.find(rule => rule.ruleId === validation.ruleId), messages: string[] = [];
    if (definition && validation) {
      if (!sheet.disableValidationPrompts && definition.showInputMessage) messages.push(...[definition.promptTitle, definition.prompt].filter(Boolean));
      if (validation.state !== "none") messages.push(validation.state === "valid" ? "当前值符合验证规则。" : validation.state === "invalid" ? "当前值不符合验证规则。" : "当前值无法校验。" + (validation.issue?.functionName ? " 未支持函数：" + validation.issue.functionName : ""));
      if (!sheet.disableValidationPrompts && validation.state === "invalid" && definition.showErrorMessage) messages.push(...[definition.errorTitle, definition.error].filter(Boolean));
    }
    chrome.validation(messages.join("\n"));
    if (notify) try { config.onSelectionChange?.(selection); } catch { /* Host-owned callback. */ }
    return selection;
  }
  async function selectRange(input: DeepReadonly<WorkbookRange>, ensureVisible = true, visibleAt?: DeepReadonly<WorkbookCoordinate>): Promise<DocumentSelection> {
    check();
    const parsed = WorkbookRangeSchema.safeParse(input);
    if (!parsed.success || parsed.data.from.row > parsed.data.to.row || parsed.data.from.column > parsed.data.to.column || typeof ensureVisible !== "boolean") throw new DocumentError("InvalidOperation", "Workbook selection requires an ordered source range.");
    if (pending) throw new DocumentError("InvalidOperation", "A replacement workbook view is still loading.");
    const requested = current(), source = requested.context.providers, selectedSheet = requested.window.frame.sheetId;
    await settled();
    if (current().context.providers !== source || current().window.frame.sheetId !== selectedSheet) throw new DocumentError("RevisionConflict", "The workbook view changed during selection.");
    const owned = current(), providers = owned.context.providers, sheet = owned.context.snapshot.metadata.workbook.sheets.find(sheet => sheet.sheetId === owned.window.frame.sheetId)!;
    const range = parsed.data.from.row === parsed.data.to.row && parsed.data.from.column === parsed.data.to.column ? new WorkbookMergeIndex(sheet.merges).at(parsed.data.from.row, parsed.data.from.column) ?? parsed.data : parsed.data;
    if (ensureVisible) {
      const dimensions = owned.window.dimensions;
      const fit = (axis: "rows" | "columns", frozen: number, extent: number, scroll: number): number => {
        const point = visibleAt ?? range.from, first = axis === "rows" ? point.row : point.column;
        if (first < frozen) return scroll;
        const dimension = dimensions[axis], fixed = dimension.position(frozen), size = Math.max(0, extent - Math.min(extent, fixed));
        if (!size) return scroll;
        const begin = dimension.position(first), end = dimension.position(first + 1), from = fixed + scroll;
        return begin < from ? begin - fixed : end > from + size ? Math.min(begin - fixed, end - fixed - size) : scroll;
      };
      const left = fit("columns", sheet.frozen.columns, desired.widthPt, desired.scrollLeftPt), top = fit("rows", sheet.frozen.rows, desired.heightPt, desired.scrollTopPt);
      if (left !== desired.scrollLeftPt || top !== desired.scrollTopPt) await render(owned.context, sheet.sheetId, { ...desired, scrollLeftPt: left, scrollTopPt: top });
    }
    if (current().context.providers !== providers || current().window.frame.sheetId !== sheet.sheetId) throw new DocumentError("RevisionConflict", "The workbook source changed during selection.");
    return selectSourceRange(range, true);
  }
  async function moveSelection(direction: "up" | "down" | "left" | "right", extend = false): Promise<DocumentSelection> {
    check();
    if (!["up", "down", "left", "right"].includes(direction) || typeof extend !== "boolean") throw new DocumentError("InvalidOperation", "Invalid workbook selection movement.");
    await settled();
    const { context, window } = current(), sheet = context.snapshot.metadata.workbook.sheets.find(sheet => sheet.sheetId === window.frame.sheetId)!;
    const first = (axis: "rows" | "columns"): number => {
      const dimensions = window.dimensions[axis], coordinate = dimensions.size(0) > 0 ? 0 : dimensions.neighbor(0, 1);
      if (coordinate === undefined) throw new DocumentError("InvalidOperation", "The worksheet has no visible coordinate on this axis."); return coordinate;
    };
    const point = navigation?.focus ?? { row: first("rows"), column: first("columns") };
    const next = navigateWorkbookSelection(sheet, window.dimensions, navigation ?? { origin: point, focus: point }, direction, extend);
    const selected = await selectRange(next.range, true, next.navigation.focus);
    if (selection === selected) navigation = next.navigation;
    return selected;
  }
  function render(context: PreviewContext, sheetId: string, viewport: WorkbookViewport): Promise<WorkbookRenderReport> {
    if (pending?.kind === "resize") { pending.cancel(); pending = undefined; chrome.busy(false); }
    const before = context.controller.current!.frame.layout;
    if (context.viewportDependentStyles && (before.widthPt !== viewport.widthPt || before.heightPt !== viewport.heightPt)) {
      prepare(context.providers, viewport, sheetId, "resize"); return settled();
    }
    context.controller.cancelPending();
    const ticket = ++generation;
    desired = Object.freeze({ widthPt: viewport.widthPt, heightPt: viewport.heightPt, scrollLeftPt: viewport.scrollLeftPt, scrollTopPt: viewport.scrollTopPt });
    const work = (async () => { if (await context.controller.render(sheetId, desired) && !disposed && ticket === generation && active === context) accepted(context); })();
    observe(work, ticket); return settled();
  }
  function update(next: WorkbookPreviewInput): void {
    check();
    const providers = workbookPreviewSource(next), snapshot = providers.cells.snapshot;
    const previousSheet = active?.snapshot.manifest.documentId === snapshot.manifest.documentId ? active.controller.current?.frame.sheetId : undefined;
    const retained = previousSheet && snapshot.metadata.workbook.sheets.some(sheet => sheet.sheetId === previousSheet && sheet.visibility === "visible") ? previousSheet : undefined;
    prepare(providers, desired, retained ?? (active ? snapshot.metadata.workbook.activeSheetId : config.sheetId ?? snapshot.metadata.workbook.activeSheetId), "update");
  }
  function prepare(providers: ReturnType<typeof workbookPreviewSource>, inputViewport: WorkbookViewport, sheetId: string, kind: "update" | "resize"): void {
    const snapshot = providers.cells.snapshot, viewport = Object.freeze({ ...inputViewport }); desired = viewport;
    pending?.cancel(); active?.controller.cancelForReplacement(); const ticket = ++generation, abort = new AbortController();
    const before = active, oldViewport = before?.controller.current?.frame.layout;
    const rows = before ? workbookRowInvalidation(before.snapshot, snapshot, workbookPreviewChangeImpact(before.providers, providers), config.limits.autoRows.maxCells ?? DEFAULT_AUTO_ROW_LIMITS.maxCells) : undefined;
    const share = rows && before?.surface.fork && oldViewport?.widthPt === viewport.widthPt && oldViewport.heightPt === viewport.heightPt
      && snapshot.styles.every(style => /^(?:\s*\.[a-zA-Z_][-\w]*\s*\{[^{}]*\}\s*)*$/.test(style.css)) ? before : undefined;
    let candidate: PreviewContext | undefined;
    const request = { kind, cancel() { abort.abort(); candidate?.dispose(); } }; pending = request; chrome.busy(true);
    const work = (async () => {
      const slot = share?.slot ?? container.ownerDocument.createElement("div");
      if (!share) {
        slot.style.cssText = "position:absolute;inset:0;opacity:0;pointer-events:none;overflow:hidden";
        slot.setAttribute("aria-hidden", "true"); chrome.canvas.append(slot);
      }
      let surface: WorkbookSurface | undefined, formatter: ReturnType<typeof createWorkbookNumberFormatter> | undefined;
      try {
        surface = share ? share.surface.fork!() : await createWorkbookSurface(slot, snapshot, { ...viewport, signal: abort.signal });
        if (share && surface.document.fonts.status === "loading") await waitForWorkbookRead(surface.document.fonts.ready, abort.signal);
        if (abort.signal.aborted) return;
        formatter = createWorkbookNumberFormatter({ dateSystem: snapshot.metadata.workbook.dateSystem,
          ...(config.locale !== undefined ? { hostLocale: config.locale } : {}), ...(snapshot.metadata.locale ? { documentLocale: snapshot.metadata.locale } : {}) });
        const windows = new WeakMap<HTMLElement, PreviewWindow>(), ownedSurface = surface, ownedFormatter = formatter;
        const controller = new WorkbookWindowController(surface, providers, formatter, config.limits, (window, content, maxNodes) => {
          const commitChrome = chrome.prepare(window, snapshot);
          const sheet = snapshot.metadata.workbook.sheets.find(sheet => sheet.sheetId === window.frame.sheetId)!, merges = new WorkbookMergeIndex(sheet.merges);
          const controls = decorateWorkbookPreview(ownedSurface, snapshot, window, content, maxNodes, attempt);
          const events = new ownedSurface.document.defaultView!.AbortController();
          windows.set(window.element, { controls, commitChrome });
          window.element.addEventListener("wheel", event => {
            if (event.ctrlKey || pending || active?.controller.current?.element !== window.element || (event.target as Element).closest('[data-office-view="workbook-preview-comments"]')) return;
            event.preventDefault();
            const factor = event.deltaMode === 1 ? 12 : event.deltaMode === 2 ? window.frame.layout.heightPt : .75;
            const x = event.shiftKey && !event.deltaX ? event.deltaY : event.deltaX, y = event.shiftKey && !event.deltaX ? 0 : event.deltaY;
            void setViewport({ ...desired, scrollLeftPt: desired.scrollLeftPt + x * factor, scrollTopPt: desired.scrollTopPt + y * factor }).catch(() => {});
          }, { passive: false, signal: events.signal });
          window.element.addEventListener("click", event => attempt(() => {
            if (active?.controller.current?.element !== window.element || (event.target as Element).closest('button,[data-office-view="workbook-preview-comments"],[data-office-view="workbook-drawings"]')) return;
            const rect = ownedSurface.view.getBoundingClientRect(), x = (event.clientX - rect.left) * .75, y = (event.clientY - rect.top) * .75;
            const cell = hitWorkbookCell(window.dimensions, window.frame.layout, merges, sheet.rightToLeft ? window.frame.layout.widthPt - x : x, y);
            if (!cell) return;
            selectSourceRange(merges.at(cell.row, cell.column) ?? { from: cell, to: cell });
          }), { signal: events.signal });
          return { get nodes() { return controls.nodes; }, get fonts() { return controls.fonts; }, dispose() { events.abort(); controls.dispose(); } };
        }, { deferPublication: true, ...(share && rows ? { measurements: share.controller.measurementSeeds(rows) } : {}) });
        let released = false;
        // Native point-based styles retain their stable indexes on resize.
        // Viewport/container units and conditional CSS need a fresh surface
        // to measure typography at the final size before atomic publication.
        const viewportDependentStyles = usesViewportStyles(snapshot);
        candidate = { snapshot, providers, viewportDependentStyles, surface, controller, windows, slot, dispose() {
          if (released) return; released = true; controller.dispose(); ownedFormatter.dispose(); ownedSurface.dispose(); if (!slot.querySelector("iframe")) slot.remove();
        } };
        if (!await controller.render(sheetId, viewport) || abort.signal.aborted || disposed || ticket !== generation) return;
        const previous = active;
        // Chrome preparation and all controls already succeeded before the
        // controller published this isolated candidate's complete window.
        previous?.controller.unpublish();
        try { controller.publish(); active = candidate; accepted(candidate); }
        catch (error) { active = previous; previous?.controller.publish(); throw error; }
        candidate = undefined;
        slot.style.opacity = "1"; slot.style.pointerEvents = "auto"; slot.removeAttribute("aria-hidden"); previous?.dispose();
        surface = undefined; formatter = undefined;
      } catch (error) { if (!abort.signal.aborted && !disposed && ticket === generation) throw error; }
      finally {
        candidate?.dispose(); surface?.dispose(); formatter?.dispose();
        if (active?.slot !== slot) slot.remove();
        if (pending === request) { pending = undefined; chrome.busy(false); }
      }
    })();
    observe(work, ticket);
  }
  try { update(input); } catch (error) { chrome.dispose(); throw error; }
  return Object.freeze({ kind: "workbook", element: chrome.element, ready, snapshot: () => current().context.snapshot,
    layout: () => reportOf(current().window), fonts: () => current().window.fonts, selection: () => { check(); return selection; }, settled, update, showSheet, setViewport,
    selectRange: (range: DeepReadonly<WorkbookRange>, ensureVisible?: boolean) => selectRange(range, ensureVisible), moveSelection,
    openComment(annotationId: string) { current().ui.controls.openComment(annotationId); }, closeComment() { current().ui.controls.closeComment(); },
    dispose() { if (disposed) return; disposed = true; ++generation; pending?.cancel(); pending = undefined; active?.dispose(); active = undefined;
      selection = undefined; navigation = undefined; chrome.dispose(); rejectReady(new DocumentError("OperationCancelled", "Workbook preview was disposed before becoming ready.")); },
  });
}
