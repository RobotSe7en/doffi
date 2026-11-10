import { DocumentError, parseContract } from "@doffi/core";
import type { DeepReadonly, DocumentOperation, DocumentSelection, TransactionPrepareOptions, WorkbookRange, WorkbookSnapshot } from "@doffi/core";
import { createWorkbookLayoutService, mountDocumentPreview } from "@doffi/render";
import type { WorkbookPreview, WorkbookReadConnection, WorkbookRenderReport } from "@doffi/render";
import type { WorkbookEditorOptions, WorkbookEditorView } from "./editor-contracts.js";
import { createWorkbookEditorUi } from "./editor-ui.js";
import { withWorkbookAsyncValidationConfirmation, workbookValidationInputKey } from "./validation-confirmation.js";
import { createWorkbookRemoteInputController } from "./remote-input.js";
import type { WorkbookRemoteInputState } from "./remote-input.js";
import { bindWorkbookAsyncTextInput } from "./text-binding.js";
import type { WorkbookAsyncTextBinding } from "./text-binding.js";
import type { WorkbookSessionCommand, WorkbookSessionConnection, WorkbookSessionReceipt } from "./session-transport.js";
import { parseWorkbookTextInput } from "./text-input.js";
import { workbookFillSeedCount, workbookSeriesAxis, workbookSeriesUnit } from "./fill-selection.js";
import type { WorkbookTextInputMode } from "./text-input.js";
import type { WorkbookFindMatch, WorkbookFindQuery, WorkbookFindResult, WorkbookReplaceOptions } from "./find-contracts.js";
import { workbookInputTextView } from "./text-projection.js";
import { selectedWorkbookSort } from "./selected-sort.js";
import { WorkbookClipboard } from "./clipboard.js";
import { WorkbookSystemClipboard } from "./system-clipboard.js";
import type { WorkbookRangeCopyOptions } from "./copy-range.js";
import { createWorkbookDrawingControls } from "./drawing-controls.js";
import { createWorkbookCommentControls } from "./comment-controls.js";
import { createWorkbookNameControls } from "./names-controls.js";
import { createWorkbookWorksheetControls } from "./worksheet-controls.js";
import { createWorkbookHyperlinkControls } from "./hyperlink-controls.js";
import { createWorkbookTableStyleControls } from "./table-style-controls.js";
import { createWorkbookTableControls } from "./table-controls.js";
import { createWorkbookFilterControls } from "./filter-controls.js";
import { createWorkbookRulesControls } from "./rules-controls.js";
import type { WorkbookControlsHost } from "./controls-host.js";

export interface WorkbookRemoteEditorView extends Omit<WorkbookEditorView, "applyInput" | "perform" | "cancelInput" | "inputState"> {
  inputState(): WorkbookRemoteInputState;
  applyInput(): Promise<WorkbookSessionReceipt | undefined>;
  perform(operations: readonly DocumentOperation[], options?: TransactionPrepareOptions): Promise<WorkbookSessionReceipt>;
  cancelInput(): Promise<void>;
  hasPendingInput(): boolean;
  layout(): WorkbookRenderReport;
  setViewport(viewport: Parameters<WorkbookPreview["setViewport"]>[0]): Promise<WorkbookRenderReport>;
}
export interface WorkbookRemoteEditorOptions extends WorkbookEditorOptions {
  readonly onViewChange?: (snapshot: WorkbookSnapshot, report: WorkbookRenderReport) => void;
  readonly onPendingChange?: (pending: boolean) => void;
}

/** Host-owned connection, sole Worker session. This view owns its transient
 * input, subscriptions and read leases; disposal never closes the host client. */
export function mountWorkbookRemoteEditor(container: HTMLElement, client: WorkbookSessionConnection, options: WorkbookRemoteEditorOptions = {}): WorkbookRemoteEditorView {
  const view = container?.ownerDocument?.defaultView;
  if (!view || !container.isConnected || !client || typeof client.readInput !== "function" || typeof client.view !== "function") throw new DocumentError("InvalidOperation", "Workbook editing requires a connected container and its authorized session connection.");
  client.state();
  if (!options || typeof options !== "object" || Array.isArray(options) || Object.keys(options).some(key => !["title", "widthPt", "heightPt", "locale", "limits", "prepareOptions", "confirmValidationWarnings", "onError", "onSelectionChange", "onViewChange", "onPendingChange"].includes(key))) throw new DocumentError("InvalidOperation", "Invalid workbook editor options.");
  if (options.confirmValidationWarnings !== undefined && typeof options.confirmValidationWarnings !== "function") throw new DocumentError("InvalidOperation", "Validation confirmation requires a host callback.");
  const config = { ...options }, ui = createWorkbookEditorUi(container, config.title ?? "工作簿编辑"), events = new view.AbortController();
  const layoutService = createWorkbookLayoutService(ui.preview, { ...(config.widthPt === undefined ? {} : { widthPt: config.widthPt }), ...(config.heightPt === undefined ? {} : { heightPt: config.heightPt }),
    ...(config.locale === undefined ? {} : { locale: config.locale }), ...(config.limits ? { limits: config.limits } : {}) });
  const authorizedClient = client;
  client = Object.freeze({ ...authorizedClient, prepare: (...args: Parameters<WorkbookSessionConnection["prepare"]>) => {
    const [command, preparation, execution] = args;
    return authorizedClient.prepare(command, preparation, { ...execution, measureLayout: execution?.measureLayout ?? (request => layoutService.measureSource(request, preview?.layout().viewport)) });
  } });
  let disposed = false, choosing = 0, generation = 0, refreshQueued = false, forceCalculation = false, refreshing = false;
  let latest: Promise<void> = Promise.resolve(), interaction: Promise<void> = Promise.resolve(), refreshAbort: AbortController | undefined, commandAbort: AbortController | undefined;
  let inputTask: Promise<void> = Promise.resolve(), commandTask: Promise<void> = Promise.resolve();
  let releaseWaiters!: () => void; const released = new Promise<void>(resolve => { releaseWaiters = resolve; });
  let refreshRevision: number | undefined;
  let reportedPending: boolean | undefined;
  let commandApplying = false;
  let drawingControls: ReturnType<typeof createWorkbookDrawingControls> | undefined;
  let commentControls: ReturnType<typeof createWorkbookCommentControls> | undefined;
  let nameControls: ReturnType<typeof createWorkbookNameControls> | undefined;
  let worksheetControls: ReturnType<typeof createWorkbookWorksheetControls> | undefined;
  let hyperlinkControls: ReturnType<typeof createWorkbookHyperlinkControls> | undefined;
  let tableControls: ReturnType<typeof createWorkbookTableControls> | undefined;
  let tableStyleControls: ReturnType<typeof createWorkbookTableStyleControls> | undefined;
  let filterControls: ReturnType<typeof createWorkbookFilterControls> | undefined;
  let rulesControls: ReturnType<typeof createWorkbookRulesControls> | undefined;
  let found: { key: string; result: WorkbookFindResult; index: number; view?: WorkbookReadConnection } | undefined;
  let findSelectionScope: WorkbookFindQuery["scope"];
  let preview: WorkbookPreview | undefined, source: WorkbookReadConnection | undefined, binding: WorkbookAsyncTextBinding;
  const buttons: { button: HTMLButtonElement; available(): boolean }[] = [];
  const input = createWorkbookRemoteInputController(client, { onChange: controls });
  const clipboard = new WorkbookClipboard();
  const systemClipboard = new WorkbookSystemClipboard(async (selected, signal) => {
    const anchor = selected.anchor;
    if (anchor.kind !== "cells" || anchor.ranges.length !== 1 || selected.baseRevision !== client.state().revision) throw new DocumentError("RevisionConflict", "Clipboard selection belongs to a previous source.");
    return client.readClipboard(anchor.sheetId, { from: { ...anchor.ranges[0]!.from }, to: { ...anchor.ranges[0]!.to } }, { signal, ...(source ? { view: source } : {}), ...(config.locale === undefined ? {} : { calculationContext: { locale: config.locale } }) });
  });
  const searchFields = [ui.find, ui.replacement, ui.findField, ui.findScope, ui.matchCase, ui.entireCell, ui.replacementMode];
  function check(): void { if (disposed) throw new DocumentError("InvalidOperation", "This workbook editor has been disposed."); }
  function report(error: unknown): void {
    if (disposed) return; const problem = error instanceof Error ? error : new Error("工作簿编辑失败。");
    if (problem instanceof DocumentError && problem.code === "OperationCancelled" && input.state().dirty) { ui.status.textContent = "输入已更新，请确认当前草稿。"; return; }
    ui.error.textContent = problem instanceof DocumentError && problem.code === "RevisionConflict" ? "文档已变化。当前输入已保留，请先复制输入或取消后重新选择。"
      : problem instanceof DocumentError && problem.details.some(detail => detail.code === "merge-covered-content") ? "合并范围包含其他内容。若确认删除，请使用“合并并清除覆盖内容”。" : problem.message; ui.error.hidden = false;
    try { config.onError?.(problem); } catch { /* Host reporting does not own source changes. */ }
  }
  function validationContext(draftKey?: string, signal?: AbortSignal) {
    const expectedDraft = draftKey ?? (config.confirmValidationWarnings ? workbookValidationInputKey(input.state()) : undefined);
    const { documentId, revision: baseRevision } = client.state();
    return { documentId, baseRevision, confirm: config.confirmValidationWarnings, assertCurrent() {
      check(); const current = client.state();
      if (current.documentId !== documentId || current.revision !== baseRevision) throw new DocumentError("RevisionConflict", "Workbook changed during validation confirmation.");
      if (signal?.aborted || expectedDraft !== undefined && workbookValidationInputKey(input.state()) !== expectedDraft)
        throw new DocumentError("OperationCancelled", "Workbook input changed during validation confirmation.");
    } };
  }
  function preparation(supplied?: TransactionPrepareOptions): TransactionPrepareOptions {
    const value = supplied ?? config.prepareOptions?.() ?? {};
    return config.locale !== undefined && value.calculationContext?.locale === undefined ? { ...value, calculationContext: { ...value.calculationContext, locale: config.locale } } : value;
  }
  function pendingInput(): boolean { return binding?.hasPendingInput() ?? false; }
  function pendingForms(): boolean { return [drawingControls, commentControls, nameControls, worksheetControls, hyperlinkControls, tableControls, tableStyleControls, filterControls, rulesControls].some(controls => controls?.hasPendingInput()); }
  function controls(): void {
    if (disposed) return; const state = input.state(), busy = !!commandAbort || state.applying || state.selecting;
    // Keep the draft editable while a remote operation is preparing or its
    // own receipt is pending. The controller preserves or cancels that draft.
    ui.input.disabled = !state.selection || state.selecting || !!commandAbort;
    ui.mode.value = state.mode; ui.mode.disabled = busy || state.composing || !state.selection; ui.address.disabled = busy || refreshing || !source;
    for (const field of searchFields) field.disabled = busy;
    for (const field of [ui.insertedFormat, ui.fillFormat, ui.sortColumns, ui.sortHeader, ui.sortCase, ui.sortMethod, ui.clipboardMode, ui.clipboardOverwrite, ui.clipboardSkipBlanks]) field.disabled = busy;
    ui.clipboardMode.disabled ||= clipboard.kind === "cut"; ui.clipboardSkipBlanks.disabled ||= clipboard.kind === "cut"; ui.clipboardOverwrite.disabled ||= clipboard.kind !== "cut";
    for (const field of ui.format.fields) field.disabled = busy;
    for (const { button, available } of buttons) button.disabled = !available();
    drawingControls?.refresh();
    commentControls?.refresh();
    nameControls?.refresh();
    worksheetControls?.refresh();
    hyperlinkControls?.refresh();
    tableControls?.refresh();
    tableStyleControls?.refresh();
    filterControls?.refresh();
    rulesControls?.refresh();
    const pending = pendingInput() || state.applying || !!commandAbort || pendingForms();
    if (reportedPending !== pending) { reportedPending = pending; try { config.onPendingChange?.(pending); } catch { /* Host owns navigation decisions. */ } }
  }
  function address(range: DeepReadonly<WorkbookRange>): string {
    const cell = (point: WorkbookRange["from"]): string => { let column = point.column + 1, name = ""; while (column) { column--; name = String.fromCharCode(65 + column % 26) + name; column = Math.floor(column / 26); } return name + (point.row + 1); };
    return cell(range.from) + (range.from.row === range.to.row && range.from.column === range.to.column ? "" : ":" + cell(range.to));
  }
  function notify(): void { try { config.onSelectionChange?.(input.state().selection); } catch { /* Host callback. */ } }
  function runInteraction<T>(action: () => Promise<T>): Promise<T> { const next = interaction.then(action); interaction = next.then(() => {}, report); return next; }
  // Internal controls use the queue's reporting boundary. Public methods keep
  // the rejecting Promise for their caller; a control must not report it twice.
  function queueInteraction(action: () => Promise<unknown>): void { void runInteraction(action); }
  async function acceptSelection(selected: DocumentSelection, focus = false): Promise<void> {
    await input.select(selected); check(); binding.refresh();
    if (source?.cells.snapshot.manifest.revision === selected.baseRevision) systemClipboard.prepare(selected);
    if (selected.anchor.kind === "cells") ui.address.value = address(selected.anchor.ranges[0]!);
    controls(); notify(); if (focus) binding.focus();
  }
  async function restoreVisual(): Promise<void> {
    const anchor = input.state().selection?.anchor;
    if (!preview || anchor?.kind !== "cells" || !preview.snapshot().metadata.workbook.sheets.some(sheet => sheet.sheetId === anchor.sheetId && sheet.visibility === "visible")) return;
    choosing++;
    try { if (preview.layout().sheetId !== anchor.sheetId) await preview.showSheet(anchor.sheetId); await preview.selectRange(anchor.ranges[0]!, false); }
    finally { choosing--; }
  }
  async function sourceSettled(): Promise<WorkbookRenderReport> {
    for (;;) {
      check(); const current = latest; try { await current; } catch (error) { if (current === latest) throw error; }
      if (current !== latest) continue; check();
      if (!preview) throw new DocumentError("InvalidOperation", "The workbook editor has no accepted view yet.");
      const report = await preview.settled(); if (current === latest) return report;
    }
  }
  async function choose(sheetId: string, range: DeepReadonly<WorkbookRange>, focus = false, retainFind = false, revision = client.state().revision): Promise<DocumentSelection> {
    check(); if (pendingInput() || input.state().applying || commandAbort && !retainFind) throw new DocumentError("InvalidOperation", "请先确认或取消当前输入或操作。");
    if (!retainFind) { found = undefined; findSelectionScope = undefined; }
    await sourceSettled();
    const current = (): void => { if (client.state().revision !== revision || preview!.layout().revision !== revision) throw new DocumentError("RevisionConflict", "Workbook source changed while selecting."); };
    current(); choosing++;
    try {
      if (preview!.layout().sheetId !== sheetId) await preview!.showSheet(sheetId);
      const selected = await preview!.selectRange(range); current(); await acceptSelection(selected, focus); return selected;
    } catch (error) { try { await restoreVisual(); } catch { /* A newer source can remove the old worksheet. */ } throw error; }
    finally { choosing--; }
  }
  function sourceSelection(selected: DocumentSelection | undefined): void {
    if (disposed || choosing) return;
    found = undefined; findSelectionScope = undefined;
    void runInteraction(async () => {
      if (input.state().composing || input.state().applying || commandAbort) { await restoreVisual(); return; }
      if (pendingInput() && !await binding.apply()) { await restoreVisual(); return; }
      if (pendingInput()) { await restoreVisual(); return; }
      await sourceSettled(); const sheetId = selected?.anchor.kind === "cells" ? selected.anchor.sheetId : preview!.layout().sheetId;
      const range = selected?.anchor.kind === "cells" ? selected.anchor.ranges[0]! : { from: { row: 0, column: 0 }, to: { row: 0, column: 0 } };
      await choose(sheetId, range, true);
    });
  }
  async function refresh(force: boolean): Promise<void> {
    check(); const ticket = ++generation, abort = new view!.AbortController(); refreshAbort?.abort(); refreshAbort = abort; refreshRevision = client.state().revision; refreshing = true; controls();
    let candidate: WorkbookReadConnection | undefined;
    try {
      const context = preparation().calculationContext;
      candidate = await client.view(undefined, { signal: abort.signal, ...(source ? { previous: source } : {}), ...(force ? { calculate: true } : {}), ...(context ? { calculationContext: context } : {}) });
      if (disposed || ticket !== generation || abort.signal.aborted) return;
      choosing++;
      try {
        if (preview) { preview.update(candidate); await preview.settled(); }
        else {
          preview = mountDocumentPreview(ui.preview, candidate, { retainSelection: true, ...(config.widthPt === undefined ? {} : { widthPt: config.widthPt }), ...(config.heightPt === undefined ? {} : { heightPt: config.heightPt }),
            ...(config.locale === undefined ? {} : { locale: config.locale }), ...(config.limits === undefined ? {} : { limits: config.limits }), onSelectionChange: sourceSelection, onError: error => report(new DocumentError(error.code, error.message, error.details)) });
          await preview.ready;
        }
        if (disposed || ticket !== generation || abort.signal.aborted) return;
        const old = source; source = candidate; candidate = undefined; old?.dispose();
        const state = input.state(), anchor = state.selection?.anchor, sheetId = preview.layout().sheetId;
        const range = anchor?.kind === "cells" && anchor.sheetId === sheetId ? anchor.ranges[0]! : { from: { row: 0, column: 0 }, to: { row: 0, column: 0 } };
        const selected = await preview.selectRange(range, false);
        if (!pendingInput() && !state.applying && !state.composing) await acceptSelection(selected);
        const current = input.state(); ui.status.textContent = current.dirty || current.composing ? current.selection?.baseRevision !== client.state().revision ? "文档已更新；当前输入仍绑定原修订。" : "输入尚未确认。" : "修订 " + client.state().revision;
        try { config.onViewChange?.(source.cells.snapshot, preview.layout()); } catch { /* Host metadata cannot roll back the accepted view. */ }
      } finally { choosing--; }
    } catch (error) { if (!disposed && ticket === generation && !abort.signal.aborted) throw error; }
    finally { candidate?.dispose(); if (ticket === generation) { refreshing = false; refreshRevision = undefined; if (refreshAbort === abort) refreshAbort = undefined; controls(); } }
  }
  function scheduleRefresh(force = false): Promise<void> {
    check(); const revision = client.state().revision;
    if (!force && !refreshQueued && (refreshing && refreshRevision === revision || !refreshing && source?.cells.snapshot.manifest.revision === revision)) return latest;
    forceCalculation ||= force;
    if (!refreshQueued) {
      refreshQueued = true;
      const next = Promise.resolve().then(() => { refreshQueued = false; const force = forceCalculation; forceCalculation = false; return refresh(force); });
      latest = next; void next.catch(error => { if (next === latest) report(error); });
    }
    return latest;
  }
  function commitCommand(command: WorkbookSessionCommand, supplied?: TransactionPrepareOptions): Promise<WorkbookSessionReceipt> {
    const next = prepareCommand(command, supplied); commandTask = next.then(() => {}, () => {}); return next;
  }
  async function prepareCommand(command: WorkbookSessionCommand, supplied?: TransactionPrepareOptions): Promise<WorkbookSessionReceipt> {
    check(); if (pendingInput() || input.state().applying || commandAbort) throw new DocumentError("InvalidOperation", "请先完成当前输入或操作。");
    const abort = new view!.AbortController(); commandAbort = abort; controls();
    try { return await applyPrepared(command, abort, supplied); }
    finally { if (commandAbort === abort) commandAbort = undefined; controls(); }
  }
  async function applyPrepared(command: WorkbookSessionCommand, abort: AbortController, supplied?: TransactionPrepareOptions): Promise<WorkbookSessionReceipt> {
    let candidate: Awaited<ReturnType<WorkbookSessionConnection["prepare"]>> | undefined, consumed = false;
    try {
      const context = validationContext(undefined, abort.signal);
      candidate = await withWorkbookAsyncValidationConfirmation(options => client.prepare(command, options, { signal: abort.signal, onProgress: progress => { if (!disposed) ui.status.textContent = "正在准备更改 · " + progress.completed + (progress.total ? " / " + progress.total : ""); } }), preparation(supplied), context);
      if (disposed || abort.signal.aborted) throw new DocumentError("OperationCancelled", "Workbook command was cancelled before application.");
      commandApplying = true; controls(); const receipt = await client.commit(candidate); consumed = true; if (!disposed) {
        ui.error.hidden = true; void scheduleRefresh();
        if (candidate.clipboardDiagnostics?.length) ui.status.textContent = "已粘贴；部分外部样式不受支持（" + candidate.clipboardDiagnostics.length + " 项）。";
      } return receipt;
    } finally {
      if (candidate && !consumed) try { await client.discard(candidate); } catch { /* Host disposal or another source can release the lease. */ }
      commandApplying = false; controls();
    }
  }
  async function perform(operations: readonly DocumentOperation[], supplied?: TransactionPrepareOptions): Promise<WorkbookSessionReceipt> {
    check(); const { documentId, revision } = client.state();
    const value = { protocolVersion: 1 as const, documentId, baseRevision: revision, operations: [...operations] };
    return commitCommand({ kind: "transaction", transaction: parseContract("document-transaction", value) }, supplied);
  }
  async function cancelInput(): Promise<void> {
    check(); input.cancel(); binding.refresh(); await sourceSettled();
    const anchor = input.state().selection?.anchor, sheetId = anchor?.kind === "cells" && source!.cells.snapshot.metadata.workbook.sheets.some(sheet => sheet.sheetId === anchor.sheetId && sheet.visibility === "visible") ? anchor.sheetId : preview!.layout().sheetId;
    const range = anchor?.kind === "cells" && anchor.sheetId === sheetId ? anchor.ranges[0]! : { from: { row: 0, column: 0 }, to: { row: 0, column: 0 } };
    await choose(sheetId, range, true); ui.error.hidden = true;
  }
  async function history(direction: "undo" | "redo"): Promise<void> {
    check(); if (pendingInput() || input.state().applying || commandAbort) throw new DocumentError("InvalidOperation", "请先确认或取消当前输入。");
    await client[direction](); await scheduleRefresh(); await sourceSettled(); binding.focus(); ui.error.hidden = true;
  }
  async function recalculate(): Promise<WorkbookRenderReport> {
    check(); const revision = client.state().revision; await scheduleRefresh(true); const report = await sourceSettled();
    if (report.revision !== revision) throw new DocumentError("RevisionConflict", "Workbook source changed while recalculating.");
    return report;
  }
  async function fill(axis: "row" | "column", series = false): Promise<void> {
    check(); if (pendingInput() || input.state().applying || commandAbort) throw new DocumentError("InvalidOperation", "请先完成当前输入或操作。");
    const selected = input.state().selection, anchor = selected?.anchor;
    if (!selected || anchor?.kind !== "cells") throw new DocumentError("InvalidOperation", "请先选择填充区域。");
    const range = anchor.ranges[0]!, count = range.to[axis] - range.from[axis] + 1;
    if (count < 2) throw new DocumentError("InvalidOperation", axis === "row" ? "向下填充需要至少两行。" : "向右填充需要至少两列。");
    const snapshot = source?.cells.snapshot;
    if (!snapshot || snapshot.manifest.documentId !== selected.documentId || snapshot.manifest.revision !== selected.baseRevision)
      throw new DocumentError("RevisionConflict", "Fill selection belongs to a previous source.");
    const merges = snapshot.metadata.workbook.sheets.find(sheet => sheet.sheetId === anchor.sheetId)!.merges, format = ui.fillFormat.value as "source" | "preserve";
    let seeds = series ? workbookSeriesUnit(range, merges, axis) : workbookFillSeedCount(range, merges, axis);
    if (range.from[axis] + seeds > range.to[axis]) return;
    if (series) {
      const at = { ...range.from, [axis]: range.from[axis] + seeds }, value = await client.readInput(anchor.sheetId, at), parsed = parseWorkbookTextInput(value.text, at);
      if (parsed.kind !== "value" || parsed.value.type !== "number" && parsed.value.type !== "blank") throw new DocumentError("InvalidOperation", "数列填充的种子必须是数字或空白。");
      if (parsed.value.type === "number") seeds *= 2;
    }
    const target = { from: { ...range.from, [axis]: range.from[axis] + seeds }, to: { ...range.to } }; if (target.from[axis] > target.to[axis]) return;
    await commitCommand({ kind: "fill", documentId: selected.documentId, revision: selected.baseRevision, sheetId: anchor.sheetId, axis, mode: series ? "series" : "copy", format,
      source: { from: { ...range.from }, to: { ...range.to, [axis]: range.from[axis] + seeds - 1 } }, target });
    await sourceSettled();
  }
  async function changeMerge(unmerge = false, clear = false): Promise<void> {
    check(); const selected = input.state().selection, anchor = selected?.anchor;
    if (!selected || anchor?.kind !== "cells" || anchor.ranges.length !== 1) throw new DocumentError("InvalidOperation", "请选择一个合并区域。");
    if (selected.baseRevision !== client.state().revision) throw new DocumentError("RevisionConflict", "Merge selection belongs to a previous source.");
    const target = { sheetId: anchor.sheetId, range: anchor.ranges[0]! };
    await perform([unmerge ? { op: "unmergeWorkbookCells", ...target } : { op: "mergeWorkbookCells", ...target, coveredContent: clear ? "clear" : "reject" }]);
    await sourceSettled(); binding.focus();
  }
  async function changeFreeze(clear = false): Promise<void> {
    check(); const selected = input.state().selection, anchor = selected?.anchor;
    if (!selected || anchor?.kind !== "cells") throw new DocumentError("InvalidOperation", "请先选择冻结位置。");
    if (selected.baseRevision !== client.state().revision) throw new DocumentError("RevisionConflict", "Freeze selection belongs to a previous source.");
    const at = anchor.ranges[0]!.from;
    if (!clear) {
      const offset = preview!.sourceOffset(at), viewport = preview!.layout().viewport;
      if (offset.revision !== selected.baseRevision || offset.sheetId !== anchor.sheetId) throw new DocumentError("RevisionConflict", "Freeze geometry belongs to a previous view.");
      if (offset.leftPt >= viewport.widthPt || offset.topPt >= viewport.heightPt) throw new DocumentError("InvalidOperation", "冻结区域将占满当前视口，请选择更靠前的位置。");
    }
    await perform([{ op: "updateWorkbookSheet", sheetId: anchor.sheetId, changes: { frozen: { rows: clear ? 0 : at.row, columns: clear ? 0 : at.column } } }]);
    await sourceSettled(); binding.focus();
  }
  async function changeAxis(axis: "row" | "column", action: "insert" | "delete"): Promise<void> {
    check(); const selected = input.state().selection, anchor = selected?.anchor;
    if (!selected || anchor?.kind !== "cells" || anchor.ranges.length !== 1) throw new DocumentError("InvalidOperation", "请选择一个行列区域。");
    if (selected.baseRevision !== client.state().revision) throw new DocumentError("RevisionConflict", "Axis selection belongs to a previous source.");
    const range = anchor.ranges[0]!;
    await perform([{ op: "editWorkbookAxis", sheetId: anchor.sheetId, axis, action, index: range.from[axis], count: range.to[axis] - range.from[axis] + 1,
      ...(action === "insert" ? { formatSource: ui.insertedFormat.value as "before" | "after" | "none" } : {}) }]);
    await sourceSettled(); binding.focus();
  }
  async function sortSelection(descending: boolean): Promise<void> {
    check(); await sourceSettled();
    const selected = input.state().selection, target = selectedWorkbookSort(source!.cells.snapshot, selected,
      { columns: ui.sortColumns.value, header: ui.sortHeader.checked, caseSensitive: ui.sortCase.checked, method: ui.sortMethod.value }, descending);
    await commitCommand({ kind: "sort", documentId: selected!.documentId, revision: selected!.baseRevision, ...target, ...(config.locale === undefined ? {} : { locale: config.locale }) });
    await sourceSettled(); binding.focus();
  }
  async function copySelection(kind: "copy" | "cut"): Promise<void> {
    check(); clipboard.capture(source!.cells.snapshot, input.state().selection, kind);
    if (kind === "cut") { ui.clipboardMode.value = "contents"; ui.clipboardSkipBlanks.checked = false; }
    ui.status.textContent = kind === "cut" ? "已选择剪切来源，请选择目标后粘贴。" : "已复制区域，请选择目标后粘贴。"; controls();
    binding.focus(); ui.input.ownerDocument.execCommand(kind);
  }
  async function pasteSelection(): Promise<void> {
    check(); await sourceSettled(); const plan = clipboard.plan(source!.cells.snapshot, input.state().selection), mode = ui.clipboardMode.value as WorkbookRangeCopyOptions["mode"];
    const receipt = plan.source.kind === "cut" ? await perform([{ op: "moveWorkbookRange", sheetId: plan.source.sheetId, range: plan.source.range, targetSheetId: plan.sheetId, target: plan.target,
      overwrittenContent: ui.clipboardOverwrite.checked ? "clear" : "reject" }]) : await commitCommand({ kind: "copy", documentId: plan.source.documentId, revision: plan.source.baseRevision,
        sheetId: plan.source.sheetId, source: plan.source.range, targetSheetId: plan.sheetId, target: plan.target, mode, skipBlanks: ui.clipboardSkipBlanks.checked });
    clipboard.applied(receipt.revision, plan.sheetId, plan.target, mode!); controls(); await sourceSettled(); binding.focus();
  }
  async function formatSelection(): Promise<void> {
    check(); const snapshot = source!.cells.snapshot, operations = ui.format.operations(snapshot, input.state().selection);
    if (!operations.length) { ui.status.textContent = "请选择要修改的格式。"; return; }
    await sourceSettled(); await commitCommand({ kind: "transaction", transaction: { protocolVersion: 1, documentId: snapshot.manifest.documentId,
      baseRevision: snapshot.manifest.revision, operations } }); await sourceSettled(); binding.focus();
  }
  async function formatSelectedText(): Promise<void> {
    check(); const selected = input.state().selection;
    if (!selected || selected.anchor.kind !== "cells") throw new DocumentError("InvalidOperation", "请先在公式栏选中文字。");
    const selection = { from: ui.input.selectionStart, to: ui.input.selectionEnd, text: ui.input.value }, properties = ui.format.fontProperties();
    if (!Object.keys(properties).length) throw new DocumentError("InvalidOperation", "请先选择要修改的字体设置。");
    await commitCommand({ kind: "formatText", documentId: selected.documentId, revision: selected.baseRevision, sheetId: selected.anchor.sheetId,
      at: selected.anchor.ranges[0]!.from, selection, properties });
    await sourceSettled(); binding.focus(false); if (ui.input.value === selection.text) ui.input.setSelectionRange(selection.from, selection.to);
  }
  function findQuery(): WorkbookFindQuery {
    const anchor = input.state().selection?.anchor, scope = ui.findScope.value, sheetId = preview!.layout().sheetId;
    if (scope === "selection" && anchor?.kind !== "cells") throw new DocumentError("InvalidOperation", "请先选择查找区域。");
    if (scope === "selection") findSelectionScope ??= anchor?.kind === "cells" ? anchor.ranges.map(range => ({ sheetId: anchor.sheetId, range })) : [];
    return { text: ui.find.value, field: ui.findField.value as NonNullable<WorkbookFindQuery["field"]>, matchCase: ui.matchCase.checked, entireCell: ui.entireCell.checked,
      ...(scope === "workbook" ? {} : { scope: scope === "selection" ? findSelectionScope! : [{ sheetId, range: { from: { row: 0, column: 0 }, to: { row: 1_048_575, column: 16383 } } }] }) };
  }
  function currentFind(query: WorkbookFindQuery): typeof found {
    const state = client.state(); return found?.key === JSON.stringify(query) && found.result.documentId === state.documentId && found.result.baseRevision === state.revision
      && (query.field !== "display" || found.view === source) ? found : undefined;
  }
  async function revealMatch(match: WorkbookFindMatch, field: WorkbookFindQuery["field"], revision: number): Promise<boolean> {
    const sheet = source!.cells.snapshot.metadata.workbook.sheets.find(sheet => sheet.sheetId === match.sheetId);
    if (sheet?.visibility !== "visible") return false;
    await choose(match.sheetId, { from: { row: match.row, column: match.column }, to: { row: match.row, column: match.column } }, false, true, revision);
    if (field !== "display") {
      const state = input.state(), prefix = state.text === "'" + match.text ? 1 : 0;
      ui.input.setSelectionRange(workbookInputTextView(state.text.slice(0, match.from + prefix)).length, workbookInputTextView(state.text.slice(0, match.to + prefix)).length);
    }
    return true;
  }
  async function search(replace?: "one" | "all"): Promise<void> {
    check(); if (pendingInput() || input.state().applying || commandAbort) throw new DocumentError("InvalidOperation", "请先完成当前输入或操作。");
    await sourceSettled(); const query = findQuery(), previous = currentFind(query), abort = new view!.AbortController();
    commandAbort = abort; controls(); let focusMatch = false;
    try {
      if (replace) {
        if (query.field === "display") throw new DocumentError("InvalidOperation", "显示值查找不能直接替换。");
        if (replace === "one" && (!previous || previous.index < 0)) throw new DocumentError("InvalidOperation", "请先查找一个当前匹配项。");
        const state = client.state(), receipt = await applyPrepared({ kind: "replace", documentId: state.documentId, revision: state.revision, query, replacement: ui.replacement.value,
          valueMode: ui.replacementMode.value as WorkbookReplaceOptions["valueMode"], ...(replace === "one" ? { only: { documentId: previous!.result.documentId, baseRevision: previous!.result.baseRevision, match: previous!.result.matches[previous!.index]! } } : {}) }, abort);
        found = undefined; await sourceSettled(); if (ui.findScope.value === "selection" && client.state().revision === receipt.revision) findSelectionScope = query.scope;
        if (!disposed) ui.status.textContent = receipt.changed ? "替换已应用。" : "没有需要替换的内容。";
      } else {
        let result = previous?.result;
        if (!result) {
          if (query.field === "display") { await scheduleRefresh(true); await sourceSettled(); }
          result = await client.find(query, { signal: abort.signal, ...(query.field === "display" ? { view: source! } : {}), ...(config.locale === undefined ? {} : { locale: config.locale }),
            onProgress: progress => { if (!disposed) ui.status.textContent = "正在查找 · " + progress.completed; } });
        }
        if (disposed || abort.signal.aborted) throw new DocumentError("OperationCancelled", "Workbook search was cancelled.");
        const index = result.matches.length ? previous ? (previous.index + 1) % result.matches.length : 0 : -1;
        found = { key: JSON.stringify(query), result, index, ...(query.field === "display" ? { view: source! } : {}) };
        const match = index >= 0 ? result.matches[index]! : undefined, visible = match ? await revealMatch(match, query.field, result.baseRevision) : false;
        focusMatch = visible && query.field !== "display";
        const sheet = match && source!.cells.snapshot.metadata.workbook.sheets.find(sheet => sheet.sheetId === match.sheetId);
        if (!disposed) ui.status.textContent = (index >= 0 ? `匹配 ${index + 1} / ${result.matches.length}${!visible ? ` · ${sheet?.name ?? match!.sheetId}（隐藏）!${address({ from: match!, to: match! })}` : ""}` : "没有匹配项")
          + (result.complete ? "。" : `；${result.diagnostics.length} 格无法取得显示值，结果不完整。`);
      }
    } catch (error) { if (!(error instanceof DocumentError && error.code === "OperationCancelled" && abort.signal.aborted)) throw error; if (!disposed) ui.status.textContent = "操作已取消。"; }
    finally { if (commandAbort === abort) commandAbort = undefined; controls(); if (!disposed && focusMatch) { view!.focus(); ui.input.focus({ preventScroll: true }); } }
  }
  const boundInput = { ...input, apply(options?: TransactionPrepareOptions) {
    const context = validationContext(config.confirmValidationWarnings ? workbookValidationInputKey(input.state()) : undefined);
    const next = withWorkbookAsyncValidationConfirmation(value => input.apply(value), options ?? {}, context); inputTask = next.then(() => {}, () => {}); return next;
  } };
  binding = bindWorkbookAsyncTextInput(ui.input, boundInput, {
    prepareOptions: () => preparation(), onError: report,
    onApplied() { controls(); ui.error.hidden = true; void scheduleRefresh(); if (!pendingInput()) binding.focus(); },
    onCancelled() { void runInteraction(cancelInput); },
    onNavigate(direction, extend) { void runInteraction(async () => { await sourceSettled(); if (pendingInput()) return; choosing++; try { const selected = await preview!.moveSelection(direction, extend); await acceptSelection(selected, true); } finally { choosing--; } }); },
    onHistory(direction) { void runInteraction(() => history(direction)); },
    onCopy(kind, selected, transfer) {
      const data = systemClipboard.data(selected);
      clipboard.capture(source!.cells.snapshot, selected, kind, data); transfer.setData("text/plain", data.text); transfer.setData("text/html", data.html);
      if (kind === "cut") { ui.clipboardMode.value = "contents"; ui.clipboardSkipBlanks.checked = false; }
      controls();
    },
    onClear(selected) { if (selected.anchor.kind !== "cells") return; const anchor = selected.anchor; void runInteraction(async () => {
      if (selected.baseRevision !== client.state().revision) throw new DocumentError("RevisionConflict", "Clear selection belongs to a previous source.");
      await perform([{ op: "clearWorkbookRanges", sheetId: anchor.sheetId, ranges: anchor.ranges.map(range => ({ from: { ...range.from }, to: { ...range.to } })), mode: "contents" }]);
    }); },
    onPaste(text, selected, html) { if (selected.anchor.kind !== "cells") return;
      if (clipboard.matches(text, html)) { void runInteraction(pasteSelection); return; }
      const anchor = selected.anchor; void runInteraction(async () => {
      await commitCommand({ kind: "paste", documentId: selected.documentId, revision: selected.baseRevision, sheetId: anchor.sheetId, at: anchor.ranges[0]!.from, text, ...(html === undefined ? {} : { html }) });
    }); },
  });
  const idle = (): boolean => !refreshing && !commandAbort && !input.state().applying && !input.state().selecting && !pendingInput() && source?.cells.snapshot.manifest.revision === client.state().revision;
  const auxiliaryHost: WorkbookControlsHost = { signal: events.signal, snapshot: () => source?.cells.snapshot, draftChanged: controls, prepareOptions: () => preparation(),
    ...(config.locale === undefined ? {} : { locale: config.locale }),
    selectSheet: sheetId => choose(sheetId, { from: { row: 0, column: 0 }, to: { row: 0, column: 0 } }),
    selection: () => { const anchor = input.state().selection?.anchor; return anchor?.kind === "cells" ? { sheetId: anchor.sheetId, range: anchor.ranges[0]! } : undefined; },
    idle, run: queueInteraction, perform: async (operations, options, expected = source?.cells.snapshot.manifest) => {
      const current = client.state(); if (!expected || current.documentId !== expected.documentId || current.revision !== expected.revision) throw new DocumentError("RevisionConflict", "辅助表单属于旧修订。");
      const receipt = await perform(operations, options); await sourceSettled(); return receipt; } };
  drawingControls = createWorkbookDrawingControls(ui.element.ownerDocument, ui.preview.parentNode!, auxiliaryHost);
  commentControls = createWorkbookCommentControls(ui.element.ownerDocument, ui.preview.parentNode!, auxiliaryHost);
  nameControls = createWorkbookNameControls(ui.element.ownerDocument, ui.preview.parentNode!, auxiliaryHost);
  worksheetControls = createWorkbookWorksheetControls(ui.element.ownerDocument, ui.preview.parentNode!, auxiliaryHost);
  hyperlinkControls = createWorkbookHyperlinkControls(ui.element.ownerDocument, ui.preview.parentNode!, auxiliaryHost);
  tableControls = createWorkbookTableControls(ui.element.ownerDocument, ui.preview.parentNode!, auxiliaryHost);
  tableStyleControls = createWorkbookTableStyleControls(ui.element.ownerDocument, ui.preview.parentNode!, auxiliaryHost);
  filterControls = createWorkbookFilterControls(ui.element.ownerDocument, ui.preview.parentNode!, auxiliaryHost);
  rulesControls = createWorkbookRulesControls(ui.element.ownerDocument, ui.preview.parentNode!, auxiliaryHost);
  function button(label: string, action: () => unknown, available: () => boolean, parent: HTMLElement = ui.tools): void {
    const element = ui.element.ownerDocument.createElement("button"); element.type = "button"; element.textContent = label;
    element.addEventListener("click", () => { try { const result = action(); if (result instanceof Promise) void result.catch(report); } catch (error) { report(error); } }, { signal: events.signal }); parent.append(element); buttons.push({ button: element, available });
  }
  button("确认输入", () => binding.apply(), () => !!input.state().selection && !input.state().composing && !input.state().applying && !commandAbort && !refreshing);
  button("取消输入", () => queueInteraction(cancelInput), () => !!input.state().selection && !input.state().applying && !commandAbort);
  button("撤销", () => queueInteraction(() => history("undo")), () => idle() && client.state().history.undoSteps > 0);
  button("重做", () => queueInteraction(() => history("redo")), () => idle() && client.state().history.redoSteps > 0);
  button("重新计算", recalculate, idle);
  button("向下填充", () => queueInteraction(() => fill("row")), idle);
  button("向右填充", () => queueInteraction(() => fill("column")), idle);
  button("数列填充", () => queueInteraction(() => { const anchor = input.state().selection?.anchor; return fill(anchor?.kind === "cells" && source ? workbookSeriesAxis(anchor.ranges[0]!, source.cells.snapshot.metadata.workbook.sheets.find(sheet => sheet.sheetId === anchor.sheetId)!.merges) : "row", true); }), idle);
  button("合并单元格", () => queueInteraction(() => changeMerge()), idle);
  button("合并并清除覆盖内容", () => queueInteraction(() => changeMerge(false, true)), idle);
  button("取消合并", () => queueInteraction(() => changeMerge(true)), idle);
  button("冻结选区之前", () => queueInteraction(() => changeFreeze()), idle);
  button("取消冻结", () => queueInteraction(() => changeFreeze(true)), idle);
  button("插入选区行", () => queueInteraction(() => changeAxis("row", "insert")), idle);
  button("删除选区行", () => queueInteraction(() => changeAxis("row", "delete")), idle);
  button("插入选区列", () => queueInteraction(() => changeAxis("column", "insert")), idle);
  button("删除选区列", () => queueInteraction(() => changeAxis("column", "delete")), idle);
  button("升序排序", () => queueInteraction(() => sortSelection(false)), idle);
  button("降序排序", () => queueInteraction(() => sortSelection(true)), idle);
  button("复制选区", () => queueInteraction(() => copySelection("copy")), idle);
  button("剪切选区", () => queueInteraction(() => copySelection("cut")), idle);
  button("粘贴选区", () => queueInteraction(pasteSelection), () => idle() && clipboard.available);
  button("应用单元格格式", () => queueInteraction(formatSelection), idle, ui.format.tools);
  button("应用选中文字字体", () => queueInteraction(formatSelectedText), () => idle() && !!input.state().selection && ui.input.selectionStart < ui.input.selectionEnd, ui.format.tools);
  button("取消操作", () => commandAbort?.abort(), () => !!commandAbort && !commandApplying);
  button("查找下一个", () => queueInteraction(() => search()), () => idle() && !!ui.find.value, ui.searchTools);
  button("替换当前项", () => queueInteraction(() => search("one")), () => idle() && ui.findField.value !== "display" && !!found && found.result.baseRevision === client.state().revision && found.index >= 0, ui.searchTools);
  button("全部替换", () => queueInteraction(() => search("all")), () => idle() && ui.findField.value !== "display" && !!ui.find.value, ui.searchTools);
  for (const field of [ui.find, ui.findField, ui.findScope, ui.matchCase, ui.entireCell]) for (const type of ["input", "change"]) field.addEventListener(type, () => { found = undefined; findSelectionScope = undefined; controls(); }, { signal: events.signal });
  ui.find.addEventListener("keydown", event => { if (event.key === "Enter") { event.preventDefault(); void runInteraction(() => search()); } }, { signal: events.signal });
  ui.element.addEventListener("keydown", event => { if ((event.ctrlKey || event.metaKey) && !event.altKey && !event.isComposing && ["f", "h"].includes(event.key.toLowerCase())) { event.preventDefault(); if (!commandAbort) { ui.find.focus(); ui.find.select(); } } }, { signal: events.signal });
  ui.mode.addEventListener("change", () => { try { input.setMode(ui.mode.value as WorkbookTextInputMode); binding.refresh(); controls(); } catch (error) { report(error); } }, { signal: events.signal });
  ui.address.addEventListener("keydown", event => {
    if (event.key !== "Enter") return; event.preventDefault(); const expression = "=" + ui.address.value.trim();
    void runInteraction(async () => {
      await sourceSettled(); const revision = client.state().revision, resolved = await client.resolveReference(expression, { sheetId: preview!.layout().sheetId, row: 0, column: 0 });
      if (resolved.kind !== "ranges" || resolved.areas.length !== 1) throw new DocumentError("InvalidOperation", "请输入一个有效的单元格或矩形区域地址。");
      const area = resolved.areas[0]!; await choose(area.sheetId, { from: area.from, to: area.to }, true, false, revision);
    });
  }, { signal: events.signal });
  for (const type of ["input", "keydown", "keyup", "select", "mouseup", "compositionstart", "compositionend"]) ui.input.addEventListener(type, controls, { signal: events.signal });
  const unsubscribe = client.subscribe(() => { if (!disposed) { found = undefined; findSelectionScope = undefined; void scheduleRefresh(); } }); controls(); void scheduleRefresh();
  const ready = sourceSettled().then(() => systemClipboard.settled()); void ready.catch(() => {});
  async function settled(): Promise<WorkbookRenderReport> {
    for (;;) {
      check(); const pendingInput = inputTask, pendingCommand = commandTask;
      await Promise.race([Promise.all([pendingInput, pendingCommand]), released]); check();
      const current = interaction; await Promise.race([current, released]); check();
      const result = await sourceSettled(); await systemClipboard.settled(); if (current === interaction && pendingInput === inputTask && pendingCommand === commandTask) return result;
    }
  }
  return Object.freeze({ element: ui.element, ready, settled, selection() { check(); return input.state().selection; }, inputState() { check(); return input.state(); },
    hasPendingInput() { check(); return pendingInput() || input.state().applying || !!commandAbort || pendingForms(); },
    layout() { check(); if (!preview) throw new DocumentError("InvalidOperation", "The workbook editor has no accepted view yet."); return preview.layout(); },
    setViewport: async (viewport: Parameters<WorkbookPreview["setViewport"]>[0]) => { await sourceSettled(); check(); return preview!.setViewport(viewport); },
    selectRange: (sheetId: string, range: DeepReadonly<WorkbookRange>) => runInteraction(() => choose(sheetId, range)), focusInput(selectAll = true) { check(); binding.focus(selectAll); }, applyInput: () => binding.apply(), cancelInput: () => runInteraction(cancelInput), perform,
    recalculate,
    dispose() { if (disposed) return; disposed = true; generation++; releaseWaiters(); refreshAbort?.abort(); commandAbort?.abort(); systemClipboard.dispose(); clipboard.clear(); layoutService.dispose(); unsubscribe(); events.abort(); binding.dispose(); input.dispose(); preview?.dispose(); source?.dispose(); ui.element.remove(); },
  });
}
