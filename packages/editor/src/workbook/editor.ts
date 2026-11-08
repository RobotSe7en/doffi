import { captureWorkbookModel, DocumentError, parseContract } from "@doffi/core";
import type { DeepReadonly, DocumentOperation, DocumentSelection, TransactionPrepareOptions, TransactionReceipt, WorkbookRange, WorkbookSession } from "@doffi/core";
import { createWorkbookPreviewSource, mountDocumentPreview } from "@doffi/render";
import type { WorkbookPreview, WorkbookPreviewSource, WorkbookRenderReport } from "@doffi/render";
import type { WorkbookEditorOptions, WorkbookEditorView } from "./editor-contracts.js";
import { createWorkbookEditorUi } from "./editor-ui.js";
import { withWorkbookValidationConfirmation, withWorkbookAsyncValidationConfirmation, workbookValidationInputKey } from "./validation-confirmation.js";
import { createWorkbookInputController } from "./input-controller.js";
import { bindWorkbookTextInput } from "./text-binding.js";
import { createWorkbookPasteTransaction } from "./text-input.js";
import type { WorkbookTextInputMode } from "./text-input.js";
import { createWorkbookFillTransaction } from "./fill.js";
import { workbookFillSeedCount } from "./fill-selection.js";
import { findWorkbookText } from "./find.js";
import { createWorkbookReplaceTransaction } from "./replace.js";
import type { WorkbookFindMatch, WorkbookFindQuery, WorkbookFindResult, WorkbookReplaceOptions } from "./find-contracts.js";
import { workbookInputTextView } from "./text-projection.js";
import { mountWorkbookRemoteEditor } from "./editor-remote.js";
import type { WorkbookRemoteEditorOptions, WorkbookRemoteEditorView } from "./editor-remote.js";
import type { WorkbookSessionConnection } from "./session-transport.js";
import { createWorkbookSortTransaction } from "./sort.js";
import { selectedWorkbookSort } from "./selected-sort.js";
import { createWorkbookTaskQueue } from "./task-queue.js";
import { WorkbookClipboard } from "./clipboard.js";
import { createWorkbookRangeCopyTransaction } from "./copy-range.js";
import type { WorkbookRangeCopyOptions } from "./copy-range.js";
import { createWorkbookTextFormatTransaction } from "./text-format.js";
import { createWorkbookDrawingControls } from "./drawing-controls.js";
import { createWorkbookCommentControls } from "./comment-controls.js";
import { createWorkbookNameControls } from "./names-controls.js";
import { createWorkbookWorksheetControls } from "./worksheet-controls.js";
import { createWorkbookHyperlinkControls } from "./hyperlink-controls.js";
import { createWorkbookTableControls } from "./table-controls.js";
import { createWorkbookFilterControls } from "./filter-controls.js";
import { createWorkbookRulesControls } from "./rules-controls.js";
import type { WorkbookControlsHost } from "./controls-host.js";

/** The host supplies its Worker connection or direct small-document session;
 * source cells, candidates and history remain behind that core boundary. */
export function mountWorkbookEditor(container: HTMLElement, session: WorkbookSessionConnection, options?: WorkbookRemoteEditorOptions): WorkbookRemoteEditorView;
export function mountWorkbookEditor(container: HTMLElement, session: WorkbookSession, options?: WorkbookEditorOptions): WorkbookEditorView;
export function mountWorkbookEditor(container: HTMLElement, session: WorkbookSession | WorkbookSessionConnection, options: WorkbookEditorOptions = {}): WorkbookEditorView | WorkbookRemoteEditorView {
  if (!session || typeof session !== "object") throw new DocumentError("InvalidOperation", "Workbook editing requires its owning session or authorized connection.");
  if (!("kind" in session)) return mountWorkbookRemoteEditor(container, session, options);
  return mountWorkbookDirectEditor(container, session, options);
}
function mountWorkbookDirectEditor(container: HTMLElement, session: WorkbookSession, options: WorkbookEditorOptions): WorkbookEditorView {
  const view = container?.ownerDocument?.defaultView;
  if (!view || !container.isConnected || session.kind !== "workbook") throw new DocumentError("InvalidOperation", "Workbook editing requires a connected container and its owning workbook session.");
  if (!options || typeof options !== "object" || Array.isArray(options) || Object.keys(options).some(key => !["title", "widthPt", "heightPt", "locale", "limits", "prepareOptions", "confirmValidationWarnings", "onError", "onSelectionChange", "onPendingChange"].includes(key))) throw new DocumentError("InvalidOperation", "Invalid workbook editor options.");
  let drawingControls: ReturnType<typeof createWorkbookDrawingControls> | undefined;
  let commentControls: ReturnType<typeof createWorkbookCommentControls> | undefined;
  let nameControls: ReturnType<typeof createWorkbookNameControls> | undefined;
  let worksheetControls: ReturnType<typeof createWorkbookWorksheetControls> | undefined;
  let hyperlinkControls: ReturnType<typeof createWorkbookHyperlinkControls> | undefined;
  let tableControls: ReturnType<typeof createWorkbookTableControls> | undefined;
  let filterControls: ReturnType<typeof createWorkbookFilterControls> | undefined;
  let rulesControls: ReturnType<typeof createWorkbookRulesControls> | undefined;
  if (options.confirmValidationWarnings !== undefined && typeof options.confirmValidationWarnings !== "function") throw new DocumentError("InvalidOperation", "Validation confirmation requires a host callback.");
  const config = { ...options }, ui = createWorkbookEditorUi(container, config.title ?? "工作簿编辑"), input = createWorkbookInputController(session), events = new view.AbortController();
  const clipboard = new WorkbookClipboard();
  let disposed = false, choosing = 0, refreshQueued = false, forceCalculation = false, latest: Promise<void> = Promise.resolve();
  let reportedPending: boolean | undefined;
  let interaction: Promise<void> = Promise.resolve();
  let revision = -1, generation = 0, calculationAbort: AbortController | undefined, commandAbort: AbortController | undefined, preview: WorkbookPreview;
  let renderedSource: WorkbookPreviewSource | undefined;
  let found: { readonly key: string; readonly result: WorkbookFindResult; index: number } | undefined;
  let findSelectionScope: WorkbookFindQuery["scope"];
  const searchFields = [ui.find, ui.replacement, ui.findField, ui.findScope, ui.matchCase, ui.entireCell, ui.replacementMode, ui.sortColumns, ui.sortHeader, ui.sortCase, ui.sortMethod, ui.clipboardMode, ui.clipboardOverwrite];
  const buttons: { readonly button: HTMLButtonElement; readonly available: () => boolean }[] = [];
  function check(): void { if (disposed) throw new DocumentError("InvalidOperation", "This workbook editor has been disposed."); }
  function validationContext(draftKey?: string, signal?: AbortSignal) {
    const expectedDraft = draftKey ?? (config.confirmValidationWarnings ? workbookValidationInputKey(input.state()) : undefined);
    const { documentId, revision: baseRevision } = session.snapshot().manifest;
    return { documentId, baseRevision, confirm: config.confirmValidationWarnings, assertCurrent() {
      check(); const current = session.snapshot().manifest;
      if (current.documentId !== documentId || current.revision !== baseRevision) throw new DocumentError("RevisionConflict", "Workbook changed during validation confirmation.");
      if (signal?.aborted || expectedDraft !== undefined && workbookValidationInputKey(input.state()) !== expectedDraft)
        throw new DocumentError("OperationCancelled", "Workbook input changed during validation confirmation.");
    } };
  }
  function preparation(input?: TransactionPrepareOptions): TransactionPrepareOptions {
    const supplied = input ?? config.prepareOptions?.() ?? {};
    return config.locale !== undefined && supplied.calculationContext?.locale === undefined ? { ...supplied, calculationContext: { ...supplied.calculationContext, locale: config.locale } } : supplied;
  }
  function runInteraction<T>(action: () => Promise<T>): Promise<T> { const work = interaction.then(action); interaction = work.then(() => {}, report); return work; }
  function pendingInput(): boolean { return binding.hasPendingInput(); }
  function pendingForms(): boolean { return [drawingControls, commentControls, nameControls, worksheetControls, hyperlinkControls, tableControls, filterControls, rulesControls].some(controls => controls?.hasPendingInput()); }
  function report(error: unknown): void {
    if (disposed) return; const problem = error instanceof Error ? error : new Error("工作簿编辑失败。");
    ui.error.textContent = problem instanceof DocumentError && problem.code === "RevisionConflict" ? "文档已变化。当前输入已保留，请先复制输入或取消后重新选择。"
      : problem instanceof DocumentError && problem.details.some(detail => detail.code === "merge-covered-content") ? "合并范围包含其他内容。若确认删除，请使用“合并并清除覆盖内容”。" : problem.message; ui.error.hidden = false;
    try { config.onError?.(problem); } catch { /* Host callback does not alter the edit. */ }
  }
  function controls(): void {
    if (disposed) return;
    for (const { button, available } of buttons) button.disabled = !available();
    const state = input.state(); ui.mode.value = state.mode; ui.mode.disabled = Boolean(commandAbort) || state.composing || !state.selection;
    ui.address.disabled = Boolean(commandAbort) || revision < 0;
    for (const field of searchFields) field.disabled = Boolean(commandAbort);
    ui.clipboardMode.disabled ||= clipboard.kind === "cut";
    ui.clipboardOverwrite.disabled ||= clipboard.kind !== "cut";
    ui.insertedFormat.disabled = Boolean(commandAbort);
    for (const field of ui.format.fields) field.disabled = Boolean(commandAbort);
    drawingControls?.refresh();
    commentControls?.refresh();
    nameControls?.refresh();
    worksheetControls?.refresh();
    hyperlinkControls?.refresh();
    tableControls?.refresh();
    filterControls?.refresh();
    rulesControls?.refresh();
    const pending = pendingInput() || Boolean(commandAbort) || pendingForms();
    if (reportedPending !== pending) { reportedPending = pending; try { config.onPendingChange?.(pending); } catch { /* Reporting cannot change source. */ } }
  }
  function notify(): void { try { config.onSelectionChange?.(input.state().selection); } catch { /* Host-owned callback. */ } }
  function address(range: DeepReadonly<WorkbookRange>): string {
    const cell = (point: WorkbookRange["from"]) => { let column = point.column + 1, name = ""; while (column) { column--; name = String.fromCharCode(65 + column % 26) + name; column = Math.floor(column / 26); } return name + (point.row + 1); };
    return cell(range.from) + (range.from.row === range.to.row && range.from.column === range.to.column ? "" : ":" + cell(range.to));
  }
  function acceptSelection(selected: DocumentSelection, focus: boolean): void {
    input.select(selected); binding.refresh(); ui.input.disabled = false;
    if (selected.anchor.kind === "cells") ui.address.value = address(selected.anchor.ranges[0]!);
    controls(); notify(); if (focus) binding.focus();
  }
  function readCurrentSelection(): void {
    if (pendingInput()) return;
    const selected = input.state().selection, snapshot = session.snapshot();
    if (!selected || selected.anchor.kind !== "cells") return;
    const sheetId = selected.anchor.sheetId;
    if (!snapshot.metadata.workbook.sheets.some(sheet => sheet.sheetId === sheetId)) return;
    acceptSelection({ ...selected, baseRevision: snapshot.manifest.revision }, ui.input === ui.input.ownerDocument.activeElement || ui.element === ui.element.ownerDocument.activeElement);
  }
  async function choose(sheetId: string, range: DeepReadonly<WorkbookRange>, focus = false, retainFind = false): Promise<DocumentSelection> {
    check(); if (pendingInput()) throw new DocumentError("InvalidOperation", "请先确认或取消当前单元格输入。");
    if (!retainFind) { found = undefined; findSelectionScope = undefined; }
    await sourceSettled(); check();
    choosing++;
    try {
      if (preview.layout().sheetId !== sheetId) await preview.showSheet(sheetId);
      const selected = await preview.selectRange(range); check(); acceptSelection(selected, focus); return selected;
    } finally { choosing--; }
  }
  async function restoreSelection(): Promise<void> {
    const selected = input.state().selection;
    if (selected?.anchor.kind !== "cells") return;
    const anchor = selected.anchor; await sourceSettled(); check();
    if (!preview.snapshot().metadata.workbook.sheets.some(sheet => sheet.sheetId === anchor.sheetId)) return;
    choosing++;
    try { if (preview.layout().sheetId !== anchor.sheetId) await preview.showSheet(anchor.sheetId); await preview.selectRange(anchor.ranges[0]!, false); if (pendingInput()) binding.focus(false); }
    finally { choosing--; }
  }
  function sourceSelection(selected: DocumentSelection | undefined): void {
    if (disposed || choosing) return;
    found = undefined; findSelectionScope = undefined;
    commandAbort?.abort();
    try {
      if (!selected) {
        const previous = input.state().selection, sheetId = preview.layout().sheetId;
        if (previous?.anchor.kind !== "cells" || previous.anchor.sheetId === sheetId) return;
        if (pendingInput()) { report(new DocumentError("InvalidOperation", "请先确认或取消当前单元格输入。")); void runInteraction(() => restoreSelection()); }
        else void runInteraction(() => choose(sheetId, { from: { row: 0, column: 0 }, to: { row: 0, column: 0 } }));
        return;
      }
      if (selected.anchor.kind !== "cells") return;
      if (pendingInput()) {
        if (!binding.apply()) { void runInteraction(() => restoreSelection()); return; }
        const anchor = selected.anchor; void runInteraction(() => choose(anchor.sheetId, anchor.ranges[0]!, true));
      } else { acceptSelection(selected, true); ui.error.hidden = true; }
    } catch (error) { report(error); }
  }
  async function refresh(force: boolean): Promise<void> {
    check(); const ticket = ++generation, abort = new view!.AbortController(); calculationAbort?.abort(); calculationAbort = abort;
    const model = captureWorkbookModel(session), snapshot = model.snapshot(); ui.status.textContent = "正在更新工作簿…";
    const prepared = preparation();
    const calculation = force || snapshot.metadata.workbook.calculationMode === "auto" ? await session.calculate({ signal: abort.signal,
      yield: () => new Promise<void>(resolve => setTimeout(resolve, 0)), ...(prepared.calculationContext ? { context: prepared.calculationContext } : {}) }) : undefined;
    if (disposed || ticket !== generation || abort.signal.aborted) return;
    const source = createWorkbookPreviewSource(model, calculation, renderedSource);
    preview.update(source); await preview.settled();
    if (disposed || ticket !== generation) return;
    renderedSource = source;
    revision = snapshot.manifest.revision;
    const state = input.state();
    if (!pendingInput()) {
      const anchor = state.selection?.anchor, sheetId = preview.layout().sheetId;
      const range = anchor?.kind === "cells" && anchor.sheetId === sheetId ? anchor.ranges[0]! : { from: { row: 0, column: 0 }, to: { row: 0, column: 0 } };
      choosing++;
      try { const selected = await preview.selectRange(range, false); if (!disposed && ticket === generation && !pendingInput()) acceptSelection(selected, false); }
      finally { choosing--; }
    }
    ui.status.textContent = pendingInput() && state.selection?.baseRevision !== revision ? "文档已更新；当前输入仍绑定原修订。" : "修订 " + revision + (snapshot.metadata.workbook.calculationMode === "manual" && !force ? " · 手动计算" : ""); controls();
  }
  function scheduleRefresh(force = false): Promise<void> {
    check(); forceCalculation ||= force; calculationAbort?.abort();
    if (!refreshQueued) {
      refreshQueued = true;
      const work = Promise.resolve().then(() => { refreshQueued = false; const force = forceCalculation; forceCalculation = false; return refresh(force); });
      latest = work; void work.catch(error => { if (work === latest) report(error); });
    }
    return latest;
  }
  async function sourceSettled(): Promise<WorkbookRenderReport> {
    check();
    for (;;) { const work = latest; try { await work; } catch (error) { if (work === latest) throw error; } if (work === latest) { const result = await preview.settled(); if (work === latest) return result; } }
  }
  async function settled(): Promise<WorkbookRenderReport> {
    for (;;) { const work = interaction; await work; const result = await sourceSettled(); if (work === interaction) return result; }
  }
  function perform(operations: readonly DocumentOperation[], preparedOptions?: TransactionPrepareOptions): TransactionReceipt {
    check(); if (pendingInput()) throw new DocumentError("InvalidOperation", "请先确认或取消当前单元格输入。");
    const manifest = session.snapshot().manifest;
    const value = { protocolVersion: 1 as const, documentId: manifest.documentId, baseRevision: manifest.revision, operations: [...operations] };
    const transaction = config.confirmValidationWarnings ? parseContract("document-transaction", value) : value;
    const candidate = withWorkbookValidationConfirmation(options => session.prepare(transaction, options), preparation(preparedOptions), validationContext());
    const receipt = session.commit(candidate);
    readCurrentSelection(); controls(); ui.error.hidden = true; return receipt;
  }
  function cancelInput(): void { check(); input.cancel(); binding.refresh(); readCurrentSelection(); controls(); ui.error.hidden = true; }
  function history(direction: "undo" | "redo"): void {
    check(); if (pendingInput()) throw new DocumentError("InvalidOperation", "请先确认或取消当前单元格输入。");
    session[direction](); readCurrentSelection(); controls(); ui.error.hidden = true;
  }
  async function changeMerge(unmerge = false, clear = false): Promise<void> {
    check(); const selected = input.state().selection, anchor = selected?.anchor;
    if (!selected || anchor?.kind !== "cells" || anchor.ranges.length !== 1) throw new DocumentError("InvalidOperation", "请选择一个合并区域。");
    if (selected.baseRevision !== session.snapshot().manifest.revision) throw new DocumentError("RevisionConflict", "Merge selection belongs to a previous source.");
    const target = { sheetId: anchor.sheetId, range: anchor.ranges[0]! };
    perform([unmerge ? { op: "unmergeWorkbookCells", ...target } : { op: "mergeWorkbookCells", ...target, coveredContent: clear ? "clear" : "reject" }]);
    await sourceSettled(); binding.focus();
  }
  async function changeFreeze(clear = false): Promise<void> {
    check(); const selected = input.state().selection, anchor = selected?.anchor;
    if (!selected || anchor?.kind !== "cells") throw new DocumentError("InvalidOperation", "请先选择冻结位置。");
    if (selected.baseRevision !== session.snapshot().manifest.revision) throw new DocumentError("RevisionConflict", "Freeze selection belongs to a previous source.");
    const at = anchor.ranges[0]!.from;
    if (!clear) {
      const offset = preview.sourceOffset(at), viewport = preview.layout().viewport;
      if (offset.revision !== selected.baseRevision || offset.sheetId !== anchor.sheetId) throw new DocumentError("RevisionConflict", "Freeze geometry belongs to a previous view.");
      if (offset.leftPt >= viewport.widthPt || offset.topPt >= viewport.heightPt) throw new DocumentError("InvalidOperation", "冻结区域将占满当前视口，请选择更靠前的位置。");
    }
    perform([{ op: "updateWorkbookSheet", sheetId: anchor.sheetId, changes: { frozen: { rows: clear ? 0 : at.row, columns: clear ? 0 : at.column } } }]);
    await sourceSettled(); binding.focus();
  }
  async function sortSelection(descending: boolean): Promise<void> {
    check(); if (pendingInput() || commandAbort) throw new DocumentError("InvalidOperation", "请先完成当前输入或操作。");
    await sourceSettled();
    const selected = input.state().selection, target = selectedWorkbookSort(session.snapshot(), selected,
      { columns: ui.sortColumns.value, header: ui.sortHeader.checked, caseSensitive: ui.sortCase.checked, method: ui.sortMethod.value }, descending);
    const abort = new view!.AbortController(), tasks = createWorkbookTaskQueue(), prepared = preparation(); commandAbort = abort; controls();
    let candidate: ReturnType<WorkbookSession["prepare"]> | undefined, committed = false;
    try {
      const calculation = await session.calculate({ signal: abort.signal, yield: () => tasks.yield(), ...(prepared.calculationContext ? { context: prepared.calculationContext } : {}) });
      const transaction = await createWorkbookSortTransaction(session, target.sheetId, target.sort, { calculation, signal: abort.signal, yield: () => tasks.yield(),
        ...(config.locale === undefined ? {} : { locale: config.locale }), onProgress: progress => { if (!disposed) ui.status.textContent = "正在排序 · " + progress.completed; } });
      candidate = await withWorkbookAsyncValidationConfirmation(options => session.prepareAsync(transaction, options, { signal: abort.signal, yield: () => tasks.yield() }), prepared, validationContext(undefined, abort.signal));
      if (disposed || abort.signal.aborted) throw new DocumentError("OperationCancelled", "Sorting was cancelled before application.");
      session.commit(candidate); committed = true; if (!disposed) ui.error.hidden = true;
    } finally {
      if (candidate && !committed) try { session.discard(candidate); } catch { /* Host disposal or another source may release it. */ }
      tasks.dispose(); if (commandAbort === abort) commandAbort = undefined; controls();
    }
    await sourceSettled(); binding.focus();
  }
  function copySelection(kind: "copy" | "cut"): void {
    check(); if (pendingInput() || commandAbort) throw new DocumentError("InvalidOperation", "请先完成当前输入或操作。");
    clipboard.capture(session.snapshot(), input.state().selection, kind); if (kind === "cut") ui.clipboardMode.value = "contents";
    ui.status.textContent = kind === "cut" ? "已选择剪切来源，请选择目标后粘贴。" : "已复制区域，请选择目标后粘贴。"; controls();
  }
  async function formatSelection(): Promise<void> {
    check(); if (pendingInput() || commandAbort) throw new DocumentError("InvalidOperation", "请先完成当前输入或操作。");
    const snapshot = session.snapshot(), operations = ui.format.operations(snapshot, input.state().selection);
    if (!operations.length) { ui.status.textContent = "请选择要修改的格式。"; return; }
    await sourceSettled();
    const abort = new view!.AbortController(); commandAbort = abort; controls();
    try { await commitCommand({ protocolVersion: 1, documentId: snapshot.manifest.documentId, baseRevision: snapshot.manifest.revision, operations }, abort); readCurrentSelection(); }
    finally { if (commandAbort === abort) commandAbort = undefined; controls(); }
    await sourceSettled(); binding.focus();
  }
  async function formatSelectedText(): Promise<void> {
    check(); if (pendingInput() || commandAbort) throw new DocumentError("InvalidOperation", "请先完成当前输入或操作。");
    const selected = input.state().selection;
    if (!selected || selected.anchor.kind !== "cells") throw new DocumentError("InvalidOperation", "请先在公式栏选中文字。");
    const selection = { from: ui.input.selectionStart, to: ui.input.selectionEnd, text: ui.input.value }, properties = ui.format.fontProperties();
    if (!Object.keys(properties).length) throw new DocumentError("InvalidOperation", "请先选择要修改的字体设置。");
    const transaction = createWorkbookTextFormatTransaction(session, selected.anchor.sheetId, selected.anchor.ranges[0]!.from, selection, properties);
    if (transaction.baseRevision !== selected.baseRevision) throw new DocumentError("RevisionConflict", "选中文字的选区已过期。");
    const abort = new view!.AbortController(); commandAbort = abort; controls();
    try { await commitCommand(transaction, abort); readCurrentSelection(); }
    finally { if (commandAbort === abort) commandAbort = undefined; controls(); }
    await sourceSettled(); binding.focus(false); if (ui.input.value === selection.text) ui.input.setSelectionRange(selection.from, selection.to);
  }
  async function pasteSelection(): Promise<void> {
    check(); if (pendingInput() || commandAbort) throw new DocumentError("InvalidOperation", "请先完成当前输入或操作。");
    await sourceSettled(); const plan = clipboard.plan(session.snapshot(), input.state().selection), mode = ui.clipboardMode.value as NonNullable<WorkbookRangeCopyOptions["mode"]>;
    const abort = new view!.AbortController(), tasks = createWorkbookTaskQueue(), prepared = preparation(); commandAbort = abort; controls();
    let candidate: ReturnType<WorkbookSession["prepare"]> | undefined, committed = false;
    try {
      const calculation = plan.source.kind === "copy" && mode === "values" ? await session.calculate({ signal: abort.signal, yield: () => tasks.yield(), ...(prepared.calculationContext ? { context: prepared.calculationContext } : {}) }) : undefined;
      const transaction = plan.source.kind === "cut" ? { protocolVersion: 1 as const, documentId: plan.source.documentId, baseRevision: plan.source.baseRevision, operations: [
        { op: "moveWorkbookRange" as const, sheetId: plan.source.sheetId, range: plan.source.range, targetSheetId: plan.sheetId, target: plan.target, overwrittenContent: ui.clipboardOverwrite.checked ? "clear" as const : "reject" as const },
      ] } : await createWorkbookRangeCopyTransaction(session, plan.source.sheetId, plan.source.range, plan.target, { targetSheetId: plan.sheetId, mode, ...(calculation ? { calculation } : {}), signal: abort.signal, yield: () => tasks.yield() });
      candidate = await withWorkbookAsyncValidationConfirmation(options => session.prepareAsync(transaction, options, { signal: abort.signal, yield: () => tasks.yield() }), prepared, validationContext(undefined, abort.signal));
      if (disposed || abort.signal.aborted) throw new DocumentError("OperationCancelled", "Pasting was cancelled before application.");
      const receipt = session.commit(candidate); committed = true; clipboard.applied(receipt.revision, plan.sheetId, plan.target, mode!); if (!disposed) ui.error.hidden = true;
    } finally {
      if (candidate && !committed) try { session.discard(candidate); } catch { /* Host disposal or another source may release it. */ }
      tasks.dispose(); if (commandAbort === abort) commandAbort = undefined; controls();
    }
    await sourceSettled(); binding.focus();
  }
  async function changeAxis(axis: "row" | "column", action: "insert" | "delete"): Promise<void> {
    check(); if (pendingInput() || commandAbort) throw new DocumentError("InvalidOperation", "请先完成当前输入或操作。");
    const selected = input.state().selection, anchor = selected?.anchor;
    if (!selected || anchor?.kind !== "cells" || anchor.ranges.length !== 1) throw new DocumentError("InvalidOperation", "请选择一个行列区域。");
    if (selected.baseRevision !== session.snapshot().manifest.revision) throw new DocumentError("RevisionConflict", "Axis selection belongs to a previous source.");
    const range = anchor.ranges[0]!, abort = new view!.AbortController(); commandAbort = abort; ui.input.disabled = true; controls(); ui.error.hidden = true;
    try {
      await commitCommand({ protocolVersion: 1, documentId: selected.documentId, baseRevision: selected.baseRevision, operations: [
        { op: "editWorkbookAxis", sheetId: anchor.sheetId, axis, action, index: range.from[axis], count: range.to[axis] - range.from[axis] + 1,
          ...(action === "insert" ? { formatSource: ui.insertedFormat.value as "before" | "after" | "none" } : {}) },
      ] }, abort);
      readCurrentSelection(); await sourceSettled();
    } finally { if (commandAbort === abort) commandAbort = undefined; if (!disposed) { ui.input.disabled = !input.state().selection; controls(); } }
    binding.focus();
  }
  async function fill(axis: "row" | "column", series = false): Promise<void> {
    check(); if (pendingInput() || commandAbort) throw new DocumentError("InvalidOperation", "请先完成当前输入或操作。");
    const selected = input.state().selection;
    if (selected?.anchor.kind !== "cells") throw new DocumentError("InvalidOperation", "请先选择填充区域。");
    const range = selected.anchor.ranges[0]!, count = range.to[axis] - range.from[axis] + 1;
    if (series && range.from.row !== range.to.row && range.from.column !== range.to.column) throw new DocumentError("InvalidOperation", "数列填充请选择单行或单列。");
    if (count < 2) throw new DocumentError("InvalidOperation", axis === "row" ? "向下填充需要至少两行。" : "向右填充需要至少两列。");
    const model = captureWorkbookModel(session), manifest = model.snapshot().manifest, sheetId = selected.anchor.sheetId;
    if (selected.documentId !== manifest.documentId || selected.baseRevision !== manifest.revision) throw new DocumentError("RevisionConflict", "Fill selection belongs to a previous source.");
    let seeds = series ? 1 : workbookFillSeedCount(range, model.snapshot().metadata.workbook.sheets.find(sheet => sheet.sheetId === sheetId)!.merges, axis);
    if (series) {
      const next = { ...range.from, [axis]: range.from[axis] + 1 }, value = model.getCell(sheetId, next.row, next.column);
      if (value?.formula || value && value.value.type !== "number" && value.value.type !== "blank") throw new DocumentError("InvalidOperation", "数列填充的种子必须是数字或空白。");
      if (value?.value.type === "number") seeds = 2;
    }
    const source = { from: { ...range.from }, to: { ...range.to, [axis]: range.from[axis] + seeds - 1 } };
    const target = { from: { ...range.from, [axis]: range.from[axis] + seeds }, to: { ...range.to } };
    if (target.from[axis] > target.to[axis]) return;
    const abort = new view!.AbortController(); commandAbort = abort; ui.input.disabled = true; controls(); ui.error.hidden = true;
    try {
      const transaction = await createWorkbookFillTransaction(session, sheetId, source, target, { mode: series ? "series" : "copy", axis, signal: abort.signal,
        yield: () => new Promise<void>(resolve => setTimeout(resolve, 0)), onProgress: progress => { if (!disposed) ui.status.textContent = `正在填充 ${progress.completed} / ${progress.total}`; } });
      if (disposed || abort.signal.aborted) throw new DocumentError("OperationCancelled", "Workbook fill was cancelled.");
      const receipt = await commitCommand(transaction, abort); readCurrentSelection();
      if (!receipt.changed) ui.status.textContent = "选区没有变化。";
    } catch (error) {
      if (!(error instanceof DocumentError && error.code === "OperationCancelled" && abort.signal.aborted)) throw error;
      if (!disposed) ui.status.textContent = "操作已取消。";
    } finally { if (commandAbort === abort) commandAbort = undefined; if (!disposed) { ui.input.disabled = !input.state().selection; controls(); } }
  }
  async function commitCommand(transaction: Parameters<WorkbookSession["prepare"]>[0], abort: AbortController): Promise<TransactionReceipt> {
    const context = validationContext(undefined, abort.signal);
    const candidate = await withWorkbookAsyncValidationConfirmation(options => session.prepareAsync(transaction, options, { signal: abort.signal,
      yield: () => new Promise<void>(resolve => setTimeout(resolve, 0)), onProgress: progress => {
        if (!disposed) ui.status.textContent = `正在${progress.phase === "calculation" ? "计算" : progress.phase === "validation" ? "检查更改" : "准备更改"} · ${progress.completed}${progress.total ? ` / ${progress.total}` : ""}`;
      } }), preparation(), context);
    if (disposed || abort.signal.aborted) {
      try { session.discard(candidate); } catch { /* A host-disposed session has already released its candidates. */ }
      throw new DocumentError("OperationCancelled", "Workbook command was cancelled before publication.");
    }
    try { return session.commit(candidate); }
    catch (error) { try { session.discard(candidate); } catch { /* A concurrent commit or disposal may already have released it. */ } throw error; }
  }
  function findQuery(): WorkbookFindQuery {
    const selected = input.state().selection, scope = ui.findScope.value, sheetId = preview.layout().sheetId;
    if (scope === "selection" && selected?.anchor.kind !== "cells") throw new DocumentError("InvalidOperation", "请先选择查找区域。");
    const anchor = selected?.anchor;
    const ranges = anchor?.kind === "cells" ? anchor.ranges.map(range => ({ sheetId: anchor.sheetId, range })) : [];
    if (scope === "selection") findSelectionScope ??= ranges;
    return { text: ui.find.value, field: ui.findField.value as NonNullable<WorkbookFindQuery["field"]>, matchCase: ui.matchCase.checked, entireCell: ui.entireCell.checked,
      ...(scope === "workbook" ? {} : { scope: scope === "selection" ? findSelectionScope! : [{ sheetId, range: { from: { row: 0, column: 0 }, to: { row: 1_048_575, column: 16383 } } }] }) };
  }
  function findKey(query: WorkbookFindQuery): string { return JSON.stringify(query); }
  function currentFind(query: WorkbookFindQuery): typeof found {
    const snapshot = session.snapshot(); return found?.key === findKey(query) && found.result.documentId === snapshot.manifest.documentId && found.result.baseRevision === snapshot.manifest.revision ? found : undefined;
  }
  async function revealMatch(match: WorkbookFindMatch, field: WorkbookFindQuery["field"]): Promise<string> {
    const sheet = session.snapshot().metadata.workbook.sheets.find(sheet => sheet.sheetId === match.sheetId);
    if (sheet && sheet.visibility !== "visible") return ` · ${sheet.name}（隐藏）!${address({ from: match, to: match })}`;
    await choose(match.sheetId, { from: { row: match.row, column: match.column }, to: { row: match.row, column: match.column } }, false, true);
    if (field === "display") return "";
    const state = input.state(), cell = session.getCell(match.sheetId, match.row, match.column);
    const raw = cell?.formula?.expression ?? (!cell || cell.value.type === "blank" ? "" : cell.value.type === "boolean" ? cell.value.value ? "TRUE" : "FALSE" : String(cell.value.value));
    const prefix = state.text === "'" + raw ? 1 : 0;
    binding.focus(false); ui.input.setSelectionRange(workbookInputTextView(state.text.slice(0, match.from + prefix)).length, workbookInputTextView(state.text.slice(0, match.to + prefix)).length);
    return "";
  }
  async function search(replace: "all" | "one" | undefined = undefined): Promise<void> {
    check(); if (pendingInput() || commandAbort) throw new DocumentError("InvalidOperation", "请先完成当前输入或操作。");
    await sourceSettled(); check();
    const query = findQuery(), key = findKey(query), previous = currentFind(query), replacement = ui.replacement.value;
    const valueMode = ui.replacementMode.value as WorkbookReplaceOptions["valueMode"];
    if (replace === "one" && (!previous || previous.index < 0)) throw new DocumentError("InvalidOperation", "请先查找并选中要替换的匹配项。");
    const abort = new view!.AbortController(); commandAbort = abort; ui.input.disabled = true; controls(); ui.error.hidden = true;
    try {
      const scan = { signal: abort.signal, yield: () => new Promise<void>(resolve => setTimeout(resolve, 0)), onProgress: (progress: { readonly cells: number; readonly matches: number }) => {
        if (!disposed) ui.status.textContent = `正在${replace ? "替换" : "查找"} · 已检查 ${progress.cells} 格 · 匹配 ${progress.matches} 项`;
      }, ...(config.locale === undefined ? {} : { locale: config.locale }) };
      if (replace) {
        const transaction = await createWorkbookReplaceTransaction(session, query, replacement, { ...scan, ...(valueMode === undefined ? {} : { valueMode }),
          ...(replace === "one" && previous ? { only: { documentId: previous.result.documentId, baseRevision: previous.result.baseRevision, match: previous.result.matches[previous.index]! } } : {}) });
        if (disposed || abort.signal.aborted) throw new DocumentError("OperationCancelled", "Workbook replacement was cancelled.");
        const count = transaction.operations.reduce((sum, operation) => sum + (operation.op === "setWorkbookCells" ? operation.cells.length : 0), 0);
        const receipt = await commitCommand(transaction, abort); found = undefined;
        readCurrentSelection();
        await sourceSettled();
        if (!disposed) {
          if (ui.findScope.value === "selection" && session.snapshot().manifest.revision === receipt.revision) findSelectionScope = query.scope;
          ui.status.textContent = receipt.changed ? `已替换 ${count} 个单元格。` : "没有需要替换的内容。";
        }
      } else {
        let result = query.field === "display" ? undefined : previous?.result;
        if (!result) {
          const prepared = preparation(), calculation = query.field === "display" ? await session.calculate({ signal: abort.signal, yield: scan.yield, ...(prepared.calculationContext ? { context: prepared.calculationContext } : {}) }) : undefined;
          if (calculation) {
            if (disposed || abort.signal.aborted) throw new DocumentError("OperationCancelled", "Workbook search was cancelled.");
            const model = captureWorkbookModel(session), anchor = input.state().selection?.anchor, ticket = ++generation;
            const source = createWorkbookPreviewSource(model, calculation, renderedSource);
            choosing++;
            try {
              preview.update(source); await preview.settled();
              if (disposed || abort.signal.aborted || ticket !== generation) throw new DocumentError("OperationCancelled", "Workbook display search was superseded.");
              renderedSource = source;
              if (!disposed && !abort.signal.aborted && anchor?.kind === "cells" && preview.layout().sheetId === anchor.sheetId) await preview.selectRange(anchor.ranges[0]!, false);
            } finally { choosing--; }
          }
          result = await findWorkbookText(session, query, { ...scan, ...(calculation === undefined ? {} : { calculation }) });
        }
        if (disposed || abort.signal.aborted) throw new DocumentError("OperationCancelled", "Workbook search was cancelled.");
        const index = result.matches.length ? previous ? (previous.index + 1) % result.matches.length : 0 : -1;
        found = { key, result, index };
        const location = index >= 0 ? await revealMatch(result.matches[index]!, query.field) : "";
        if (disposed || abort.signal.aborted) throw new DocumentError("OperationCancelled", "Workbook search was cancelled.");
        ui.status.textContent = `${index >= 0 ? `匹配 ${index + 1} / ${result.matches.length}${location}` : "没有匹配项"}${result.complete ? "。" : `；${result.diagnostics.length} 格无法取得显示值，结果不完整。`}`;
      }
    } catch (error) {
      if (!(error instanceof DocumentError && error.code === "OperationCancelled" && abort.signal.aborted)) throw error;
      if (!disposed) ui.status.textContent = "操作已取消。";
    } finally { if (commandAbort === abort) commandAbort = undefined; if (!disposed) { ui.input.disabled = !input.state().selection; controls(); } }
  }
  const boundInput = { ...input, apply(options?: TransactionPrepareOptions) {
    const context = validationContext(config.confirmValidationWarnings ? workbookValidationInputKey(input.state()) : undefined);
    return withWorkbookValidationConfirmation(value => input.apply(value), options ?? {}, context);
  } };
  const binding = bindWorkbookTextInput(ui.input, boundInput, {
    prepareOptions: () => preparation(),
    onApplied() { controls(); ui.error.hidden = true; }, onCancelled: cancelInput,
    onNavigate(direction, extend) { void runInteraction(async () => {
      found = undefined; findSelectionScope = undefined;
      await sourceSettled(); if (pendingInput()) return; choosing++;
      try { const selected = await preview.moveSelection(direction, extend); check(); acceptSelection(selected, true); } finally { choosing--; }
    }); },
    onHistory: history,
    onClear(selected) {
      if (selected.anchor.kind !== "cells") return;
      const manifest = session.snapshot().manifest;
      if (selected.documentId !== manifest.documentId || selected.baseRevision !== manifest.revision) throw new DocumentError("RevisionConflict", "The selected clear range belongs to a previous revision.");
      perform([{ op: "clearWorkbookRanges", sheetId: selected.anchor.sheetId, ranges: selected.anchor.ranges.map(range => ({ from: { ...range.from }, to: { ...range.to } })), mode: "contents" }]);
    },
    onPaste(text, selected) {
      if (selected.anchor.kind !== "cells") return;
      const transaction = createWorkbookPasteTransaction(session.snapshot(), selected.anchor.sheetId, selected.anchor.ranges[0]!.from, text);
      const context = validationContext(config.confirmValidationWarnings ? workbookValidationInputKey(input.state()) : undefined);
      const candidate = withWorkbookValidationConfirmation(options => session.prepare({ ...transaction, documentId: selected.documentId, baseRevision: selected.baseRevision }, options), preparation(), context);
      const receipt = session.commit(candidate);
      readCurrentSelection(); controls();
      if (receipt.changed) { const cells = transaction.operations[0]; if (cells?.op === "setWorkbookCells") {
        const last = cells.cells.at(-1)!, anchor = selected.anchor; void runInteraction(() => choose(anchor.sheetId, { from: anchor.ranges[0]!.from, to: { row: last.row, column: last.column } }, true));
      } }
    },
    onError(error) { report(error); controls(); },
  });
  try {
    renderedSource = createWorkbookPreviewSource(session);
    preview = mountDocumentPreview(ui.preview, renderedSource, { title: "工作簿网格", retainSelection: true, ...(config.widthPt === undefined ? {} : { widthPt: config.widthPt }),
      ...(config.heightPt === undefined ? {} : { heightPt: config.heightPt }), ...(config.locale === undefined ? {} : { locale: config.locale }),
      ...(config.limits === undefined ? {} : { limits: config.limits }), onError: error => report(new DocumentError(error.code, error.message, error.details)), onSelectionChange: sourceSelection });
  } catch (error) { binding.dispose(); input.dispose(); ui.element.remove(); throw error; }
  const unsubscribe = session.subscribe(event => { if (event.type === "documentChanged" || event.type === "annotationChanged") { found = undefined; findSelectionScope = undefined; commandAbort?.abort(); void scheduleRefresh(); } });
  const button = (label: string, action: () => unknown, available: () => boolean, parent: HTMLElement = ui.tools): void => {
    const button = ui.element.ownerDocument.createElement("button"); button.type = "button"; button.textContent = label;
    button.addEventListener("click", () => { try { action(); } catch (error) { report(error); } }, { signal: events.signal }); parent.append(button); buttons.push({ button, available });
  };
  const idle = (): boolean => !commandAbort && !pendingInput() && revision === session.snapshot().manifest.revision;
  const auxiliaryHost: WorkbookControlsHost = { signal: events.signal, snapshot: () => session.snapshot(), draftChanged: controls, prepareOptions: () => preparation(),
    ...(config.locale === undefined ? {} : { locale: config.locale }),
    selectSheet: sheetId => choose(sheetId, { from: { row: 0, column: 0 }, to: { row: 0, column: 0 } }),
    selection: () => { const anchor = input.state().selection?.anchor; return anchor?.kind === "cells" ? { sheetId: anchor.sheetId, range: anchor.ranges[0]! } : undefined; },
    idle: () => idle() && revision >= 0, run: action => { void runInteraction(action); }, perform: async (operations, options, expected) => {
      const current = session.snapshot().manifest; if (expected && (current.documentId !== expected.documentId || current.revision !== expected.revision)) throw new DocumentError("RevisionConflict", "辅助表单属于旧修订。");
      const receipt = perform(operations, options); await sourceSettled(); return receipt; } };
  drawingControls = createWorkbookDrawingControls(ui.element.ownerDocument, ui.preview.parentNode!, auxiliaryHost);
  commentControls = createWorkbookCommentControls(ui.element.ownerDocument, ui.preview.parentNode!, auxiliaryHost);
  nameControls = createWorkbookNameControls(ui.element.ownerDocument, ui.preview.parentNode!, auxiliaryHost);
  worksheetControls = createWorkbookWorksheetControls(ui.element.ownerDocument, ui.preview.parentNode!, auxiliaryHost);
  hyperlinkControls = createWorkbookHyperlinkControls(ui.element.ownerDocument, ui.preview.parentNode!, auxiliaryHost);
  tableControls = createWorkbookTableControls(ui.element.ownerDocument, ui.preview.parentNode!, auxiliaryHost);
  filterControls = createWorkbookFilterControls(ui.element.ownerDocument, ui.preview.parentNode!, auxiliaryHost);
  rulesControls = createWorkbookRulesControls(ui.element.ownerDocument, ui.preview.parentNode!, auxiliaryHost);
  button("确认输入", () => binding.apply(), () => !commandAbort && Boolean(input.state().selection) && !input.state().composing);
  button("取消输入", cancelInput, () => !commandAbort && Boolean(input.state().selection));
  button("撤销", () => history("undo"), () => idle() && session.history().undoSteps > 0);
  button("重做", () => history("redo"), () => idle() && session.history().redoSteps > 0);
  button("重新计算", () => { void scheduleRefresh(true).catch(() => {}); }, idle);
  button("向下填充", () => { void runInteraction(() => fill("row")); }, idle);
  button("向右填充", () => { void runInteraction(() => fill("column")); }, idle);
  button("数列填充", () => { void runInteraction(() => { const anchor = input.state().selection?.anchor; return fill(anchor?.kind === "cells" && anchor.ranges[0]!.from.column === anchor.ranges[0]!.to.column ? "row" : "column", true); }); }, idle);
  button("合并单元格", () => { void runInteraction(() => changeMerge()); }, idle);
  button("合并并清除覆盖内容", () => { void runInteraction(() => changeMerge(false, true)); }, idle);
  button("取消合并", () => { void runInteraction(() => changeMerge(true)); }, idle);
  button("冻结选区之前", () => { void runInteraction(() => changeFreeze()); }, idle);
  button("取消冻结", () => { void runInteraction(() => changeFreeze(true)); }, idle);
  button("插入选区行", () => { void runInteraction(() => changeAxis("row", "insert")); }, idle);
  button("删除选区行", () => { void runInteraction(() => changeAxis("row", "delete")); }, idle);
  button("插入选区列", () => { void runInteraction(() => changeAxis("column", "insert")); }, idle);
  button("删除选区列", () => { void runInteraction(() => changeAxis("column", "delete")); }, idle);
  button("升序排序", () => { void runInteraction(() => sortSelection(false)); }, idle);
  button("降序排序", () => { void runInteraction(() => sortSelection(true)); }, idle);
  button("复制选区", () => copySelection("copy"), idle);
  button("剪切选区", () => copySelection("cut"), idle);
  button("粘贴选区", () => { void runInteraction(pasteSelection); }, () => idle() && clipboard.available);
  button("应用单元格格式", () => { void runInteraction(formatSelection); }, idle, ui.format.tools);
  button("应用选中文字字体", () => { void runInteraction(formatSelectedText); }, () => idle() && !!input.state().selection && ui.input.selectionStart < ui.input.selectionEnd, ui.format.tools);
  button("取消操作", () => commandAbort?.abort(), () => Boolean(commandAbort));
  button("查找下一个", () => { void runInteraction(() => search()); }, () => idle() && Boolean(ui.find.value), ui.searchTools);
  button("替换当前项", () => { void runInteraction(() => search("one")); }, () => idle() && ui.findField.value !== "display" && Boolean(found && found.index >= 0), ui.searchTools);
  button("全部替换", () => { void runInteraction(() => search("all")); }, () => idle() && ui.findField.value !== "display" && Boolean(ui.find.value), ui.searchTools);
  const searchSettings = (): string => JSON.stringify([ui.find.value, ui.findField.value, ui.findScope.value, ui.matchCase.checked, ui.entireCell.checked]);
  let lastSearchSettings = searchSettings();
  for (const field of [ui.find, ui.findField, ui.findScope, ui.matchCase, ui.entireCell]) for (const event of ["input", "change"]) field.addEventListener(event, () => {
    const current = searchSettings(); if (current !== lastSearchSettings) { found = undefined; findSelectionScope = undefined; lastSearchSettings = current; } controls();
  }, { signal: events.signal });
  ui.find.addEventListener("keydown", event => { if (event.key === "Enter") { event.preventDefault(); void runInteraction(() => search()); } }, { signal: events.signal });
  ui.element.addEventListener("keydown", event => {
    if (!(event.ctrlKey || event.metaKey) || event.altKey || event.isComposing || !["f", "h"].includes(event.key.toLowerCase())) return;
    event.preventDefault(); if (!commandAbort) { ui.find.focus(); ui.find.select(); }
  }, { signal: events.signal });
  for (const type of ["input", "keydown", "keyup", "select", "mouseup", "compositionstart", "compositionend"]) ui.input.addEventListener(type, controls, { signal: events.signal });
  ui.mode.addEventListener("change", () => { try { input.setMode(ui.mode.value as WorkbookTextInputMode); binding.refresh(); controls(); } catch (error) { report(error); } }, { signal: events.signal });
  ui.address.addEventListener("keydown", event => {
    if (event.key !== "Enter") return; event.preventDefault();
    try {
      const sheetId = preview.layout().sheetId, resolved = session.resolveReference("=" + ui.address.value.trim(), { sheetId, row: 0, column: 0 });
      if (resolved.kind !== "ranges" || resolved.areas.length !== 1) throw new DocumentError("InvalidOperation", "请输入一个有效的单元格或矩形区域地址。");
      const area = resolved.areas[0]!; void runInteraction(() => choose(area.sheetId, { from: area.from, to: area.to }, true));
    } catch (error) { report(error); }
  }, { signal: events.signal });
  controls(); void scheduleRefresh(); const ready = sourceSettled().then(() => {}); void ready.catch(() => {});
  return Object.freeze({ element: ui.element, ready, settled, selection() { check(); return input.state().selection; }, inputState() { check(); return input.state(); }, hasPendingInput() { check(); return pendingInput() || Boolean(commandAbort) || pendingForms(); },
    selectRange: (sheetId: string, range: DeepReadonly<WorkbookRange>) => runInteraction(() => choose(sheetId, range)), focusInput(selectAll = true) { check(); binding.focus(selectAll); }, applyInput: binding.apply, cancelInput, perform,
    recalculate: async () => { commandAbort?.abort(); await scheduleRefresh(true); return settled(); },
    dispose() { if (disposed) return; disposed = true; ++generation; calculationAbort?.abort(); commandAbort?.abort(); unsubscribe(); events.abort(); binding.dispose(); input.dispose(); preview.dispose(); renderedSource = undefined; ui.element.remove(); },
  });
}
