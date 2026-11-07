import { captureWorkbookModel, DocumentError } from "@doffi/core";
import type { DeepReadonly, DocumentOperation, DocumentSelection, TransactionPrepareOptions, TransactionReceipt, WorkbookRange, WorkbookSession } from "@doffi/core";
import { createWorkbookPreviewSource, mountDocumentPreview } from "@doffi/render";
import type { WorkbookPreview, WorkbookPreviewSource, WorkbookRenderReport } from "@doffi/render";
import type { WorkbookEditorOptions, WorkbookEditorView } from "./editor-contracts.js";
import { createWorkbookEditorUi } from "./editor-ui.js";
import { createWorkbookInputController } from "./input-controller.js";
import { bindWorkbookTextInput } from "./text-binding.js";
import { createWorkbookPasteTransaction } from "./text-input.js";
import type { WorkbookTextInputMode } from "./text-input.js";
import { createWorkbookFillTransaction } from "./fill.js";
import { findWorkbookText } from "./find.js";
import { createWorkbookReplaceTransaction } from "./replace.js";
import type { WorkbookFindMatch, WorkbookFindQuery, WorkbookFindResult, WorkbookReplaceOptions } from "./find-contracts.js";
import { workbookInputTextView } from "./text-projection.js";

/** Direct small-document host integration. The host owns the sole session;
 * source cells, candidates and history remain behind its core boundary. */
export function mountWorkbookEditor(container: HTMLElement, session: WorkbookSession, options: WorkbookEditorOptions = {}): WorkbookEditorView {
  const view = container?.ownerDocument?.defaultView;
  if (!view || !container.isConnected || session.kind !== "workbook") throw new DocumentError("InvalidOperation", "Workbook editing requires a connected container and its owning workbook session.");
  if (!options || typeof options !== "object" || Array.isArray(options) || Object.keys(options).some(key => !["title", "widthPt", "heightPt", "locale", "limits", "prepareOptions", "onError", "onSelectionChange"].includes(key))) throw new DocumentError("InvalidOperation", "Invalid workbook editor options.");
  const config = { ...options }, ui = createWorkbookEditorUi(container, config.title ?? "工作簿编辑"), input = createWorkbookInputController(session), events = new view.AbortController();
  let disposed = false, choosing = 0, refreshQueued = false, forceCalculation = false, latest: Promise<void> = Promise.resolve();
  let interaction: Promise<void> = Promise.resolve();
  let revision = -1, generation = 0, calculationAbort: AbortController | undefined, commandAbort: AbortController | undefined, preview: WorkbookPreview;
  let renderedSource: WorkbookPreviewSource | undefined;
  let found: { readonly key: string; readonly result: WorkbookFindResult; index: number } | undefined;
  let findSelectionScope: WorkbookFindQuery["scope"];
  const searchFields = [ui.find, ui.replacement, ui.findField, ui.findScope, ui.matchCase, ui.entireCell, ui.replacementMode];
  const buttons: { readonly button: HTMLButtonElement; readonly available: () => boolean }[] = [];
  function check(): void { if (disposed) throw new DocumentError("InvalidOperation", "This workbook editor has been disposed."); }
  function preparation(input?: TransactionPrepareOptions): TransactionPrepareOptions {
    const supplied = input ?? config.prepareOptions?.() ?? {};
    return config.locale !== undefined && supplied.calculationContext?.locale === undefined ? { ...supplied, calculationContext: { ...supplied.calculationContext, locale: config.locale } } : supplied;
  }
  function runInteraction<T>(action: () => Promise<T>): Promise<T> { const work = interaction.then(action); interaction = work.then(() => {}, report); return work; }
  function pendingInput(): boolean { return binding.hasPendingInput(); }
  function report(error: unknown): void {
    if (disposed) return; const problem = error instanceof Error ? error : new Error("工作簿编辑失败。");
    ui.error.textContent = problem instanceof DocumentError && problem.code === "RevisionConflict" ? "文档已变化。当前输入已保留，请先复制输入或取消后重新选择。" : problem.message; ui.error.hidden = false;
    try { config.onError?.(problem); } catch { /* Host callback does not alter the edit. */ }
  }
  function controls(): void {
    if (disposed) return;
    for (const { button, available } of buttons) button.disabled = !available();
    const state = input.state(); ui.mode.value = state.mode; ui.mode.disabled = Boolean(commandAbort) || state.composing || !state.selection;
    ui.address.disabled = Boolean(commandAbort) || revision < 0;
    for (const field of searchFields) field.disabled = Boolean(commandAbort);
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
    const receipt = session.commit(session.prepare({ protocolVersion: 1, documentId: manifest.documentId, baseRevision: manifest.revision, operations: [...operations] }, preparation(preparedOptions)));
    readCurrentSelection(); controls(); ui.error.hidden = true; return receipt;
  }
  function cancelInput(): void { check(); input.cancel(); binding.refresh(); readCurrentSelection(); controls(); ui.error.hidden = true; }
  function history(direction: "undo" | "redo"): void {
    check(); if (pendingInput()) throw new DocumentError("InvalidOperation", "请先确认或取消当前单元格输入。");
    session[direction](); readCurrentSelection(); controls(); ui.error.hidden = true;
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
    let seeds = 1;
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
    const candidate = await session.prepareAsync(transaction, preparation(), { signal: abort.signal,
      yield: () => new Promise<void>(resolve => setTimeout(resolve, 0)), onProgress: progress => {
        if (!disposed) ui.status.textContent = `正在${progress.phase === "calculation" ? "计算" : progress.phase === "validation" ? "检查更改" : "准备更改"} · ${progress.completed}${progress.total ? ` / ${progress.total}` : ""}`;
      } });
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
  const binding = bindWorkbookTextInput(ui.input, input, {
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
      const receipt = session.commit(session.prepare({ ...transaction, documentId: selected.documentId, baseRevision: selected.baseRevision }, preparation()));
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
  const idle = (): boolean => !commandAbort && !pendingInput();
  button("确认输入", () => binding.apply(), () => !commandAbort && Boolean(input.state().selection) && !input.state().composing);
  button("取消输入", cancelInput, () => !commandAbort && Boolean(input.state().selection));
  button("撤销", () => history("undo"), () => idle() && session.history().undoSteps > 0);
  button("重做", () => history("redo"), () => idle() && session.history().redoSteps > 0);
  button("重新计算", () => { void scheduleRefresh(true).catch(() => {}); }, idle);
  button("向下填充", () => { void runInteraction(() => fill("row")); }, idle);
  button("向右填充", () => { void runInteraction(() => fill("column")); }, idle);
  button("数列填充", () => { void runInteraction(() => { const anchor = input.state().selection?.anchor; return fill(anchor?.kind === "cells" && anchor.ranges[0]!.from.column === anchor.ranges[0]!.to.column ? "row" : "column", true); }); }, idle);
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
  for (const type of ["input", "keydown", "compositionstart", "compositionend"]) ui.input.addEventListener(type, controls, { signal: events.signal });
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
  return Object.freeze({ element: ui.element, ready, settled, selection() { check(); return input.state().selection; }, inputState() { check(); return input.state(); },
    selectRange: (sheetId: string, range: DeepReadonly<WorkbookRange>) => runInteraction(() => choose(sheetId, range)), focusInput(selectAll = true) { check(); binding.focus(selectAll); }, applyInput: binding.apply, cancelInput, perform,
    recalculate: async () => { commandAbort?.abort(); await scheduleRefresh(true); return settled(); },
    dispose() { if (disposed) return; disposed = true; ++generation; calculationAbort?.abort(); commandAbort?.abort(); unsubscribe(); events.abort(); binding.dispose(); input.dispose(); preview.dispose(); renderedSource = undefined; ui.element.remove(); },
  });
}
