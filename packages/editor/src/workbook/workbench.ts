import { createWorkbookCalculator, DocumentError, parseContract } from "@doffi/core";
import type { DeepReadonly, TransactionPrepareOptions, TransactionReceipt, WorkbookEditProposal, WorkbookInteractionRequest, WorkbookInteractionRequestInput, WorkbookInteractionScope, WorkbookSession, WorkbookSnapshot } from "@doffi/core";
import { createWorkbookPreviewSource, mountDocumentPreview } from "@doffi/render";
import type { WorkbookPreview, WorkbookPreviewSource, WorkbookReadConnection } from "@doffi/render";
import { mountWorkbookEditor } from "./editor.js";
import type { WorkbookEditorOptions, WorkbookEditorView } from "./editor-contracts.js";
import type { WorkbookRemoteEditorView } from "./editor-remote.js";
import { checkSessionMessage, freezeSessionValue, sessionTransportConfig } from "./session-transport.js";
import type { WorkbookSessionConnection, WorkbookSessionReceipt } from "./session-transport.js";
import { createWorkbookTaskQueue } from "./task-queue.js";

export type { WorkbookWorkbenchCandidate, WorkbookWorkbenchEvent, WorkbookWorkbenchOptions, WorkbookWorkbenchView } from "./workbench-contracts.js";
import type { WorkbookWorkbenchCandidate, WorkbookWorkbenchEvent, WorkbookWorkbenchOptions, WorkbookWorkbenchView } from "./workbench-contracts.js";
import { renderWorkbookProposalDiff } from "./workbench-diff.js";

type Job = { readonly abort: AbortController; readonly ticket: number; guard(): void };
const equal = (a: unknown, b: unknown) => JSON.stringify(a) === JSON.stringify(b);

/** Host callbacks own model execution. This view owns only drafts, immutable
 * read leases and the same isolated core candidate used by ordinary edits. */
export function mountWorkbookWorkbench(container: HTMLElement, session: WorkbookSession | WorkbookSessionConnection, options: WorkbookWorkbenchOptions = {}): WorkbookWorkbenchView {
  const document = container?.ownerDocument, window = document?.defaultView;
  if (!window || !container.isConnected) throw new DocumentError("InvalidOperation", "Workbook workbench requires a connected container.");
  if (!options || typeof options !== "object" || Array.isArray(options) || Object.keys(options).some(key => !["title", "widthPt", "heightPt", "locale", "limits", "prepareOptions", "onError", "onSelectionChange", "onPendingChange", "scope", "allowGlobalScope", "readRanges", "supportedOperations", "contextLimits", "onSubmitAnnotations", "onEvent"].includes(key))) throw new DocumentError("InvalidOperation", "Invalid workbook workbench options.");
  const config = { ...options }, messageLimits = sessionTransportConfig({ maxMessageBytes: Math.min(Number.MAX_SAFE_INTEGER, (options.contextLimits?.maxBytes ?? 4 * 1024 * 1024) * 8) }), element = document.createElement("div"), shadow = element.attachShadow({ mode: "open" }), events = new window.AbortController(), tasks = createWorkbookTaskQueue();
  const style = document.createElement("style"); style.textContent = ":host{display:block;color:#172638;font:14px/1.5 system-ui}section{border:1px solid #ccd6df;padding:12px;margin:10px 0}button,input,select{font:inherit}button,select{padding:5px 8px}button{cursor:pointer}button:disabled{cursor:default;opacity:.5}label{display:inline-flex;gap:6px;margin:5px 8px 5px 0}select[multiple]{min-width:160px;max-height:100px}.actions{display:flex;gap:7px;flex-wrap:wrap}pre{white-space:pre-wrap;overflow-wrap:anywhere;max-height:300px;overflow:auto}table{border-collapse:collapse;width:100%}th,td{padding:5px;border:1px solid #d8e1e9;text-align:left;vertical-align:top}td pre{max-width:320px;margin:0}.error{color:#a12626}.candidate{position:relative}[hidden]{display:none!important}";
  const editorHost = document.createElement("div"), panel = document.createElement("section"), choices = document.createElement("div"), actions = document.createElement("div"), status = document.createElement("p"), error = document.createElement("p"), diff = document.createElement("section"), candidateHost = document.createElement("div");
  panel.setAttribute("aria-label", "工作簿请求与提案"); choices.setAttribute("aria-label", "选择提交的工作簿批注"); actions.className = "actions"; status.setAttribute("role", "status"); status.setAttribute("aria-label", "工作簿协作状态"); error.setAttribute("role", "alert"); error.className = "error"; error.hidden = true; diff.setAttribute("aria-label", "工作簿提案差异"); diff.hidden = true; candidateHost.className = "candidate";
  const heading = document.createElement("h2"); heading.textContent = "工作簿协作"; panel.append(heading, choices);
  function field<T extends HTMLElement>(title: string, control: T): T { const label = document.createElement("label"); control.setAttribute("aria-label", title); label.append(title, control); panel.append(label); return control; }
  const scopeField = field("请求写入范围", document.createElement("select"));
  for (const [value, title] of [["anchors", "批注锚定区域"], ["selection", "当前选区"], ["tables", "指定智能表"], ...config.allowGlobalScope ? [["global", "整个工作簿"]] : []]) { const option = document.createElement("option"); option.value = value!; option.textContent = title!; scopeField.append(option); }
  scopeField.hidden = config.scope !== undefined;
  const tables = field("允许修改的智能表", document.createElement("select")), objects = field("允许修改的绘图", document.createElement("select")); tables.multiple = true; objects.multiple = true;
  const createObjects = field("允许在当前表新增绘图", document.createElement("input")); createObjects.type = "checkbox";
  const readInputs = field("额外只读输入区域", document.createElement("input")); readInputs.placeholder = "例如 A20 或 Sheet2!A1:B10"; readInputs.maxLength = 8191;
  if (config.scope) for (const control of [scopeField, tables, objects, createObjects]) control.parentElement!.hidden = true;
  panel.append(actions, status, error); shadow.append(style, editorHost, panel, diff); container.append(element);
  let disposed = false, sequence = 0, busy: Job | undefined, applying = false, choiceEvents = new window.AbortController();
  let diffEvents = new window.AbortController(), captionHtml = "", captionNodes = new Map<string, string>();
  let editor!: WorkbookEditorView | WorkbookRemoteEditorView;
  let sourceSnapshot: WorkbookSnapshot | undefined = "kind" in session ? session.snapshot() : undefined;
  let activeRequest: WorkbookInteractionRequest | undefined, activeResources: TransactionPrepareOptions | undefined;
  let candidate: WorkbookWorkbenchCandidate | undefined, candidatePreview: WorkbookPreview | undefined, candidateLease: WorkbookReadConnection | undefined;
  let candidateSource: WorkbookPreviewSource | undefined, candidateBefore: WorkbookSnapshot | undefined;
  const selected = new Map<string, boolean>();
  function state() { const manifest = "kind" in session ? session.snapshot().manifest : session.state(); return { documentId: manifest.documentId, revision: manifest.revision }; }
  function check(): void { if (disposed) throw new DocumentError("InvalidOperation", "This workbook workbench has been disposed."); }
  function pendingInput(): boolean { const value = editor.inputState(); return editor.hasPendingInput() || "applying" in value && value.applying === true; }
  function notify(value: WorkbookWorkbenchEvent): void { if (!disposed) try { config.onEvent?.(Object.freeze(value)); } catch { /* Host reporting never writes source. */ } }
  function report(reason: unknown): void { if (disposed) return; const problem = reason instanceof Error ? reason : new Error("工作簿协作失败。"); error.textContent = problem.message; error.hidden = false; try { config.onError?.(problem); } catch { /* Reporting only. */ } }
  function buttons(): void {
    if (disposed) return; const available = [...selected.values()].some(Boolean), current = state();
    generate.disabled = !available || !!busy || !!candidate || pendingInput(); submit.disabled = generate.disabled || !config.onSubmitAnnotations;
    download.disabled = !activeRequest; file.disabled = !activeRequest || !!busy || !!candidate || pendingInput(); cancel.disabled = !busy || applying;
    apply.disabled = !candidate || !!busy || candidate.baseRevision !== current.revision; reject.disabled = !candidate || !!busy;
    for (const control of [scopeField, tables, objects, createObjects]) control.disabled = !!busy || !!candidate || pendingInput();
    tables.disabled ||= scopeField.value !== "tables"; objects.disabled ||= scopeField.value === "global"; createObjects.disabled ||= scopeField.value === "global";
    readInputs.disabled = !!busy || !!candidate || pendingInput();
    for (const checkbox of choices.querySelectorAll("input")) checkbox.disabled = !!busy || !!candidate || pendingInput();
    editorHost.inert = !!busy || !!candidate; editorHost.setAttribute("aria-busy", String(!!busy));
  }
  async function run<Value>(action: (job: Job) => Promise<Value>): Promise<Value> {
    check(); if (busy || pendingInput()) throw new DocumentError("InvalidOperation", "请先完成或取消当前输入／操作。");
    const abort = new window!.AbortController(), ticket = ++sequence, expected = state(), job: Job = { abort, ticket, guard() { if (disposed || abort.signal.aborted || ticket !== sequence) throw new DocumentError("OperationCancelled", "工作簿协作操作已取消。"); if (!equal(state(), expected)) throw new DocumentError("RevisionConflict", "工作簿已变化，请重新生成请求。"); } };
    busy = job; error.hidden = true; buttons();
    try { return await action(job); } finally { if (!disposed && busy === job) { busy = undefined; buttons(); } }
  }
  function action(title: string, work: () => Promise<unknown> | void): HTMLButtonElement { const button = document.createElement("button"); button.type = "button"; button.textContent = title; actions.append(button); button.addEventListener("click", () => { try { void Promise.resolve(work()).catch(report); } catch (reason) { report(reason); } }, { signal: events.signal }); return button; }
  const editorOptions: WorkbookEditorOptions = { ...(config.title !== undefined ? { title: config.title } : {}), ...(config.widthPt !== undefined ? { widthPt: config.widthPt } : {}), ...(config.heightPt !== undefined ? { heightPt: config.heightPt } : {}), ...(config.locale !== undefined ? { locale: config.locale } : {}), ...(config.limits ? { limits: config.limits } : {}), ...(config.prepareOptions ? { prepareOptions: config.prepareOptions } : {}), onError: report,
    onPendingChange: pending => { try { config.onPendingChange?.(pending); } catch { /* Host reporting only. */ } if (editor) buttons(); },
    onSelectionChange: selection => { try { config.onSelectionChange?.(selection); } catch { /* Host reporting only. */ } notify({ type: "selectionChanged", selection }); if (editor) buttons(); } };
  editor = "kind" in session ? mountWorkbookEditor(editorHost, session, editorOptions) : mountWorkbookEditor(editorHost, session, { ...editorOptions, onViewChange(snapshot) { sourceSnapshot = snapshot; if (editor) update(); } });
  const generate = action("生成工作簿请求", () => buildRequest()), submit = action("提交工作簿批注", () => submitAnnotations());
  const download = action("下载工作簿请求", () => {
    check(); if (!activeRequest) throw new DocumentError("InvalidOperation", "请先生成请求。");
    const url = window!.URL.createObjectURL(new Blob([JSON.stringify(activeRequest, null, 2)], { type: "application/json" })), anchor = document.createElement("a"); anchor.href = url; anchor.download = "workbook.request.json"; anchor.click(); setTimeout(() => window!.URL.revokeObjectURL(url), 0);
  });
  const file = document.createElement("input"); file.type = "file"; file.accept = ".json,application/json"; file.setAttribute("aria-label", "导入工作簿提案 JSON"); actions.append(file);
  const cancel = action("取消协作操作", () => { if (!applying) busy?.abort.abort(); }), apply = action("整批应用工作簿提案", () => applyProposal()), reject = action("放弃工作簿提案", () => rejectProposal());
  file.addEventListener("change", () => { const selectedFile = file.files?.[0], request = activeRequest; file.value = ""; if (!selectedFile || !request) return;
    void run(async job => { if (selectedFile.size > (config.contextLimits?.maxBytes ?? 4 * 1024 * 1024)) throw new DocumentError("ResourceLimitExceeded", "工作簿提案文件超过大小限制。"); const text = await selectedFile.text(); job.guard(); const parsed: unknown = JSON.parse(text); checkSessionMessage(parsed, messageLimits); const proposal = parseContract("workbook-edit-proposal", parsed); await load(request, proposal, activeResources, job); }).catch(report);
  }, { signal: events.signal });
  for (const field of [scopeField, tables, objects, createObjects]) field.addEventListener("change", buttons, { signal: events.signal });
  function caption(annotation: WorkbookSnapshot["annotations"][number]): string {
    if (annotation.instruction !== undefined) return annotation.instruction;
    const html = sourceSnapshot?.content?.html ?? "";
    if (captionHtml !== html) { captionHtml = html; const template = document.createElement("template"); template.innerHTML = html; captionNodes = new Map([...template.content.querySelectorAll("[data-node-id]")].map(node => [node.getAttribute("data-node-id")!, node.textContent ?? ""])); }
    return captionNodes.get(annotation.bodyNodeId!) ?? "工作簿批注";
  }
  function update(): void {
    if (disposed || !sourceSnapshot) return;
    if ("kind" in session) sourceSnapshot = session.snapshot();
    choiceEvents.abort(); choiceEvents = new window!.AbortController(); choices.replaceChildren();
    const annotations = sourceSnapshot.annotations.filter(annotation => annotation.status === "open" && annotation.baseRevision === sourceSnapshot!.manifest.revision);
    for (const id of selected.keys()) if (!annotations.some(annotation => annotation.annotationId === id)) selected.delete(id);
    for (const annotation of annotations) { if (!selected.has(annotation.annotationId)) selected.set(annotation.annotationId, true); const label = document.createElement("label"), checkbox = document.createElement("input"), text = caption(annotation); checkbox.type = "checkbox"; checkbox.checked = selected.get(annotation.annotationId)!; checkbox.value = annotation.annotationId; checkbox.setAttribute("aria-label", "提交工作簿批注：" + text); checkbox.addEventListener("change", () => { selected.set(annotation.annotationId, checkbox.checked); buttons(); }, { signal: choiceEvents.signal }); label.append(checkbox, text); choices.append(label); }
    function list(control: HTMLSelectElement, entries: readonly (readonly [string, string])[]) { const retained = new Set([...control.selectedOptions].map(option => option.value)); control.replaceChildren(); for (const [id, title] of entries) { const option = document.createElement("option"); option.value = id; option.textContent = title; option.selected = retained.has(id); control.append(option); } }
    list(tables, sourceSnapshot.metadata.workbook.sheets.flatMap(sheet => (sheet.tables ?? []).map(table => [table.tableId, sheet.name + " · " + table.name] as const)));
    list(objects, (sourceSnapshot.metadata.workbook.objects ?? []).map((object, index) => [object.nodeId, object.name ?? ({ text: "文本框", geometry: "形状", image: "图片", chart: "图表" }[object.kind] + " " + (index + 1))] as const));
    if (activeRequest && activeRequest.baseRevision !== state().revision) status.textContent = "请求已过期，请基于当前修订重新生成。"; buttons();
  }
  function chosenScope(ids: readonly string[], override?: DeepReadonly<WorkbookInteractionScope>): DeepReadonly<WorkbookInteractionScope> {
    if (override ?? config.scope) return (override ?? config.scope)!;
    if (scopeField.value === "global") { if (!config.allowGlobalScope) throw new DocumentError("InvalidOperation", "宿主未启用全局请求。"); return { kind: "global" }; }
    const objectIds = [...objects.selectedOptions].map(option => option.value), selection = editor.selection(), anchor = selection?.anchor;
    const newObjectSheetIds = createObjects.checked ? [anchor?.kind === "cells" ? anchor.sheetId : sourceSnapshot!.metadata.workbook.activeSheetId] : [];
    if (scopeField.value === "tables") return { kind: "tables", tableIds: [...tables.selectedOptions].map(option => option.value), objectIds, newObjectSheetIds };
    const ranges = scopeField.value === "selection" ? anchor?.kind === "cells" ? anchor.ranges.map(range => ({ sheetId: anchor.sheetId, range })) : [] : sourceSnapshot!.annotations.filter(annotation => ids.includes(annotation.annotationId)).flatMap(annotation => annotation.anchor.kind === "cells" ? annotation.anchor.ranges.map(range => ({ sheetId: annotation.anchor.kind === "cells" ? annotation.anchor.sheetId : "", range })) : []);
    return { kind: "ranges", ranges, objectIds, newObjectSheetIds };
  }
  async function build(ids: readonly string[] | undefined, scope: DeepReadonly<WorkbookInteractionScope> | undefined, job: Job): Promise<WorkbookInteractionRequest> {
    if (candidate) throw new DocumentError("InvalidOperation", "请先应用或放弃当前提案。"); await editor.settled(); job.guard();
    if ("kind" in session) sourceSnapshot = session.snapshot();
    const annotationIds = ids ?? [...selected].filter(([, value]) => value).map(([id]) => id), resources = config.prepareOptions?.(), calculationContext = { ...resources?.calculationContext, ...(resources?.calculationContext?.locale === undefined && config.locale ? { locale: config.locale } : {}) };
    const readRanges = [...config.readRanges ?? []], reference = readInputs.value.trim();
    if (reference) {
      const anchor = editor.selection()?.anchor, origin = { sheetId: anchor?.kind === "cells" ? anchor.sheetId : sourceSnapshot!.metadata.workbook.activeSheetId, row: 0, column: 0 };
      const bound = await session.resolveReference(reference.startsWith("=") ? reference : "=" + reference, origin); job.guard();
      if (bound.kind !== "ranges") throw new DocumentError("InvalidOperation", "请输入可解析的单元格或区域作为只读输入。");
      readRanges.push(...bound.areas.map(area => ({ sheetId: area.sheetId, range: { from: { ...area.from }, to: { ...area.to } } })));
    }
    const input: WorkbookInteractionRequestInput = { kind: "workbook", annotationIds, scope: chosenScope(annotationIds, scope), ...(readRanges.length ? { readRanges } : {}), ...(config.supportedOperations ? { supportedOperations: config.supportedOperations } : {}), ...(config.contextLimits ? { limits: config.contextLimits } : {}), ...(resources?.assets ? { assets: resources.assets.map(resource => resource.descriptor) } : {}), ...(calculationContext ? { calculationContext } : {}) };
    const progress = (value: { readonly completed: number }) => { if (!disposed && busy === job) status.textContent = "正在准备工作簿上下文 · " + value.completed; };
    const request = "kind" in session ? await session.buildInteractionRequestAsync(input, { signal: job.abort.signal, yield: tasks.yield, onProgress: progress }) : await session.buildInteractionRequest(input, { signal: job.abort.signal, onProgress: progress }); job.guard();
    activeRequest = request; activeResources = { ...resources, calculationContext: { ...(request.context.calculationContext.now ? { now: request.context.calculationContext.now } : {}), ...(request.context.calculationContext.locale !== undefined ? { locale: request.context.calculationContext.locale } : {}) } }; status.textContent = "请求已生成 · 修订 " + request.baseRevision + " · " + request.context.cells.length + " 个来源单元格 · " + request.context.dependencies.length + " 项只读依赖"; buttons(); return request;
  }
  const buildRequest = (ids?: readonly string[], scope?: DeepReadonly<WorkbookInteractionScope>) => run(job => build(ids, scope, job));
  async function submitAnnotations(ids?: readonly string[]): Promise<void> {
    return run(async job => {
      if (!config.onSubmitAnnotations) throw new DocumentError("HostSubmissionFailed", "宿主尚未配置工作簿提交入口。");
      const request = await build(ids, undefined, job); status.textContent = "请求已交给宿主，等待提案。";
      const response = await new Promise<void | WorkbookEditProposal>((resolve, reject) => {
        const stop = () => reject(new DocumentError("OperationCancelled", "工作簿提交已取消。")); job.abort.signal.addEventListener("abort", stop, { once: true });
        Promise.resolve().then(() => config.onSubmitAnnotations!(request, { signal: job.abort.signal })).then(resolve, () => reject(new DocumentError("HostSubmissionFailed", "宿主未能处理工作簿请求。"))).finally(() => job.abort.signal.removeEventListener("abort", stop));
      }); job.guard(); notify({ type: "annotationsSubmitted", request });
      if (response) await load(request, response, activeResources, job);
    });
  }
  async function discard(change: WorkbookWorkbenchCandidate): Promise<void> { try { if ("kind" in session && "model" in change) session.discard(change); else if (!("kind" in session) && !("model" in change)) await session.discard(change); } catch { /* A host revision/disposal may already have released it. */ } }
  function clearPreview(): void { diffEvents.abort(); if (!disposed) diffEvents = new window!.AbortController(); candidatePreview?.dispose(); candidatePreview = undefined; candidateLease?.dispose(); candidateLease = undefined; candidateSource = undefined; candidateBefore = undefined; candidateHost.replaceChildren(); diff.replaceChildren(); diff.hidden = true; }
  async function load(request: WorkbookInteractionRequest, proposal: WorkbookEditProposal, resources: TransactionPrepareOptions | undefined, job: Job): Promise<WorkbookWorkbenchCandidate> {
    checkSessionMessage({ request, proposal }, messageLimits); request = freezeSessionValue(parseContract("workbook-interaction-request", request)); proposal = freezeSessionValue(parseContract("workbook-edit-proposal", proposal));
    if (candidate) throw new DocumentError("InvalidOperation", "请先应用或放弃当前提案。"); await editor.settled(); job.guard();
    const base = state(); if (request.documentId !== base.documentId || request.baseRevision !== base.revision) throw new DocumentError("RevisionConflict", "提案请求属于旧修订，请重新生成。");
    let owned: WorkbookWorkbenchCandidate | undefined, lease: WorkbookReadConnection | undefined;
    try {
      candidateBefore = "kind" in session ? session.snapshot() : sourceSnapshot;
      owned = "kind" in session ? await session.prepareProposalAsync(request, proposal, resources, { signal: job.abort.signal, yield: tasks.yield }) : await session.prepare({ kind: "proposal", request, proposal }, resources, { signal: job.abort.signal }); job.guard();
      const prepared = owned, review = "kind" in session && "model" in prepared ? session.getProposalReview(prepared) : "proposalReview" in prepared ? prepared.proposalReview : undefined;
      if (!review) throw new DocumentError("InvalidProposal", "工作簿候选缺少核对结果。");
      if ("kind" in session && "model" in prepared) { const calculator = createWorkbookCalculator(prepared.model); try { const frozen = request.context.calculationContext, calculation = await calculator.calculate({ signal: job.abort.signal, yield: async () => { job.guard(); await tasks.yield(); job.guard(); }, context: { ...(frozen.now ? { now: frozen.now } : {}), ...(frozen.locale !== undefined ? { locale: frozen.locale } : {}) } }); candidateSource = createWorkbookPreviewSource(prepared.model, calculation); } finally { calculator.dispose(); } }
      else if (!("kind" in session) && !("model" in prepared)) { lease = await session.view(prepared, { signal: job.abort.signal, calculate: true }); candidateLease = lease; candidateSource = lease; }
      job.guard(); if (!candidateSource || !candidateBefore || !equal(state(), base)) throw new DocumentError("RevisionConflict", "候选准备期间工作簿已变化。");
      diff.replaceChildren(candidateHost); diff.hidden = false;
      candidatePreview = mountDocumentPreview(candidateHost, candidateSource, { title: "工作簿候选预览", ...(config.widthPt !== undefined ? { widthPt: config.widthPt } : {}), ...(config.heightPt !== undefined ? { heightPt: config.heightPt } : {}), ...(config.locale !== undefined ? { locale: config.locale } : {}), ...(config.limits ? { limits: config.limits } : {}) });
      await candidatePreview.ready; job.guard(); await renderWorkbookProposalDiff(diff, { candidate: prepared, before: candidateBefore, source: candidateSource, request, proposal, review, candidateHost, signal: diffEvents.signal, onError: report }); job.guard(); candidate = prepared; activeRequest = request;
      status.textContent = "提案已准备；核对来源和候选预览后可整批应用。"; buttons(); notify({ type: "proposalReady", proposal, candidate: prepared, review }); return prepared;
    } catch (reason) { clearPreview(); lease?.dispose(); if (owned) await discard(owned); throw reason; }
  }
  const loadProposal = (request: WorkbookInteractionRequest, proposal: WorkbookEditProposal, resources?: TransactionPrepareOptions) => run(job => load(request, proposal, resources ?? (activeRequest?.requestId === request.requestId ? activeResources : undefined), job));
  async function applyProposal(): Promise<TransactionReceipt | WorkbookSessionReceipt> {
    return run(async job => { const owned = candidate; if (!owned) throw new DocumentError("InvalidOperation", "当前没有候选提案。"); if (owned.baseRevision !== state().revision) throw new DocumentError("RevisionConflict", "候选已过期，请重新生成请求。"); job.guard(); applying = true;
      buttons();
      try {
        const receipt = "kind" in session && "model" in owned ? session.commit(owned) : !("kind" in session) && !("model" in owned) ? await session.commit(owned) : (() => { throw new DocumentError("InvalidOperation", "候选不属于当前宿主。"); })();
        candidate = undefined;
        // Once the atomic boundary returns, disposal/rendering failure cannot
        // turn an authoritative committed receipt into a failed edit.
        if (!disposed) { clearPreview(); await editor.settled().catch(report); if (!disposed) { update(); status.textContent = receipt.changed ? "提案已整批应用，可通过编辑器撤销。" : "提案没有净修改，批注保持原状态。"; } }
        return receipt;
      } finally { applying = false; }
    });
  }
  const rejectProposal = () => run(async () => { const owned = candidate; candidate = undefined; clearPreview(); if (owned) await discard(owned); status.textContent = "已放弃提案，工作簿来源保持原状。"; buttons(); });
  const unsubscribe = session.subscribe(() => { if (disposed) return; if ("kind" in session) sourceSnapshot = session.snapshot(); if (!applying && candidate && candidate.baseRevision !== state().revision) { const owned = candidate; candidate = undefined; clearPreview(); void discard(owned); status.textContent = "工作簿已变化，旧候选已释放，请重新生成请求。"; } update(); });
  void editor.ready.then(() => { if (!disposed) update(); }).catch(report); update();
  return { element, editor, ready: editor.ready, settled: editor.settled, request: () => activeRequest, buildRequest, submitAnnotations, loadProposal, applyProposal, rejectProposal, cancel: () => { if (!applying) busy?.abort.abort(); },
    dispose() { if (disposed) return; disposed = true; sequence++; busy?.abort.abort(); unsubscribe(); events.abort(); choiceEvents.abort(); tasks.dispose(); const owned = candidate; candidate = undefined; clearPreview(); if (owned) void discard(owned); editor.dispose(); selected.clear(); captionNodes.clear(); captionHtml = ""; sourceSnapshot = undefined; activeRequest = undefined; activeResources = undefined; element.remove(); },
  };
}
