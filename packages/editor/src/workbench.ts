import { DocumentError, parseContract } from "@doffi/core";
import type {
  AnnotationSubmissionHandler, CandidateChange, DocumentInteractionRequest, ContentDocumentModel, DocumentSelection,
  DocumentSession, EditProposal, InteractionRequestInput, TransactionReceipt,
  WorkbookSession,
} from "@doffi/core";
import { mountDocumentInteraction } from "./interaction.js";
import type { DocumentInteractionOptions, DocumentInteractionView } from "./interaction.js";
import { mountDocumentPreview } from "./preview.js";
import type { DocumentPreview } from "./preview.js";
import { mountDocumentTextEditor } from "./rich-text.js";
import type { DocumentTextEditor } from "./rich-text.js";
import { mountPresentationWorkbench } from "./presentation-workbench.js";
import type { PresentationWorkbenchOptions,PresentationWorkbenchView } from "./presentation-workbench.js";
import { createLayoutFieldTransaction } from "@doffi/render";
import { workspaceControls } from "./workspace.js";
import { mountWorkbookWorkbench } from "./workbook/workbench.js";
import type { WorkbookWorkbenchOptions, WorkbookWorkbenchView } from "./workbook/workbench.js";
import type { WorkbookSessionConnection } from "./workbook/session-transport.js";

function immutable<Value>(value: Value): Value {
  if (value && typeof value === "object") {
    for (const child of Object.values(value)) immutable(child);
    Object.freeze(value);
  }
  return value;
}

export type DocumentWorkbenchEvent =
  | { readonly type: "selectionChanged"; readonly selection: DocumentSelection | undefined }
  | { readonly type: "annotationsSubmitted"; readonly request: DocumentInteractionRequest }
  | { readonly type: "proposalReady"; readonly proposal: EditProposal; readonly candidate: CandidateChange };
export interface DocumentWorkbenchOptions extends DocumentInteractionOptions {
  readonly format?:"document";
  readonly editable?: boolean;
  readonly onSubmitAnnotations?: AnnotationSubmissionHandler;
  readonly supportedOperations?: InteractionRequestInput["supportedOperations"];
  readonly contextLimits?: InteractionRequestInput["limits"];
  readonly onEvent?: (event: DocumentWorkbenchEvent) => void;
}
export interface DocumentWorkbenchView extends DocumentInteractionView {
  request(): DocumentInteractionRequest | undefined;
  buildRequest(annotationIds?: readonly string[]): DocumentInteractionRequest;
  submitAnnotations(annotationIds?: readonly string[]): Promise<void>;
  loadProposal(request: DocumentInteractionRequest, proposal: EditProposal): CandidateChange;
  applyProposal(): TransactionReceipt;
  rejectProposal(): void;
}

/** Host-owned session; no model execution or authoritative file storage is introduced here. */
export function mountDocumentWorkbench(container:HTMLElement,session:DocumentSession,options?:DocumentWorkbenchOptions):DocumentWorkbenchView;
export function mountDocumentWorkbench(container:HTMLElement,session:DocumentSession,options:PresentationWorkbenchOptions&{readonly format:"presentation"}):PresentationWorkbenchView;
export function mountDocumentWorkbench(container:HTMLElement,session:WorkbookSession|WorkbookSessionConnection,options:WorkbookWorkbenchOptions&{readonly format:"workbook"}):WorkbookWorkbenchView;
export function mountDocumentWorkbench(container: HTMLElement, session: DocumentSession|WorkbookSession|WorkbookSessionConnection, optionsInput: DocumentWorkbenchOptions|(PresentationWorkbenchOptions&{readonly format:"presentation"})|(WorkbookWorkbenchOptions&{readonly format:"workbook"}) = {}): DocumentWorkbenchView|PresentationWorkbenchView|WorkbookWorkbenchView {
  if (optionsInput.format === "workbook") {
    if ("kind" in session && session.kind !== "workbook") throw new DocumentError("UnsupportedFormat", "The workbook workbench requires its workbook session.");
    const { format: _format, ...options } = optionsInput; return mountWorkbookWorkbench(container, session, options);
  }
  if (!("kind" in session) || session.kind === "workbook") throw new DocumentError("UnsupportedFormat", "Workbook workbenches require their explicit format adapter.");
  if(optionsInput.format==="presentation")return mountPresentationWorkbench(container,session,optionsInput);
  if(optionsInput.format!==undefined&&optionsInput.format!=="document")throw new DocumentError("UnsupportedFormat","Unsupported workbench format.");
  return mountContentWorkbench(container, session, optionsInput);
}
function mountContentWorkbench(container: HTMLElement, session: DocumentSession, options: DocumentWorkbenchOptions): DocumentWorkbenchView {
  if (session.snapshot().metadata.presentation) throw new DocumentError("UnsupportedFeature", "Presentation workbench requires its format adapter.");
  const document = container.ownerDocument, element = document.createElement("div"), shadow = element.attachShadow({ mode: "open" });
  const style = document.createElement("style");
  style.textContent = ":host{display:block;font:14px system-ui;color:#20252b}button,input{font:inherit}button{padding:6px 10px;cursor:pointer}button:disabled{cursor:default}section{margin:16px 0;padding:12px;border:1px solid #d5dbe2}h2{font-size:16px}.actions{display:flex;gap:8px;flex-wrap:wrap;margin:10px 0}label{display:block;margin:8px 0}pre{white-space:pre-wrap;overflow-wrap:anywhere;max-height:240px;overflow:auto}li{margin:12px 0}.error{color:#a00000}.candidate{max-height:60vh;overflow:auto}button:focus-visible,input:focus-visible{outline:2px solid #245aa5;outline-offset:2px}";
  const workspace = options.appearance === "workspace";
  if (workspace) style.textContent += workspaceControls + `
    :host{background:#f2f5f1}.editing-controls{display:flex;gap:6px;padding:10px 20px 0;background:#fcfdfb}
    section{margin:0;padding:10px 20px;border:0;border-bottom:1px solid #e1e8df;background:#fcfdfb}
    section[aria-label="请求与提案"]{display:flex;align-items:center;flex-wrap:wrap;gap:8px 12px}
    h2{font-size:11px;font-weight:550;margin:0;color:#91a187}.actions{gap:5px;margin:0;flex-wrap:wrap}
    button{font-size:11px}.actions>button:nth-child(2){background:#eaf2e6;color:#4c7850;border-color:#d8e5ce}
    label{margin:0;font-size:10px;color:#90a084}.proposal-import{position:relative}.proposal-import input{position:absolute;width:1px;height:1px;opacity:0;clip-path:inset(50%)}
    details{font-size:10px;color:#8b9c7f}details summary{cursor:pointer}details>div{position:absolute;z-index:20;max-width:350px;max-height:200px;overflow:auto;padding:12px;background:white;border:1px solid #dfe7da;border-radius:6px;box-shadow:0 5px 20px #23432112}details label{display:flex;align-items:flex-start;gap:6px;margin:8px 0;line-height:1.7}
    p[aria-label="协作状态"]:empty{display:none}p[aria-label="协作状态"]{font-size:10px;color:#809774;margin:0}
    section[aria-label="提案差异"]{width:100%;max-height:65vh;overflow:auto;border:1px solid #d7e4d0;border-radius:6px;background:#f5f8f2;margin:8px 0;padding:15px}
    section[aria-label="提案差异"] h2{font-size:13px;color:#547846}.candidate{background:#eef2eb}.error{width:100%;margin:0}
    @media(max-width:680px){.editing-controls{padding:9px 12px 0}section{padding:9px 12px}.actions{gap:4px}.actions button{font-size:10px;padding:5px 6px}}
  `;
  const interactionHost = document.createElement("div"), collaboration = document.createElement("section");
  collaboration.setAttribute("aria-label", "请求与提案");
  const heading = document.createElement("h2"); heading.textContent = "协作修改";
  const choices = document.createElement("div"); choices.setAttribute("aria-label", "选择提交的批注");
  const actions = document.createElement("div"); actions.className = "actions";
  const status = document.createElement("p"); status.setAttribute("aria-live", "polite"); status.setAttribute("aria-label", "协作状态");
  const error = document.createElement("p"); error.className = "error"; error.setAttribute("role", "alert"); error.hidden = true;
  const diff = document.createElement("section"); diff.setAttribute("aria-label", "提案差异"); diff.hidden = true;
  if (workspace) { const details = document.createElement("details"), summary = document.createElement("summary"); summary.textContent = "选择提交的批注"; details.append(summary, choices); collaboration.append(heading, actions, details, status, error, diff); }
  else collaboration.append(heading, choices, actions, status, error, diff);
  const editingControls = document.createElement("div"), editingHost = document.createElement("div"); editingHost.hidden = true;
  editingControls.className = "editing-controls";
  if (workspace) shadow.append(style, editingControls, collaboration, editingHost, interactionHost);
  else shadow.append(style, editingControls, editingHost, interactionHost, collaboration);
  container.append(element);
  const controller = new AbortController();
  let choiceController = new AbortController(), disposed = false, pending = false, generation = 0, importSequence = 0;
  let activeRequest: DocumentInteractionRequest | undefined, deliveredRequest: DocumentInteractionRequest | undefined, candidate: CandidateChange | undefined;
  let candidatePreview: DocumentPreview | undefined, candidateReady: Promise<void> = Promise.resolve();
  let textEditor: DocumentTextEditor | undefined;
  const selected = new Map<string, boolean>();
  let importControl: HTMLButtonElement | undefined;

  function ensureActive(): void { if (disposed) throw new DocumentError("InvalidOperation", "The workbench has been disposed."); }
  function notify(event: DocumentWorkbenchEvent): void {
    if (disposed) return;
    try { options.onEvent?.(Object.freeze(event)); } catch { /* Hosts own their callback errors. */ }
  }
  function report(reason: unknown): void {
    if (disposed) return;
    const value = reason instanceof Error ? reason : new Error("文档协作操作失败。");
    error.textContent = value.message; error.hidden = false;
    options.onError?.(value);
  }
  function attempt(action: () => void): void {
    if (disposed) return;
    error.hidden = true;
    try { action(); } catch (reason) { report(reason); }
  }
  function button(label: string, action: () => void, parent: HTMLElement = actions): HTMLButtonElement {
    const control = document.createElement("button"); control.type = "button"; control.textContent = label;
    control.addEventListener("click", () => attempt(action), { signal: controller.signal }); parent.append(control); return control;
  }
  const interaction = mountDocumentInteraction(interactionHost, session, {
    ...(options.layout ? { layout: options.layout } : {}), ...(options.limits ? { limits: options.limits } : {}),
    ...(options.title ? { title: options.title } : {}),
    ...(options.appearance ? { appearance: options.appearance } : {}),
    onSelectionChange: (selection) => { options.onSelectionChange?.(selection); notify({ type: "selectionChanged", selection }); },
    onError: report,
  });
  if (options.editable !== false) {
    const edit = button("编辑正文", () => {
      if (textEditor) {
        textEditor.dispose(); textEditor = undefined; editingHost.hidden = true; interactionHost.hidden = false; edit.textContent = "编辑正文";
      } else {
        textEditor = mountDocumentTextEditor(editingHost, session, { ...(options.layout ? { layout: options.layout } : {}), ...(options.limits ? { limits: options.limits } : {}), ...(options.appearance ? { appearance: options.appearance } : {}), onError: report });
        editingHost.hidden = false; interactionHost.hidden = true; edit.textContent = "返回预览与批注";
        void textEditor.ready.catch(report);
      }
    }, editingControls);
    button("更新页码与目录", () => {
      const layout = textEditor?.layout() ?? interaction.layout();
      if (!layout?.pages.length) throw new DocumentError("InvalidOperation", "请等待页面布局完成。");
      const prepared = session.prepare(createLayoutFieldTransaction(session, layout));
      try { session.commit(prepared); } catch (reason) { session.discard(prepared); throw reason; }
    }, editingControls);
  }
  function update(): void {
    if (disposed) return;
    choiceController.abort(); choiceController = new AbortController(); choices.replaceChildren();
    const annotations = session.getAnnotations().filter((annotation) => annotation.status === "open");
    for (const id of selected.keys()) if (!annotations.some((annotation) => annotation.annotationId === id)) selected.delete(id);
    for (const annotation of annotations) {
      if (!selected.has(annotation.annotationId)) selected.set(annotation.annotationId, true);
      const label = document.createElement("label"), checkbox = document.createElement("input"); checkbox.type = "checkbox";
      checkbox.checked = selected.get(annotation.annotationId)!; checkbox.value = annotation.annotationId;
      checkbox.setAttribute("aria-label", "提交批注：" + (annotation.instruction ?? session.getNode(annotation.bodyNodeId!)!.text));
      checkbox.addEventListener("change", () => { selected.set(annotation.annotationId, checkbox.checked); updateButtons(); }, { signal: choiceController.signal });
      label.append(checkbox, document.createTextNode(annotation.instruction ?? session.getNode(annotation.bodyNodeId!)!.text)); choices.append(label);
    }
    if (activeRequest && activeRequest.baseRevision !== session.snapshot().manifest.revision)
      status.textContent = "请求已过期，可保留查看；请基于当前修订重新生成请求。";
    updateButtons();
  }
  function updateButtons(): void {
    const hasSelection = [...selected.values()].some(Boolean);
    generate.disabled = !hasSelection || pending || Boolean(candidate); submit.disabled = !hasSelection || pending || Boolean(candidate) || !options.onSubmitAnnotations;
    download.disabled = !activeRequest;
    input.disabled = !activeRequest || Boolean(candidate);
    if (importControl) importControl.disabled = input.disabled;
    apply.disabled = !candidate || candidate.baseRevision !== session.snapshot().manifest.revision;
    reject.disabled = !candidate;
  }
  function buildRequest(annotationIds?: readonly string[]): DocumentInteractionRequest {
    ensureActive();
    if (candidate) throw new DocumentError("InvalidOperation", "请先应用或拒绝当前提案，再生成新的请求。");
    const request = session.buildInteractionRequest({ annotationIds: annotationIds ?? [...selected].filter(([, value]) => value).map(([id]) => id),
      ...(options.supportedOperations ? { supportedOperations: options.supportedOperations } : {}),
      ...(options.contextLimits ? { limits: options.contextLimits } : {}) });
    activeRequest = request; deliveredRequest = undefined;
    status.textContent = "已生成请求 · 修订 " + request.baseRevision + " · " + request.annotations.length + " 条批注";
    updateButtons(); return request;
  }
  async function submitAnnotations(annotationIds?: readonly string[]): Promise<void> {
    ensureActive();
    if (!options.onSubmitAnnotations) throw new DocumentError("HostSubmissionFailed", "The host has not configured annotation submission.");
    if (pending) throw new DocumentError("InvalidOperation", "An annotation submission is already in progress.");
    const request = buildRequest(annotationIds), current = ++generation;
    pending = true; updateButtons();
    try {
      await options.onSubmitAnnotations(request);
      if (!disposed && current === generation && activeRequest?.requestId === request.requestId
        && activeRequest.documentId === request.documentId && activeRequest.baseRevision === request.baseRevision) {
        if (!deliveredRequest || deliveredRequest.requestId !== request.requestId
          || deliveredRequest.documentId !== request.documentId || deliveredRequest.baseRevision !== request.baseRevision)
          status.textContent = request.baseRevision === session.snapshot().manifest.revision ? "请求已交给宿主，等待提案。" : "请求已交给宿主；文档已变化，返回提案将按旧修订拒绝。";
        notify({ type: "annotationsSubmitted", request });
      }
    } catch {
      if (!disposed && current === generation) throw new DocumentError("HostSubmissionFailed", "The host could not accept the annotation request.");
    } finally { if (!disposed && current === generation) { pending = false; updateButtons(); } }
  }
  function downloadRequest(): void {
    ensureActive(); if (!activeRequest) throw new DocumentError("InvalidOperation", "Generate a request first.");
    const url = URL.createObjectURL(new Blob([JSON.stringify(activeRequest, null, 2)], { type: "application/json" }));
    const anchor = document.createElement("a"); anchor.href = url; anchor.download = "document.request.json"; anchor.click();
    setTimeout(() => URL.revokeObjectURL(url), 0);
  }
  function previewModel(change: CandidateChange): ContentDocumentModel {
    return change.model;
  }
  function loadProposal(request: DocumentInteractionRequest, proposal: EditProposal): CandidateChange {
    ensureActive();
    if (candidate) throw new DocumentError("InvalidOperation", "请先应用或拒绝当前提案，再载入下一份提案。");
    const next = session.prepareProposal(request, proposal);
    let model: ContentDocumentModel;
    try { model = previewModel(next); } catch (reason) { session.discard(next); throw reason; }
    candidatePreview?.dispose(); candidatePreview = undefined;
    candidate = next; activeRequest = immutable(parseContract("interaction-request", request)); deliveredRequest = activeRequest;
    diff.replaceChildren(); diff.hidden = false;
    const title = document.createElement("h2"); title.textContent = "提案差异";
    const summary = document.createElement("p"); summary.textContent = proposal.summary;
    const list = document.createElement("ol");
    for (const change of next.changes) {
      const before = session.getNode(change.nodeId), after = model.getNode(change.nodeId), item = document.createElement("li");
      const text = document.createElement("p"); text.textContent = change.kind === "added" ? "新增内容" : change.kind === "removed" ? "删除内容" : "修改内容";
      const previous = document.createElement("pre"), following = document.createElement("pre");
      previous.setAttribute("aria-label", "修改前"); following.setAttribute("aria-label", "修改后");
      previous.textContent = before?.text ?? "（无）"; following.textContent = after?.text ?? "（无）";
      const details = document.createElement("details"), explanation = document.createElement("summary"); explanation.textContent = "查看结构与样式变化";
      const source = document.createElement("pre"); source.textContent = (before?.html ?? "（无）") + "\n→\n" + (after?.html ?? "（无）");
      details.append(explanation, source); item.append(text, previous, following, details); list.append(item);
    }
    const annotationSummary = document.createElement("p");
    annotationSummary.textContent = "批注结果：" + next.preview.annotations.filter((a) => proposal.annotationIds.includes(a.annotationId))
      .map((a) => a.status === "resolved" ? "已处理" : a.status === "needs_reanchor" ? "需要重新定位" : "待处理").join("、");
    const previewContainer = document.createElement("div"); previewContainer.className = "candidate";
    diff.append(title, summary, list, annotationSummary, previewContainer, apply, reject);
    candidatePreview = mountDocumentPreview(previewContainer, model, { title: "提案候选预览" });
    candidateReady = candidatePreview.ready;
    void candidateReady.catch(report);
    status.textContent = next.changes.length ? "提案已准备，查看差异后可整批应用。" : "提案没有净内容变化，批注将保持原状态。";
    updateButtons();
    notify({ type: "proposalReady", proposal: immutable(parseContract("edit-proposal", proposal)), candidate: next });
    return next;
  }
  function rejectProposal(): void {
    ensureActive(); if (candidate) session.discard(candidate);
    candidate = undefined; candidatePreview?.dispose(); candidatePreview = undefined; candidateReady = Promise.resolve();
    diff.hidden = true; diff.replaceChildren(); status.textContent = "已拒绝提案，文档和批注保持原状。"; updateButtons();
  }
  function applyProposal(): TransactionReceipt {
    ensureActive(); if (!candidate) throw new DocumentError("InvalidOperation", "There is no prepared proposal.");
    const receipt = session.commit(candidate);
    candidate = undefined; candidatePreview?.dispose(); candidatePreview = undefined; candidateReady = Promise.resolve();
    diff.hidden = true; diff.replaceChildren(); status.textContent = receipt.changed ? "提案已整批应用到文档会话。" : "提案没有净变化。";
    update(); return receipt;
  }
  const generate = button("生成请求", () => { buildRequest(); });
  const submit = button("统一提交批注", () => { void submitAnnotations().catch(report); });
  const download = button("下载请求 JSON", downloadRequest);
  const importLabel = document.createElement(workspace ? "div" : "label"); importLabel.textContent = "导入提案 JSON";
  const input = document.createElement("input"); input.type = "file"; input.accept = ".json,application/json"; input.setAttribute("aria-label", "导入提案 JSON");
  input.addEventListener("change", () => {
    const file = input.files?.[0], request = activeRequest, current = ++importSequence;
    input.value = ""; if (!file) return;
    if (!request) { report(new DocumentError("InvalidOperation", "请先生成或提交请求。")); return; }
    if (file.size > 8 * 1024 * 1024) { report(new DocumentError("ResourceLimitExceeded", "提案文件超过大小限制。")); return; }
    void file.text().then((text) => {
      if (disposed || importSequence !== current || activeRequest !== request) return;
      attempt(() => { loadProposal(request, parseContract("edit-proposal", JSON.parse(text))); });
    }).catch(report);
  }, { signal: controller.signal });
  if (workspace) { importLabel.textContent = ""; importLabel.className = "proposal-import"; input.hidden = true; importControl = button("导入提案 JSON", () => input.click(), importLabel); }
  importLabel.append(input); actions.append(importLabel);
  const apply = button("整批应用提案", () => { applyProposal(); }, diff), reject = button("拒绝提案", rejectProposal, diff);
  const unsubscribe = session.subscribe((event) => { if (event.type !== "error") update(); });
  update();
  return {
    element, get ready() { return Promise.all([interaction.ready, candidateReady, textEditor?.ready]).then(() => {}); },
    selection: interaction.selection, fonts: interaction.fonts, layout: () => textEditor?.layout() ?? interaction.layout(), setMode: interaction.setMode, setZoom: interaction.setZoom, focusAnnotation: interaction.focusAnnotation,
    request: () => activeRequest, buildRequest, submitAnnotations, loadProposal, applyProposal, rejectProposal,
    dispose: () => {
      if (disposed) return;
      disposed = true; generation++; unsubscribe(); controller.abort(); choiceController.abort();
      if (candidate) { try { session.discard(candidate); } catch { /* The host may have already disposed the session. */ } }
      candidate = undefined; activeRequest = undefined; deliveredRequest = undefined;
      selected.clear(); textEditor?.dispose(); textEditor = undefined; candidatePreview?.dispose(); interaction.dispose(); element.remove();
    },
  };
}
