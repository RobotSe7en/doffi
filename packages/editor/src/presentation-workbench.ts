import { DocumentError, isPresentationCommentPoint, parseContract } from "@doffi/core";
import type { CandidateChange, DocumentSession, PresentationAnnotationSubmissionHandler, PresentationEditProposal, PresentationInteractionRequest, PresentationInteractionRequestInput, TransactionPrepareOptions, TransactionReceipt } from "@doffi/core";
import { mountPreview } from "@doffi/render";
import type { DocumentPreview } from "@doffi/render";
import { mountPresentationEditor } from "./presentation-editor.js";
import type { PresentationEditorOptions, PresentationEditorView } from "./presentation-editor.js";
import type { DocumentInteractionView } from "./interaction.js";
import { workspaceControls } from "./workspace.js";
function immutable<T>(value: T): T { if (value && typeof value === "object") {
    for (const child of Object.values(value))
        immutable(child);
    Object.freeze(value);
} return value; }
export type PresentationWorkbenchEvent = {
    readonly type: "selectionChanged";
    readonly selection: ReturnType<PresentationEditorView["selection"]>;
} | {
    readonly type: "annotationsSubmitted";
    readonly request: PresentationInteractionRequest;
} | {
    readonly type: "proposalReady";
    readonly proposal: PresentationEditProposal;
    readonly candidate: CandidateChange;
};
export interface PresentationWorkbenchOptions extends PresentationEditorOptions {
    readonly onSubmitAnnotations?: PresentationAnnotationSubmissionHandler;
    readonly supportedOperations?: PresentationInteractionRequestInput["supportedOperations"];
    readonly contextLimits?: PresentationInteractionRequestInput["limits"];
    readonly proposalResources?: TransactionPrepareOptions;
    readonly onEvent?: (event: PresentationWorkbenchEvent) => void;
}
export interface PresentationWorkbenchView extends Omit<DocumentInteractionView, "setRulers" | "settled"> {
    readonly editor: PresentationEditorView;
    setReviewVisible(visible: boolean): void;
    request(): PresentationInteractionRequest | undefined;
    buildRequest(annotationIds?: readonly string[], scope?: "slides" | "global"): PresentationInteractionRequest;
    submitAnnotations(annotationIds?: readonly string[]): Promise<void>;
    loadProposal(request: PresentationInteractionRequest, proposal: PresentationEditProposal, resources?: TransactionPrepareOptions): CandidateChange;
    applyProposal(): TransactionReceipt;
    rejectProposal(): void;
}
/** PPT adapter over the host session and immutable candidates; no model runtime. */
export function mountPresentationWorkbench(container: HTMLElement, session: DocumentSession, options: PresentationWorkbenchOptions = {}): PresentationWorkbenchView {
    if (!session.snapshot().metadata.presentation)
        throw new DocumentError("UnsupportedFormat", "The presentation workbench requires a PPT session.");
    const document = container.ownerDocument, element = document.createElement("div"), shadow = element.attachShadow({ mode: "open" }), controls = new AbortController(), editorHost = document.createElement("div"), panel = document.createElement("section"), choices = document.createElement("div"), actions = document.createElement("div"), status = document.createElement("p"), error = document.createElement("p"), diff = document.createElement("section"), candidateHost = document.createElement("div");
    const style = document.createElement("style");
    style.textContent = ":host{display:block;font:14px system-ui;color:#20252b}button,input,select{font:inherit}button,select{padding:6px 10px}button{cursor:pointer}section{padding:12px;margin:12px 0;border:1px solid #ccd3dc}label{display:inline-flex;gap:5px;margin:5px}button:focus-visible,input:focus-visible,select:focus-visible{outline:2px solid #245aa5;outline-offset:2px}p{white-space:pre-wrap}.actions{display:flex;flex-wrap:wrap;gap:6px}.candidate{max-height:70vh;overflow:auto}table{border-collapse:collapse}td,th{padding:4px 8px;border:1px solid #ccd3dc}.error{color:#a00000}[hidden]{display:none!important}";
    const workspace = options.appearance === "workspace";
    if (workspace) style.textContent += workspaceControls + `
      section{margin:0;padding:10px 18px;border:0;border-bottom:1px solid #e0e6f0;background:#f8faff}section[aria-label="PPT 请求与提案"]{display:flex;align-items:center;gap:8px 14px;flex-wrap:wrap}
      h2{font-size:11px;font-weight:550;margin:0;color:#8596b0}.actions{gap:5px}.actions button,.actions select{font-size:10px;padding:5px 7px}.actions>button:nth-of-type(2){background:#eaf1ff;color:#3567b6;border-color:#cdddf7}
      div[aria-label="选择提交的批注"]{width:100%;font-size:10px;color:#8193ae}label{margin:0 15px 0 0;font-size:10px;line-height:1.8}input[type=checkbox]{accent-color:#175ce5}
      p[aria-label="协作状态"]{font-size:10px;color:#8193ae;margin:0}p[aria-label="协作状态"]:empty{display:none}.error{margin:0;width:100%}
      section[aria-label="提案差异"]{width:100%;max-height:40vh;overflow:auto;background:#f3f7ff;border:1px solid #d5e1f7;border-radius:6px}.candidate:not(:empty){width:100%;max-height:50vh;overflow:auto}
    `;
    panel.setAttribute("aria-label", "PPT 请求与提案");
    choices.setAttribute("aria-label", "选择提交的批注");
    status.setAttribute("aria-label", "协作状态");
    status.setAttribute("aria-live", "polite");
    error.setAttribute("role", "alert");
    error.className = "error";
    error.hidden = true;
    actions.className = "actions";
    diff.setAttribute("aria-label", "提案差异");
    diff.hidden = true;
    candidateHost.className = "candidate";
    const heading = document.createElement("h2");
    heading.textContent = "协作修改";
    panel.append(heading, choices, actions, status, error, diff, candidateHost);
    if (workspace) shadow.append(style, panel, editorHost);
    else shadow.append(style, editorHost, panel);
    let disposed = false, pending = false, renderingCandidate = false, generation = 0, request: PresentationInteractionRequest | undefined, candidate: CandidateChange | undefined, candidateView: DocumentPreview | undefined, candidateReady: Promise<void> = Promise.resolve(), choiceControls = new AbortController();
    const checked = new Set<string>();
    const ensure = (): void => { if (disposed)
        throw new DocumentError("InvalidOperation", "The PPT workbench has been disposed."); };
    const notify = (event: PresentationWorkbenchEvent): void => { try {
        options.onEvent?.(Object.freeze(event));
    }
    catch { /* Hosts own callback failures. */ } };
    const report = (reason: unknown): void => { if (disposed)
        return; const value = reason instanceof Error ? reason : new Error("PPT 协作操作失败。"); error.hidden = false; error.textContent = value.message; try {
        options.onError?.(value);
    }
    catch { /* Host callbacks do not change the source. */ } };
    const attempt = (action: () => void): void => { error.hidden = true; try {
        action();
    }
    catch (reason) {
        report(reason);
    } };
    const button = (label: string, action: () => void): HTMLButtonElement => { const node = document.createElement("button"); node.type = "button"; node.textContent = label; node.addEventListener("click", () => attempt(action), { signal: controls.signal }); actions.append(node); return node; };
    const editor = mountPresentationEditor(editorHost, session, { ...options, onSelectionChange: selection => { try {
            options.onSelectionChange?.(selection);
        }
        catch { /* Callback ownership stays in the host. */ } notify({ type: "selectionChanged", selection }); }, onError: report });
    container.append(element);
    const scope = document.createElement("select");
    scope.setAttribute("aria-label", "请求范围");
    for (const [value, label] of [["slides", "所选批注的完整幻灯片"], ["global", "全局文稿（含共享定义）"]]) {
        const option = document.createElement("option");
        option.value = value!;
        option.textContent = label!;
        scope.append(option);
    }
    actions.append(scope);
    const build = button("生成 PPT 请求", () => { buildRequest(); }), submit = button("提交 PPT 批注", () => { void submitAnnotations().catch(report); }), download = button("下载 PPT 请求", () => {
        ensure();
        if (!request)
            throw new DocumentError("InvalidOperation", "请先生成请求。");
        const url = URL.createObjectURL(new Blob([JSON.stringify(request, null, 2)], { type: "application/json" })), anchor = document.createElement("a");
        anchor.href = url;
        anchor.download = "presentation.request.json";
        anchor.click();
        setTimeout(() => URL.revokeObjectURL(url), 0);
    });
    const file = document.createElement("input");
    file.type = "file";
    file.accept = "application/json,.json";
    file.hidden = true;
    actions.append(file);
    button("载入 PPT 提案", () => file.click());
    file.addEventListener("change", () => { const input = file.files?.[0]; file.value = ""; if (!input)
        return; const version = generation; if (input.size > 4 * 1024 * 1024) {
        report(new DocumentError("ResourceLimitExceeded", "提案文件超过大小限制。"));
        return;
    } void input.text().then(text => { if (disposed || version !== generation)
        return; if (!request)
        throw new DocumentError("InvalidOperation", "请先生成对应请求。"); loadProposal(request, parseContract("presentation-edit-proposal", JSON.parse(text) as unknown)); }).catch(report); }, { signal: controls.signal });
    const apply = button("应用 PPT 提案", () => { applyProposal(); }), reject = button("拒绝 PPT 提案", () => rejectProposal());
    const buttons = (): void => { build.disabled = pending || !checked.size; submit.disabled = pending || !checked.size || !options.onSubmitAnnotations; download.disabled = !request; apply.disabled = !candidate || renderingCandidate || candidate.baseRevision !== session.snapshot().manifest.revision; reject.disabled = !candidate; scope.disabled = pending || !!candidate; };
    const clearCandidate = (discard = true): void => { const previous = candidate; candidate = undefined; renderingCandidate = false; candidateView?.dispose(); candidateView = undefined; candidateHost.replaceChildren(); diff.replaceChildren(); diff.hidden = true; candidateReady = Promise.resolve(); if (discard && previous)
        try {
            session.discard(previous);
        }
        catch { /* A source commit or disposal already revoked the candidate. */ } };
    const refreshChoices = (): void => {
        choiceControls.abort();
        choiceControls = new AbortController();
        choices.replaceChildren();
        const open = session.getAnnotations().filter(a => a.status === "open");
        for (const id of checked)
            if (!open.some(a => a.annotationId === id))
                checked.delete(id);
        for (const annotation of open) {
            const label = document.createElement("label"), input = document.createElement("input"), body = annotation.bodyNodeId ? session.getNode(annotation.bodyNodeId)?.text : annotation.instruction;
            input.type = "checkbox";
            input.checked = checked.has(annotation.annotationId);
            label.append(input, body ?? "批注");
            choices.append(label);
            input.addEventListener("change", () => { if (input.checked)
                checked.add(annotation.annotationId);
            else
                checked.delete(annotation.annotationId); buttons(); }, { signal: choiceControls.signal });
        }
        buttons();
    };
    const validateRegions = (input: PresentationInteractionRequest): void => { const p = session.snapshot().metadata.presentation!; for (const annotation of input.annotations) {
        const anchor = annotation.anchor;
        if (anchor.kind !== "region" || isPresentationCommentPoint(anchor))
            continue;
        const slide = p.slides.find(s => s.rootNodeId === anchor.surfaceId);
        if (!slide || editor.regionLayoutKey(slide.slideId) !== anchor.layoutKey)
            throw new DocumentError("AnchorNotFound", "区域的字体或布局已变化，请先重新定位批注。");
    } };
    function buildRequest(annotationIds: readonly string[] = [...checked], requestedScope: "slides" | "global" = scope.value as "slides" | "global"): PresentationInteractionRequest {
        ensure();
        if (candidate)
            throw new DocumentError("InvalidOperation", "请先应用或拒绝当前候选。");
        const next = session.buildInteractionRequest({ kind: "presentation", annotationIds, scope: requestedScope, ...(options.supportedOperations ? { supportedOperations: options.supportedOperations } : {}), ...(options.contextLimits ? { limits: options.contextLimits } : {}), ...(options.proposalResources?.assets ? { assets: options.proposalResources.assets.map(r => r.descriptor) } : {}) });
        validateRegions(next);
        request = next;
        status.textContent = `已冻结 ${request.context.scope.slideIds.length} 张幻灯片的${requestedScope === "global" ? "全局" : "页面"}请求。`;
        buttons();
        return request;
    }
    async function submitAnnotations(annotationIds?: readonly string[]): Promise<void> {
        ensure();
        if (!options.onSubmitAnnotations)
            throw new DocumentError("HostSubmissionFailed", "宿主未提供批注提交回调。");
        if (pending)
            throw new DocumentError("InvalidOperation", "已有请求正在处理。");
        const current = buildRequest(annotationIds), version = ++generation;
        pending = true;
        buttons();
        try {
            notify({ type: "annotationsSubmitted", request: current });
            const proposal = await options.onSubmitAnnotations(current);
            if (disposed || version !== generation)
                return;
            if (proposal)
                loadProposal(current, proposal);
            else
                status.textContent = "请求已交给宿主。";
        }
        catch (reason) {
            if (!disposed && version === generation)
                throw reason instanceof DocumentError ? reason : new DocumentError("HostSubmissionFailed", "宿主未能完成批注请求。");
        }
        finally {
            if (!disposed && version === generation) {
                pending = false;
                buttons();
            }
        }
    }
    const showDiff = (change: CandidateChange, proposal: PresentationEditProposal): void => {
        diff.replaceChildren();
        diff.hidden = false;
        const summary = document.createElement("p");
        summary.textContent = proposal.summary;
        diff.append(summary);
        const pages = document.createElement("p"), before = session.snapshot().metadata.presentation!, after = change.preview.metadata.presentation!;
        pages.textContent = "受影响幻灯片：" + (change.affectedSlideIds ?? []).map(id => { const slide = after.slides.find(s => s.slideId === id) ?? before.slides.find(s => s.slideId === id); return slide?.name ?? "第 " + Math.max(1, after.slides.findIndex(s => s.slideId === id) + 1) + " 页"; }).join("、");
        diff.append(pages);
        const note = (text: string): void => { const p = document.createElement("p"); p.textContent = text; diff.append(p); };
        if (proposal.annotationIds.length)
            note("同时处理 " + proposal.annotationIds.length + " 条请求批注。");
        if (JSON.stringify(before.dataSets) !== JSON.stringify(after.dataSets)) {
            note("图表数据已修改：");
            for (const data of after.dataSets) {
                const old = before.dataSets.find(d => d.dataSetId === data.dataSetId);
                if (JSON.stringify(old) === JSON.stringify(data))
                    continue;
                const table = document.createElement("table"), head = document.createElement("tr");
                for (const value of ["系列", ...data.categories]) {
                    const cell = document.createElement("th");
                    cell.textContent = value;
                    head.append(cell);
                }
                table.append(head);
                for (const series of data.series) {
                    const row = document.createElement("tr");
                    for (const value of [series.name, ...series.values.map((v, i) => { const previous = old?.series.find(s => s.seriesId === series.seriesId)?.values[i]; return previous === v ? String(v ?? "空") : String(previous ?? "空") + " → " + String(v ?? "空"); })]) {
                        const cell = document.createElement("td");
                        cell.textContent = value;
                        row.append(cell);
                    }
                    table.append(row);
                }
                diff.append(table);
            }
        }
        if (JSON.stringify(before.charts) !== JSON.stringify(after.charts))
            note("图表类型或外观已修改。 ");
        for (const object of after.objects) {
            const previous = before.objects.find(o => o.nodeId === object.nodeId);
            if (!previous || JSON.stringify(previous) === JSON.stringify(object))
                continue;
            const text = change.model.getNode(object.nodeId)?.text.slice(0, 80) ?? "";
            if (object.kind === "connector")
                note("连接线的端点关联或直线／折线配置已修改。");
            else if (object.kind === "table" && previous.kind === "table" && JSON.stringify(previous.grid) !== JSON.stringify(object.grid))
                note("表格列宽：" + previous.grid.join("、") + " → " + object.grid.join("、") + " pt。");
            else if (object.kind === "text" || object.kind === "geometry")
                note("文字布局或占位符配置已修改" + (text ? "：" + text : "。"));
        }
        if (JSON.stringify(before.themes) !== JSON.stringify(after.themes) || JSON.stringify(session.snapshot().styles) !== JSON.stringify(change.preview.styles))
            note("主题或共享样式已修改，请核对受影响页面的预览。");
        if (JSON.stringify([before.masters, before.layouts, before.notesMasters]) !== JSON.stringify([after.masters, after.layouts, after.notesMasters]))
            note("母版或布局定义、引用已修改。");
        if (JSON.stringify(before.size) !== JSON.stringify(after.size))
            note(`幻灯片尺寸：${before.size.width} × ${before.size.height} → ${after.size.width} × ${after.size.height} pt。`);
        for (const item of change.changes) {
            const old = session.getNode(item.nodeId), next = change.model.getNode(item.nodeId);
            if (old?.text === next?.text && old?.attributes["data-office-kind"] === undefined && next?.attributes["data-office-kind"] === undefined)
                continue;
            const detail = document.createElement("details"), heading = document.createElement("summary"), kind = next?.attributes["data-office-kind"] ?? old?.attributes["data-office-kind"] ?? "文字";
            if (kind === "presentation")
                continue;
            const label = ({ text: "文本框", geometry: "形状", image: "图片", connector: "连接线", table: "表格", chart: "图表", group: "组合", slide: "幻灯片", master: "母版", layout: "布局", "notes-master": "备注母版", notes: "备注", comment: "批注" } as Readonly<Record<string, string>>)[kind] ?? "文字";
            heading.textContent = (item.kind === "added" ? "新增" : item.kind === "removed" ? "删除" : "修改") + " · " + label;
            detail.append(heading);
            const text = document.createElement("p");
            text.textContent = (old?.text ?? "") + " → " + (next?.text ?? "");
            detail.append(text);
            for (const key of ["x", "y", "width", "height", "rotation"]) {
                const property = "data-office-" + key, a = old?.attributes[property], b = next?.attributes[property];
                if (a !== b) {
                    const row = document.createElement("p");
                    row.textContent = ({ x: "横坐标", y: "纵坐标", width: "宽度", height: "高度", rotation: "旋转" }[key]) + "：" + (a ?? "继承") + " → " + (b ?? "继承");
                    detail.append(row);
                }
            }
            diff.append(detail);
        }
    };
    function loadProposal(inputRequest: PresentationInteractionRequest, inputProposal: PresentationEditProposal, resources = options.proposalResources): CandidateChange {
        ensure();
        if (candidate)
            throw new DocumentError("InvalidOperation", "请先应用或拒绝当前候选。");
        const next = session.prepareProposal(inputRequest, inputProposal, resources);
        let view: DocumentPreview, frozenRequest: PresentationInteractionRequest, frozenProposal: PresentationEditProposal;
        try {
            frozenRequest = immutable(parseContract("presentation-interaction-request", inputRequest));
            frozenProposal = immutable(parseContract("presentation-edit-proposal", inputProposal));
            validateRegions(frozenRequest);
            view = mountPreview(candidateHost, next.model, { title: "PPT 候选只读预览", ...(options.limits ? { limits: options.limits } : {}) });
            showDiff(next, frozenProposal);
        }
        catch (reason) {
            session.discard(next);
            candidateHost.replaceChildren();
            throw reason;
        }
        request = frozenRequest;
        candidate = next;
        candidateView = view;
        renderingCandidate = true;
        candidateReady = view.ready.then(() => view.settled()).then(() => { if (candidate === next) {
            renderingCandidate = false;
            buttons();
        } });
        void candidateReady.catch(reason => { if (candidate === next) {
            clearCandidate();
            buttons();
            report(reason);
        } });
        status.textContent = "候选已准备，当前文稿尚未修改。";
        buttons();
        notify({ type: "proposalReady", proposal: frozenProposal, candidate: next });
        return next;
    }
    function applyProposal(): TransactionReceipt { ensure(); if (!candidate)
        throw new DocumentError("InvalidOperation", "没有待应用候选。"); const current = candidate; try {
        const receipt = session.commit(current);
        clearCandidate(false);
        status.textContent = "提案已一次应用，可撤销。";
        buttons();
        return receipt;
    }
    catch (reason) {
        clearCandidate();
        buttons();
        throw reason;
    } }
    function rejectProposal(): void { ensure(); clearCandidate(); status.textContent = "已拒绝候选，文稿不变。"; buttons(); }
    const unsubscribe = session.subscribe(event => { if (event.type !== "documentChanged" || disposed)
        return; ++generation; pending = false; clearCandidate(false); if (request && request.baseRevision !== event.revision)
        status.textContent = "文稿已变化，请重新生成请求。"; refreshChoices(); });
    refreshChoices();
    return { element, editor, setReviewVisible(visible) { ensure(); panel.hidden = !visible; panel.style.display = visible ? workspace ? "flex" : "block" : "none"; }, get ready() { return Promise.all([editor.ready, candidateReady]).then(() => { }); }, selection: () => editor.selection(), fonts: () => editor.fonts(), layout: () => editor.layout(), setMode: mode => editor.setMode(mode), setZoom: value => editor.setZoom(value), focusAnnotation: id => editor.focusAnnotation(id), startAnnotation: () => editor.startAnnotation(), request: () => { ensure(); return request; }, buildRequest, submitAnnotations, loadProposal, applyProposal, rejectProposal,
        dispose() { if (disposed)
            return; disposed = true; ++generation; unsubscribe(); controls.abort(); choiceControls.abort(); clearCandidate(); editor.dispose(); element.remove(); } };
}
