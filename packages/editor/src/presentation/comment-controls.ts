import { AnchorSchema, DocumentError } from "@doffi/core";
import type { DocumentEditingSession, DocumentOperation, DocumentSelection } from "@doffi/core";
import { createAnnotationPopover } from "../workspace.js";
export interface CommentControlsHost {
    readonly session: DocumentEditingSession;
    selection(): DocumentSelection | undefined;
    slideRoot(): string;
    perform(operations: DocumentOperation[]): void;
    edit(annotationId: string, replyId?: string): void;
    locate(annotationId: string): void;
    attempt(action: () => void): void;
}
export interface CommentControls {
    refresh(): void;
    start(): void;
    selectionChanged(): void;
    clearSelection(): void;
    dispose(): void;
}
const escape = (text: string): string => text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/\n/g, "<br>");
export function createCommentControls(document: Document, parent: HTMLElement, host: CommentControlsHost, floating?: { root: ShadowRoot; rect(selection: DocumentSelection): DOMRect | undefined; bounds(): DOMRect; toolbar: HTMLElement }): CommentControls {
    const pane = document.createElement("fieldset"), legend = document.createElement("legend"), list = document.createElement("div"), controls = new AbortController();
    let listControls = new AbortController();
    legend.textContent = "批注与回复";
    pane.append(legend);
    parent.append(pane);
    const field = (label: string): HTMLInputElement => { const wrap = document.createElement("label"), node = document.createElement("input"); node.type = "text"; node.setAttribute("aria-label", label); wrap.append(label, node); pane.append(wrap); return node; };
    const author = field("批注显示署名"), draft = floating ? document.createElement("textarea") : field("新批注正文"), format = document.createElement("select");
    draft.setAttribute("aria-label", "新批注正文");
    if (floating) { draft.placeholder = "写下批注或希望修改的内容…"; pane.append(draft); }
    format.setAttribute("aria-label", "新批注格式");
    for (const [value, label] of [["modern", "现代批注"], ["classic", "经典批注"]]) {
        const option = document.createElement("option");
        option.value = value!;
        option.textContent = label!;
        format.append(option);
    }
    pane.append(format);
    const button = (label: string, action: () => void, parent: HTMLElement = pane, signal = controls.signal): void => { const node = document.createElement("button"); node.type = "button"; node.textContent = label; node.addEventListener("click", () => host.attempt(action), { signal }); parent.append(node); };
    let chosen: DocumentSelection | undefined;
    const composer = floating ? document.createElement("div") : pane;
    if (floating) { composer.className = "composer"; pane.append(composer); }
    const add = (): void => {
        const identity = host.session.snapshot().manifest;
        if (floating && (!chosen || chosen.documentId !== identity.documentId || chosen.baseRevision !== identity.revision)) throw new DocumentError("RevisionConflict", "选择已过期，请重新选择要批注的内容。");
        if (floating && !draft.value.trim()) throw new DocumentError("InvalidOperation", "请填写批注内容。");
        const before = new Set(host.session.getAnnotations().map(a => a.annotationId)), anchor = (floating ? chosen : host.selection())?.anchor ?? { kind: "node" as const, nodeId: host.slideRoot() };
        host.perform([{ op: "addPresentationComment", anchor: AnchorSchema.parse(anchor), html: "<p>" + escape(draft.value) + "</p>", format: format.value as "classic" | "modern", ...(author.value ? { author: { name: author.value } } : {}) }]);
        draft.value = ""; popover?.close(); chosen = undefined;
        if (!floating) host.edit(host.session.getAnnotations().find(a => !before.has(a.annotationId))!.annotationId);
    };
    button("添加批注", add, composer);
    const popover = floating ? createAnnotationPopover(floating.root, draft as HTMLTextAreaElement, [author.parentElement!, format, draft, composer], {
        bounds: floating.bounds,
        rect: () => chosen ? floating.rect(chosen) : undefined,
        quote: () => {
            const anchor = chosen?.anchor;
            if (!anchor) return "";
            if (anchor.kind === "text") return anchor.quote;
            if (anchor.kind === "text-set") return anchor.ranges.map(range => range.quote).join(" … ");
            if (anchor.kind === "node") return host.session.getNode(anchor.nodeId)?.text || "所选对象";
            return anchor.kind === "region" ? "所选幻灯片区域" : "所选幻灯片对象";
        },
        submit: () => host.attempt(add), cancel: () => { draft.value = ""; chosen = undefined; },
    }) : undefined;
    const clearSelection = (): void => { chosen = undefined; draft.value = ""; popover?.close(); };
    const show = (): void => { const next = host.selection(); if (!next) { clearSelection(); return; } if (JSON.stringify(chosen) !== JSON.stringify(next)) draft.value = ""; chosen = next; popover?.open(); };
    if (floating) { button("新建批注", () => { if (!host.selection()) throw new DocumentError("InvalidOperation", "请先选择幻灯片内容。"); show(); draft.focus(); }, floating.toolbar); button("取消", () => { popover?.close(); chosen = undefined; draft.value = ""; }, composer); }
    pane.append(list);
    return { start() { if (!host.selection()) throw new DocumentError("InvalidOperation", "请先选择幻灯片内容。"); show(); draft.focus(); }, selectionChanged: show, clearSelection, refresh() {
            if (chosen && chosen.baseRevision !== host.session.snapshot().manifest.revision) { chosen = undefined; draft.value = ""; popover?.close(); }
            listControls.abort();
            listControls = new AbortController();
            list.replaceChildren();
            const p = host.session.snapshot().metadata.presentation!;
            for (const annotation of host.session.getAnnotations()) {
                const item = document.createElement("section"), title = document.createElement("p"), body = annotation.bodyNodeId ? host.session.getNode(annotation.bodyNodeId)?.text : annotation.instruction;
                title.textContent = (annotation.author?.name ?? "未署名") + " · " + (p.comments.find(c => c.annotationId === annotation.annotationId)?.format === "classic" ? "经典" : "现代") + " · " + ({ open: "待处理", resolved: "已解决", closed: "已关闭", needs_reanchor: "待重新定位" }[annotation.status]) + " · " + (body ?? "");
                item.append(title);
                list.append(item);
                const action = (label: string, callback: () => void): void => button(label, callback, item, listControls.signal);
                action("编辑批注正文", () => host.edit(annotation.annotationId));
                action("定位批注", () => host.locate(annotation.annotationId));
                action("删除批注", () => host.perform([{ op: "removeAnnotation", annotationId: annotation.annotationId }]));
                for (const [label, status] of [["解决批注", "resolved"], ["重新打开批注", "open"], ["关闭批注", "closed"]] as const)
                    action(label, () => host.perform([{ op: "updatePresentationComment", annotationId: annotation.annotationId, changes: { status } }]));
                action("按当前选择重新定位", () => { const anchor = host.selection()?.anchor; if (!anchor)
                    throw new Error("请先选择新的幻灯片内容。"); const status = annotation.status === "needs_reanchor" ? p.comments.find(c => c.annotationId === annotation.annotationId)?.recoveryStatus ?? "open" : annotation.status; host.perform([{ op: "updatePresentationComment", annotationId: annotation.annotationId, changes: { anchor: AnchorSchema.parse(anchor), status } }]); });
                const reply = document.createElement("input");
                reply.setAttribute("aria-label", "回复正文");
                item.append(reply);
                action("添加回复", () => { const before = new Set(annotation.replies?.map(r => r.replyId)); host.perform([{ op: "addPresentationCommentReply", annotationId: annotation.annotationId, html: "<p>" + escape(reply.value) + "</p>", ...(author.value ? { author: { name: author.value } } : {}) }]); host.edit(annotation.annotationId, host.session.getAnnotations().find(a => a.annotationId === annotation.annotationId)!.replies!.find(r => !before.has(r.replyId))!.replyId); });
                for (const entry of annotation.replies ?? []) {
                    const section = document.createElement("div");
                    section.textContent = (entry.author?.name ?? "未署名") + "：" + (entry.bodyNodeId ? host.session.getNode(entry.bodyNodeId)?.text : entry.instruction);
                    item.append(section);
                    button("编辑回复", () => host.edit(annotation.annotationId, entry.replyId), section, listControls.signal);
                    button("删除回复", () => host.perform([{ op: "removePresentationCommentReply", annotationId: annotation.annotationId, replyId: entry.replyId }]), section, listControls.signal);
                }
            }
        }, dispose() { controls.abort(); listControls.abort(); popover?.dispose(); pane.remove(); } };
}
