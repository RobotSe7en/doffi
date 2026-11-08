import { DocumentError, WorkbookNoteBoxSchema } from "@doffi/core";
import type { Annotation, DeepReadonly, DocumentOperation, WorkbookComment } from "@doffi/core";
import type { WorkbookControlsHost } from "./controls-host.js";
import { reconcileWorkbookInputText, workbookInputTextView, workbookTextSplice } from "./text-projection.js";

/** Native messages and collaboration annotations use the same owning session;
 * this panel retains only a revision-bound draft and a read-only source view. */
export function createWorkbookCommentControls(document: Document, parent: ParentNode, host: WorkbookControlsHost): { readonly element: HTMLDetailsElement; refresh(): void; hasPendingInput(): boolean } {
  const panel = document.createElement("details"), summary = document.createElement("summary"); summary.textContent = "备注与线程批注"; panel.append(summary); parent.append(panel);
  const fields: (HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement)[] = [], buttons: { element: HTMLButtonElement; scope: "new" | "thread" | "reply"; extra?: () => boolean }[] = [];
  const label = <T extends HTMLElement>(name: string, element: T): T => { const wrapper = document.createElement("label"); element.setAttribute("aria-label", name); wrapper.append(name, element); panel.append(wrapper); return element; };
  const input = (name: string, type = "text") => { const field = label(name, document.createElement("input")); field.type = type; fields.push(field); return field; };
  const select = (name: string, values: readonly (readonly [string, string])[]) => { const field = label(name, document.createElement("select")); for (const [value, text] of values) { const option = document.createElement("option"); option.value = value; option.textContent = text; field.append(option); } fields.push(field); return field; };
  const threads = select("选择工作簿批注", []), replies = select("选择批注消息", [["", "主评论"]]), format = select("批注格式", [["modern", "现代线程"], ["classic", "经典备注"]]);
  const author = input("批注作者"), date = input("批注时间"), status = select("批注状态", [["open", "打开"], ["resolved", "已解决"], ["closed", "已关闭"], ["needs_reanchor", "需要重新定位"]]);
  const text = label("批注正文", document.createElement("textarea")); text.maxLength = 2 * 1024 * 1024; fields.push(text);
  const clearMentions = input("替换正文时清除原提及关联", "checkbox"), visible = input("显示经典备注框", "checkbox"), printable = input("打印经典备注框", "checkbox");
  const message = document.createElement("p"); message.setAttribute("role", "status"); message.setAttribute("aria-label", "批注编辑状态"); panel.append(message);
  let selected: string | undefined, replyId: string | undefined, revision = -1, documentId = "", raw = "", dirty = false, composing: HTMLElement | undefined, applying = false, key = "", cached = "", nodes = new Map<string, Element>();
  const draftGroups = new Set<string>();
  function sourceText(id: string): string {
    const snapshot = host.snapshot(), html = snapshot?.content?.html ?? "";
    if (html !== cached) { cached = html; const template = document.createElement("template"); template.innerHTML = html; nodes = new Map([...template.content.querySelectorAll("[data-node-id]")].map(element => [element.getAttribute("data-node-id")!, element])); }
    const root = nodes.get(id); if (!root) throw new DocumentError("InvalidDocument", "批注正文来源不存在。");
    const parts: string[] = [], pending: Node[] = [root];
    while (pending.length) { const node = pending.pop()!; if (node.nodeType === 3) parts.push(node.textContent ?? ""); else if (node.nodeType === 1 && (node as Element).localName === "br") parts.push("\n"); else for (let index = node.childNodes.length - 1; index >= 0; index--) pending.push(node.childNodes[index]!); }
    return parts.join("");
  }
  const thread = () => host.snapshot()?.annotations.find(annotation => annotation.annotationId === selected);
  const binding = () => host.snapshot()?.metadata.workbook.comments?.find(comment => comment.annotationId === selected);
  function current(): DeepReadonly<Annotation> {
    const snapshot = host.snapshot(), annotation = thread(); if (!snapshot || !annotation) throw new DocumentError("InvalidOperation", "请选择批注。");
    if (snapshot.manifest.documentId !== documentId || snapshot.manifest.revision !== revision) throw new DocumentError("RevisionConflict", "批注表单属于旧修订，请重新读取批注。"); return annotation;
  }
  const body = (value: { readonly instruction?: string | undefined; readonly bodyNodeId?: string | undefined }) => value.instruction ?? sourceText(value.bodyNodeId!);
  function load(preserveReply = false): void {
    const snapshot = host.snapshot(); selected = threads.value || undefined; const annotation = thread(), native = binding(), oldReply = preserveReply ? replies.value : "";
    dirty = false; draftGroups.clear(); clearMentions.checked = false; replies.replaceChildren(); const root = document.createElement("option"); root.value = ""; root.textContent = "主评论"; replies.append(root);
    if (!snapshot || !annotation) { replyId = undefined; raw = text.value = author.value = date.value = ""; format.value = "modern"; status.value = "open"; visible.checked = printable.checked = false; revision = snapshot?.manifest.revision ?? -1; documentId = snapshot?.manifest.documentId ?? ""; refresh(); return; }
    for (const [index, reply] of (annotation.replies ?? []).entries()) { const option = document.createElement("option"); option.value = reply.replyId; option.textContent = (reply.author?.name || "回复") + " · " + (index + 1); replies.append(option); }
    replies.value = annotation.replies?.some(reply => reply.replyId === oldReply) ? oldReply : ""; replyId = replies.value || undefined;
    const value = replyId ? annotation.replies!.find(reply => reply.replyId === replyId)! : annotation;
    raw = body(value); text.value = workbookInputTextView(raw); author.value = value.author?.name ?? ""; date.value = value.date ?? "";
    format.value = native?.format ?? "modern"; status.value = annotation.status; visible.checked = native?.format === "classic" && (native.box?.visible ?? false); printable.checked = native?.format === "classic" && (native.box?.printable ?? false);
    revision = snapshot.manifest.revision; documentId = snapshot.manifest.documentId; refresh();
  }
  const html = (text: string) => "<p>" + text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/\r/g, "&#13;").replace(/\n/g, "<br>") + "</p>";
  async function apply(operations: readonly DocumentOperation[]): Promise<void> {
    const groups = new Set<string>(operations.length ? [] : ["text"]); let remove = false;
    for (const operation of operations) {
      if (operation.op === "replaceWorkbookCommentText" || operation.op === "setWorkbookCommentBody") groups.add("text");
      else if (operation.op === "addWorkbookComment" || operation.op === "addWorkbookCommentReply") { groups.add("text"); groups.add("identity"); if (operation.op === "addWorkbookComment") groups.add("format"); }
      else if (operation.op === "setWorkbookNoteBox") groups.add("box");
      else if (operation.op === "updateWorkbookComment" || operation.op === "updateWorkbookCommentReply") {
        if ("author" in operation.changes || "date" in operation.changes) groups.add("identity");
        if ("status" in operation.changes) groups.add("status"); if ("format" in operation.changes) groups.add("format");
      } else if (operation.op === "removeAnnotation" || operation.op === "removeWorkbookCommentReply") remove = true;
    }
    await host.perform(operations, undefined, { documentId, revision });
    if (remove) draftGroups.clear(); else for (const group of groups) draftGroups.delete(group);
    dirty = draftGroups.size > 0; if (!dirty) revision = -1;
  }
  function button(name: string, action: () => Promise<void>, scope: "new" | "thread" | "reply" = "thread", extra?: () => boolean) {
    const element = document.createElement("button"); element.type = "button"; element.textContent = name; panel.append(element); buttons.push({ element, scope, ...(extra ? { extra } : {}) });
    element.addEventListener("click", () => host.run(async () => {
      if (!host.idle() || composing || applying) throw new DocumentError("InvalidOperation", "请先完成当前输入或操作。");
      const current = host.snapshot()?.manifest; if (!current || current.documentId !== documentId || current.revision !== revision) throw new DocumentError("RevisionConflict", "批注表单属于旧修订，请重新读取批注。");
      applying = true;
      try { const submitted = action(); refresh(); host.draftChanged?.(); await submitted; } finally { applying = false; refresh(); host.draftChanged?.(); }
    }), { signal: host.signal });
  }
  button("在选区添加批注", async () => { const selection = host.selection(); if (!selection) throw new DocumentError("InvalidOperation", "请先选择单元格区域。");
    await apply([{ op: "addWorkbookComment", anchor: { kind: "cells", sheetId: selection.sheetId, ranges: [{ from: { ...selection.range.from }, to: { ...selection.range.to } }] }, html: html(text.value), format: format.value as "modern" | "classic",
      ...(author.value ? { author: { name: author.value } } : {}), ...(date.value ? { date: date.value } : {}) }]); }, "new");
  button("添加线程回复", async () => { const root = current(); await apply([{ op: "addWorkbookCommentReply", annotationId: root.annotationId, html: html(text.value), ...(author.value ? { author: { name: author.value } } : {}), ...(date.value ? { date: date.value } : {}), ...(replyId ? { parentReplyId: replyId } : {}) }]); }, "thread", () => binding()?.format !== "classic");
  button("保存批注文字", async () => { const root = current(), after = reconcileWorkbookInputText(raw, text.value), edit = workbookTextSplice(raw, after);
    await apply(edit ? [{ op: "replaceWorkbookCommentText", annotationId: root.annotationId, ...(replyId ? { replyId } : {}), ...edit, ...(clearMentions.checked ? { mentions: [] } : {}) }] : []); });
  button("替换整段批注正文", async () => { const root = current(), after = reconcileWorkbookInputText(raw, text.value), edit = workbookTextSplice(raw, after);
    await apply([{ op: "setWorkbookCommentBody", annotationId: root.annotationId, ...(replyId ? { replyId } : {}), html: html(after), ...(clearMentions.checked ? { mentions: [] } : edit ? { textEdits: [edit] } : {}) }]); });
  button("更新批注作者时间与状态", async () => { const root = current();
    await apply([replyId ? { op: "updateWorkbookCommentReply", annotationId: root.annotationId, replyId, changes: { author: author.value ? { name: author.value } : null, date: date.value || null } }
      : { op: "updateWorkbookComment", annotationId: root.annotationId, changes: { author: author.value ? { name: author.value } : null, date: date.value || null, status: status.value as "open", format: format.value as "modern" | "classic" } }]); });
  button("将批注重新锚定到选区", async () => { const root = current(), selection = host.selection(); if (!selection) throw new DocumentError("InvalidOperation", "请选择新的单元格区域。");
    await apply([{ op: "updateWorkbookComment", annotationId: root.annotationId, changes: { anchor: { kind: "cells", sheetId: selection.sheetId, ranges: [{ from: { ...selection.range.from }, to: { ...selection.range.to } }] }, status: status.value === "needs_reanchor" ? "open" : status.value as "open" } }]); });
  button("应用经典备注框显示", async () => { const root = current(), native = binding(); if (native?.format !== "classic") throw new DocumentError("InvalidOperation", "请选择经典备注。");
    const anchor = root.anchor; if (anchor.kind !== "cells" || !anchor.ranges.length) throw new DocumentError("InvalidOperation", "请先重新定位备注。");
    const from = anchor.ranges[0]!.from, box: Extract<WorkbookComment, { format: "classic" }>["box"] = native.box ? WorkbookNoteBoxSchema.parse({ ...native.box, visible: visible.checked, printable: printable.checked }) : {
      anchor: { kind: "one-cell", from: { ...from, offsetXPt: 20, offsetYPt: 0 }, widthPt: 180, heightPt: 90 }, visible: visible.checked, printable: printable.checked,
      locked: true, lockText: false, defaultSize: true, disabled: false, autoFill: true, autoLine: true, autoScale: false, justifyLastLine: false };
    await apply([{ op: "setWorkbookNoteBox", annotationId: root.annotationId, box }]); }, "thread", () => binding()?.format === "classic");
  button("删除所选回复", async () => { const root = current(); if (!replyId) throw new DocumentError("InvalidOperation", "请选择回复。"); await apply([{ op: "removeWorkbookCommentReply", annotationId: root.annotationId, replyId }]); }, "reply");
  button("删除整条批注", async () => { const root = current(); await apply([{ op: "removeAnnotation", annotationId: root.annotationId }]); });
  const reload = document.createElement("button"); reload.type = "button"; reload.textContent = "重新读取批注"; panel.append(reload);
  const discard = document.createElement("button"); discard.type = "button"; discard.textContent = "放弃批注表单草稿"; panel.append(discard);
  for (const button of [reload, discard]) button.addEventListener("click", () => { if (!composing && !applying && host.idle()) { load(true); host.draftChanged?.(); } }, { signal: host.signal });
  for (const field of [threads, replies]) field.addEventListener("change", () => {
    if (dirty || composing || applying) { threads.value = selected ?? ""; replies.value = replyId ?? ""; message.textContent = "请先应用或放弃批注草稿，再切换线程或消息。"; return; }
    load(field === replies); host.draftChanged?.();
  }, { signal: host.signal });
  const changed = (field: HTMLElement) => { dirty = true; draftGroups.add(field === text || field === clearMentions ? "text" : field === author || field === date ? "identity" : field === status ? "status" : field === format ? "format" : "box"); refresh(); host.draftChanged?.(); };
  for (const field of fields.filter(field => field !== threads && field !== replies)) field.addEventListener("input", () => changed(field), { signal: host.signal });
  panel.addEventListener("compositionstart", event => { if (event.target instanceof document.defaultView!.HTMLInputElement || event.target instanceof document.defaultView!.HTMLTextAreaElement) { composing = event.target; changed(composing); } }, { signal: host.signal });
  panel.addEventListener("compositionend", () => { composing = undefined; refresh(); host.draftChanged?.(); }, { signal: host.signal });
  function refresh(): void {
    const snapshot = host.snapshot(), nextKey = JSON.stringify([snapshot?.manifest.documentId, snapshot?.annotations.map(annotation => [annotation.annotationId, annotation.author?.name, annotation.status])]);
    if (nextKey !== key) {
      key = nextKey; const previous = threads.value; threads.replaceChildren(); const empty = document.createElement("option"); empty.value = ""; empty.textContent = "请选择批注"; threads.append(empty);
      for (const [index, annotation] of (snapshot?.annotations ?? []).entries()) { const option = document.createElement("option"); option.value = annotation.annotationId; option.textContent = (annotation.author?.name || "批注") + " · " + (index + 1); threads.append(option); }
      if (dirty && selected && !snapshot?.annotations.some(annotation => annotation.annotationId === selected)) { const old = document.createElement("option"); old.value = selected; old.textContent = "已移除，草稿保留"; threads.append(old); }
      threads.value = snapshot?.annotations.some(annotation => annotation.annotationId === previous) || dirty && previous === selected ? previous : ""; if (!threads.value && !dirty) selected = undefined;
    }
    if (snapshot && (revision !== snapshot.manifest.revision || documentId !== snapshot.manifest.documentId) && !dirty && !composing) { load(true); return; }
    const idle = host.idle() && !applying, stale = revision >= 0 && (snapshot?.manifest.documentId !== documentId || snapshot.manifest.revision !== revision), root = thread();
    message.textContent = stale ? "文档已变化。批注草稿已保留，请重新读取批注后编辑。" : dirty ? "批注草稿尚未保存。" : root?.status === "needs_reanchor" ? "批注需要显式重新定位；原正文和原生来源已保留。" : "";
    for (const field of fields) field.disabled = (!idle || !!composing) && field !== composing; reload.disabled = !idle || !root || !!composing; discard.disabled = !idle || !dirty || !!composing;
    status.disabled ||= !root || !!replyId; format.disabled ||= !!replyId; clearMentions.disabled ||= !root; visible.disabled ||= binding()?.format !== "classic"; printable.disabled ||= binding()?.format !== "classic";
    for (const button of buttons) button.element.disabled = !idle || !!composing || stale || button.scope !== "new" && !root || button.scope === "reply" && !replyId || button.extra?.() === false;
  }
  refresh(); return { element: panel, refresh, hasPendingInput: () => dirty || !!composing || applying };
}
