import { DocumentError, WorkbookNoteBoxSchema } from "@doffi/core";
import type { Annotation, DeepReadonly, DocumentOperation, WorkbookComment } from "@doffi/core";
import type { WorkbookControlsHost } from "./controls-host.js";
import { reconcileWorkbookInputText, workbookInputTextView, workbookTextSplice } from "./text-projection.js";

/** Native messages and collaboration annotations use the same owning session;
 * this panel retains only a revision-bound draft and a read-only source view. */
export function createWorkbookCommentControls(document: Document, parent: ParentNode, host: WorkbookControlsHost): { readonly element: HTMLDetailsElement; refresh(): void } {
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
  let selected: string | undefined, replyId: string | undefined, revision = -1, documentId = "", raw = "", dirty = false, composing = false, key = "", cached = "", nodes = new Map<string, Element>();
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
    dirty = false; clearMentions.checked = false; replies.replaceChildren(); const root = document.createElement("option"); root.value = ""; root.textContent = "主评论"; replies.append(root);
    if (!snapshot || !annotation) { replyId = undefined; raw = ""; revision = snapshot?.manifest.revision ?? -1; documentId = snapshot?.manifest.documentId ?? ""; refresh(); return; }
    for (const [index, reply] of (annotation.replies ?? []).entries()) { const option = document.createElement("option"); option.value = reply.replyId; option.textContent = (reply.author?.name || "回复") + " · " + (index + 1); replies.append(option); }
    replies.value = annotation.replies?.some(reply => reply.replyId === oldReply) ? oldReply : ""; replyId = replies.value || undefined;
    const value = replyId ? annotation.replies!.find(reply => reply.replyId === replyId)! : annotation;
    raw = body(value); text.value = workbookInputTextView(raw); author.value = value.author?.name ?? ""; date.value = value.date ?? "";
    format.value = native?.format ?? "modern"; status.value = annotation.status; visible.checked = native?.format === "classic" && (native.box?.visible ?? false); printable.checked = native?.format === "classic" && (native.box?.printable ?? false);
    revision = snapshot.manifest.revision; documentId = snapshot.manifest.documentId; refresh();
  }
  const html = (text: string) => "<p>" + text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/\r/g, "&#13;").replace(/\n/g, "<br>") + "</p>";
  async function apply(operations: readonly DocumentOperation[]): Promise<void> { const expected = host.snapshot()?.manifest; await host.perform(operations, undefined, expected); dirty = false; revision = -1; }
  function button(name: string, action: () => Promise<void>, scope: "new" | "thread" | "reply" = "thread", extra?: () => boolean) {
    const element = document.createElement("button"); element.type = "button"; element.textContent = name; panel.append(element); buttons.push({ element, scope, ...(extra ? { extra } : {}) });
    element.addEventListener("click", () => host.run(async () => { if (!host.idle() || composing) throw new DocumentError("InvalidOperation", "请先完成当前输入或操作。"); await action(); refresh(); }), { signal: host.signal });
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
  const reload = document.createElement("button"); reload.type = "button"; reload.textContent = "重新读取批注"; panel.append(reload); reload.addEventListener("click", () => load(true), { signal: host.signal });
  threads.addEventListener("change", () => load(), { signal: host.signal }); replies.addEventListener("change", () => load(true), { signal: host.signal });
  for (const field of fields.filter(field => field !== threads && field !== replies)) field.addEventListener("input", () => { dirty = true; }, { signal: host.signal });
  text.addEventListener("compositionstart", () => { composing = true; dirty = true; refresh(); }, { signal: host.signal });
  text.addEventListener("compositionend", () => { composing = false; refresh(); }, { signal: host.signal });
  function refresh(): void {
    const snapshot = host.snapshot(), nextKey = JSON.stringify([snapshot?.manifest.documentId, snapshot?.annotations.map(annotation => [annotation.annotationId, annotation.author?.name, annotation.status])]);
    if (nextKey !== key) {
      key = nextKey; const previous = threads.value; threads.replaceChildren(); const empty = document.createElement("option"); empty.value = ""; empty.textContent = "请选择批注"; threads.append(empty);
      for (const [index, annotation] of (snapshot?.annotations ?? []).entries()) { const option = document.createElement("option"); option.value = annotation.annotationId; option.textContent = (annotation.author?.name || "批注") + " · " + (index + 1); threads.append(option); }
      threads.value = snapshot?.annotations.some(annotation => annotation.annotationId === previous) ? previous : ""; if (!threads.value) selected = undefined;
    }
    if (selected && snapshot && (revision !== snapshot.manifest.revision || documentId !== snapshot.manifest.documentId) && !dirty) { load(true); return; }
    const idle = host.idle(), stale = !!selected && (snapshot?.manifest.documentId !== documentId || snapshot.manifest.revision !== revision), root = thread();
    message.textContent = stale ? "文档已变化。批注草稿已保留，请重新读取批注后编辑。" : root?.status === "needs_reanchor" ? "批注需要显式重新定位；原正文和原生来源已保留。" : "";
    for (const field of fields) field.disabled = !idle && !(field === text && composing); reload.disabled = !idle || !root || composing;
    for (const button of buttons) button.element.disabled = !idle || composing || button.scope !== "new" && (!root || stale) || button.scope === "reply" && !replyId || button.extra?.() === false;
  }
  return { element: panel, refresh };
}
