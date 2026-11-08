import { DocumentError } from "@doffi/core";
import type { DeepReadonly, DocumentOperation, WorkbookSheet } from "@doffi/core";
import type { WorkbookControlsHost, WorkbookControlsIdentity } from "./controls-host.js";

type Action = "properties" | "copy" | "insert" | "move" | "remove";
interface SheetDraft {
  readonly identity: WorkbookControlsIdentity;
  readonly sheetId: string;
  readonly action: Action;
  name: string; visibility: WorkbookSheet["visibility"]; before: string; dirty: boolean;
  conflicts: Map<string, { original: string; name: string }>;
}
/** Drafts belong to an action, worksheet and revision. Source ownership and
 * all reference/name decisions remain in the owning session transaction. */
export function createWorkbookWorksheetControls(document: Document, parent: ParentNode, host: WorkbookControlsHost): { readonly element: HTMLDetailsElement; refresh(): void; hasPendingInput(): boolean } {
  const panel = document.createElement("details"), summary = document.createElement("summary"); summary.textContent = "工作表管理"; panel.append(summary); parent.append(panel);
  const fields: (HTMLInputElement | HTMLSelectElement)[] = [];
  const label = <T extends HTMLInputElement | HTMLSelectElement>(text: string, field: T) => { const wrapper = document.createElement("label"); field.setAttribute("aria-label", text); wrapper.append(text, field); panel.append(wrapper); fields.push(field); return field; };
  const sheets = label("管理工作表", document.createElement("select")), action = label("工作表操作", document.createElement("select"));
  const option = (select: HTMLSelectElement, value: string, text: string) => { const item = document.createElement("option"); item.value = value; item.textContent = text; select.append(item); };
  for (const [value, text] of [["properties", "名称与可见性"], ["copy", "复制工作表"], ["insert", "新建工作表"], ["move", "移动工作表"], ["remove", "删除工作表"]]) option(action, value!, text!);
  const name = label("工作表名称", document.createElement("input")), visibility = label("工作表可见性", document.createElement("select")), before = label("放在工作表之前", document.createElement("select"));
  for (const [value, text] of [["visible", "显示"], ["hidden", "隐藏"], ["veryHidden", "深度隐藏"]]) option(visibility, value!, text!); name.maxLength = 31;
  const conflicts = document.createElement("div"), status = document.createElement("p"); status.setAttribute("aria-label", "工作表编辑状态"); status.setAttribute("role", "status"); panel.append(conflicts, status);
  const drafts = new Map<string, SheetDraft>(); let draft: SheetDraft | undefined, selected = "", actionValue: Action = "properties", listKey = "", composing: HTMLElement | undefined, saving = false;
  const key = () => JSON.stringify([selected, actionValue]);
  const original = (): DeepReadonly<WorkbookSheet> | undefined => host.snapshot()?.metadata.workbook.sheets.find(sheet => sheet.sheetId === selected);
  const stale = () => !!draft && (host.snapshot()?.manifest.documentId !== draft.identity.documentId || host.snapshot()?.manifest.revision !== draft.identity.revision);
  const invalid = (text: string): never => { throw new DocumentError("InvalidOperation", text); };
  function load(reset = false): void {
    const snapshot = host.snapshot(); if (!snapshot) return;
    selected = sheets.value; actionValue = action.value as Action; let value = reset ? undefined : drafts.get(key());
    if (!value) {
      const sheet = original(), used = new Set(snapshot.metadata.workbook.sheets.map(sheet => sheet.name.toUpperCase())); let proposed = sheet?.name ?? "Sheet";
      if (actionValue === "copy" || actionValue === "insert") { const base = actionValue === "copy" ? proposed : "Sheet"; let serial = 2; do { const suffix = " (" + serial++ + ")"; proposed = base.slice(0, 31 - suffix.length) + suffix; } while (used.has(proposed.toUpperCase())); }
      value = { identity: { documentId: snapshot.manifest.documentId, revision: snapshot.manifest.revision }, sheetId: selected, action: actionValue, name: proposed, visibility: sheet?.visibility ?? "visible", before: "", dirty: false, conflicts: new Map() }; drafts.set(key(), value);
    }
    draft = value; name.value = value.name; visibility.value = value.visibility; before.value = value.before; renderConflicts(); refresh();
  }
  function change(): void { if (!draft) return; Object.assign(draft, { name: name.value, visibility: visibility.value, before: before.value, dirty: true }); refresh(); host.draftChanged?.(); }
  function renderConflicts(): void {
    conflicts.replaceChildren();
    for (const [nameId, value] of draft?.conflicts ?? []) {
      const wrapper = document.createElement("label"), field = document.createElement("input"); field.maxLength = 255; field.value = value.name; field.setAttribute("aria-label", "复制名称冲突：" + value.original); wrapper.append("为复制的全局名称“" + value.original + "”指定新名称", field); conflicts.append(wrapper);
      field.addEventListener("input", () => { if (!draft) return; draft.conflicts.get(nameId)!.name = field.value; draft.dirty = true; refresh(); host.draftChanged?.(); }, { signal: host.signal });
      field.addEventListener("compositionstart", () => { composing = field; change(); }, { signal: host.signal });
      field.addEventListener("compositionend", () => { composing = undefined; change(); }, { signal: host.signal });
    }
  }
  const button = (text: string, callback: () => Promise<void>) => {
    const element = document.createElement("button"); element.type = "button"; element.textContent = text; panel.append(element);
    element.addEventListener("click", () => host.run(async () => { if (saving || composing || !host.idle()) invalid("请先完成当前输入或操作。"); saving = true; refresh(); try { await callback(); } finally { saving = false; refresh(); host.draftChanged?.(); } }), { signal: host.signal }); return element;
  };
  const save = button("执行工作表操作", async () => {
    if (!draft || stale()) throw new DocumentError("RevisionConflict", "工作表草稿属于旧修订，请重新读取。");
    const submitted = draft, snapshot = host.snapshot()!, position = submitted.before ? { beforeSheetId: submitted.before } : {};
    let operation: DocumentOperation;
    switch (submitted.action) {
      case "properties": operation = { op: "updateWorkbookSheet", sheetId: submitted.sheetId, changes: { name: submitted.name, visibility: submitted.visibility } }; break;
      case "copy": operation = { op: "copyWorkbookSheet", sheetId: submitted.sheetId, name: submitted.name, ...position, ...(submitted.conflicts.size ? { nameConflicts: [...submitted.conflicts].map(([nameId, value]) => ({ nameId, name: value.name })) } : {}) }; break;
      case "insert": operation = { op: "insertWorkbookSheet", name: submitted.name, ...position }; break;
      case "move": operation = { op: "moveWorkbookSheet", sheetId: submitted.sheetId, ...position }; break;
      case "remove": operation = { op: "removeWorkbookSheet", sheetId: submitted.sheetId }; break;
    }
    try { await host.perform([operation], undefined, submitted.identity); }
    catch (error) {
      if (error instanceof DocumentError) for (const detail of error.details) if (detail.code === "copy-name-conflict" && typeof detail.path[0] === "string") {
        submitted.conflicts.set(detail.path[0], { original: detail.message, name: "" }); submitted.dirty = true; renderConflicts();
      }
      throw error;
    }
    drafts.delete(key()); const after = host.snapshot()!;
    const added = submitted.action === "copy" || submitted.action === "insert" ? after.metadata.workbook.sheets.find(sheet => !snapshot.metadata.workbook.sheets.some(previous => previous.sheetId === sheet.sheetId)) : undefined;
    selected = added?.sheetId ?? (after.metadata.workbook.sheets.some(sheet => sheet.sheetId === submitted.sheetId) ? submitted.sheetId : after.metadata.workbook.activeSheetId ?? after.metadata.workbook.sheets.find(sheet => sheet.visibility === "visible")!.sheetId);
    draft = undefined; listKey = ""; refresh();
    if (added?.visibility === "visible") await host.selectSheet?.(added.sheetId);
  });
  const reload = button("放弃当前工作表草稿并重新读取", async () => {
    drafts.delete(key()); if (!original()) { selected = host.snapshot()?.metadata.workbook.sheets[0]?.sheetId ?? ""; listKey = ""; draft = undefined; refresh(); } else load(true);
  });
  for (const field of [sheets, action]) field.addEventListener("change", () => { load(); host.draftChanged?.(); }, { signal: host.signal });
  for (const field of [name, visibility, before]) field.addEventListener("input", change, { signal: host.signal });
  for (const field of fields) {
    field.addEventListener("compositionstart", () => { composing = field; change(); }, { signal: host.signal });
    field.addEventListener("compositionend", () => { composing = undefined; change(); }, { signal: host.signal });
  }
  function refresh(): void {
    const snapshot = host.snapshot(), values = snapshot?.metadata.workbook.sheets ?? [], next = JSON.stringify([snapshot?.manifest.documentId, values.map(sheet => [sheet.sheetId, sheet.name, sheet.visibility])]);
    if (next !== listKey) {
      listKey = next; sheets.replaceChildren(); before.replaceChildren(); option(before, "", "末尾");
      for (const sheet of values) { option(sheets, sheet.sheetId, sheet.name); option(before, sheet.sheetId, sheet.name); }
      if (!selected) selected = host.selection()?.sheetId ?? values[0]?.sheetId ?? "";
      if (selected && !values.some(sheet => sheet.sheetId === selected)) option(sheets, selected, (draft?.name ?? selected) + " · 已移除，草稿保留");
      sheets.value = selected; if (draft) before.value = draft.before;
    }
    if (snapshot && (!draft || !draft.dirty && stale())) { load(true); return; }
    const blocked = !host.idle() || saving || !!composing;
    for (const field of [...fields, ...conflicts.querySelectorAll("input")]) field.disabled = blocked && field !== composing;
    name.disabled ||= actionValue !== "properties" && actionValue !== "copy" && actionValue !== "insert";
    visibility.disabled ||= actionValue !== "properties"; before.disabled ||= actionValue !== "copy" && actionValue !== "insert" && actionValue !== "move";
    save.disabled = blocked || !draft || stale() || actionValue !== "insert" && !original(); reload.disabled = blocked || !draft;
    const dirty = [...drafts.values()].filter(value => value.dirty).length;
    status.textContent = stale() ? "文档已变化。工作表草稿已保留，请重新读取后编辑。" : dirty ? "工作表草稿尚未保存（" + dirty + "）。" : "";
  }
  refresh(); return { element: panel, refresh, hasPendingInput: () => !!composing || saving || [...drafts.values()].some(value => value.dirty) };
}
