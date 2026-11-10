import { DocumentError, parseWorkbookFormula, WorkbookDefinedNameSchema } from "@doffi/core";
import type { DeepReadonly, WorkbookDefinedName } from "@doffi/core";
import type { WorkbookControlsHost, WorkbookControlsIdentity } from "./controls-host.js";

interface NameDraft {
  readonly identity: WorkbookControlsIdentity;
  readonly original?: DeepReadonly<WorkbookDefinedName>;
  name: string; expression: string; origin: string; scope: string; description: string; hidden: boolean; dirty: boolean;
}
const address = (row: number, column: number): string => {
  let letters = ""; for (let value = column + 1; value; value = Math.floor((value - 1) / 26)) letters = String.fromCharCode(65 + (value - 1) % 26) + letters;
  return letters + (row + 1);
};
/** Name forms keep revision-bound drafts, while the owning session resolves
 * bindings, validates names and rewrites consumers in its normal transaction. */
export function createWorkbookNameControls(document: Document, parent: ParentNode, host: WorkbookControlsHost): { readonly element: HTMLDetailsElement; refresh(): void; hasPendingInput(): boolean } {
  const panel = document.createElement("details"), summary = document.createElement("summary"); summary.textContent = "名称管理"; panel.append(summary); parent.append(panel);
  const fields: (HTMLInputElement | HTMLTextAreaElement | HTMLSelectElement)[] = [];
  const label = <T extends HTMLInputElement | HTMLTextAreaElement | HTMLSelectElement>(name: string, field: T): T => { const wrapper = document.createElement("label"); field.setAttribute("aria-label", name); wrapper.append(name, field); panel.append(wrapper); fields.push(field); return field; };
  const choice = label("选择定义名称", document.createElement("select")), name = label("定义名称", document.createElement("input")), scope = label("名称作用域", document.createElement("select"));
  const expression = label("名称引用位置或公式", document.createElement("textarea")), origin = label("名称相对引用基准单元格", document.createElement("input")), description = label("名称说明", document.createElement("input")), hidden = label("隐藏名称", document.createElement("input"));
  name.maxLength = 255; expression.maxLength = 8192; origin.maxLength = 16; description.maxLength = 32767; hidden.type = "checkbox";
  const status = document.createElement("p"); status.setAttribute("role", "status"); status.setAttribute("aria-label", "名称编辑状态"); panel.append(status);
  const drafts = new Map<string, NameDraft>(); let current: NameDraft | undefined, selected = "", key = "", composing: HTMLElement | undefined, saving = false, serial = 0;
  const fail = (message: string): never => { throw new DocumentError("InvalidOperation", message); };
  const source = () => host.snapshot()?.metadata.workbook.names?.find(value => value.nameId === selected);
  const stale = () => !!current && (host.snapshot()?.manifest.documentId !== current.identity.documentId || host.snapshot()?.manifest.revision !== current.identity.revision);
  function load(reset = false): void {
    const snapshot = host.snapshot(); if (!snapshot) return;
    selected = choice.value; let draft = reset ? undefined : drafts.get(selected);
    if (!draft) {
      const original = source(), selection = host.selection(), at = selection?.range.from ?? { row: 0, column: 0 }, sheet = snapshot.metadata.workbook.sheets.find(sheet => sheet.sheetId === selection?.sheetId);
      const cell = (point: typeof at) => "$" + address(0, point.column).slice(0, -1) + "$" + (point.row + 1);
      draft = { identity: { documentId: snapshot.manifest.documentId, revision: snapshot.manifest.revision }, ...(original ? { original } : {}),
        name: original?.name ?? "", scope: original?.sheetId ?? "", expression: original?.expression ?? (sheet && selection ? "='" + sheet.name.replaceAll("'", "''") + "'!" + cell(at) + ":" + cell(selection.range.to) : "=0"),
        origin: address(original?.origin.row ?? at.row, original?.origin.column ?? at.column), description: original?.description ?? "", hidden: original?.hidden ?? false, dirty: false };
      drafts.set(selected, draft);
    }
    current = draft; name.value = draft.name; scope.value = draft.scope; expression.value = draft.expression; origin.value = draft.origin; description.value = draft.description; hidden.checked = draft.hidden;
    refresh();
  }
  function changed(): void {
    if (!current) return; Object.assign(current, { name: name.value, expression: expression.value, origin: origin.value, scope: scope.value, description: description.value, hidden: hidden.checked, dirty: true });
    refresh(); host.draftChanged?.();
  }
  const button = (text: string, action: () => Promise<void>): HTMLButtonElement => {
    const button = document.createElement("button"); button.type = "button"; button.textContent = text; panel.append(button);
    button.addEventListener("click", () => host.run(async () => {
      if (saving || composing || !host.idle()) fail("请先完成当前输入或操作。"); saving = true; refresh();
      try { await action(); } finally { saving = false; refresh(); host.draftChanged?.(); }
    }), { signal: host.signal }); return button;
  };
  const save = button("保存定义名称", async () => {
    if (!current || stale()) throw new DocumentError("RevisionConflict", "名称表单属于旧修订，请重新读取。");
    const snapshot = host.snapshot()!, parsed = parseWorkbookFormula("=" + current.origin, { row: 0, column: 0 });
    if (parsed.kind !== "cell" || parsed.sheet !== undefined) throw new DocumentError("InvalidOperation", "请输入一个基准单元格，例如 A1。");
    const original = current.original; let nameId = original?.nameId;
    if (!nameId) do { nameId = "editor-name-" + ++serial; } while (snapshot.metadata.workbook.names?.some(name => name.nameId === nameId));
    const value = WorkbookDefinedNameSchema.parse({ ...original, nameId, name: current.name, expression: current.expression, origin: { row: parsed.row, column: parsed.column },
      ...(current.scope ? { sheetId: current.scope } : {}), ...(current.hidden || original?.hidden !== undefined ? { hidden: current.hidden } : {}), ...(current.description || original?.description !== undefined ? { description: current.description } : {}) });
    const oldKey = selected; await host.perform([{ op: "setWorkbookDefinedName", name: value }], undefined, current.identity);
    drafts.delete(oldKey); selected = nameId; key = ""; current = undefined; refresh(); choice.value = nameId; load(true);
  });
  const remove = button("删除定义名称", async () => {
    if (!current?.original || stale() || !source()) throw new DocumentError("RevisionConflict", "请重新读取要删除的名称。");
    await host.perform([{ op: "removeWorkbookDefinedName", nameId: current.original.nameId }], undefined, current.identity);
    drafts.delete(selected); selected = ""; key = ""; current = undefined; refresh(); choice.value = ""; load();
  });
  const reload = button("放弃当前名称草稿并重新读取", async () => { drafts.delete(selected); load(true); });
  choice.addEventListener("change", () => { load(); host.draftChanged?.(); }, { signal: host.signal });
  for (const field of fields.filter(field => field !== choice)) field.addEventListener("input", changed, { signal: host.signal });
  for (const field of fields) {
    field.addEventListener("compositionstart", () => { composing = field; changed(); }, { signal: host.signal });
    field.addEventListener("compositionend", () => { composing = undefined; changed(); }, { signal: host.signal });
  }
  function refresh(): void {
    const snapshot = host.snapshot(), names = snapshot?.metadata.workbook.names ?? [], sheets = snapshot?.metadata.workbook.sheets ?? [];
    const next = JSON.stringify([snapshot?.manifest.documentId, names.map(name => [name.nameId, name.name, name.sheetId]), sheets.map(sheet => [sheet.sheetId, sheet.name])]);
    if (key !== next) {
      key = next; choice.replaceChildren(); scope.replaceChildren();
      const option = (select: HTMLSelectElement, value: string, text: string) => { const item = document.createElement("option"); item.value = value; item.textContent = text; select.append(item); };
      option(choice, "", "新建名称"); option(scope, "", "工作簿"); for (const sheet of sheets) option(scope, sheet.sheetId, sheet.name);
      for (const value of names) option(choice, value.nameId, value.name + " · " + (sheets.find(sheet => sheet.sheetId === value.sheetId)?.name ?? "工作簿"));
      if (selected && !names.some(name => name.nameId === selected)) option(choice, selected, (current?.name ?? selected) + " · 已移除，草稿保留");
      choice.value = selected; if (current) scope.value = current.scope;
    }
    if (snapshot && (!current || !current.dirty && stale())) { load(true); return; }
    const idle = host.idle(), blocked = !idle || saving || !!composing;
    for (const field of fields) field.disabled = blocked && field !== composing;
    scope.disabled ||= !!current?.original;
    save.disabled = blocked || !current || stale(); remove.disabled = blocked || !current?.original || stale() || !source(); reload.disabled = blocked || !current;
    const dirty = [...drafts.values()].filter(draft => draft.dirty).length;
    status.textContent = stale() ? "文档已变化。名称草稿已保留，请重新读取后编辑。" : dirty ? "名称草稿尚未保存（" + dirty + "）。" : "";
  }
  refresh(); return { element: panel, refresh, hasPendingInput: () => !!composing || saving || [...drafts.values()].some(draft => draft.dirty) };
}
