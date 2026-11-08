import { DocumentError, WorkbookAutoFilterSchema } from "@doffi/core";
import type { WorkbookAutoFilter, WorkbookTable } from "@doffi/core";
import type { WorkbookControlsHost, WorkbookControlsIdentity } from "./controls-host.js";
import { workbookControlAddress, workbookControlRange, workbookControlRangeText } from "./controls-range.js";
import { createWorkbookFilterCriterionControls } from "./filter-criterion-controls.js";

interface FilterDraft {
  readonly identity: WorkbookControlsIdentity; readonly sheetId: string; readonly table?: WorkbookTable;
  filter: WorkbookAutoFilter; range: string; column: number; dirty: boolean; columnDirty: boolean;
}
/** A sheet filter and a table filter use their existing source operations. The
 * form never materializes hidden rows or stores a second filtered cell set. */
export function createWorkbookFilterControls(document: Document, parent: ParentNode, host: WorkbookControlsHost): { readonly element: HTMLDetailsElement; refresh(): void; hasPendingInput(): boolean } {
  const panel = document.createElement("details"), summary = document.createElement("summary"); summary.textContent = "筛选管理"; panel.append(summary); parent.append(panel);
  const fields: (HTMLInputElement | HTMLSelectElement)[] = [];
  const field = <T extends HTMLInputElement | HTMLSelectElement>(label: string, input: T): T => { const wrapper = document.createElement("label"); input.setAttribute("aria-label", label); wrapper.append(label, input); panel.append(wrapper); fields.push(input); return input; };
  const choice = field("筛选所属区域", document.createElement("select")), range = field("筛选区域", document.createElement("input")), column = field("筛选列", document.createElement("input")), hiddenButton = field("隐藏此列筛选按钮", document.createElement("input")), showButton = field("显示此列筛选按钮", document.createElement("input"));
  hiddenButton.type = showButton.type = "checkbox"; range.maxLength = 32; column.maxLength = 4; column.placeholder = "例如 A";
  const status = document.createElement("p"); status.setAttribute("role", "status"); status.setAttribute("aria-label", "筛选编辑状态"); panel.append(status);
  let draft: FilterDraft | undefined, selected = "", listKey = "", composing: HTMLElement | undefined, saving = false;
  const tables = () => host.snapshot()?.metadata.workbook.sheets.flatMap(sheet => (sheet.tables ?? []).map(table => ({ sheet, table }))) ?? [];
  const stale = () => !!draft && (host.snapshot()?.manifest.documentId !== draft.identity.documentId || host.snapshot()?.manifest.revision !== draft.identity.revision);
  function changed(): void { if (!draft) return; draft.dirty = true; refresh(); host.draftChanged?.(); }
  const criteria = createWorkbookFilterCriterionControls(document, panel, host.signal, () => { if (draft) draft.columnDirty = true; changed(); });
  function stash(): void {
    if (!draft?.columnDirty) return;
    const value = criteria.value(), original = draft.filter.columns.find(value => value.column === draft!.column);
    if (original || value !== undefined || hiddenButton.checked || !showButton.checked) {
      const target = original ?? { column: draft.column, hiddenButton: false, showButton: true };
      target.hiddenButton = hiddenButton.checked; target.showButton = showButton.checked;
      if (value === undefined) delete target.criterion; else target.criterion = value;
      if (!original) draft.filter.columns.push(target);
    }
    draft.columnDirty = false;
  }
  function readColumn(): void {
    const value = draft?.filter.columns.find(value => value.column === draft!.column);
    column.value = workbookControlAddress({ row: 0, column: draft?.column ?? 0 }).slice(0, -1); hiddenButton.checked = value?.hiddenButton ?? false; showButton.checked = value?.showButton ?? true;
    criteria.load(value?.criterion); if (draft) draft.columnDirty = false;
  }
  function load(): void {
    const snapshot = host.snapshot(), owner = tables().find(({ table }) => table.tableId === selected), selection = host.selection(); draft = undefined;
    const sheet = owner?.sheet ?? snapshot?.metadata.workbook.sheets.find(sheet => sheet.sheetId === selection?.sheetId);
    if (!snapshot || !sheet) return;
    const original = owner?.table.autoFilter ?? (!owner ? sheet.autoFilter : undefined), bounds = owner ? { from: { ...owner.table.range.from }, to: { row: Math.max(owner.table.range.from.row, owner.table.range.to.row - Number(owner.table.totals)), column: owner.table.range.to.column } } : selection?.range ?? { from: { row: 0, column: 0 }, to: { row: 1, column: 0 } };
    const filter = original ? JSON.parse(JSON.stringify(original)) as WorkbookAutoFilter : { range: { from: { ...bounds.from }, to: { ...bounds.to } }, columns: [] };
    draft = { identity: { documentId: snapshot.manifest.documentId, revision: snapshot.manifest.revision }, sheetId: sheet.sheetId, ...(owner ? { table: JSON.parse(JSON.stringify(owner.table)) as WorkbookTable } : {}),
      filter, range: workbookControlRangeText(filter.range), column: filter.range.from.column, dirty: false, columnDirty: false };
    range.value = draft.range; readColumn();
  }
  function resize(clear = false): WorkbookAutoFilter {
    stash(); const current = draft!; current.filter.range = workbookControlRange(current.range);
    current.filter.columns = current.filter.columns.filter(column => column.column >= current.filter.range.from.column && column.column <= current.filter.range.to.column);
    if (clear) for (const column of current.filter.columns) delete column.criterion;
    // A changed range must not silently clip an existing sort. Its source is
    // preserved and the shared definition validator checks compatibility.
    return WorkbookAutoFilterSchema.parse(current.filter);
  }
  async function perform(filter: WorkbookAutoFilter | undefined): Promise<void> {
    const current = draft!;
    if (current.table) { const table = { ...current.table }; if (filter) table.autoFilter = filter; else delete table.autoFilter; await host.perform([{ op: "setWorkbookTable", sheetId: current.sheetId, table }], undefined, current.identity); }
    else await host.perform([{ op: "setWorkbookAutoFilter", sheetId: current.sheetId, filter: filter ?? null }], undefined, current.identity);
    draft = undefined;
  }
  const button = (label: string, action: () => Promise<void>, requiresCurrent = true) => {
    const button = document.createElement("button"); button.type = "button"; button.textContent = label; panel.append(button);
    button.addEventListener("click", () => host.run(async () => {
      if (!host.idle() || saving || composing) throw new DocumentError("InvalidOperation", "请先完成当前输入或操作。");
      if (requiresCurrent && (!draft || stale())) throw new DocumentError("RevisionConflict", "筛选草稿属于旧修订，请重新读取。");
      saving = true; try { const submitted = action(); refresh(); host.draftChanged?.(); await submitted; } finally { saving = false; refresh(); host.draftChanged?.(); }
    }), { signal: host.signal }); return button;
  };
  const chooseColumn = button("读取筛选列", async () => {
    const at = workbookControlRange(column.value.trim() + "1"), bounds = workbookControlRange(draft!.range);
    if (at.from.row !== 0 || at.to.row !== 0 || at.from.column !== at.to.column || at.from.column < bounds.from.column || at.from.column > bounds.to.column) throw new DocumentError("InvalidOperation", "请输入筛选区域中的列名，例如 A。");
    stash(); draft!.column = at.from.column; readColumn();
  });
  const save = button("应用筛选条件", async () => { await perform(resize()); });
  const clear = button("清除所有筛选条件", async () => { await perform(resize(true)); });
  const remove = button("移除区域筛选", async () => { await perform(undefined); });
  const reload = button("放弃筛选草稿并重新读取", async () => { if (selected && !tables().some(({ table }) => table.tableId === selected)) selected = ""; draft = undefined; listKey = ""; refresh(); load(); }, false);
  choice.addEventListener("change", () => { if (draft?.dirty || saving || composing) { choice.value = selected; return; } selected = choice.value; draft = undefined; load(); refresh(); }, { signal: host.signal });
  range.addEventListener("input", () => { if (draft) draft.range = range.value; changed(); }, { signal: host.signal });
  for (const input of [hiddenButton, showButton]) input.addEventListener("input", () => { if (draft) draft.columnDirty = true; changed(); }, { signal: host.signal });
  panel.addEventListener("compositionstart", event => { if (!(event.target instanceof document.defaultView!.HTMLElement)) return; composing = event.target; if (criteria.element.contains(composing) || composing === hiddenButton || composing === showButton) { if (draft) draft.columnDirty = true; } changed(); }, { capture: true, signal: host.signal });
  panel.addEventListener("compositionend", () => { composing = undefined; changed(); }, { capture: true, signal: host.signal });
  function refresh(): void {
    const snapshot = host.snapshot(), all = tables(), selection = host.selection(), next = JSON.stringify([snapshot?.manifest.documentId, all.map(({ sheet, table }) => [sheet.sheetId, sheet.name, table.tableId, table.name])]);
    if (next !== listKey) {
      listKey = next; choice.replaceChildren(); const option = (value: string, text: string) => { const option = document.createElement("option"); option.value = value; option.textContent = text; choice.append(option); };
      option("", "普通区域筛选"); for (const { sheet, table } of all) option(table.tableId, table.name + " · " + sheet.name);
      if (selected && !all.some(({ table }) => table.tableId === selected)) option(selected, "已移除的智能表 · 草稿保留"); choice.value = selected;
    }
    const followsSelection = draft && !draft.table && selection && (draft.sheetId !== selection.sheetId || !snapshot?.metadata.workbook.sheets.find(sheet => sheet.sheetId === draft!.sheetId)?.autoFilter && draft.range !== workbookControlRangeText(selection.range));
    if (!saving && !composing && (!draft || !draft.dirty && (stale() || followsSelection))) load();
    const blocked = !host.idle() || saving || !!composing, old = stale();
    for (const field of fields) field.disabled = (blocked || !draft) && field !== composing;
    choice.disabled = blocked || !!draft?.dirty; range.disabled ||= !!draft?.table; criteria.refresh(blocked || !draft, composing);
    chooseColumn.disabled = save.disabled = clear.disabled = remove.disabled = blocked || !draft || old; reload.disabled = blocked;
    const sheet = snapshot?.metadata.workbook.sheets.find(sheet => sheet.sheetId === draft?.sheetId);
    status.textContent = old ? "文档已变化。筛选草稿已保留，请重新读取后编辑。" : !draft ? "请选择工作表区域。" : (sheet?.name ?? draft.sheetId) + " · " + (draft.table?.name ?? draft.range) + (draft.dirty ? " · 筛选草稿尚未保存。" : "");
  }
  refresh(); return { element: panel, refresh, hasPendingInput: () => !!draft?.dirty || saving || !!composing };
}
