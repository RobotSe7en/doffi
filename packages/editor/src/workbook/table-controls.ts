import { DocumentError, WorkbookTableSchema } from "@doffi/core";
import type { WorkbookTable } from "@doffi/core";
import type { WorkbookControlsHost, WorkbookControlsIdentity } from "./controls-host.js";
import { workbookControlRange, workbookControlRangeText } from "./controls-range.js";

interface TableDraft {
  readonly identity: WorkbookControlsIdentity; readonly sheetId: string; readonly existing: boolean;
  table: WorkbookTable; range: string; dirty: boolean;
}
/** The form edits source table definitions; actual headers, formulas, reference
 * rewrites and exceptions remain owned by the shared core transaction. */
export function createWorkbookTableControls(document: Document, parent: ParentNode, host: WorkbookControlsHost): { readonly element: HTMLDetailsElement; refresh(): void; hasPendingInput(): boolean } {
  const panel = document.createElement("details"), summary = document.createElement("summary"); summary.textContent = "智能表管理"; panel.append(summary); parent.append(panel);
  const fields: (HTMLInputElement | HTMLSelectElement)[] = [];
  const field = <T extends HTMLInputElement | HTMLSelectElement>(label: string, input: T): T => { const wrapper = document.createElement("label"); input.setAttribute("aria-label", label); wrapper.append(label, input); panel.append(wrapper); fields.push(input); return input; };
  const text = (label: string, maximum: number) => { const input = field(label, document.createElement("input")); input.maxLength = maximum; return input; };
  const check = (label: string) => { const input = field(label, document.createElement("input")); input.type = "checkbox"; return input; };
  const option = (select: HTMLSelectElement, value: string, label: string) => { const item = document.createElement("option"); item.value = value; item.textContent = label; select.append(item); };
  const choice = field("选择智能表", document.createElement("select")), name = text("智能表名称", 255), objectName = text("智能表对象名称", 255), description = text("智能表说明", 255), range = text("智能表区域", 32);
  const header = check("智能表包含标题行"), totals = check("智能表显示汇总行"), filter = check("智能表显示筛选按钮"), style = field("智能表样式", document.createElement("select"));
  const firstColumn = check("智能表突出首列"), lastColumn = check("智能表突出末列"), rowStripes = check("智能表行条纹"), columnStripes = check("智能表列条纹");
  const column = field("选择智能表列", document.createElement("select")), columnName = text("智能表列名", 255), calculated = text("智能表计算列公式", 8192), totalFunction = field("智能表列汇总方式", document.createElement("select")), totalFormula = text("智能表列自定义汇总公式", 8192), totalLabel = text("智能表列汇总文字", 32767);
  for (const [value, label] of [["none", "无"], ["sum", "求和"], ["average", "平均值"], ["count", "非空计数"], ["countNums", "数值计数"], ["min", "最小值"], ["max", "最大值"], ["stdDev", "样本标准差"], ["var", "样本方差"], ["custom", "自定义公式"]]) option(totalFunction, value!, label!);
  const status = document.createElement("p"); status.setAttribute("role", "status"); status.setAttribute("aria-label", "智能表编辑状态"); panel.append(status);
  let draft: TableDraft | undefined, selected = "", composing: HTMLElement | undefined, saving = false, serial = 0, listKey = "", styleKey = "";
  const tables = () => host.snapshot()?.metadata.workbook.sheets.flatMap(sheet => (sheet.tables ?? []).map(table => ({ sheet, table }))) ?? [];
  const stale = () => !!draft && (host.snapshot()?.manifest.documentId !== draft.identity.documentId || host.snapshot()?.manifest.revision !== draft.identity.revision);
  const fail = (message: string): never => { throw new DocumentError("InvalidOperation", message); };
  function fresh(kind: string): string {
    const used = new Set(tables().flatMap(({ table }) => [table.tableId, ...table.columns.map(column => column.columnId)]));
    let id: string; do { id = "editor-table-" + kind + "-" + ++serial; } while (used.has(id)); return id;
  }
  function loadColumns(): void {
    const previous = column.value; column.replaceChildren();
    for (const value of draft?.table.columns ?? []) option(column, value.columnId, value.name);
    if (draft?.table.columns.some(value => value.columnId === previous)) column.value = previous;
    readColumn();
  }
  function readColumn(): void {
    const value = draft?.table.columns.find(value => value.columnId === column.value);
    columnName.value = value?.name ?? ""; calculated.value = value?.calculatedFormula ?? ""; totalFunction.value = value?.totalFunction ?? "none"; totalFormula.value = value?.totalFormula ?? ""; totalLabel.value = value?.totalLabel ?? "";
  }
  function load(): void {
    const snapshot = host.snapshot(), original = tables().find(({ table }) => table.tableId === selected), selection = host.selection(); draft = undefined;
    const sheetId = original?.sheet.sheetId ?? selection?.sheetId;
    if (!snapshot || !sheetId || !snapshot.metadata.workbook.sheets.some(sheet => sheet.sheetId === sheetId)) return;
    const at = selection?.range ?? { from: { row: 0, column: 0 }, to: { row: 1, column: 0 } };
    const table: WorkbookTable = original ? JSON.parse(JSON.stringify(original.table)) as WorkbookTable : {
      tableId: fresh("id"), name: "", range: { from: { ...at.from }, to: { ...at.to } }, header: true, totals: false, totalsPreviouslyShown: false,
      columns: Array.from({ length: at.to.column - at.from.column + 1 }, (_, index) => ({ columnId: fresh("column"), name: "列" + (index + 1), totalFunction: "none" })),
      autoFilter: { range: { from: { ...at.from }, to: { ...at.to } }, columns: [] },
      rowStripes: true, columnStripes: false, showFirstColumn: false, showLastColumn: false,
    };
    draft = { identity: { documentId: snapshot.manifest.documentId, revision: snapshot.manifest.revision }, sheetId, existing: !!original, table, range: workbookControlRangeText(table.range), dirty: false };
    name.value = table.name; objectName.value = table.objectName ?? ""; description.value = table.description ?? ""; range.value = draft.range;
    header.checked = table.header; totals.checked = table.totals; filter.checked = !!table.autoFilter; firstColumn.checked = table.showFirstColumn; lastColumn.checked = table.showLastColumn; rowStripes.checked = table.rowStripes; columnStripes.checked = table.columnStripes;
    style.value = table.style?.kind === "builtin" ? table.style.name : table.style?.kind === "custom" ? "custom:" + table.style.styleId : ""; loadColumns();
  }
  function changed(): void {
    if (!draft) return; const table = draft.table;
    table.name = name.value; draft.range = range.value; table.header = header.checked; table.totals = totals.checked; table.totalsPreviouslyShown ||= totals.checked;
    for (const [key, input] of [["objectName", objectName], ["description", description]] as const) if (input.value !== (table[key] ?? "")) { if (input.value) table[key] = input.value; else delete table[key]; }
    table.showFirstColumn = firstColumn.checked; table.showLastColumn = lastColumn.checked; table.rowStripes = rowStripes.checked; table.columnStripes = columnStripes.checked;
    if (!style.value) delete table.style;
    else table.style = style.value.startsWith("custom:") ? { kind: "custom", styleId: style.value.slice(7) } : { kind: "builtin", name: style.value };
    const value = table.columns.find(value => value.columnId === column.value);
    if (value) {
      value.name = columnName.value; value.totalFunction = totalFunction.value as typeof value.totalFunction;
      for (const [key, input] of [["calculatedFormula", calculated], ["totalFormula", totalFormula], ["totalLabel", totalLabel]] as const) if (input.value !== (value[key] ?? "")) { if (input.value) value[key] = input.value; else delete value[key]; }
      const current = column.selectedOptions[0]; if (current) current.textContent = value.name;
    }
    draft.dirty = true; refresh(); host.draftChanged?.();
  }
  function resize(): void {
    if (!draft) return; const table = draft.table, next = workbookControlRange(draft.range), old = table.range;
    if (draft.existing && (next.from.row !== old.from.row || next.from.column !== old.from.column)) fail("调整区域时请保留智能表起点；移动整张表请使用剪切。");
    table.columns = Array.from({ length: next.to.column - next.from.column + 1 }, (_, index) => table.columns[index] ?? { columnId: fresh("column"), name: "列" + (index + 1), totalFunction: "none" });
    table.range = next;
    // Retain predicates for surviving columns. A resized table's filter covers
    // its current body; shrinking cannot leave out-of-range sort definitions.
    if (filter.checked) {
      table.autoFilter ??= { range: next, columns: [] };
      table.autoFilter.range = { from: { ...next.from }, to: { row: Math.max(next.from.row, next.to.row - Number(table.totals)), column: next.to.column } };
      table.autoFilter.columns = table.autoFilter.columns.filter(value => value.column >= next.from.column && value.column <= next.to.column);
    } else delete table.autoFilter;
    for (const owner of [table, table.autoFilter]) if (owner?.sort) {
      const sort = owner.sort; sort.keys = sort.keys.filter(key => key.column >= next.from.column && key.column <= next.to.column);
      sort.range.to.row = Math.min(sort.range.to.row, next.to.row); sort.range.to.column = Math.min(sort.range.to.column, next.to.column);
      if (!sort.keys.length || sort.range.from.row > sort.range.to.row || sort.range.from.column > sort.range.to.column) delete owner.sort;
    }
    loadColumns();
  }
  const button = (label: string, action: () => Promise<void>) => {
    const button = document.createElement("button"); button.type = "button"; button.textContent = label; panel.append(button);
    button.addEventListener("click", () => host.run(async () => {
      if (!host.idle() || composing || saving) fail("请先完成当前输入或操作。"); saving = true;
      try { const submitted = action(); refresh(); host.draftChanged?.(); await submitted; } finally { saving = false; refresh(); host.draftChanged?.(); }
    }), { signal: host.signal }); return button;
  };
  const adjust = button("按智能表区域调整列", async () => { if (!draft) return; resize(); draft.dirty = true; refresh(); });
  const save = button("保存智能表", async () => {
    if (!draft || stale()) throw new DocumentError("RevisionConflict", "智能表草稿属于旧修订，请重新读取。");
    const submitted = draft; resize(); const table = WorkbookTableSchema.parse(submitted.table);
    await host.perform([{ op: "setWorkbookTable", sheetId: submitted.sheetId, table }], undefined, submitted.identity); selected = table.tableId; draft = undefined; listKey = "";
  });
  const remove = button("智能表转为普通区域", async () => {
    if (!draft?.existing || stale()) throw new DocumentError("RevisionConflict", "请重新读取要转换的智能表。");
    await host.perform([{ op: "removeWorkbookTable", sheetId: draft.sheetId, tableId: draft.table.tableId }], undefined, draft.identity); selected = ""; draft = undefined; listKey = "";
  });
  const reload = button("放弃智能表草稿并重新读取", async () => { selected = tables().some(({ table }) => table.tableId === selected) ? selected : ""; draft = undefined; listKey = ""; refresh(); load(); });
  choice.addEventListener("change", () => { if (draft?.dirty || composing || saving) { choice.value = selected; return; } selected = choice.value; draft = undefined; load(); refresh(); host.draftChanged?.(); }, { signal: host.signal });
  column.addEventListener("change", readColumn, { signal: host.signal });
  for (const input of fields.filter(input => input !== choice && input !== column)) {
    input.addEventListener("input", changed, { signal: host.signal });
    input.addEventListener("compositionstart", () => { composing = input; changed(); }, { signal: host.signal });
    input.addEventListener("compositionend", () => { composing = undefined; changed(); }, { signal: host.signal });
  }
  function refresh(): void {
    const snapshot = host.snapshot(), all = tables(), styles = snapshot?.metadata.workbook.tableStyles ?? [];
    const nextStyles = JSON.stringify(styles.map(style => [style.styleId, style.name]));
    if (styleKey !== nextStyles) {
      const previous = style.value; styleKey = nextStyles; style.replaceChildren(); option(style, "", "无表样式");
      for (const [family, count] of [["Light", 21], ["Medium", 28], ["Dark", 11]] as const) for (let index = 1; index <= count; index++) option(style, "TableStyle" + family + index, "TableStyle" + family + index);
      for (const value of styles) option(style, "custom:" + value.styleId, value.name); style.value = previous;
    }
    const next = JSON.stringify([snapshot?.manifest.documentId, all.map(({ sheet, table }) => [sheet.sheetId, sheet.name, table.tableId, table.name])]);
    if (next !== listKey) {
      listKey = next; choice.replaceChildren(); option(choice, "", "从选区创建智能表");
      for (const { sheet, table } of all) option(choice, table.tableId, table.name + " · " + sheet.name);
      if (selected && !all.some(({ table }) => table.tableId === selected)) option(choice, selected, "已移除的智能表 · 草稿保留"); choice.value = selected;
    }
    const selection = host.selection(), followsSelection = draft && !draft.existing && selection && (draft.sheetId !== selection.sheetId || draft.range !== workbookControlRangeText(selection.range));
    if (!saving && !composing && (!draft || !draft.dirty && (stale() || followsSelection))) load();
    const blocked = !host.idle() || saving || !!composing, old = stale();
    for (const input of fields) input.disabled = (blocked || !draft) && input !== composing;
    choice.disabled = blocked || !!draft?.dirty; adjust.disabled = save.disabled = blocked || !draft || old;
    remove.disabled = save.disabled || !draft?.existing; reload.disabled = blocked;
    const sheet = snapshot?.metadata.workbook.sheets.find(sheet => sheet.sheetId === draft?.sheetId);
    status.textContent = old ? "文档已变化。智能表草稿已保留，请重新读取后编辑。" : !draft ? "请选择工作表区域。" : (sheet?.name ?? draft.sheetId) + " · " + draft.range + (draft.dirty ? " · 智能表草稿尚未保存。" : "");
  }
  refresh(); return { element: panel, refresh, hasPendingInput: () => !!draft?.dirty || !!composing || saving };
}
