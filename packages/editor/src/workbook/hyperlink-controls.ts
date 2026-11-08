import { DocumentError } from "@doffi/core";
import type { DeepReadonly, WorkbookHyperlink, WorkbookRange } from "@doffi/core";
import type { WorkbookControlsHost, WorkbookControlsIdentity } from "./controls-host.js";

interface LinkDraft {
  readonly identity: WorkbookControlsIdentity; readonly sheetId: string; readonly range: WorkbookRange;
  readonly original?: DeepReadonly<WorkbookHyperlink>;
  url: string; location: string; tooltip: string; dirty: boolean;
}
const overlaps = (left: DeepReadonly<WorkbookRange>, right: DeepReadonly<WorkbookRange>) => left.from.row <= right.to.row && left.to.row >= right.from.row && left.from.column <= right.to.column && left.to.column >= right.from.column;
const address = (point: DeepReadonly<WorkbookRange>["from"]) => { let column = point.column + 1, text = ""; while (column) { column--; text = String.fromCharCode(65 + column % 26) + text; column = Math.floor(column / 26); } return text + (point.row + 1); };
/** A hyperlink form owns a selected source range and revision. Navigation can
 * leave its draft visible, but cannot silently change the write destination. */
export function createWorkbookHyperlinkControls(document: Document, parent: ParentNode, host: WorkbookControlsHost): { readonly element: HTMLDetailsElement; refresh(): void; hasPendingInput(): boolean } {
  const panel = document.createElement("details"), summary = document.createElement("summary"); summary.textContent = "单元格超链接"; panel.append(summary); parent.append(panel);
  const fields: HTMLInputElement[] = [];
  const field = (label: string, maximum: number) => { const wrapper = document.createElement("label"), input = document.createElement("input"); input.maxLength = maximum; input.setAttribute("aria-label", label); wrapper.append(label, input); panel.append(wrapper); fields.push(input); return input; };
  const url = field("超链接外部地址", 8192), location = field("超链接工作簿内位置", 8192), tooltip = field("超链接屏幕提示", 32767);
  url.placeholder = "https://example.com 或 mailto: 地址"; location.placeholder = "例如 数据!A1 或定义名称";
  const status = document.createElement("p"); status.setAttribute("role", "status"); status.setAttribute("aria-label", "超链接编辑状态"); panel.append(status);
  let draft: LinkDraft | undefined, composing: HTMLInputElement | undefined, saving = false;
  const selectionKey = (value: { readonly sheetId: string; readonly range: DeepReadonly<WorkbookRange> } | undefined) => value ? JSON.stringify([value.sheetId, value.range.from.row, value.range.from.column, value.range.to.row, value.range.to.column]) : "";
  const stale = () => !!draft && (host.snapshot()?.manifest.documentId !== draft.identity.documentId || host.snapshot()?.manifest.revision !== draft.identity.revision);
  const links = () => draft ? host.snapshot()?.metadata.workbook.sheets.find(sheet => sheet.sheetId === draft!.sheetId)?.hyperlinks?.filter(link => overlaps(link.range, draft!.range)) ?? [] : [];
  function load(): void {
    const snapshot = host.snapshot(), selection = host.selection(); draft = undefined;
    if (snapshot && selection && snapshot.metadata.workbook.sheets.some(sheet => sheet.sheetId === selection.sheetId)) {
      const sheet = snapshot.metadata.workbook.sheets.find(sheet => sheet.sheetId === selection.sheetId)!, matches = sheet.hyperlinks?.filter(link => overlaps(link.range, selection.range)) ?? [];
      const identity = (link: DeepReadonly<WorkbookHyperlink>) => JSON.stringify([link.url, link.location, link.tooltip, link.display]);
      const original = matches.length && matches.every(link => identity(link) === identity(matches[0]!)) ? matches[0] : undefined;
      draft = { identity: { documentId: snapshot.manifest.documentId, revision: snapshot.manifest.revision }, sheetId: selection.sheetId, range: { from: { ...selection.range.from }, to: { ...selection.range.to } },
        ...(original ? { original } : {}), url: original?.url ?? "", location: original?.location ?? "", tooltip: original?.tooltip ?? "", dirty: false };
    }
    url.value = draft?.url ?? ""; location.value = draft?.location ?? ""; tooltip.value = draft?.tooltip ?? "";
  }
  function changed(): void { if (!draft) return; Object.assign(draft, { url: url.value, location: location.value, tooltip: tooltip.value, dirty: true }); refresh(); host.draftChanged?.(); }
  const button = (text: string, callback: () => Promise<void>) => {
    const element = document.createElement("button"); element.type = "button"; element.textContent = text; panel.append(element);
    element.addEventListener("click", () => host.run(async () => {
      if (!host.idle() || composing || saving || !draft) throw new DocumentError("InvalidOperation", "请先选择区域并完成当前输入或操作。");
      if (stale()) throw new DocumentError("RevisionConflict", "超链接草稿属于旧修订，请重新读取。");
      saving = true;
      try { const submitted = callback(); refresh(); host.draftChanged?.(); await submitted; } finally { saving = false; refresh(); host.draftChanged?.(); }
    }), { signal: host.signal }); return element;
  };
  const save = button("设置选区超链接", async () => {
    const submitted = draft!; if (!submitted.url && !submitted.location) throw new DocumentError("InvalidOperation", "请填写外部地址或工作簿内位置。");
    await host.perform([{ op: "setWorkbookHyperlink", sheetId: submitted.sheetId, range: submitted.range, target: {
      ...(submitted.url ? { url: submitted.url } : {}), ...(submitted.location ? { location: submitted.location } : {}),
      ...(submitted.tooltip || submitted.original?.tooltip !== undefined ? { tooltip: submitted.tooltip } : {}),
      ...(submitted.original?.display !== undefined ? { display: submitted.original.display } : {}),
    } }], undefined, submitted.identity); draft = undefined;
  });
  const remove = button("移除选区超链接", async () => { const submitted = draft!; await host.perform([{ op: "setWorkbookHyperlink", sheetId: submitted.sheetId, range: submitted.range, target: null }], undefined, submitted.identity); draft = undefined; });
  const reload = document.createElement("button"); reload.type = "button"; reload.textContent = "放弃超链接草稿并读取选区"; panel.append(reload);
  reload.addEventListener("click", () => { if (!host.idle() || composing || saving) return; load(); refresh(); host.draftChanged?.(); }, { signal: host.signal });
  for (const input of fields) {
    input.addEventListener("input", changed, { signal: host.signal });
    input.addEventListener("compositionstart", () => { composing = input; changed(); }, { signal: host.signal });
    input.addEventListener("compositionend", () => { composing = undefined; changed(); }, { signal: host.signal });
  }
  function refresh(): void {
    if (!draft?.dirty && !composing && !saving && (stale() || selectionKey(draft) !== selectionKey(host.selection()) || !draft && !!host.selection())) load();
    const blocked = !host.idle() || !!composing || saving, old = stale(), present = links(), moved = !!draft && selectionKey(draft) !== selectionKey(host.selection());
    for (const input of fields) input.disabled = (blocked || !draft) && input !== composing;
    save.disabled = blocked || !draft || old; remove.disabled = save.disabled || !present.length; reload.disabled = blocked || !host.selection();
    const sheet = host.snapshot()?.metadata.workbook.sheets.find(sheet => sheet.sheetId === draft?.sheetId), target = draft ? (sheet?.name ?? draft.sheetId) + "!" + address(draft.range.from) + ":" + address(draft.range.to) : "请选择单元格区域。";
    status.textContent = old ? "文档已变化。超链接草稿已保留，请重新读取后编辑。" : target + (moved ? " · 草稿仍对应此原区域，当前选区已变化。" : draft?.dirty ? " · 超链接草稿尚未保存。" : present.length && !draft?.original ? " · 选区包含不同的超链接。" : "");
  }
  refresh(); return { element: panel, refresh, hasPendingInput: () => !!draft?.dirty || !!composing || saving };
}
