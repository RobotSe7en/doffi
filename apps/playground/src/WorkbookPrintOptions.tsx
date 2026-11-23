import type { WorkbookSnapshot } from "@doffi/core";
import type { WorkbookRemoteEditorView } from "@doffi/editor";
import type { WorkbookPrintSettings } from "@doffi/render/contracts";

export type WorkbookSelectedCells = Extract<NonNullable<ReturnType<WorkbookRemoteEditorView["selection"]>>["anchor"], { kind: "cells" }>;

export function WorkbookPrintOptions({ format, snapshot, selection, onSubmit, onClose }: {
  format: "pdf" | "png";
  snapshot: WorkbookSnapshot;
  selection: WorkbookSelectedCells | undefined;
  onSubmit(print: WorkbookPrintSettings, pageIndex?: number): void;
  onClose(): void;
}) {
  const sheets = snapshot.metadata.workbook.sheets.filter(sheet => sheet.visibility === "visible");
  return (
    <form className="export-report workbook-print-options" aria-label="工作簿打印设置" onSubmit={event => {
      event.preventDefault();
      const data = new FormData(event.currentTarget), scope = String(data.get("scope"));
      const targets = scope === "selection" && selection ? [{ sheetId: selection.sheetId, areas: selection.ranges }]
        : scope.startsWith("sheet:") ? [{ sheetId: scope.slice(6) }] : undefined;
      onSubmit({
        ...(targets ? { sheets: targets } : {}),
        formulaValues: data.has("cached") ? "cached" : "current",
      }, format === "png" ? Number(data.get("page")) - 1 : undefined);
    }}>
      <label>打印范围
        <select name="scope" defaultValue="all">
          <option value="all">全部可见工作表</option>
          {sheets.map(sheet => <option key={sheet.sheetId} value={"sheet:" + sheet.sheetId}>{sheet.name}</option>)}
          {selection && <option value="selection">当前选区</option>}
        </select>
      </label>
      {format === "png" && <label>打印页码 <input name="page" type="number" min={1} max={256} defaultValue={1} required /></label>}
      <label><input name="cached" type="checkbox" /> 使用已保存的公式值</label>
      <button type="submit">生成 {format.toUpperCase()}</button>
      <button type="button" onClick={onClose}>关闭打印设置</button>
    </form>
  );
}
