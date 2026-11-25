import { DocumentError } from "@doffi/core";
import type { DocumentOperation, DocumentSelection, WorkbookSnapshot } from "@doffi/core";
import type { EditorCommandRegistry } from "../commands.js";

/** Ribbon commands use the same range transactions as the format dialog. */
export function registerWorkbookAppearanceCommands(registry: EditorCommandRegistry, host: {
  snapshot(): WorkbookSnapshot;
  selection(): DocumentSelection | undefined;
  style(): Readonly<Record<string, string>> | undefined;
  enabled(): boolean;
  run(operations: DocumentOperation[]): Promise<unknown>;
}): void {
  const rangeOperation = (properties: Record<string, string>): Extract<DocumentOperation, { op: "formatWorkbookRanges" }> => {
    const selected = host.selection(), snapshot = host.snapshot();
    if (!selected || selected.anchor.kind !== "cells") throw new DocumentError("InvalidOperation", "请先选择单元格。");
    if (selected.documentId !== snapshot.manifest.documentId || selected.baseRevision !== snapshot.manifest.revision) throw new DocumentError("RevisionConflict", "格式选区已过期。");
    return { op: "formatWorkbookRanges", sheetId: selected.anchor.sheetId, ranges: selected.anchor.ranges.map(range => ({ from: { ...range.from }, to: { ...range.to } })), properties };
  };
  for (const [id, label, value] of [["verticalTop", "顶端对齐", "top"], ["verticalCenter", "垂直居中", "center"], ["verticalBottom", "底端对齐", "bottom"]])
    registry.register(id!, label!, () => host.run([rangeOperation({ "--office-cell-vertical": value! })]), host.enabled, () => (host.style()?.["--office-cell-vertical"] || "bottom") === value);
  registry.register("wrapText", "自动换行", () => host.run([rangeOperation({ "--office-cell-wrap": String(!["true", "1"].includes(host.style()?.["--office-cell-wrap"] ?? "")) })]), host.enabled,
    () => ["true", "1"].includes(host.style()?.["--office-cell-wrap"] ?? ""));
  for (const [id, label, direction] of [["decreaseCellIndent", "减少缩进量", -1], ["increaseCellIndent", "增加缩进量", 1]] as const)
    registry.register(id, label, () => host.run([rangeOperation({ "--office-cell-indent": String(Math.max(0, Math.min(250, (Number(host.style()?.["--office-cell-indent"]) || 0) + direction))) })]), host.enabled);
  const formats = [["numberComma", "千位分隔样式", "#,##0.00"], ["numberGeneral", "常规", "General"], ["numberNumber", "数值", "0.00"], ["numberCurrency", "货币", "¥#,##0.00"],
    ["numberAccounting", "会计专用", '_ ¥* #,##0.00_ ;_ ¥* -#,##0.00_ ;_ ¥* "-"??_ ;_ @_ '], ["numberShortDate", "短日期", "yyyy/m/d"],
    ["numberLongDate", "长日期", "yyyy年m月d日"], ["numberTime", "时间", "h:mm:ss"], ["numberPercent", "百分比", "0.00%"], ["numberFraction", "分数", "# ?/?"], ["numberScientific", "科学记数", "0.00E+00"], ["numberText", "文本", "@"]];
  for (const [id, label, code] of formats) registry.register(id!, label!, () => {
    const operations: DocumentOperation[] = [], source = host.snapshot();
    const current = source.metadata.workbook.numberFormats.find(format => "code" in format && format.code === code);
    let formatId = current?.formatId;
    if (!formatId) {
      let serial = 1; const ids = new Set(source.metadata.workbook.numberFormats.map(format => format.formatId));
      do { formatId = "office-ribbon-format-" + serial++; } while (ids.has(formatId));
      operations.push({ op: "setWorkbookNumberFormat", format: { formatId, code: code! } });
    }
    operations.push({ ...rangeOperation({}), numberFormatId: formatId });
    return host.run(operations);
  }, host.enabled);
}
