import { DocumentError, WorkbookRangeSchema } from "@doffi/core";
import type { WorkbookChangeImpact, WorkbookSnapshot } from "@doffi/core";

export interface WorkbookRowInterval { readonly from: number; readonly to: number }
const record = (value: unknown): value is Record<string, unknown> => value !== null && typeof value === "object" && !Array.isArray(value);
/** These are trusted-provider invalidations, not grants of access. A stale
 * hint or an incompatible snapshot takes the complete measurement path. */
export function workbookRowInvalidation(previous: WorkbookSnapshot, next: WorkbookSnapshot, input: WorkbookChangeImpact | undefined, maxRanges: number): ReadonlyMap<string, readonly WorkbookRowInterval[]> | undefined {
  if (input === undefined) return;
  if (!record(input) || Object.keys(input).some(key => !["documentId", "fromRevision", "revision", "sourceRanges", "calculatedRanges"].includes(key))
    || !Array.isArray(input.sourceRanges) || !Array.isArray(input.calculatedRanges)) throw new DocumentError("InvalidDocument", "Workbook row invalidation has an invalid envelope.");
  if (input.documentId !== previous.manifest.documentId || input.documentId !== next.manifest.documentId || input.fromRevision !== previous.manifest.revision || input.revision !== next.manifest.revision) return;
  if (input.sourceRanges.length + input.calculatedRanges.length > maxRanges) return;
  if (JSON.stringify(previous.metadata) !== JSON.stringify(next.metadata) || previous.content?.html !== next.content?.html
    || previous.styles.length !== next.styles.length || previous.styles.some((style, index) => style.path !== next.styles[index]!.path || style.css !== next.styles[index]!.css)) return;
  const rows = new Map<string, WorkbookRowInterval[]>();
  const sheets = new Set(next.metadata.workbook.sheets.map(sheet => sheet.sheetId));
  for (const ranges of [input.sourceRanges, input.calculatedRanges]) for (const value of ranges) {
    if (!record(value) || Object.keys(value).some(key => !["sheetId", "range"].includes(key)) || typeof value.sheetId !== "string" || !value.sheetId || value.sheetId.length > 256) throw new DocumentError("InvalidDocument", "Workbook row invalidation requires a worksheet identity.");
    if (!sheets.has(value.sheetId)) throw new DocumentError("InvalidDocument", "Workbook row invalidation references an unknown worksheet.");
    const parsed = WorkbookRangeSchema.safeParse(value.range);
    if (!parsed.success || parsed.data.from.row > parsed.data.to.row || parsed.data.from.column > parsed.data.to.column) throw new DocumentError("InvalidDocument", "Workbook row invalidation has an invalid source rectangle.");
    let values = rows.get(value.sheetId); if (!values) { values = []; rows.set(value.sheetId, values); }
    values.push({ from: parsed.data.from.row, to: parsed.data.to.row });
  }
  for (const [sheetId, intervals] of rows) {
    intervals.sort((a, b) => a.from - b.from || a.to - b.to); const merged: WorkbookRowInterval[] = [];
    for (const interval of intervals) {
      const last = merged.at(-1);
      if (last && interval.from <= last.to + 1) merged[merged.length - 1] = { from: last.from, to: Math.max(last.to, interval.to) };
      else merged.push(interval);
    }
    rows.set(sheetId, merged);
  }
  return rows;
}
