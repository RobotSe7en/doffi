import { DocumentError } from "@doffi/core";
import type { DocumentSelection, WorkbookRange, WorkbookSnapshot } from "@doffi/core";
export interface WorkbookClipboardSource {
  readonly documentId: string; readonly baseRevision: number; readonly sheetId: string; readonly range: WorkbookRange; readonly kind: "copy" | "cut";
}
/** An editor-local reference to its own authorized session. It is never put on
 * the system clipboard, accepted from another editor, or resolved by document
 * IDs. A cut does not remove source data before the destination is committed. */
export class WorkbookClipboard {
  private source: WorkbookClipboardSource | undefined;
  get available(): boolean { return this.source !== undefined; }
  get kind(): "copy" | "cut" | undefined { return this.source?.kind; }
  capture(snapshot: WorkbookSnapshot, selection: DocumentSelection | undefined, kind: "copy" | "cut"): void {
    const target = this.selection(snapshot, selection);
    this.source = Object.freeze({ documentId: snapshot.manifest.documentId, baseRevision: snapshot.manifest.revision, ...target, kind });
  }
  plan(snapshot: WorkbookSnapshot, selection: DocumentSelection | undefined): { source: WorkbookClipboardSource; sheetId: string; target: WorkbookRange["from"] } {
    const source = this.source;
    if (!source) throw new DocumentError("InvalidOperation", "请先复制或剪切一个区域。");
    if (source.documentId !== snapshot.manifest.documentId || source.baseRevision !== snapshot.manifest.revision) throw new DocumentError("RevisionConflict", "复制或剪切来源已变化，请重新选择来源。");
    const target = this.selection(snapshot, selection);
    return { source, sheetId: target.sheetId, target: target.range.from };
  }
  applied(revision: number, targetSheetId: string, at: WorkbookRange["from"], mode: string): void {
    const source = this.source; if (!source) return;
    const range = source.range, target = { from: at, to: { row: at.row + range.to.row - range.from.row, column: at.column + range.to.column - range.from.column } };
    const overlaps = targetSheetId === source.sheetId && range.from.row <= target.to.row && target.from.row <= range.to.row && range.from.column <= target.to.column && target.from.column <= range.to.column;
    // The copy command only writes its destination. Its own nonoverlapping
    // content/formula/format paste proves that the source spelling is retained.
    // Value snapshots and any external revision require a fresh explicit copy.
    if (source.kind === "cut" || overlaps || mode === "values") this.source = undefined;
    else this.source = Object.freeze({ ...source, baseRevision: revision });
  }
  private selection(snapshot: WorkbookSnapshot, selection: DocumentSelection | undefined): { sheetId: string; range: WorkbookRange } {
    const anchor = selection?.anchor;
    if (!selection || anchor?.kind !== "cells" || anchor.ranges.length !== 1) throw new DocumentError("InvalidOperation", "请选择一个单元格区域。");
    if (selection.documentId !== snapshot.manifest.documentId || selection.baseRevision !== snapshot.manifest.revision) throw new DocumentError("RevisionConflict", "Clipboard selection belongs to a previous source.");
    if (!snapshot.metadata.workbook.sheets.some(sheet => sheet.sheetId === anchor.sheetId)) throw new DocumentError("InvalidOperation", "剪贴区域所在工作表不存在。");
    return { sheetId: anchor.sheetId, range: Object.freeze({ from: Object.freeze({ ...anchor.ranges[0]!.from }), to: Object.freeze({ ...anchor.ranges[0]!.to }) }) };
  }
}
