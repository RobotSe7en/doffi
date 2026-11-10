import { DocumentError, parseContract } from "@doffi/core";
import type { DocumentSelection } from "@doffi/core";
export type WorkbookInputSelection = DocumentSelection & { readonly anchor: Extract<DocumentSelection["anchor"], { kind: "cells" }> };
/** IDs guard revision consistency; the actual source owner checks membership. */
export function copyWorkbookInputSelection(input: DocumentSelection, identity: { readonly documentId: string; readonly revision: number }): WorkbookInputSelection {
  const invalid = (): never => { throw new DocumentError("InvalidOperation", "Text input requires one ordered worksheet rectangle."); };
  if (!input || typeof input !== "object" || Array.isArray(input) || Object.keys(input).some(key => !["documentId", "baseRevision", "anchor"].includes(key))
    || !input.anchor || input.anchor.kind !== "cells" || !Array.isArray(input.anchor.ranges) || input.anchor.ranges.length !== 1) invalid();
  const anchor = parseContract("anchor", input.anchor); if (anchor.kind !== "cells") return invalid();
  const range = anchor.ranges[0]!; if (range.from.row > range.to.row || range.from.column > range.to.column) invalid();
  if (input.documentId !== identity.documentId || input.baseRevision !== identity.revision) throw new DocumentError("RevisionConflict", "Input selection belongs to another workbook revision.");
  return Object.freeze({ documentId: input.documentId, baseRevision: input.baseRevision,
    anchor: Object.freeze({ kind: "cells", sheetId: anchor.sheetId, ranges: Object.freeze([Object.freeze({ from: Object.freeze({ ...range.from }), to: Object.freeze({ ...range.to }) })]) }) });
}
