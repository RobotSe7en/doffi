import type { WorkbookLayoutMeasurer } from '@doffi/core';

/** Test-only point metrics for the final host-ownership write after the editor
 * and its real font measurer have been released. The accepted-source audit and
 * all actual canvas/font assertions happen before this independent probe; its
 * geometry is deliberately not native or font-layout acceptance evidence. */
export const hostOwnershipProbeLayout: WorkbookLayoutMeasurer = request => {
  const snapshot = request.model.snapshot(), anchors = [...snapshot.metadata.workbook.objects?.map(object => object.anchor) ?? [],
    ...snapshot.metadata.workbook.comments?.flatMap(comment => comment.format === 'classic' ? [comment.box?.anchor, comment.commentPrBox?.anchor].filter(anchor => anchor !== undefined) : []) ?? []];
  // Accepted markers may have offsets larger than these arbitrary default
  // cells. Give each probe cell enough room to keep normalized forward-marker
  // boxes positive; these dimensions still make no font/geometry claim.
  let row = 15, column = 45;
  for (const anchor of anchors) if (anchor.kind !== 'absolute') {
    row = Math.max(row, anchor.from.offsetYPt + 1, anchor.kind === 'two-cell' ? anchor.to.offsetYPt + 1 : 0);
    column = Math.max(column, anchor.from.offsetXPt + 1, anchor.kind === 'two-cell' ? anchor.to.offsetXPt + 1 : 0);
    if (anchor.kind === 'two-cell') {
      if (anchor.to.row > anchor.from.row) row = Math.max(row, (anchor.from.offsetYPt - anchor.to.offsetYPt) / (anchor.to.row - anchor.from.row) + 1);
      if (anchor.to.column > anchor.from.column) column = Math.max(column, (anchor.from.offsetXPt - anchor.to.offsetXPt) / (anchor.to.column - anchor.from.column) + 1);
    }
  }
  return { sourceDigest: request.sourceDigest, fontEnvironment: 'host-ownership-probe-synthetic-points', sheets: request.sheetIds.map(sheetId => ({ sheetId,
    rows: { defaultSizePt: row, overrides: [] }, columns: { defaultSizePt: column, overrides: [] } })) };
};
