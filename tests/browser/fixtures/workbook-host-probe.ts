import type { WorkbookLayoutMeasurer } from '@doffi/core';

/** Test-only point metrics for the final host-ownership write after the editor
 * and its real font measurer have been released. The accepted-source audit and
 * all actual canvas/font assertions happen before this independent probe; its
 * geometry is deliberately not native or font-layout acceptance evidence. */
export const hostOwnershipProbeLayout: WorkbookLayoutMeasurer = request => ({ sourceDigest: request.sourceDigest,
  fontEnvironment: 'host-ownership-probe-synthetic-points', sheets: request.sheetIds.map(sheetId => ({ sheetId,
    rows: { defaultSizePt: 15, overrides: [] }, columns: { defaultSizePt: 45, overrides: [] } })) });
