import type { DeepReadonly, WorkbookModel, WorkbookRangeChange } from '../runtime.js';
import type { WorkbookAutoFilter, WorkbookCoordinate, WorkbookSheet } from '../contracts/workbook.js';
import { resolveWorkbookCellFormat } from '../workbook/cell-format.js';

const lists = (sheet: DeepReadonly<WorkbookSheet> | undefined) => [...sheet?.autoFilter ? [sheet.autoFilter] : [], ...sheet?.tables?.flatMap(table => table.autoFilter ? [table.autoFilter] : []) ?? []]
  .filter(filter => filter.columns.some(column => column.criterion?.kind === 'values' && column.criterion.values.length));
/** Source number formats participate; CSS paint and conditional display cannot
 * make predicates recursively depend on their own visibility. Dimensions are
 * only a scan hint: their segmentation never proves an effective format change. */
export function* workbookFilterFormatImpact(previous: WorkbookModel, model: WorkbookModel, sourceRanges: readonly WorkbookRangeChange[], step: () => void): Generator<void, WorkbookRangeChange[]> {
  const before = previous.snapshot(), after = model.snapshot(), output: WorkbookRangeChange[] = [];
  if (![before, after].some(snapshot => snapshot.metadata.workbook.sheets.some(sheet => lists(sheet).length))) return output;
  const definitions = [new Map<string, string>(), new Map<string, string>()];
  for (const [index, snapshot] of [before, after].entries()) for (const format of snapshot.metadata.workbook.numberFormats) {
    step(); const { formatId, ...definition } = format; definitions[index]!.set(formatId, JSON.stringify(definition)); yield;
  }
  const changed = new Set([...definitions[0]!.keys()].filter(id => definitions[0]!.get(id) !== definitions[1]!.get(id)));
  for (const sheet of after.metadata.workbook.sheets) {
    const old = before.metadata.workbook.sheets.find(item => item.sheetId === sheet.sheetId), filters = [...lists(old), ...lists(sheet)]; if (!old || !filters.length) continue;
    const identity = (item: DeepReadonly<WorkbookSheet>) => JSON.stringify([item.defaultNumberFormatId, item.rowDimensions.map(span => [span.from, span.to, span.numberFormatId]), item.columnDimensions.map(span => [span.from, span.to, span.numberFormatId])]);
    const calendar = before.metadata.workbook.dateSystem !== after.metadata.workbook.dateSystem;
    const metadataChanged = identity(old) !== identity(sheet) || calendar;
    const owns = (filter: DeepReadonly<WorkbookAutoFilter>, row: number, column: number) => row > filter.range.from.row && row <= filter.range.to.row
      && filter.columns.some(item => { step(); return item.column === column && item.criterion?.kind === 'values' && item.criterion.values.length; });
    const signature = (source: WorkbookModel, owner: DeepReadonly<WorkbookSheet>, index: number, at: WorkbookCoordinate) => {
      const id = resolveWorkbookCellFormat(owner, at, source.getCell(owner.sheetId, at.row, at.column)).numberFormatId;
      return id === undefined ? 'default' : definitions[index]!.get(id);
    };
    let affected = false;
    const compare = (at: WorkbookCoordinate) => {
      step(); if (!filters.some(filter => owns(filter, at.row, at.column))) return;
      if (calendar || signature(previous, old, 0, at) !== signature(model, sheet, 1, at)) affected = true;
    };
    if (!metadataChanged && !changed.size) for (const range of sourceRanges) {
      if (range.sheetId !== sheet.sheetId) continue;
      for (let row = range.range.from.row; row <= range.range.to.row; row++) for (let column = range.range.from.column; column <= range.range.to.column; column++) { compare({ row, column }); yield; }
    }
    else for (const source of [previous, model]) for (const chunk of source.snapshot().manifest.chunks) {
      step(); if (chunk.sheetId !== sheet.sheetId) continue;
      for (const offset of [0, 128]) for (const cell of source.readWindow({ sheetId: sheet.sheetId, row: chunk.row + offset, column: chunk.column, rows: 128, columns: 32 })) { compare(cell); yield; }
    }
    if (affected) for (const filter of filters) if (filter.range.from.row < filter.range.to.row) output.push({ sheetId: sheet.sheetId, range: { from: { row: filter.range.from.row + 1, column: 0 }, to: { row: filter.range.to.row, column: 16383 } } });
  }
  return output;
}
