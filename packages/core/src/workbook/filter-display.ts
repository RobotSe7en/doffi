import type { WorkbookModel } from '../runtime.js';
import type { WorkbookValue } from '../contracts/workbook.js';
import type { DeepReadonly } from '../runtime.js';
import type { WorkbookReferenceLocation } from './reference.js';
import type { WorkbookScalarContext } from './scalar.js';
import { resolveWorkbookCellFormat } from './cell-format.js';
import { createWorkbookNumberFormatter } from './number-formatter.js';

/** Intrinsic format text for value-list filters; no screen width, fonts or
 * conditional rules enter calculation. Native scalar spellings remain valid. */
export function createWorkbookFilterDisplay(model: WorkbookModel, context: WorkbookScalarContext, account: (units: number) => void, step: () => void): {
  text(at: WorkbookReferenceLocation, value: DeepReadonly<WorkbookValue>): string | undefined;
  dispose(): void;
} {
  const snapshot = model.snapshot(), workbook = snapshot.metadata.workbook;
  const formats = new Map<string, (typeof workbook.numberFormats)[number] | undefined>();
  const formatter = createWorkbookNumberFormatter({ dateSystem: workbook.dateSystem, ...(context.locale ? { hostLocale: context.locale } : {}), ...(snapshot.metadata.locale ? { documentLocale: snapshot.metadata.locale } : {}) });
  return {
    text(at, value) {
      if (value.type !== 'number' && value.type !== 'text') return undefined;
      const sheet = workbook.sheets.find(sheet => { step(); return sheet.sheetId === at.sheetId; })!, id = resolveWorkbookCellFormat(sheet, at, model.getCell(at.sheetId, at.row, at.column)).numberFormatId;
      if (id !== undefined && !formats.has(id)) {
        const definition = workbook.numberFormats.find(format => { step(); return format.formatId === id; });
        if (formats.size >= 256) formats.delete(formats.keys().next().value!); formats.set(id, definition);
      }
      const definition = id === undefined ? undefined : formats.get(id); account(definition && 'code' in definition ? definition.code.length : 1);
      const result = formatter.format(value.value, definition);
      if (result.kind !== 'parts') return undefined;
      // Fill and reserved spacing have no intrinsic string. Explicit literal
      // spaces remain text, as they do in the source number format.
      const text = result.parts.map(part => part.kind === 'text' ? part.text : '').join(''); account(text.length); return text;
    },
    dispose() { formatter.dispose(); formats.clear(); },
  };
}
