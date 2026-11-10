import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { createWorkbook, DEFAULT_WORKBOOK_RULE_LIMITS, openDocument, readWorkbookRules } from '@doffi/core';
import type { WorkbookRuleReadResponse } from '@doffi/core';
import { importXlsx } from '../../office/src/xlsx-import.js';
import { checkedWorkbookRules } from '../src/workbook/rule-response.js';
import { modelRuleSource } from '../src/workbook/rule-source.js';

describe('workbook rule response acceptance', () => {
  it('lets port traffic cancel a current-value read and keeps later independent reads usable', async () => {
    const model = openDocument(createWorkbook({ documentId: 'rule-tasks' }), { kind: 'workbook' }), source = modelRuleSource(model);
    const coordinates = Array.from({ length: 512 }, (_, row) => ({ row, column: 0 }));
    const request = { documentId: 'rule-tasks', revision: 0, sheetId: 'sheet-1', coordinates, includeCurrent: true, ...DEFAULT_WORKBOOK_RULE_LIMITS };
    const abort = new AbortController(), reading = source.read(request, abort.signal);
    setImmediate(() => abort.abort());
    await expect(reading).rejects.toMatchObject({ code: 'OperationCancelled' });
    const result = await source.read({ ...request, coordinates: coordinates.slice(0, 64) }, new AbortController().signal);
    expect(result.cells).toHaveLength(64);
    expect(result.cells.every(cell => cell.current.kind === 'value' && cell.current.value.type === 'blank')).toBe(true);
    expect(model.snapshot().manifest.revision).toBe(0);
  });
  it('requires exact priority order and admits omission only after an actual stopping result', async () => {
    const files = new Map(createWorkbook({ documentId: 'priority' })), metadata = JSON.parse(new TextDecoder().decode(files.get('metadata.json')!));
    const origin = { row: 0, column: 0 }, ranges = [{ from: origin, to: origin }];
    metadata.workbook.sheets[0].conditionalRules = [3, 1, 2].map(priority => ({ ruleId: 'rule-' + priority, priority, kind: 'expression', expression: '=TRUE', origin, ranges, stopIfTrue: priority === 2 }));
    files.set('metadata.json', new TextEncoder().encode(JSON.stringify(metadata))); const model = openDocument(files, { kind: 'workbook' });
    const request = { documentId: 'priority', revision: 0, sheetId: 'sheet-1', coordinates: [origin], ...DEFAULT_WORKBOOK_RULE_LIMITS }, result = await readWorkbookRules(model, request);
    expect(result.cells[0]!.conditional.map(rule => rule.ruleId)).toEqual(['rule-1', 'rule-2']);
    expect(checkedWorkbookRules(result, request, model.snapshot())).toEqual(result);
    const altered = structuredClone(result); (altered.cells[0]!.conditional as unknown[]).reverse();
    expect(() => checkedWorkbookRules(altered, request, model.snapshot())).toThrowError(expect.objectContaining({ code: 'InvalidDocument' }));
    expect(() => checkedWorkbookRules({ ...result, cells: [{ ...result.cells[0]!, conditional: result.cells[0]!.conditional.slice(0, 1) }] }, request, model.snapshot())).toThrowError(expect.objectContaining({ code: 'InvalidDocument' }));
  });
  it('captures valid derived values and rejects incomplete, malformed, reordered or oversized replies', async () => {
    const model = openDocument(await importXlsx(readFileSync('examples/xlsx/06-rules-validation.xlsx'), { documentId: 'rules-response' }), { kind: 'workbook' });
    const request = { documentId: 'rules-response', revision: 0, sheetId: 'sheet-1', coordinates: [{ row: 1, column: 0 }, { row: 2, column: 2 }, { row: 2, column: 4 }], ...DEFAULT_WORKBOOK_RULE_LIMITS };
    const result = await readWorkbookRules(model, request), accepted = checkedWorkbookRules(result, request, model.snapshot());
    expect(accepted).toEqual(result); expect(accepted).not.toBe(result); expect(Object.isFrozen(accepted.cells[0]!.current)).toBe(true);
    const invalid: ((value: WorkbookRuleReadResponse) => void)[] = [
      value => { (value as unknown as { cells: unknown[] }).cells.pop(); },
      value => { (value as unknown as { cells: unknown[] }).cells[1] = value.cells[0]; },
      value => { (value.cells[0] as { row: number }).row = -1; (value.cells[0] as { column: number }).column = 16_384; },
      value => { (value.cells[0]!.conditional as unknown[]).pop(); },
      value => { (value.cells[0]!.conditional[0] as { ruleId: string }).ruleId = 'not-this-rule'; },
      value => { (value.cells[0]!.conditional[0] as { state: string }).state = 'unavailable'; },
      value => { (value.cells[1]!.conditional[0]!.scale as { ratio: number }).ratio = Infinity; },
      value => { (value.cells[1]!.conditional[0]!.scale as { toStop: number }).toStop = 3; },
      value => { (value.cells[2]!.validation as { state: string }).state = 'none'; },
      value => { (value as { steps: number }).steps = -1; },
      value => { (value as unknown as { unexpected: boolean }).unexpected = true; },
      value => { (value.cells[0]!.current as unknown as { value: unknown }).value = { type: 'number', value: NaN }; },
    ];
    for (const alter of invalid) {
      const broken = structuredClone(result); alter(broken);
      expect(() => checkedWorkbookRules(broken, request, model.snapshot())).toThrowError(expect.objectContaining({ code: 'InvalidDocument' }));
    }
    expect(() => checkedWorkbookRules({ ...result, revision: 1 }, request, model.snapshot())).toThrowError(expect.objectContaining({ code: 'RevisionConflict' }));
    expect(() => checkedWorkbookRules(result, { ...request, maxCodeUnits: 1 }, model.snapshot())).toThrowError(expect.objectContaining({ code: 'ResourceLimitExceeded' }));
  });
});
