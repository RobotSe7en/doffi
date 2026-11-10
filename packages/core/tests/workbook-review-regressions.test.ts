import { describe, expect, it } from 'vitest';
import { createWorkbookScalarEvaluator, DEFAULT_WORKBOOK_RULE_LIMITS, readWorkbookRules } from '../src/index.js';
import { scalarWorkbook } from './workbook-scalar-fixture.js';

const at = { sheetId: 'sheet-1', row: 0, column: 0 };
describe('independent PR25 core review regressions', () => {
  it('charges dependency text across the complete rule batch', async () => {
    const model = scalarWorkbook([0, 1].flatMap(row => [
      { row, column: 0, value: { type: 'number' as const, value: 0 } },
      { row, column: 1, value: { type: 'text' as const, value: 'x'.repeat(60) } },
    ]), metadata => {
      metadata.workbook.sheets[0]!.conditionalRules = [{ ruleId: 'long', kind: 'expression', priority: 1, stopIfTrue: false,
        origin: { row: 0, column: 0 }, ranges: [{ from: { row: 0, column: 0 }, to: { row: 1, column: 0 } }], expression: '=LEN(B1)>0' }];
    });
    const before = model.serialize(), query = { documentId: model.snapshot().manifest.documentId, revision: 0, sheetId: 'sheet-1',
      coordinates: [{ row: 0, column: 0 }, { row: 1, column: 0 }], ...DEFAULT_WORKBOOK_RULE_LIMITS };
    await expect(readWorkbookRules(model, { ...query, maxCodeUnits: 100 })).rejects.toMatchObject({ code: 'ResourceLimitExceeded' });
    expect((await readWorkbookRules(model, { ...query, maxCodeUnits: 512 })).cells.map(cell => cell.conditional[0]!.state)).toEqual(['match', 'match']);
    expect(model.serialize()).toEqual(before);
  });
  it('reports consumed scalar text and honors remaining and zero text budgets', () => {
    const model = scalarWorkbook([{ row: 0, column: 0, value: { type: 'number', value: 1 } }, { row: 0, column: 1, value: { type: 'text', value: 'x'.repeat(60) } }]);
    const evaluator = createWorkbookScalarEvaluator(model);
    try {
      expect(evaluator.cell(at, { maxCodeUnits: 0 })).toMatchObject({ kind: 'value', codeUnits: 0 });
      expect(() => evaluator.evaluate('=1', at, { maxCodeUnits: 0 })).toThrowError(expect.objectContaining({ code: 'ResourceLimitExceeded' }));
      expect(() => evaluator.evaluate('=LEN(B1)', at, { maxCodeUnits: 67 })).toThrowError(expect.objectContaining({ code: 'ResourceLimitExceeded' }));
      expect(evaluator.evaluate('=LEN(B1)', at, { maxCodeUnits: 68 })).toMatchObject({ kind: 'value', value: { value: 60 }, codeUnits: 68 });
    } finally { evaluator.dispose(); }
  });
  it('retains Excel global-name bindings inside unions, intersections and ranges', () => {
    const model = scalarWorkbook([1, 9, 0].map((value, row) => ({ row, column: 0, value: { type: 'number' as const, value } })), metadata => {
      const origin = { row: 0, column: 0 }, sheet = metadata.workbook.sheets[0]!.name.replaceAll("'", "''");
      metadata.workbook.names = [{ nameId: 'inner-global', name: 'Inner', origin, expression: '=$A$1' },
        { nameId: 'inner-local', name: 'Inner', sheetId: 'sheet-1', origin, expression: '=$A$2' },
        ...['=SUM(Inner)', '=SUM((Inner,$A$3))', '=SUM(Inner $A$1)', '=SUM(Inner:$A$3)', `=SUM(('${sheet}'!Inner,$A$3))`]
          .map((expression, index) => ({ nameId: 'outer-' + index, name: 'Outer' + index, origin, expression }))];
    });
    const evaluator = createWorkbookScalarEvaluator(model);
    try {
      for (const [index, value] of [1, 1, 1, 10].entries()) expect(evaluator.evaluate('=Outer' + index, at)).toMatchObject({ kind: 'value', value: { type: 'number', value } });
      expect(evaluator.evaluate('=Outer4', at)).toMatchObject({ kind: 'value', value: { type: 'number', value: 9 } });
    } finally { evaluator.dispose(); }
  });
  it('does not diagnose a cycle when the same global name is evaluated in a different bang context', () => {
    const model = scalarWorkbook([{ row: 0, column: 0, value: { type: 'number', value: 1 } }], metadata => {
      metadata.workbook.names = [{ nameId: 'context', name: 'ContextValue', origin: { row: 0, column: 0 }, expression: '=IF(ISERROR(!$A$1),7,[0]!ContextValue)' }];
    });
    const evaluator = createWorkbookScalarEvaluator(model);
    try { expect(evaluator.evaluate('=ContextValue', at)).toMatchObject({ kind: 'value', value: { type: 'number', value: 7 } }); }
    finally { evaluator.dispose(); }
  });
});
