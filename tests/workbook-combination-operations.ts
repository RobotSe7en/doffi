import type { WorkbookDocumentMetadata, WorkbookOperation, WorkbookSession } from '@doffi/core';

/** One user command crossing definitions, content, drawing and thread boundaries. */
export function mixedEditingOperations(snapshot: ReturnType<WorkbookSession['snapshot']>): WorkbookOperation[] {
  const workbook = structuredClone(snapshot.metadata.workbook) as WorkbookDocumentMetadata['workbook'];
  const sheet = workbook.sheets[0]!, table = sheet.tables![0]!, rule = sheet.conditionalRules![0]!;
  const shape = workbook.objects!.find(object => object.kind === 'geometry')!, chart = workbook.charts![0]!;
  const modern = workbook.comments!.find(comment => comment.format === 'modern' && comment.replies!.length > 0)!;
  return [
    { op: 'setWorkbookCells', sheetId: sheet.sheetId, cells: [{ row: 1, column: 1, source: { kind: 'value', value: { type: 'number', value: 6 } } }] },
    { op: 'formatWorkbookRanges', sheetId: sheet.sheetId, ranges: [{ from: { row: 1, column: 0 }, to: { row: 1, column: 0 } }], properties: { 'font-size': '18pt', 'white-space': 'pre-wrap', color: '#224488' } },
    { op: 'setWorkbookTable', sheetId: sheet.sheetId, table: { ...table, autoFilter: { range: table.autoFilter!.range, columns: [
      { column: 1, hiddenButton: false, showButton: true, criterion: { kind: 'custom', and: true, tests: [{ operator: 'greaterThan', value: '2' }] } },
      { column: 2, hiddenButton: false, showButton: true, criterion: { kind: 'custom', and: true, tests: [{ operator: 'lessThan', value: '30' }] } },
    ] } } },
    { op: 'setWorkbookDefinedName', name: { nameId: 'current-amounts', name: 'CurrentAmounts', expression: '=SalesTable[金额]', origin: { row: 0, column: 0 } } },
    { op: 'setWorkbookConditionalRule', sheetId: sheet.sheetId, rule: { ...rule, priority: 3, ...(rule.kind === 'cellIs' ? { formulas: ['=40'] } : {}) } },
    { op: 'setWorkbookObjectText', nodeId: shape.nodeId, html: '<p>组合<strong>禰󠄀😀</strong></p>' },
    { op: 'setWorkbookChartTitle', nodeId: chart.nodeId, target: 'title', source: { html: '<p>当前<strong>销售</strong>😀</p>' } },
    { op: 'addWorkbookCommentReply', annotationId: modern.annotationId, html: '<p>复核<strong>当前值</strong>禰󠄀</p>' },
  ];
}
