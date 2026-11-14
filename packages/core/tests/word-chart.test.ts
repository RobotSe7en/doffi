import { expect, test } from 'vitest';
import { createWordDocument, createDocumentSession, openDocument } from '@doffi/core';
import type { DocumentOperation } from '@doffi/core';
test('chart data edits are atomic, reject invalid axes and vectors, copy independently and undo through shared history', () => {
    const session = createDocumentSession(openDocument(createWordDocument({ documentId: 'chart-core' }), { kind: 'document' })), apply = (operations: DocumentOperation[]) => session.commit(session.prepare({ protocolVersion: 1, documentId: 'chart-core', baseRevision: session.snapshot().manifest.revision, operations }));
    apply([{ op: 'insertWordDrawing', nodeId: 'paragraph-1', offset: 0, geometry: { x: 0, y: 0, width: 300, height: 200, rotation: 0, flipH: false, flipV: false }, layout: { mode: 'inline' }, object: { kind: 'chart', chart: { kind: 'column', legend: 'none', showLabels: false, categories: ['甲', '乙'], series: [{ seriesId: 's', name: '数据', values: [1, 2] }] } } }]);
    const original = session.snapshot().metadata.word!.drawings![0]!;
    if (original.kind !== 'chart')
        throw Error('chart missing');
    const before = session.snapshot();
    expect(() => apply([{ op: 'setWordChart', nodeId: original.nodeId, changes: { minimum: 10, maximum: 2 } }])).toThrow(/maximum/);
    expect(() => apply([{ op: 'setWordChart', nodeId: original.nodeId, changes: { categories: ['短'] } }])).toThrow(/vectors/);
    expect(session.snapshot()).toEqual(before);
    apply([{ op: 'copyWordDrawing', nodeId: original.nodeId }]);
    const copy = session.snapshot().metadata.word!.drawings!.find(drawing => drawing.nodeId !== original.nodeId)!;
    apply([{ op: 'setWordChart', nodeId: copy.nodeId, changes: { series: [{ seriesId: 's', name: '修改', values: [9, 8] }] } }]);
    expect(session.snapshot().metadata.word!.drawings!.find(drawing => drawing.nodeId === original.nodeId)).toEqual(original);
    session.undo();
    expect(session.snapshot().metadata.word!.drawings!.find(drawing => drawing.nodeId === copy.nodeId)?.kind).toBe('chart');
    session.undo();
    expect(session.snapshot().metadata.word!.drawings).toHaveLength(1);
});
