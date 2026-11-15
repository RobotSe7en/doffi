import { expect, test } from 'vitest';
import { createDocumentSession, openDocument, SourceStyleResolver, DEFAULT_DOCUMENT_READ_LIMITS } from '@doffi/core';
import type { ContentDocumentModel, DocumentOperation, SourceStyleNode } from '@doffi/core';
import { exportDocx, importDocx } from '@doffi/office';
import { body, mutateDocx, styles } from './regression-fixtures.js';

function fixture() {
    const bytes = mutateDocx('01-text-numbering.docx', {
        'word/styles.xml': () => styles('<w:style w:type="paragraph" w:styleId="Normal" w:default="1"/><w:style w:type="character" w:styleId="Italic"><w:rPr><w:i/></w:rPr></w:style>'),
        'word/document.xml': source => body(source, '<w:p><w:r><w:rPr><w:color w:val="AA0000"/><w:sz w:val="64"/><w:highlight w:val="yellow"/><w:u w:val="single"/><w:vertAlign w:val="superscript"/></w:rPr><w:t>ABC</w:t></w:r></w:p>'),
    });
    const session = createDocumentSession(openDocument(importDocx(bytes, { documentId: 'run-inherit', unsupportedContent: 'reject' }), { kind: 'document' })), nodes = [session.snapshot().metadata.word!.stories[0]!.rootNodeId];
    for (let index = 0; index < nodes.length; index++) nodes.push(...session.getNode(nodes[index]!)!.childNodeIds);
    const paragraph = nodes.find(id => session.getNode(id)!.tagName === 'p')!, root = session.getNode(nodes[0]!)!.parentNodeId!;
    const apply = (operations: DocumentOperation[]) => session.commit(session.prepare({ protocolVersion: 1, documentId: 'run-inherit', baseRevision: session.snapshot().manifest.revision, operations }));
    return { session, paragraph, root, apply };
}
function runs(model: ContentDocumentModel): { text: string; css: Readonly<Record<string, string>> }[] {
    const resolver = new SourceStyleResolver(model.snapshot().styles.map(sheet => sheet.css), DEFAULT_DOCUMENT_READ_LIMITS,new Set(model.snapshot().metadata.word!.revisions.filter(revision=>['insert','delete','moveFrom','moveTo'].includes(revision.kind)).map(revision=>revision.nodeId))), result: { text: string; css: Readonly<Record<string, string>> }[] = [];
    const visit = (id: string, parent?: SourceStyleNode): void => { const node = model.getNode(id)!, projected = { tag: node.tagName, attrs: node.attributes, parent }; if (node.text) result.push({ text: node.text, css: resolver.get(projected) }); for (const child of node.childNodeIds) visit(child, projected); };
    visit(model.snapshot().metadata.word!.stories[0]!.rootNodeId); return result;
}
const selected = (model: ContentDocumentModel) => runs(model).find(run => run.text === 'B')!.css;
function reopenTwice(model: ContentDocumentModel, verify: (model: ContentDocumentModel) => void): void {
    let bytes = exportDocx(model).bytes;
    for (let pass = 0; pass < 2; pass++) { const reopened = openDocument(importDocx(bytes, { documentId: 'reopened', unsupportedContent: 'reject' }), { kind: 'document' }); verify(reopened); bytes = exportDocx(reopened).bytes; }
}
test('partial direct color and size changes retain native superscript, highlight and an automatic-color underline through two reopens', () => {
    const { session, paragraph, apply } = fixture();
    apply([{ op: 'formatWordText', nodeId: paragraph, from: 1, to: 2, expectedText: 'B', properties: { color: '#0000aa', 'font-size': '12pt' } }]);
    const verify = (model: ContentDocumentModel) => { expect(selected(model)).toMatchObject({ color: '#0000aa', 'font-size': '12pt', 'vertical-align': 'super', 'background-color': '#ffff00', 'text-decoration-line': 'underline' }); expect(runs(model).find(run => run.text.includes('A') && run.css['font-size'] === '32pt')!.css).toMatchObject({ color: '#aa0000', 'font-size': '32pt', 'vertical-align': 'super' }); };
    verify(session); reopenTwice(session, verify); session.undo(); expect(runs(session)[0]!.text).toBe('ABC'); session.redo(); verify(session);
});
test('explicit baseline, transparent highlight and new decoration replace only the selected native run properties', () => {
    const { session, paragraph, apply } = fixture();
    apply([{ op: 'formatWordText', nodeId: paragraph, from: 1, to: 2, expectedText: 'B', properties: { 'vertical-align': 'baseline', 'background-color': 'transparent', 'text-decoration-style': 'dotted', 'text-decoration-color': '#008800' } }]);
    const verify = (model: ContentDocumentModel) => expect(selected(model)).toMatchObject({ 'vertical-align': 'baseline', 'background-color': 'transparent', 'text-decoration-line': 'underline', 'text-decoration-style': 'dotted', 'text-decoration-color': '#008800' });
    verify(session); reopenTwice(session, verify);
});
test('overlapping tracked run edits retain the original native inheritance when rejecting before and after native reopen', () => {
    const { session, paragraph, root, apply } = fixture();
    apply([{ op: 'setWordReview', nodeId: root, trackChanges: true, author: 'A' }, { op: 'formatWordText', nodeId: paragraph, from: 1, to: 2, expectedText: 'B', properties: { 'vertical-align': 'baseline', color: '#0000aa' } }]);
    apply([{ op: 'setWordReview', nodeId: root, trackChanges: true, author: 'B' }, { op: 'formatWordText', nodeId: paragraph, from: 1, to: 2, expectedText: 'B', properties: { 'background-color': 'transparent', 'text-decoration-line': 'none' } }]);
    const before = session.snapshot(); expect(before.metadata.word!.revisions).toHaveLength(1); expect(before.metadata.word!.revisions[0]!.author).toBe('B');
    const reopened = createDocumentSession(openDocument(importDocx(exportDocx(session).bytes, { documentId: 'native-review', unsupportedContent: 'reject' }), { kind: 'document' }));
    const nativeRoot = reopened.getNode(reopened.snapshot().metadata.word!.stories[0]!.rootNodeId)!.parentNodeId!;
    reopened.commit(reopened.prepare({ protocolVersion: 1, documentId: 'native-review', baseRevision: 0, operations: [{ op: 'reviewWordRevisions', nodeId: nativeRoot, decision: 'reject' }] }));
    const verify = (model: ContentDocumentModel) => expect(selected(model)).toMatchObject({ color: '#aa0000', 'font-size': '32pt', 'vertical-align': 'super', 'background-color': '#ffff00', 'text-decoration-line': 'underline' });
    verify(reopened); reopenTwice(reopened, verify);
    apply([{ op: 'reviewWordRevisions', nodeId: root, decision: 'reject' }]); verify(session); reopenTwice(session, verify); session.undo(); expect(session.snapshot().html).toBe(before.html);
});
test('applying an additive character style retains native vertical alignment and highlight', () => {
    const { session, paragraph, apply } = fixture(); apply([{ op: 'applyWordCharacterStyle', nodeId: paragraph, from: 1, to: 2, expectedText: 'B', styleId: 'Italic' }]);
    const verify = (model: ContentDocumentModel) => expect(selected(model)).toMatchObject({ 'font-style': 'italic', 'vertical-align': 'super', 'background-color': '#ffff00', 'text-decoration-line': 'underline' }); verify(session); reopenTwice(session, verify);
});

 test('tracked replacement and continued insertion retain effective native alignment, highlight and source decoration',()=>{
 const{session,paragraph,root,apply}=fixture();apply([{op:'setWordReview',nodeId:root,trackChanges:true,author:'A'},{op:'trackWordText',nodeId:paragraph,from:1,to:2,expectedText:'B',text:'X',properties:{'font-size':'12pt',color:'#0000aa'}}]);
 const verify=(model:ContentDocumentModel)=>{expect(runs(model).filter(run=>run.text==='X').at(-1)!.css).toMatchObject({'vertical-align':'super','background-color':'#ffff00','text-decoration-line':'underline',color:'#0000aa'});expect(runs(model).filter(run=>run.text==='B').at(-1)!.css).toMatchObject({'vertical-align':'super','font-size':'32pt','background-color':'#ffff00'});};verify(session);reopenTwice(session,verify);
 apply([{op:'trackWordText',nodeId:paragraph,from:3,to:3,expectedText:'',text:'Y',properties:{'font-size':'8pt'}}]);expect(session.getNode(paragraph)!.text).toBe('ABXYC');expect(runs(session).filter(run=>run.text==='Y').at(-1)!.css['vertical-align']).toBe('super');reopenTwice(session,model=>expect(model.getNode(model.snapshot().metadata.word!.stories[0]!.rootNodeId)!.text).toBe('ABXYC'));session.undo();verify(session);apply([{op:'reviewWordRevisions',nodeId:root,decision:'reject'}]);expect(session.getNode(paragraph)!.text).toBe('ABC');reopenTwice(session,model=>expect(runs(model).filter(run=>run.text==='B'||run.text==='ABC').at(-1)!.css['vertical-align']).toBe('super'));
 });
 test('removing the last text of a current author pending insertion cancels its empty revision and is undoable',()=>{const{session,paragraph,root,apply}=fixture();apply([{op:'setWordReview',nodeId:root,trackChanges:true,author:'A'},{op:'trackWordText',nodeId:paragraph,from:1,to:1,expectedText:'',text:'X'}]);const before=session.snapshot();apply([{op:'trackWordText',nodeId:paragraph,from:1,to:2,expectedText:'X',text:''}]);expect(session.getNode(paragraph)!.text).toBe('ABC');expect(session.snapshot().metadata.word!.revisions).toEqual([]);reopenTwice(session,model=>expect(model.snapshot().metadata.word!.revisions).toEqual([]));session.undo();expect(session.snapshot().html).toBe(before.html);});
