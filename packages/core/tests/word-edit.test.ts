import {expect,test}from 'vitest';
import {createWordDocument,createDocumentSession,openDocument}from '@doffi/core';
const create=()=>createDocumentSession(openDocument(createWordDocument({documentId:'word-edits'}),{kind:'document'}));
test('paragraph formatting and style application are atomic and share undo history',()=>{
 const session=create();const apply=(operations:Parameters<typeof session.prepare>[0]['operations'])=>session.commit(session.prepare({protocolVersion:1,documentId:'word-edits',baseRevision:session.snapshot().manifest.revision,operations}));
 apply([{op:'replaceText',nodeId:'paragraph-1',from:0,to:0,expectedText:'',text:'标题文字😀'}]);
 apply([{op:'setWordParagraph',nodeId:'paragraph-1',styleId:'Heading1',properties:{'text-align':'justify','margin-left':'24pt','line-height':'2'}}]);
 expect(session.getNode('paragraph-1')!.tagName).toBe('h1');expect(session.getNode('paragraph-1')!.html).toContain('text-align:justify');
 session.undo();expect(session.getNode('paragraph-1')!.tagName).toBe('p');session.redo();expect(session.getNode('paragraph-1')!.text).toBe('标题文字😀');
 const before=session.snapshot();expect(()=>apply([{op:'setWordParagraph',nodeId:'paragraph-1',properties:{'margin-top':'8pt'}},{op:'setWordParagraph',nodeId:'paragraph-1',styleId:'missing'}])).toThrow();expect(session.snapshot()).toEqual(before);
});
test('numbering definitions and paragraph references are one transaction',()=>{
 const session=create();session.commit(session.prepare({protocolVersion:1,documentId:'word-edits',baseRevision:0,operations:[{op:'setWordNumbering',nodeId:'document-root',numberingId:'decimal',levels:[{level:0,start:5,format:'decimal',text:'%1.',suffix:'tab',cssClass:'numbering'}]},{op:'setWordList',nodeId:'paragraph-1',numberingId:'decimal',level:0}]}));
 expect(session.getNode('paragraph-1')!.attributes['data-word-numbering']).toBe('decimal');session.undo();expect(session.snapshot().metadata.word!.numbering).toEqual([]);
});
test('shared Word definitions cannot be written through a paragraph-scoped operation',()=>{
 const session=create();expect(()=>session.prepare({protocolVersion:1,documentId:'word-edits',baseRevision:0,operations:[{op:'setWordStyle',nodeId:'paragraph-1',style:{styleId:'secret',kind:'paragraph',cssClass:'secret'}}]})).toThrow();expect(session.history().undoSteps).toBe(0);
});
