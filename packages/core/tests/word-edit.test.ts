import {expect,test}from 'vitest';
import {createWordDocument,createDocumentSession,openDocument,getWordTextDefaults,SourceStyleResolver,DEFAULT_DOCUMENT_READ_LIMITS}from '@doffi/core';
import type {SourceStyleNode,DocumentOperation}from '@doffi/core';
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
test('reset formatting resolves authored defaults and keeps partial ranges, Unicode and one-step history',()=>{
 const s=create(),apply=(operations:DocumentOperation[])=>s.commit(s.prepare({protocolVersion:1,documentId:'word-edits',baseRevision:s.snapshot().manifest.revision,operations}));
 apply([{op:'replaceText',nodeId:'paragraph-1',from:0,to:0,expectedText:'',text:'A禰󠄀😀Z'},{op:'setWordParagraph',nodeId:'paragraph-1',styleId:'Heading1'},{op:'formatWordText',nodeId:'paragraph-1',from:0,to:7,expectedText:'A禰󠄀😀Z',properties:{'font-size':'32pt','font-style':'italic','text-decoration':'underline','font-family':'"Liberation Mono"'}}]);
 const before=s.snapshot();apply([{op:'resetWordText',nodeId:'paragraph-1',from:1,to:6,expectedText:'禰󠄀😀'}]);
 const resolver=new SourceStyleResolver(s.snapshot().styles.map(style=>style.css),DEFAULT_DOCUMENT_READ_LIMITS),values:Record<string,string>[]=[];const visit=(id:string,parent?:SourceStyleNode):void=>{const item=s.getNode(id)!,node={tag:item.tagName,attrs:item.attributes,parent};values.push(resolver.get(node));for(const child of item.childNodeIds)visit(child,node);};visit('paragraph-1');expect(values.some(value=>value['font-size']==='32pt'&&value['font-style']==='italic')).toBe(true);expect(values.some(value=>value['font-size']==='18pt'&&value['font-style']==='normal')).toBe(true);expect(s.getNode('paragraph-1')!.text).toBe('A禰󠄀😀Z');s.undo();expect(s.snapshot().html).toBe(before.html);
 apply([{op:'resetWordText',nodeId:'paragraph-1',from:0,to:7,expectedText:'A禰󠄀😀Z',resetParagraph:true}]);expect(s.getNode('paragraph-1')!.tagName).toBe('p');expect(getWordTextDefaults(s,'paragraph-1')['font-size']).toBe('12pt');expect(s.getNode('paragraph-1')!.text).toBe('A禰󠄀😀Z');
});
