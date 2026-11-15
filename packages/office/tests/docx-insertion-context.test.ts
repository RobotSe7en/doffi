import { expect, test } from 'vitest';
import { createDocumentSession, openDocument, SourceStyleResolver, DEFAULT_DOCUMENT_READ_LIMITS } from '@doffi/core';
import type { ContentDocumentModel, DocumentOperation, SourceStyleNode } from '@doffi/core';
import { exportDocx, importDocx } from '@doffi/office';
import { body, mutateDocx, styles } from './regression-fixtures.js';

function setup() {
    const bytes=mutateDocx('01-text-numbering.docx',{'word/styles.xml':()=>styles('<w:style w:type="paragraph" w:styleId="Normal" w:default="1"/>'),'word/document.xml':source=>body(source,'<w:p><w:r><w:rPr><w:b/><w:color w:val="880000"/></w:rPr><w:t>A</w:t></w:r><w:r><w:rPr><w:i/><w:color w:val="000088"/></w:rPr><w:t>B</w:t></w:r></w:p><w:p><w:r><w:t>OTHER</w:t></w:r></w:p>')}),session=createDocumentSession(openDocument(importDocx(bytes,{documentId:'insertion-context',unsupportedContent:'reject'}),{kind:'document'})),ids=[session.snapshot().metadata.word!.stories[0]!.rootNodeId];
    for(let i=0;i<ids.length;i++)ids.push(...session.getNode(ids[i]!)!.childNodeIds);
    const paragraph=ids.find(id=>session.getNode(id)!.tagName==='p'&&session.getNode(id)!.text==='AB')!,left=ids.filter(id=>session.getNode(id)!.text==='A').at(-1)!,right=ids.filter(id=>session.getNode(id)!.text==='B').at(-1)!,other=ids.find(id=>session.getNode(id)!.text==='OTHER')!,root=session.getNode(ids[0]!)!.parentNodeId!,apply=(operations:DocumentOperation[])=>session.commit(session.prepare({protocolVersion:1,documentId:'insertion-context',baseRevision:session.snapshot().manifest.revision,operations}));
    apply([{op:'setWordReview',nodeId:root,trackChanges:true,author:'A'}]);return{session,paragraph,left,right,other,root,apply};
}
function appearance(model:ContentDocumentModel,text:string){
    const word=model.snapshot().metadata.word!,resolver=new SourceStyleResolver(model.snapshot().styles.map(sheet=>sheet.css),DEFAULT_DOCUMENT_READ_LIMITS,new Set(word.revisions.filter(item=>item.kind==='insert'||item.kind==='delete').map(item=>item.nodeId)));let result:Readonly<Record<string,string>>|undefined;
    const visit=(id:string,parent?:SourceStyleNode):void=>{const node=model.getNode(id)!,projected={tag:node.tagName,attrs:node.attributes,parent};if(node.text===text)result=resolver.get(projected);for(const child of node.childNodeIds)visit(child,projected);};visit(word.stories[0]!.rootNodeId);return result!;
}
for(const context of ['default','left','right','paragraph']as const)test('tracked boundary insertion inherits the declared source context: '+context,()=>{
    const {session,paragraph,left,right,apply}=setup(),expected=context==='right'?{color:'#000088','font-style':'italic','font-weight':'normal'}:context==='paragraph'?{'font-style':'normal','font-weight':'normal'}:{color:'#880000','font-weight':'bold','font-style':'normal'};
    apply([{op:'trackWordText',nodeId:paragraph,from:1,to:1,expectedText:'',text:'X',...(context==='default'?{}:{insertionContextNodeId:context==='left'?left:context==='right'?right:paragraph})}]);expect(appearance(session,'X')).toMatchObject(expected);expect(session.getNode(paragraph)!.text).toBe('AXB');
    let bytes=exportDocx(session).bytes;for(let i=0;i<2;i++){const native=openDocument(importDocx(bytes,{documentId:'context-native',unsupportedContent:'reject'}),{kind:'document'});expect(appearance(native,'X')).toMatchObject(expected);bytes=exportDocx(native).bytes;}
    session.undo();expect(session.getNode(paragraph)!.text).toBe('AB');
});
test('typing into the right run after a pending insertion keeps independent format and review ownership',()=>{
    const{session,paragraph,left,right,apply}=setup();apply([{op:'trackWordText',nodeId:paragraph,from:1,to:1,expectedText:'',text:'X',insertionContextNodeId:left}]);const first=session.snapshot().metadata.word!.revisions[0]!;
    apply([{op:'trackWordText',nodeId:paragraph,from:2,to:2,expectedText:'',text:'Y',insertionContextNodeId:right}]);expect(session.snapshot().metadata.word!.revisions).toHaveLength(2);expect(appearance(session,'Y')).toMatchObject({color:'#000088','font-style':'italic'});
    apply([{op:'reviewWordRevision',nodeId:first.nodeId,revisionId:first.revisionId!,decision:'reject'}]);expect(session.getNode(paragraph)!.text).toBe('AYB');expect(appearance(session,'Y')).toMatchObject({color:'#000088','font-style':'italic'});openDocument(importDocx(exportDocx(session).bytes,{documentId:'context-native',unsupportedContent:'reject'}),{kind:'document'});
});
test('an insertion context outside its target or boundary fails without a partial revision',()=>{
    const{session,paragraph,other,right,apply}=setup(),before=session.snapshot();for(const context of [other,right,'missing'])expect(()=>apply([{op:'trackWordText',nodeId:paragraph,from:0,to:0,expectedText:'',text:'X',insertionContextNodeId:context}])).toThrowError(expect.objectContaining({code:'InvalidOperation'}));expect(session.snapshot()).toEqual(before);
});

test('an explicit empty run retains its format while default insertion after a native break stays outside the break',()=>{
 const{session,paragraph,apply}=setup();apply([{op:'setWordReview',nodeId:session.getNode(session.snapshot().metadata.word!.stories[0]!.rootNodeId)!.parentNodeId!,trackChanges:false},{op:'replaceFragment',nodeId:paragraph,html:'<span data-word-run="true" style="color:#008800"></span><br><em>B</em>'}]);const empty=session.getNode(paragraph)!.childNodeIds[0]!,lineBreak=session.getNode(paragraph)!.childNodeIds[1]!,root=session.getNode(session.snapshot().metadata.word!.stories[0]!.rootNodeId)!.parentNodeId!;
 apply([{op:'setWordReview',nodeId:root,trackChanges:true,author:'A'},{op:'trackWordText',nodeId:paragraph,from:0,to:0,expectedText:'',text:'X',insertionContextNodeId:empty}]);expect(appearance(session,'X').color).toBe('#008800');apply([{op:'trackWordText',nodeId:paragraph,from:2,to:2,expectedText:'',text:'Y'}]);expect(session.getNode(paragraph)!.text).toBe('X\nYB');expect(appearance(session,'Y')['font-style']).toBe('normal');const before=session.snapshot();expect(()=>apply([{op:'trackWordText',nodeId:paragraph,from:2,to:2,expectedText:'',text:'Z',insertionContextNodeId:lineBreak}])).toThrowError(expect.objectContaining({code:'InvalidOperation'}));expect(session.snapshot()).toEqual(before);openDocument(importDocx(exportDocx(session).bytes,{documentId:'break-context',unsupportedContent:'reject'}),{kind:'document'});
});
