import {expect,test} from 'vitest';
import {createDocumentSession,openDocument} from '@doffi/core';
import {exportDocx,importDocx} from '@doffi/office';
import {strFromU8,unzipSync} from 'fflate';
import {mutateDocx,body,styles} from './regression-fixtures.js';

function fixture(){
 const bytes=mutateDocx('01-text-numbering.docx',{'word/styles.xml':()=>styles('<w:style w:type="paragraph" w:styleId="Normal" w:default="1"/>'),'word/document.xml':source=>body(source,'<w:p><w:pPr><w:numPr><w:numId w:val="0"/></w:numPr></w:pPr><w:r><w:t>关闭编号</w:t></w:r></w:p>')}),session=createDocumentSession(openDocument(importDocx(bytes,{documentId:'off',unsupportedContent:'reject'}),{kind:'document'})),nodes=[session.snapshot().metadata.word!.stories[0]!.rootNodeId];
 for(let index=0;index<nodes.length;index++)nodes.push(...session.getNode(nodes[index]!)!.childNodeIds);
 return{session,paragraph:nodes.find(id=>session.getNode(id)!.tagName==='p')!,root:session.getNode(nodes[0]!)!.parentNodeId!};
}
test('explicit numbering cancellation survives later formatting and two strict native reopens',()=>{
 const{session,paragraph}=fixture();expect(session.getNode(paragraph)!.attributes['data-word-numbering-off']).toBe('true');
 session.commit(session.prepare({protocolVersion:1,documentId:'off',baseRevision:0,operations:[{op:'setWordParagraph',nodeId:paragraph,properties:{'margin-left':'18pt'}}]}));
 let bytes=exportDocx(session).bytes;for(let pass=0;pass<2;pass++){expect(strFromU8(unzipSync(bytes)['word/document.xml']!)).toContain('<w:numId w:val="0"');const model=openDocument(importDocx(bytes,{documentId:'off-reopen',unsupportedContent:'reject'}),{kind:'document'});expect(model.snapshot().html).toContain('data-word-numbering-off="true"');expect(model.snapshot().html).not.toContain('data-word-numbering="');bytes=exportDocx(model).bytes;}
 session.undo();expect(session.getNode(paragraph)!.attributes['data-word-numbering-off']).toBe('true');
});
test('tracked list removal preserves the old list and explicit off state through rejection and native reopen',()=>{
 const{session,paragraph,root}=fixture(),definition=session.snapshot().metadata.word!.numbering[0]!;
 session.commit(session.prepare({protocolVersion:1,documentId:'off',baseRevision:0,operations:[{op:'setWordList',nodeId:paragraph,numberingId:definition.numberingId,level:definition.levels[0]!.level}]}));expect(session.getNode(paragraph)!.attributes['data-word-numbering-off']).toBeUndefined();
 session.commit(session.prepare({protocolVersion:1,documentId:'off',baseRevision:1,operations:[{op:'setWordReview',nodeId:root,trackChanges:true,author:'A'},{op:'setWordList',nodeId:paragraph,numberingId:null}]}));
 expect(session.snapshot().metadata.word!.revisions[0]!.before!.attributes!['data-word-numbering']).toBe(definition.numberingId);expect(session.getNode(paragraph)!.attributes['data-word-numbering-off']).toBe('true');
 const bytes=exportDocx(session).bytes,reopened=createDocumentSession(openDocument(importDocx(bytes,{documentId:'off-native',unsupportedContent:'reject'}),{kind:'document'})),revision=reopened.snapshot().metadata.word!.revisions[0]!;
 reopened.commit(reopened.prepare({protocolVersion:1,documentId:'off-native',baseRevision:0,operations:[{op:'reviewWordRevision',nodeId:revision.nodeId,revisionId:revision.revisionId!,decision:'reject'}]}));expect(reopened.getNode(revision.nodeId)!.attributes['data-word-numbering']).toBeDefined();expect(reopened.getNode(revision.nodeId)!.attributes['data-word-numbering-off']).toBeUndefined();expect(()=>importDocx(exportDocx(reopened).bytes,{documentId:'reviewed',unsupportedContent:'reject'})).not.toThrow();
 session.undo();expect(session.getNode(paragraph)!.attributes['data-word-numbering']).toBe(definition.numberingId);
});
