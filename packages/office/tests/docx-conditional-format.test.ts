import {expect,test}from'vitest';
import{unzipSync,strFromU8}from'fflate';
import{openDocument,createDocumentSession}from'@doffi/core';
import type{DocumentOperation}from'@doffi/core';
import{importDocx,exportDocx}from'@doffi/office';
import{conditionalFormatFixture}from'./table-style-fixture.js';
function setup(bytes=conditionalFormatFixture()){
 const session=createDocumentSession(openDocument(importDocx(bytes,{documentId:'cnf',unsupportedContent:'reject'}),{kind:'document'})),apply=(operations:DocumentOperation[])=>session.commit(session.prepare({protocolVersion:1,documentId:'cnf',baseRevision:session.snapshot().manifest.revision,operations}));return{session,apply};
}
test('native conditional format caches retain presence, conflicting mask/named OFF flags and vertical merge continuation paragraphs through two reopens',()=>{
 let bytes=conditionalFormatFixture();for(let pass=0;pass<3;pass++){
 const{session}=setup(bytes),word=session.snapshot().metadata.word!;expect(word.conditionalFormats).toHaveLength(4);expect(word.conditionalFormats!.find(item=>session.getNode(item.nodeId)!.tagName==='tr')!.format).toEqual({val:'100000000000',firstRow:false});expect(word.conditionalFormats!.find(item=>session.getNode(item.nodeId)!.tagName==='p')!.format).toEqual({});expect(word.tables[0]!.continuationFormats).toMatchObject([{row:1,column:0,format:{val:'000000000000',lastRow:true},paragraphs:[{firstRow:false},null]}]);
 bytes=exportDocx(session).bytes;expect(strFromU8(unzipSync(bytes)['word/document.xml']!).match(/<w:cnfStyle/g)).toHaveLength(6);session.dispose();
 }
});
test('conditional format history survives native row/cell/paragraph changes and rejecting restores original flags',()=>{
 const{session,apply}=setup(),word=session.snapshot().metadata.word!,root=session.getNode(word.stories.find(story=>story.kind==='main')!.rootNodeId)!.parentNodeId!,cell=word.conditionalFormats!.find(item=>session.getNode(item.nodeId)!.tagName==='td')!.nodeId,paragraph=word.conditionalFormats!.find(item=>session.getNode(item.nodeId)!.tagName==='p')!.nodeId;
 apply([{op:'setWordReview',nodeId:root,trackChanges:true,author:'甲'},{op:'formatWordCell',nodeId:cell,properties:{'background-color':'#00ff00'}},{op:'setWordParagraph',nodeId:paragraph,properties:{'text-align':'right'}},{op:'formatWordTable',nodeId:word.tables[0]!.nodeId,allowRowSplit:false}]);
 const reopened=createDocumentSession(openDocument(importDocx(exportDocx(session).bytes,{documentId:'cnf-reopen',unsupportedContent:'reject'}),{kind:'document'}));expect(reopened.snapshot().metadata.word!.revisions.find(item=>item.property==='row')!.before!.conditionalFormat).toEqual({val:'100000000000',firstRow:false});expect(reopened.snapshot().metadata.word!.revisions.find(item=>item.property==='cell')!.before!.conditionalFormat).toMatchObject({lastColumn:false});const snapshot=reopened.snapshot(),main=snapshot.metadata.word!.stories.find(story=>story.kind==='main')!;reopened.commit(reopened.prepare({protocolVersion:1,documentId:'cnf-reopen',baseRevision:0,operations:[{op:'reviewWordRevisions',nodeId:reopened.getNode(main.rootNodeId)!.parentNodeId!,decision:'reject'}]}));expect(reopened.snapshot().metadata.word!.conditionalFormats).toHaveLength(4);
});
test('cloned header caches are isolated and vertical continuation coordinates follow insertion, actual row review and undo',()=>{
 const{session,apply}=setup(conditionalFormatFixture(true)),word=session.snapshot().metadata.word!,header=word.stories.find(story=>story.kind==='header')!,section=session.snapshot().metadata.sections![0]!;apply([{op:'cloneWordStory',nodeId:section.rootNodeId,storyId:header.storyId,variant:'default'}]);
 const current=session.snapshot(),cloned=current.metadata.sections![0]!.headers!.default!,originalTable=word.tables.find(table=>session.getNode(table.nodeId)!.parentNodeId===header.rootNodeId)!,cloneTable=current.metadata.word!.tables.find(table=>session.getNode(table.nodeId)!.parentNodeId===cloned)!;
 apply([{op:'editWordTable',nodeId:cloneTable.nodeId,change:{kind:'insertRows',at:0,count:1}}]);expect(session.snapshot().metadata.word!.tables.find(table=>table.nodeId===originalTable.nodeId)!.continuationFormats![0]!.row).toBe(1);expect(session.snapshot().metadata.word!.tables.find(table=>table.nodeId===cloneTable.nodeId)!.continuationFormats![0]!.row).toBe(2);session.undo();expect(session.snapshot().metadata.word!.tables.find(table=>table.nodeId===cloneTable.nodeId)!.continuationFormats![0]!.row).toBe(1);
 session.redo();const root=session.getNode(word.stories.find(story=>story.kind==='main')!.rootNodeId)!.parentNodeId!;apply([{op:'setWordReview',nodeId:root,trackChanges:true,author:'甲'},{op:'editWordTable',nodeId:cloneTable.nodeId,change:{kind:'deleteRows',at:0,count:1}}]);apply([{op:'reviewWordRevisions',nodeId:root,decision:'accept'}]);expect(session.snapshot().metadata.word!.tables.find(table=>table.nodeId===cloneTable.nodeId)!.continuationFormats![0]!.row).toBe(1);expect(session.snapshot().metadata.word!.tables.find(table=>table.nodeId===originalTable.nodeId)!.continuationFormats![0]!.row).toBe(1);
});
