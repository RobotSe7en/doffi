import{expect,test}from'vitest';
import{createWordDocument,openDocument,createDocumentSession}from'@doffi/core';
import type{DocumentOperation}from'@doffi/core';
function create(){const session=createDocumentSession(openDocument(createWordDocument({documentId:'tables'}),{kind:'document'}));const apply=(operations:DocumentOperation[])=>session.commit(session.prepare({protocolVersion:1,documentId:'tables',baseRevision:session.snapshot().manifest.revision,operations}));apply([{op:'insertWordTable',nodeId:'paragraph-1',position:'after',rows:3,columns:3}]);const tableId=session.snapshot().metadata.word!.tables[0]!.nodeId;return{session,apply,tableId};}
test('Word table insertion, merged structure changes, resizing and undo preserve identities and content',()=>{
 const{session,apply,tableId}=create();apply([{op:'editWordTable',nodeId:tableId,change:{kind:'setCells',row:0,column:0,values:[['A','B','C'],['D','E','F'],['G','H','I']]}}]);
 const before=session.snapshot();apply([{op:'editWordTable',nodeId:tableId,change:{kind:'mergeCells',row:0,column:0,rows:2,columns:2}}]);expect(session.getNode(tableId)!.html).toContain('rowspan="2"');expect(session.getNode(tableId)!.text).toContain('ABDE');
 apply([{op:'editWordTable',nodeId:tableId,change:{kind:'insertColumns',at:1,count:1}},{op:'editWordTable',nodeId:tableId,change:{kind:'insertRows',at:1,count:1}}]);expect(session.snapshot().metadata.word!.tables[0]!.grid).toHaveLength(4);
 session.undo();session.undo();expect(session.snapshot().html).toBe(before.html);
 apply([{op:'editWordTable',nodeId:tableId,change:{kind:'resize',grid:[100,120,140]}}]);expect(session.getNode(tableId)!.html).toContain('width:360pt');
});
test('merging then splitting keeps original content and creates a rectangular native grid',()=>{
 const{session,apply,tableId}=create();apply([{op:'editWordTable',nodeId:tableId,change:{kind:'mergeCells',row:0,column:0,rows:2,columns:3}}]);apply([{op:'editWordTable',nodeId:tableId,change:{kind:'splitCell',row:0,column:0}}]);expect(session.getNode(tableId)!.html).not.toContain('rowspan');expect(session.getNode(tableId)!.html).not.toContain('colspan');expect(openDocument(session.serialize(),{kind:'document'}).snapshot()).toEqual(session.snapshot());
});
test('invalid Word table batches are atomic and header and cell formatting share history',()=>{
 const{session,apply,tableId}=create();const before=session.snapshot();expect(()=>apply([{op:'formatWordTable',nodeId:tableId,repeatHeaderRows:1},{op:'editWordTable',nodeId:tableId,change:{kind:'deleteRows',at:0,count:3}}])).toThrow();expect(session.snapshot()).toEqual(before);
 apply([{op:'formatWordTable',nodeId:tableId,repeatHeaderRows:1,allowRowSplit:false}]);expect(session.getNode(tableId)!.html).toContain('data-word-repeat-header="true"');session.undo();expect(session.snapshot().html).toBe(before.html);
});
test('arbitrary cell subdivisions refine the native grid and preserve neighbouring cells and contents',()=>{
 const{session,apply,tableId}=create();apply([{op:'editWordTable',nodeId:tableId,change:{kind:'setCells',row:0,column:0,values:[['A','B','C'],['D','E','F'],['G','H','I']]}},{op:'editWordTable',nodeId:tableId,change:{kind:'mergeCells',row:0,column:0,rows:2,columns:2}}]);const before=session.snapshot(),text=session.getNode(tableId)!.text,width=before.metadata.word!.tables[0]!.grid.reduce((a,b)=>a+b,0);
 apply([{op:'editWordTable',nodeId:tableId,change:{kind:'splitCell',row:0,column:0,rows:3,columns:3}}]);const snapshot=session.snapshot(),grid=snapshot.metadata.word!.tables[0]!.grid;expect(grid).toHaveLength(5);expect(grid.reduce((a,b)=>a+b,0)).toBeCloseTo(width);expect(session.getNode(tableId)!.text).toBe(text);expect(openDocument(session.serialize(),{kind:'document'}).snapshot()).toEqual(snapshot);session.undo();expect(session.snapshot().html).toBe(before.html);session.redo();expect(session.getNode(tableId)!.text).toBe(text);
});
