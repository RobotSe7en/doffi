import{expect,test}from'vitest';
import{unzipSync,zipSync,strFromU8,strToU8}from'fflate';
import{createWordDocument,openDocument,createDocumentSession}from'@doffi/core';
import{importDocx,exportDocx}from'@doffi/office';
test('column balancing settings edit without losing retained compatibility flags or their official sequence',()=>{
 const original=openDocument(createWordDocument({documentId:'column-flags'}),{kind:'document'}),files=unzipSync(exportDocx(original).bytes);
 files['word/settings.xml']=strToU8(strFromU8(files['word/settings.xml']!).replace('</w:settings>','<w:compat><w:useSingleBorderforContiguousCells/><w:noColumnBalance/><w:doNotUseHTMLParagraphAutoSpacing/></w:compat></w:settings>'));
 const session=createDocumentSession(openDocument(importDocx(zipSync(files),{documentId:'column-flags'}),{kind:'document'})),root=session.getNode(session.snapshot().metadata.word!.stories.find(story=>story.kind==='main')!.rootNodeId)!.parentNodeId!;
 expect(session.snapshot().metadata.word!.noColumnBalance).toBe(true);session.commit(session.prepare({protocolVersion:1,documentId:'column-flags',baseRevision:0,operations:[{op:'setWordOptions',nodeId:root,noColumnBalance:false}]}));
 let bytes=exportDocx(session).bytes;for(let pass=0;pass<2;pass++){
  const xml=strFromU8(unzipSync(bytes)['word/settings.xml']!);expect(xml.match(/<w:noColumnBalance/g)).toHaveLength(1);expect(xml).toContain('<w:noColumnBalance w:val="0"');expect(xml.indexOf('useSingleBorderforContiguousCells')).toBeLessThan(xml.indexOf('noColumnBalance'));expect(xml.indexOf('noColumnBalance')).toBeLessThan(xml.indexOf('doNotUseHTMLParagraphAutoSpacing'));
  const reopened=openDocument(importDocx(bytes,{documentId:'column-reopen'}),{kind:'document'});expect(reopened.snapshot().metadata.word!.noColumnBalance).toBe(false);bytes=exportDocx(reopened).bytes;
 }
 session.undo();expect(session.snapshot().metadata.word!.noColumnBalance).toBe(true);
});
