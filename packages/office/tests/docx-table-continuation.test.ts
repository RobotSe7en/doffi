import{expect,test}from'vitest';
import{strFromU8,unzipSync}from'fflate';
import{createDocumentSession,openDocument}from'@doffi/core';
import{exportDocx,importDocx,importDocxWithReport}from'@doffi/office';
import{body,mutateDocx}from'./regression-fixtures.js';
function fixture(properties:string,content='<w:p/>'):Uint8Array{return mutateDocx('02-table-merges.docx',{'word/document.xml':source=>body(source,'<w:p><w:r><w:t>BEFORE</w:t></w:r></w:p>'+source.match(/<w:tbl>[\s\S]*?<\/w:tbl>/)![0].replace(/<w:tc><w:tcPr><w:vMerge\/><\/w:tcPr>[\s\S]*?<\/w:tc>/,'<w:tc><w:tcPr>'+(properties.startsWith('<w:tcW')?properties+'<w:vMerge/>':'<w:vMerge/>'+properties)+'</w:tcPr>'+content+'</w:tc>')+'<w:p><w:r><w:t>AFTER</w:t></w:r></w:p>')});}
test.each([
 ['shading','<w:shd w:fill="ABC123"/>','<w:p/>','ABC123'],
 ['width','<w:tcW w:w="1440" w:type="dxa"/>','<w:p/>','1440'],
 ['border','<w:tcBorders><w:bottom w:val="double" w:sz="12" w:color="123ABC"/></w:tcBorders>','<w:p/>','123ABC'],
 ['paragraph format','','<w:p><w:pPr><w:jc w:val="right"/></w:pPr></w:p>','w:val="right"'],
 ['run format','','<w:p><w:r><w:rPr><w:color w:val="FED123"/></w:rPr><w:t/></w:r></w:p>','FED123'],
 ['unknown cell property','<w:cellIns w:id="41" w:author="author"/>','<w:p/>','cellIns'],
])('preserves independent vertical continuation %s instead of discarding or attaching it to a wrong cell',(_kind,properties,content,retained)=>{
 const bytes=fixture(properties,content);expect(()=>importDocx(bytes,{documentId:'strict',unsupportedContent:'reject'})).toThrowError(expect.objectContaining({code:'UnsupportedFeature'}));
 const imported=importDocxWithReport(bytes,{documentId:'continuation'});expect(imported.report.preserved.some(item=>item.kind==='block')).toBe(true);const session=createDocumentSession(openDocument(imported.files,{kind:'document'})),main=session.snapshot().metadata.word!.stories.find(story=>story.kind==='main')!,paragraph=session.getNode(main.rootNodeId)!.childNodeIds.flatMap(id=>session.getNode(id)!.tagName==='section'?session.getNode(id)!.childNodeIds:[id]).find(id=>session.getNode(id)!.text==='BEFORE')!;
 session.commit(session.prepare({protocolVersion:1,documentId:'continuation',baseRevision:0,operations:[{op:'replaceText',nodeId:paragraph,from:0,to:6,expectedText:'BEFORE',text:'EDITED'}]}));let output=exportDocx(session).bytes;
 for(let pass=0;pass<2;pass++){const xml=strFromU8(unzipSync(output)['word/document.xml']!);expect(xml).toContain(retained);expect(xml).toContain('EDITED');expect(xml).toContain('AFTER');output=exportDocx(openDocument(importDocx(output,{documentId:'continuation-reopen'}),{kind:'document'})).bytes;}
 session.undo();expect(session.getNode(paragraph)!.text).toBe('BEFORE');
});
test('multiple unformatted continuation paragraphs retain their native count without making the table opaque',()=>{
 let bytes=fixture('','<w:p/><w:p/><w:p/>');for(let pass=0;pass<3;pass++){const model=openDocument(importDocx(bytes,{documentId:'empty-continuation',unsupportedContent:'reject'}),{kind:'document'});expect(model.snapshot().metadata.word!.nativeContent).toBeUndefined();expect(model.snapshot().metadata.word!.tables[0]!.continuationFormats).toEqual([{row:2,column:0,paragraphs:[null,null,null]}]);bytes=exportDocx(model).bytes;}
});
