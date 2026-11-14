import { readFileSync } from 'node:fs';
import { SaxesParser } from 'saxes';
import { expect, test } from 'vitest';
import { unzipSync, zipSync, strToU8, strFromU8 } from 'fflate';
import { createWordDocument, createDocumentSession, openDocument } from '@doffi/core';
import { importDocx, importDocxWithReport, exportDocx, readDocumentZip, writeDocumentZip } from '@doffi/office';
const base = () => unzipSync(new Uint8Array(readFileSync(new URL('../../../examples/docx/01-text-numbering.docx', import.meta.url))));
function fixture(content: string): Uint8Array {
  const files=base();files['word/document.xml']=strToU8(strFromU8(files['word/document.xml']!).replace('<w:body>','<w:body>'+content));return zipSync(files);
}
function mainText(model: ReturnType<typeof openDocument>): string {
  if(model.kind!=='document')throw new Error('Expected Word');const main=model.snapshot().metadata.word!.stories.find(s=>s.kind==='main')!;return model.getNode(main.rootNodeId)!.text;
}
test('preserves unknown inline content while surrounding text remains editable across two native saves',()=>{
  const bytes=fixture('<w:p><w:r><w:t>before</w:t></w:r><m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:r><m:rPr><m:brk m:alnAt="1"/></m:rPr><m:t>x+1</m:t></m:r></m:oMath><w:r><w:t>after</w:t></w:r></w:p>');
  const {files,report}=importDocxWithReport(bytes,{documentId:'native'});
  expect(report.preserved).toHaveLength(1);expect(report.preserved[0]!.kind).toBe('inline');
  const model=openDocument(files,{kind:'document'}),session=createDocumentSession(model);
  const run=[...model.snapshot().html.matchAll(/data-node-id="([^"]+)"[^>]*>before</g)][0]?.[1];expect(run).toBeDefined();
  session.commit(session.prepare({protocolVersion:1,documentId:'native',baseRevision:0,operations:[{op:'replaceText',nodeId:run!,from:0,to:6,expectedText:'before',text:'edited'}]}));
  session.undo();session.redo();
  let exported=exportDocx(session).bytes;
  for(let pass=0;pass<2;pass++){
    const output=unzipSync(exported),xml=strFromU8(output['word/document.xml']!);
    expect(xml).toContain('oMath');expect(xml).toContain('x+1');expect(xml).toContain('edited');expect(xml).not.toContain('transient:native');
    const reopened=openDocument(importDocx(exported,{documentId:'reopened'}),{kind:'document'});
    expect(mainText(reopened)).toContain('edited\ufffcafter');exported=exportDocx(reopened).bytes;
  }
  const saved=openDocument(readDocumentZip(writeDocumentZip(session.serialize())),{kind:'document'});expect(mainText(saved)).toContain('edited\ufffcafter');
});
test('preserves unsupported paragraph properties rather than losing the paragraph',()=>{
  const bytes=fixture('<w:p><w:pPr><w:overflowPunct w:val="true"/></w:pPr><w:r><w:t>悬挂标点</w:t></w:r></w:p>');
  const model=openDocument(importDocx(bytes,{documentId:'native-props'}),{kind:'document'});
  expect(mainText(model)).toContain('悬挂标点');
  const output=unzipSync(exportDocx(model).bytes);expect(strFromU8(output['word/document.xml']!)).toContain('overflowPunct');
  expect(()=>importDocx(bytes,{documentId:'strict',unsupportedContent:'reject'})).toThrowError(expect.objectContaining({code:'UnsupportedFeature'}));
});
test('preserves an unknown block and its embedded binary relationship',()=>{
  const files=base();files['word/document.xml']=strToU8(strFromU8(files['word/document.xml']!).replace('<w:body>','<w:body><x:widget xmlns:x="urn:vendor" r:id="nativeId"/>'));
  files['word/_rels/document.xml.rels']=strToU8(strFromU8(files['word/_rels/document.xml.rels']!).replace('</Relationships>','<Relationship Id="nativeId" Type="urn:vendor:data" Target="media/opaque.bin"/></Relationships>'));
  files['[Content_Types].xml']=strToU8(strFromU8(files['[Content_Types].xml']!).replace('</Types>','<Override PartName="/word/media/opaque.bin" ContentType="application/octet-stream"/></Types>'));
  files['word/media/opaque.bin']=new Uint8Array([0,1,2,3,255]);
  const model=openDocument(importDocx(zipSync(files),{documentId:'binary'}),{kind:'document'});
  const output=unzipSync(exportDocx(model).bytes);
  expect(Object.values(output).some(b=>b.length===5&&b[4]===255)).toBe(true);
  expect(strFromU8(output['word/document.xml']!)).toContain('widget');
  expect(()=>importDocx(zipSync(files),{documentId:'strict',unsupportedContent:'reject'})).toThrow();
});
test('malformed XML remains rejected in preserving mode',()=>{
  const files=base();files['word/document.xml']=strToU8('<!DOCTYPE a [<!ENTITY e SYSTEM "file:///etc/passwd">]><a>&e;</a>');
  expect(()=>importDocx(zipSync(files),{documentId:'bad'})).toThrow();
});
test('preserves section columns and document background at their native owners',()=>{
  const files=base();let source=strFromU8(files['word/document.xml']!);
  source=source.replace('<w:body>','<w:background w:color="BEEFFF"/><w:body>').replace('<w:sectPr>','<w:sectPr><w:cols w:num="2"/>');files['word/document.xml']=strToU8(source);
  const model=openDocument(importDocx(zipSync(files),{documentId:'layout'}),{kind:'document'}),output=unzipSync(exportDocx(model).bytes),xml=strFromU8(output['word/document.xml']!);
  expect(xml).toContain('background');expect(xml).toContain('BEEFFF');expect(xml).toContain('cols');expect(xml).toContain('num="2"');
  expect(()=>importDocx(exportDocx(model).bytes,{documentId:'layout-again'})).not.toThrow();
});
test('keeps recovered run properties while preserving only its unsupported native child',()=>{
  const bytes=fixture('<w:p><w:r><w:rPr><w:unknownProperty/></w:rPr><w:pict><x:shape xmlns:x="urn:vendor"/></w:pict></w:r></w:p>');
  const model=openDocument(importDocx(bytes,{documentId:'nested'}),{kind:'document'});
  const out=strFromU8(unzipSync(exportDocx(model).bytes)['word/document.xml']!);
  expect(out).toContain('unknownProperty');expect(out).toContain('shape');expect(model.snapshot().metadata.word!.nativeContent!.fragments).toHaveLength(2);
});
test('unsupported run children retain editable adjacent text and balanced field controls',()=>{
 const bytes=fixture('<w:p><w:r><w:rPr><w:unknownProperty/></w:rPr><w:fldChar w:fldCharType="begin"/><w:instrText> IF 1 = 1 "yes" "no" </w:instrText><w:fldChar w:fldCharType="separate"/><w:t>before</w:t><w:pict><x:shape xmlns:x="urn:vendor"/></w:pict><w:t>after</w:t><w:fldChar w:fldCharType="end"/></w:r></w:p>');
 const model=openDocument(importDocx(bytes,{documentId:'narrow-native'}),{kind:'document'});expect(mainText(model)).toContain('before\ufffcafter');const session=createDocumentSession(model),run=[...model.snapshot().html.matchAll(/data-node-id="([^"]+)"[^>]*>before</g)][0]?.[1];expect(run).toBeDefined();session.commit(session.prepare({protocolVersion:1,documentId:'narrow-native',baseRevision:0,operations:[{op:'replaceText',nodeId:run!,from:0,to:6,expectedText:'before',text:'edited'}]}));const result=exportDocx(session),xml=strFromU8(unzipSync(result.bytes)['word/document.xml']!);expect(xml).toContain('unknownProperty');const stack:string[]=[];let found=false;const parser=new SaxesParser({xmlns:true});parser.on('opentag',tag=>{if(tag.local==='pict'){expect(stack.at(-1)).toBe('r');found=true;}stack.push(tag.local);});parser.on('closetag',()=>{stack.pop();});parser.write(xml).close();expect(found).toBe(true);expect(xml).not.toContain('transient:native');expect(mainText(openDocument(importDocx(result.bytes,{documentId:'narrow-again'}),{kind:'document'}))).toContain('edited\ufffcafter');
});
test('keeps recovered style properties inside the owning style',()=>{
  const files=base();files['word/styles.xml']=strToU8(strFromU8(files['word/styles.xml']!).replace('</w:style>','<w:pPr><w:unknownStyleProperty/></w:pPr></w:style>'));
  const model=openDocument(importDocx(zipSync(files),{documentId:'style'}),{kind:'document'});
  const out=strFromU8(unzipSync(exportDocx(model).bytes)['word/styles.xml']!);
  expect(out).toContain('unknownStyleProperty');
  expect(()=>importDocx(exportDocx(model).bytes,{documentId:'style-again'})).not.toThrow();
});
test('rejects partial native object edits atomically but permits whole object deletion and undo',()=>{
  const model=openDocument(importDocx(fixture('<w:p><m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:r><m:rPr><m:brk m:alnAt="1"/></m:rPr><m:t>x</m:t></m:r></m:oMath></w:p>'),{documentId:'atomic'}),{kind:'document'}),session=createDocumentSession(model);
  const object=model.snapshot().metadata.word!.nativeContent!.fragments.find(f=>f.kind==='inline')!;
  expect(()=>session.prepare({protocolVersion:1,documentId:'atomic',baseRevision:0,operations:[{op:'replaceText',nodeId:object.nodeId!,from:0,to:1,expectedText:'\ufffc',text:'broken'}]})).toThrow();
  expect(session.snapshot()).toEqual(model.snapshot());
  session.commit(session.prepare({protocolVersion:1,documentId:'atomic',baseRevision:0,operations:[{op:'removeNode',nodeId:object.nodeId!}]}));
  expect(strFromU8(unzipSync(exportDocx(session).bytes)['word/document.xml']!)).not.toContain('oMath');
  session.undo();expect(strFromU8(unzipSync(exportDocx(session).bytes)['word/document.xml']!)).toContain('oMath');
});
test('modern comment extension attributes cannot enter a placeholder retry loop',()=>{
  const input=new Uint8Array(readFileSync(new URL('../../../examples/docx/08-comments-revisions.docx',import.meta.url)));
  const model=openDocument(importDocx(input,{documentId:'classic'}),{kind:'document'}),parts=unzipSync(exportDocx(model).bytes);
  parts['word/commentsExtended.xml']=strToU8(strFromU8(parts['word/commentsExtended.xml']!).replace('<w15:commentEx ','<w15:commentEx w15:future="retained" '));
  const imported=openDocument(importDocx(zipSync(parts),{documentId:'modern-extension'}),{kind:'document'});
  expect(imported.snapshot().metadata.word!.nativeContent!.fragments.some(fragment=>fragment.ownerName==='commentEx')).toBe(true);
  expect(mainText(imported)).toBe(mainText(model));
  const exported=exportDocx(imported).bytes;expect(strFromU8(unzipSync(exported)['word/commentsExtended.xml']!)).toContain('future="retained"');
  expect(()=>importDocx(exported,{documentId:'modern-extension-again'})).not.toThrow();
});
test('more than 1024 unsupported native content nodes are accepted within normal document limits',()=>{
  const content='<w:p>'+Array.from({length:1025},(_,index)=>'<x:widget xmlns:x="urn:future" index="'+index+'"/>').join('')+'</w:p>';
  const model=openDocument(importDocx(fixture(content),{documentId:'many-native'}),{kind:'document'});
  expect(model.snapshot().metadata.word!.nativeContent!.fragments).toHaveLength(1025);
});

test('unmodeled document settings and native compatibility flags survive repeated native export',()=>{
 const session=createDocumentSession(openDocument(createWordDocument({documentId:'settings-native'}),{kind:'document'})),files=unzipSync(exportDocx(session).bytes);files['word/settings.xml']=strToU8(strFromU8(files['word/settings.xml']!).replace('</w:settings>','<w:trackRevisions/><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>'));let bytes:Uint8Array=zipSync(files);
 for(let i=0;i<2;i++){const model=openDocument(importDocxWithReport(bytes,{documentId:'preserved-settings'}).files,{kind:'document'});bytes=exportDocx(model).bytes;const settings=strFromU8(unzipSync(bytes)['word/settings.xml']!);expect(settings).toMatch(/:trackRevisions\b/);expect(settings).toContain('compatibilityMode');expect(settings).toContain('15');}
});
