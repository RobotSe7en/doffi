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

test('track settings and unmodeled native compatibility flags survive repeated native export',()=>{
 const session=createDocumentSession(openDocument(createWordDocument({documentId:'settings-native'}),{kind:'document'})),files=unzipSync(exportDocx(session).bytes);files['word/settings.xml']=strToU8(strFromU8(files['word/settings.xml']!).replace(/<w:trackRevisions[^>]*\/>/,'').replace('</w:settings>','<w:trackRevisions/><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>'));let bytes:Uint8Array=zipSync(files);
 for(let i=0;i<2;i++){const model=openDocument(importDocxWithReport(bytes,{documentId:'preserved-settings'}).files,{kind:'document'});expect(model.snapshot().metadata.word!.trackChanges).toBe(true);bytes=exportDocx(model).bytes;const settings=strFromU8(unzipSync(bytes)['word/settings.xml']!);expect(settings).toMatch(/:trackRevisions\b/);expect(settings).toContain('compatibilityMode');expect(settings).toContain('15');}
});

test('unknown compatibility choices preserve their complete alternate and fallback with a declared raster cache',()=>{
 const files=unzipSync(new Uint8Array(readFileSync(new URL('../../../examples/docx/03-images.docx',import.meta.url))));
 const source=strFromU8(files['word/document.xml']!),drawing=source.match(/<w:drawing[\s>][\s\S]*?<\/w:drawing>/)![0];
 files['word/document.xml']=strToU8(source.replace(drawing,'<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:future"><mc:Choice Requires="x"><x:editableObject x:id="keep"/></mc:Choice><mc:Fallback>'+drawing+'</mc:Fallback></mc:AlternateContent>'));
 let model=openDocument(importDocx(zipSync(files),{documentId:'alternate'}),{kind:'document'});
 for(let pass=0;pass<2;pass++){
  const native=model.snapshot().metadata.word!.nativeContent!.fragments.find(item=>item.kind!=='property')!;expect(native.preview).toBeDefined();expect(model.snapshot().manifest.assets.some(asset=>asset.path===native.preview!.assetPath&&asset.mimeType==='image/png')).toBe(true);
  const output=exportDocx(model).bytes,xml=strFromU8(unzipSync(output)['word/document.xml']!);const frameIds=[...xml.matchAll(/<(?:[\w.-]+:)?docPr\s[^>]*?\bid="(\d+)"/g)].map(match=>match[1]);expect(new Set(frameIds).size).toBe(frameIds.length);expect(xml).toContain('AlternateContent');expect(xml).toContain('editableObject');expect(xml).toContain('Fallback');expect(xml).not.toContain('transient:native');model=openDocument(importDocx(output,{documentId:'alternate'}),{kind:'document'});
 }
});

test('Strict namespace text documents remain editable and export to the declared Transitional target',()=>{const files=base();for(const[path,bytes]of Object.entries(files))if(path.endsWith('.xml')||path.endsWith('.rels'))files[path]=strToU8(strFromU8(bytes).replaceAll('http://schemas.openxmlformats.org/wordprocessingml/2006/main','http://purl.oclc.org/ooxml/wordprocessingml/main').replaceAll('http://schemas.openxmlformats.org/officeDocument/2006/relationships','http://purl.oclc.org/ooxml/officeDocument/relationships'));const model=openDocument(importDocx(zipSync(files),{documentId:'strict-namespaces',unsupportedContent:'reject'}),{kind:'document'});expect(mainText(model)).toContain('第一项');expect(()=>importDocx(exportDocx(model).bytes,{documentId:'strict-reopen',unsupportedContent:'reject'})).not.toThrow();});

 test('repeated native properties stay distinct through repeated native saves',()=>{let bytes=fixture('<w:p><w:pPr><x:setting xmlns:x="urn:vendor" x:id="first"/><x:setting xmlns:x="urn:vendor" x:id="second"/></w:pPr><w:r><w:t>重复</w:t></w:r></w:p>');for(let pass=0;pass<2;pass++){const model=openDocument(importDocx(bytes,{documentId:'repeated'}),{kind:'document'});bytes=exportDocx(model).bytes;const xml=strFromU8(unzipSync(bytes)['word/document.xml']!);expect(xml).toContain('id="first"');expect(xml).toContain('id="second"');}});

test('native residual paths distinguish repeated numbering levels',()=>{const files=base(),source=strFromU8(files['word/numbering.xml']!);let serial=0;files['word/numbering.xml']=strToU8(source.replaceAll('</w:lvl>',()=>'<x:keep xmlns:x="urn:vendor" x:value="level-'+serial+++'"/></w:lvl>'));let bytes:Uint8Array=zipSync(files);for(let pass=0;pass<2;pass++){const model=openDocument(importDocx(bytes,{documentId:'level-paths'}),{kind:'document'});bytes=exportDocx(model).bytes;const xml=strFromU8(unzipSync(bytes)['word/numbering.xml']!),values:string[][]=[];let level:string[]|undefined;const parser=new SaxesParser({xmlns:true});parser.on('opentag',tag=>{if(tag.local==='lvl')level=[];if(tag.local==='keep')level?.push(Object.values(tag.attributes).find(attr=>attr.local==='value')!.value);});parser.on('closetag',tag=>{if(tag.local==='lvl'&&level){values.push(level);level=undefined;}});parser.write(xml).close();expect(values.filter(values=>values.length)).toEqual(Array.from({length:serial},(_,index)=>['level-'+index]));}});

test('native connector endpoints follow regenerated editable frame identities across native reopens',()=>{
 const session=createDocumentSession(openDocument(createWordDocument({documentId:'native-links'}),{kind:'document'})),geometry={x:0,y:0,width:80,height:40,rotation:0,flipH:false,flipV:false};
 const insert=(object:import('@doffi/core').DocumentOperation):void=>{session.commit(session.prepare({protocolVersion:1,documentId:'native-links',baseRevision:session.snapshot().manifest.revision,operations:[object]}));};
 insert({op:'insertWordDrawing',nodeId:'paragraph-1',offset:0,geometry,layout:{mode:'floating',wrap:'none'},object:{kind:'shape',preset:'rect'}});insert({op:'insertWordDrawing',nodeId:'paragraph-1',offset:1,geometry,layout:{mode:'floating',wrap:'none'},object:{kind:'shape',preset:'rect'}});
 const [start,end]=session.snapshot().metadata.word!.drawings!;insert({op:'insertWordDrawing',nodeId:'paragraph-1',offset:2,geometry,layout:{mode:'floating',wrap:'none'},object:{kind:'connector',connector:{routing:'straight',startNodeId:start!.nodeId,endNodeId:end!.nodeId}}});
 const files=unzipSync(exportDocx(session).bytes);let source=strFromU8(files['word/document.xml']!);source=source.replace(/\bid="([123])"/g,(_,id:string)=>'id="'+(Number(id)+40)+'"').replace('</wps:cNvCnPr><wps:spPr>','</wps:cNvCnPr><wps:spPr><x:future xmlns:x="urn:future"/>');files['word/document.xml']=strToU8(source);let bytes:Uint8Array=zipSync(files);
 for(let pass=0;pass<2;pass++){const model=openDocument(importDocx(bytes,{documentId:'remap'}),{kind:'document'});expect(model.snapshot().metadata.word!.drawings).toHaveLength(2);bytes=exportDocx(model).bytes;const xml=strFromU8(unzipSync(bytes)['word/document.xml']!);expect(xml).toMatch(/<a:stCxn[^>]*\bid="1"/);expect(xml).toMatch(/<a:endCxn[^>]*\bid="2"/);expect(xml).not.toContain('dn:identity');expect(xml).toContain('future');}
});

test('many unknown run properties recover together within a small declared work budget',()=>{const count=1800,bytes=fixture(Array.from({length:count},()=>'<w:p><w:r><w:rPr><x:future xmlns:x="urn:future"/></w:rPr><w:t>保留</w:t></w:r></w:p>').join(''));const result=importDocxWithReport(bytes,{documentId:'batched-native',limits:{maxNodes:50000}});expect(result.report.preserved).toHaveLength(count);expect(mainText(openDocument(result.files,{kind:'document',limits:{maxNodes:50000}}))).toContain('保留保留');});

test('extension attributes and revision provenance survive editable property changes without replacing their new values',()=>{let model=openDocument(importDocx(fixture('<w:p xmlns:x="urn:extension" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x" x:paragraph="keep" w:rsidR="00FF0012"><w:pPr x:property="keep"><w:jc w:val="left" x:alignment="keep"/></w:pPr><w:r x:run="keep"><w:rPr x:font="keep"/><w:t>扩展属性</w:t></w:r></w:p>'),{documentId:'attributes'}),{kind:'document'});const session=createDocumentSession(model),paragraph=[...model.snapshot().html.matchAll(/<p[^>]*data-node-id="([^"]+)"[^>]*>[\s\S]*?扩展属性/g)][0]![1]!;session.commit(session.prepare({protocolVersion:1,documentId:'attributes',baseRevision:0,operations:[{op:'setWordParagraph',nodeId:paragraph,properties:{'text-align':'right'}}]}));let bytes=exportDocx(session).bytes;for(let pass=0;pass<2;pass++){const xml=strFromU8(unzipSync(bytes)['word/document.xml']!);for(const value of ['paragraph="keep"','property="keep"','alignment="keep"','run="keep"','font="keep"','rsidR="00FF0012"'])expect(xml).toContain(value);expect(xml).toMatch(/<w:jc[^>]*w:val="right"/);expect(xml).not.toContain('transient:native');model=openDocument(importDocx(bytes,{documentId:'attributes'}),{kind:'document'});bytes=exportDocx(model).bytes;}});

test.each(['pict','vendor']as const)('many %s elements recover as a grammar class while preserving adjacent editable text',kind=>{const count=900,content=kind==='pict'?'<w:pict><v:shape xmlns:v="urn:schemas-microsoft-com:vml" id="_x0000_s1"/></w:pict>':'<x:widget xmlns:x="urn:vendor"/>';const bytes=fixture(Array.from({length:count},(_,index)=>'<w:p><w:r><w:t>前</w:t>'+content.replace('_x0000_s1','_x0000_s'+(1024+index))+'<w:t>后</w:t></w:r></w:p>').join(''));const result=importDocxWithReport(bytes,{documentId:'many-native',limits:{maxNodes:50000}});expect(result.report.preserved).toHaveLength(count);const model=openDocument(result.files,{kind:'document',limits:{maxNodes:50000}});expect(mainText(model)).toContain('前\ufffc后');});

test('property compatibility alternatives stay native properties instead of becoming nested run placeholders',()=>{const bytes=fixture('<w:p><w:r><w:rPr><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:future"><mc:Choice Requires="x"><x:bold/></mc:Choice><mc:Fallback><w:b/></mc:Fallback></mc:AlternateContent></w:rPr><w:t>可编辑</w:t></w:r></w:p>');let model=openDocument(importDocx(bytes,{documentId:'property-alternate'}),{kind:'document'});for(let pass=0;pass<2;pass++){expect(mainText(model)).toContain('可编辑');expect(mainText(model)).not.toContain('\ufffc');const output=exportDocx(model).bytes;expect(strFromU8(unzipSync(output)['word/document.xml']!)).toContain('AlternateContent');model=openDocument(importDocx(output,{documentId:'property-alternate'}),{kind:'document'});}});

test('VML type references and OLE shape aliases follow renumbered native identities',()=>{let bytes=fixture('<w:p><w:r><w:object xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"><v:shapetype id="_x0000_t75" o:spt="75"/><v:shape id="PictureName" o:spid="_x0000_s1025" type="#_x0000_t75"/><o:OLEObject ShapeID="_x0000_s1025"/></w:object></w:r></w:p>');for(let pass=0;pass<2;pass++){const model=openDocument(importDocx(bytes,{documentId:'vml'}),{kind:'document'});bytes=exportDocx(model).bytes;const xml=strFromU8(unzipSync(bytes)['word/document.xml']!),ids:{type?:string;shape?:string;alias?:string;reference?:string;ole?:string}={};const parser=new SaxesParser({xmlns:true});parser.on('opentag',tag=>{const attrs=Object.fromEntries(Object.values(tag.attributes).map(attr=>[attr.local,attr.value]));if(tag.local==='shapetype')ids.type=attrs.id!;if(tag.local==='shape'){ids.shape=attrs.id!;ids.alias=attrs.spid!;ids.reference=attrs.type!;}if(tag.local==='OLEObject')ids.ole=attrs.ShapeID!;});parser.write(xml).close();expect(ids.shape).not.toBe('PictureName');expect(ids.alias).toBe(ids.shape);expect(ids.ole).toBe(ids.shape);expect(ids.reference).toBe('#'+ids.type);}});

test('valid numeric fitText is preserved instead of being rejected as a malformed Boolean',()=>{const bytes=fixture('<w:p><w:r><w:rPr><w:fitText w:val="2000" w:id="7"/></w:rPr><w:t>宽度</w:t></w:r></w:p>'),model=openDocument(importDocx(bytes,{documentId:'fit-text'}),{kind:'document'});expect(mainText(model)).toContain('宽度');expect(strFromU8(unzipSync(exportDocx(model).bytes)['word/document.xml']!)).toMatch(/<w:fitText[^>]*w:val="2000"/);});
