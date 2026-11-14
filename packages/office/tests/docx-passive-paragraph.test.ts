import { expect, test } from 'vitest';
import { createDocumentSession, openDocument } from '@doffi/core';
import { importDocx, exportDocx } from '@doffi/office';
import { strFromU8, unzipSync } from 'fflate';
import { mutateDocx, body, styles } from './regression-fixtures.js';

function fixture(properties = '<w:suppressLineNumbers/><w:bidi w:val="0"/><w:textDirection w:val="lrTb"/><w:divId w:val="42"/>') {
    return mutateDocx('01-text-numbering.docx', {
        'word/styles.xml': () => styles('<w:style w:type="paragraph" w:styleId="Normal" w:default="1"/>'),
        'word/document.xml': source => body(source,'<w:p><w:pPr><w:pStyle w:val="Normal"/>'+properties+'<w:spacing w:before="60"/></w:pPr><w:r><w:t>段落属性</w:t></w:r></w:p>'),
    });
}
function current(bytes: Uint8Array, strict = true) {
    const session=createDocumentSession(openDocument(importDocx(bytes,{documentId:'paragraph-flags',...(strict?{unsupportedContent:'reject' as const}:{})}),{kind:'document'})),nodes=[session.snapshot().metadata.word!.stories[0]!.rootNodeId];
    for(let index=0;index<nodes.length;index++)nodes.push(...session.getNode(nodes[index]!)!.childNodeIds);
    return{session,paragraph:nodes.find(id=>session.getNode(id)!.tagName==='p')!,root:session.getNode(nodes[0]!)!.parentNodeId!};
}
test('passive native paragraph flags survive tracked format editing, undo and two strict reopens',()=>{
    const{session,paragraph,root}=current(fixture()),before=session.snapshot();
    session.commit(session.prepare({protocolVersion:1,documentId:'paragraph-flags',baseRevision:0,operations:[{op:'setWordReview',nodeId:root,trackChanges:true,author:'A'},{op:'setWordParagraph',nodeId:paragraph,properties:{'margin-left':'18pt'}}]}));
    let bytes=exportDocx(session).bytes;
    for(let pass=0;pass<2;pass++){
        const xml=strFromU8(unzipSync(bytes)['word/document.xml']!);for(const[name,value]of [['suppressLineNumbers','1'],['bidi','0'],['textDirection','lrTb'],['divId','42']])expect(xml).toMatch(new RegExp('<w:'+name+'\\b[^>]*w:val="'+value+'"'));
        const reopened=current(bytes);expect(reopened.session.snapshot().metadata.word!.revisions[0]!.property).toBe('paragraph');expect(reopened.session.getNode(reopened.paragraph)!.attributes.style).toContain('margin-left:18pt');bytes=exportDocx(reopened.session).bytes;
    }
    session.undo();expect(session.snapshot().html).toBe(before.html);
});
test('native bidi preservation retains surrounding editable paragraph properties and accepts later independent formatting',()=>{
    const{session,paragraph}=current(fixture('<w:bidi/>'),false),fragments=session.snapshot().metadata.word!.nativeContent!.fragments;
    expect(fragments.some(item=>item.propertyName?.endsWith(' bidi'))).toBe(true);expect(fragments.some(item=>item.propertyName?.endsWith(' pPr'))).toBe(false);expect(session.getNode(paragraph)!.attributes.style).toContain('margin-top:3pt');
    session.commit(session.prepare({protocolVersion:1,documentId:'paragraph-flags',baseRevision:0,operations:[{op:'setWordParagraph',nodeId:paragraph,properties:{'margin-left':'18pt'}}]}));
    let bytes=exportDocx(session).bytes;for(let pass=0;pass<2;pass++){const xml=strFromU8(unzipSync(bytes)['word/document.xml']!);expect(xml).toMatch(/<w:bidi(?:\s|\/|>)/);expect(xml).toContain('w:left="360"');expect(xml.indexOf('<w:bidi')).toBeLessThan(xml.indexOf('<w:spacing'));const reopened=current(bytes,false);expect(reopened.session.getNode(reopened.paragraph)!.attributes.style).toContain('margin-left:18pt');bytes=exportDocx(reopened.session).bytes;}
});
for(const properties of ['<w:suppressLineNumbers w:val="maybe"/>','<w:bidi w:val="maybe"/>','<w:divId w:val="text"/>','<w:divId w:val="1.5"/>','<w:divId w:val="-1"/>'])test('invalid passive paragraph values remain malformed input '+properties,()=>{
    expect(()=>importDocx(fixture(properties),{documentId:'invalid-paragraph'})).toThrowError(expect.objectContaining({code:'InvalidDocument'}));
});
