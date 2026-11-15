import{body,mutateDocx,styles}from'./regression-fixtures.js';
export function nativeWordIdFixture():Uint8Array{return mutateDocx('01-text-numbering.docx',{
 'word/styles.xml':()=>styles('<w:style w:type="paragraph" w:styleId="Normal" w:default="1"/>'),
 'word/document.xml':source=>body(source,'<w:customXml w:element="payload" w:uri="urn:doffi:test"><w:p><w:bookmarkStart w:id="0" w:name="Hidden"/><w:permStart w:id="0" w:edGrp="everyone"/><w:ins w:id="0" w:author="Original" w:date="2026-11-15T00:00:00Z"><w:r><w:t>PRIVATE</w:t></w:r></w:ins><w:permEnd w:id="0"/><w:bookmarkEnd w:id="0"/></w:p></w:customXml><w:p><w:bookmarkStart w:id="7" w:name="Visible"/><w:r><w:t>VISIBLE</w:t></w:r><w:bookmarkEnd w:id="7"/></w:p>'),
});}
