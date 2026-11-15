import{body,mutateDocx,styles}from'./regression-fixtures.js';
export function contextualSpacingFixture(flag='1'):Uint8Array{
 const p=(value:string,style='Normal',extra='')=>'<w:p><w:pPr><w:pStyle w:val="'+style+'"/>'+extra+'</w:pPr><w:r><w:t>'+value+'</w:t></w:r></w:p>';
 return mutateDocx('01-text-numbering.docx',{'word/styles.xml':()=>styles('<w:style w:type="paragraph" w:styleId="Normal" w:default="1"><w:pPr><w:spacing w:before="240" w:after="360"/><w:contextualSpacing w:val="'+flag+'"/></w:pPr><w:rPr><w:sz w:val="24"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Other"><w:basedOn w:val="Normal"/></w:style>'),'word/document.xml':source=>body(source,p('A')+p('B')+p('C','Other')+'<w:tbl><w:tblPr><w:tblW w:type="dxa" w:w="3600"/></w:tblPr><w:tblGrid><w:gridCol w:w="3600"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:type="dxa" w:w="3600"/></w:tcPr>'+p('Y')+p('Z')+'</w:tc></w:tr></w:tbl>'+p('D'))});
}
