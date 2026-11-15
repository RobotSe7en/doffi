import {mutateDocx,body,styles} from './regression-fixtures.js';
export function rowGapFixture(alignment=false){
 const cell=(text:string)=>'<w:tc><w:tcPr><w:tcW w:type="dxa" w:w="1440"/></w:tcPr><w:p><w:r><w:t>'+text+'</w:t></w:r></w:p></w:tc>';
 return mutateDocx('01-text-numbering.docx',{'word/styles.xml':()=>styles('<w:style w:type="paragraph" w:styleId="Normal" w:default="1"/>'),'word/document.xml':source=>body(source,'<w:tbl><w:tblPr><w:tblW w:type="dxa" w:w="4320"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid>'+('<w:gridCol w:w="1440"/>'.repeat(3))+'</w:tblGrid><w:tr>'+['A','B','C'].map(cell).join('')+'</w:tr><w:tr><w:trPr>'+ (alignment?'<w:jc w:val="center"/>':'') +'<w:gridBefore w:val="1"/><w:gridAfter w:val="1"/><w:wBefore w:w="1440" w:type="dxa"/><w:wAfter w:w="25%" w:type="pct"/></w:trPr>'+cell('X')+'</w:tr></w:tbl>')});
}
