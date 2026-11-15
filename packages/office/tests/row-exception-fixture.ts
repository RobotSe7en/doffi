import{strFromU8,strToU8,unzipSync,zipSync}from'fflate';
import{rowSpacingFixture}from'./row-spacing-fixture.js';
export function rowExceptionFixture():Uint8Array{
 const files=unzipSync(rowSpacingFixture());files['word/document.xml']=strToU8(strFromU8(files['word/document.xml']!).replaceAll('<w:tr>','<w:tr><w:tblPrEx><w:tblCellSpacing w:w="80" w:type="dxa"/></w:tblPrEx>'));return zipSync(files);
}
