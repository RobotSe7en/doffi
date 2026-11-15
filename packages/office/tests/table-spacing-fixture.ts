import{strFromU8,strToU8,unzipSync,zipSync}from'fflate';
import{tableBidiFixture}from'./table-bidi-fixture.js';
export function tableSpacingFixture(value='120',type='dxa'):Uint8Array{
 const files=unzipSync(tableBidiFixture('<w:bidiVisual w:val="0"/>'));files['word/document.xml']=strToU8(strFromU8(files['word/document.xml']!).replace('<w:tblInd','<w:tblCellSpacing w:w="'+value+'" w:type="'+type+'"/><w:tblInd'));return zipSync(files);
}
