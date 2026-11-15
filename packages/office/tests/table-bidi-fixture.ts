import { body, mutateDocx, styles } from "./regression-fixtures.js";
export function tableBidiFixture(override = ""): Uint8Array {
  return mutateDocx("02-table-merges.docx", {
    "word/styles.xml": () =>
      styles(
        '<w:style w:type="paragraph" w:styleId="Normal" w:default="1"/><w:style w:type="table" w:styleId="TableNormal" w:default="1"/><w:style w:type="table" w:styleId="RTL"><w:tblPr><w:bidiVisual/><w:tblStyleRowBandSize w:val="2"/></w:tblPr></w:style>',
      ),
    "word/document.xml": (source) =>
      body(
        source,
        '<w:tbl><w:tblPr><w:tblStyle w:val="RTL"/>' +
          override +
          '<w:tblStyleRowBandSize w:val="2"/><w:tblInd w:w="600" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="1600"/><w:gridCol w:w="2400"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcBorders><w:left w:val="single" w:sz="16" w:color="000088"/><w:right w:val="double" w:sz="8" w:color="880000"/></w:tcBorders><w:tcMar><w:left w:w="180" w:type="dxa"/><w:right w:w="60" w:type="dxa"/></w:tcMar></w:tcPr><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p><w:tbl><w:tblPr><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="400"/><w:gridCol w:w="400"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>L</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>R</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/></w:tc></w:tr><w:tr><w:tc><w:p><w:r><w:t>C</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>D</w:t></w:r></w:p></w:tc></w:tr></w:tbl>',
      ),
  });
}
