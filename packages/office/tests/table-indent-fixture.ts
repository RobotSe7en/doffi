import { body, mutateDocx } from "./regression-fixtures.js";

/** Native values are twips; no SDK model or renderer produces this source. */
export function tableIndentFixture(options: {
  mode?: number;
  styled?: boolean;
  rtl?: boolean;
  cellOverride?: boolean;
  nested?: boolean;
  ignoredIndent?: boolean;
} = {}): Uint8Array {
  const indent = '<w:tblInd w:w="240" w:type="dxa"/>';
  const margins = '<w:tblCellMar><w:left w:w="140" w:type="dxa"/><w:right w:w="180" w:type="dxa"/></w:tblCellMar>';
  return mutateDocx("02-table-merges.docx", {
    "word/settings.xml": xml => xml.replace("</w:settings>", options.mode === undefined ? "</w:settings>"
      : '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="' + options.mode + '"/></w:compat></w:settings>'),
    "word/styles.xml": xml => xml.replace("</w:styles>", options.styled
      ? '<w:style w:type="table" w:styleId="UnrelatedBase"><w:name w:val="Custom Base"/><w:tblPr>' + indent + margins + '</w:tblPr></w:style>'
        + '<w:style w:type="table" w:styleId="DerivedGeometry"><w:name w:val="Custom Geometry"/><w:basedOn w:val="UnrelatedBase"/></w:style></w:styles>' : "</w:styles>"),
    "word/document.xml": xml => body(xml, '<w:p><w:r><w:t>REFERENCE</w:t></w:r></w:p><w:tbl><w:tblPr>'
      + (options.styled ? '<w:tblStyle w:val="DerivedGeometry"/>' : indent + margins)
      + (options.ignoredIndent ? '<w:tblInd w:w="5000" w:type="pct"/>' : "")
      + (options.rtl ? '<w:bidiVisual/>' : "")
      + '<w:tblW w:w="4000" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr>'
      + '<w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid><w:tr>'
      + '<w:tc><w:tcPr>' + (options.cellOverride ? '<w:tcMar><w:left w:w="100" w:type="dxa"/></w:tcMar>' : "")
      + '</w:tcPr><w:p><w:r><w:t>LEADING</w:t></w:r></w:p>'
      + (options.nested ? '<w:tbl><w:tblPr><w:tblW w:w="1000" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>INNER</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/>' : "")
      + '</w:tc>'
      + '<w:tc><w:p><w:r><w:t>TRAILING</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'),
  });
}
