import { body, mutateDocx } from "./regression-fixtures.js";

/** Independent native source: no SDK export supplies its tables or declarations. */
export function rowExceptionScopeFixture(mode: 14 | 15, kind: "indent" | "width" | "align" | "first-indent" | "first-width" | "first-align", merged = false, directAlignment = false): Uint8Array {
  const exception = (row: number) => {
    const first = kind.startsWith("first-"), active = first ? row === 0 : row > 0;
    if (!active) return "";
    const property = kind.endsWith("indent")
      ? `<w:tblInd w:w="${row === 2 ? -120 : 600}" w:type="dxa"/>`
      : kind.endsWith("width") ? `<w:tblW w:w="${row === 2 ? 6000 : 3600}" w:type="dxa"/>`
      : `<w:tblW w:w="3600" w:type="dxa"/><w:jc w:val="${row === 2 ? "right" : "center"}"/>`;
    const margin = kind === "indent" ? `<w:tblCellMar><w:left w:w="${row === 1 ? 220 : 60}" w:type="dxa"/></w:tblCellMar>` : "";
    return directAlignment && kind.endsWith("align")
      ? `<w:tblPrEx>${property.replace(/<w:jc[^>]*\/>/, "")}${margin}</w:tblPrEx><w:trPr><w:jc w:val="${row === 2 ? "right" : "center"}"/></w:trPr>`
      : `<w:tblPrEx>${property}${margin}</w:tblPrEx>`;
  };
  return mutateDocx("02-table-merges.docx", {
    "word/styles.xml": () => '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="DengXian" w:hAnsi="DengXian" w:eastAsia="DengXian" w:cs="DengXian"/><w:sz w:val="24"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults></w:styles>',
    "word/settings.xml": () => `<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="${mode}"/></w:compat></w:settings>`,
    "word/document.xml": xml => body(xml, '<w:p><w:r><w:t>REFERENCE</w:t></w:r></w:p><w:tbl><w:tblPr><w:tblW w:w="4800" w:type="dxa"/><w:tblInd w:w="240" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="140" w:type="dxa"/><w:right w:w="140" w:type="dxa"/></w:tblCellMar><w:tblBorders>'
      + ["top", "left", "bottom", "right", "insideH", "insideV"].map(side => `<w:${side} w:val="single" w:sz="8" w:color="335577"/>`).join("")
      + '</w:tblBorders></w:tblPr><w:tblGrid>' + '<w:gridCol w:w="1600"/>'.repeat(3) + '</w:tblGrid>'
      + [0, 1, 2].map(row => '<w:tr>' + exception(row) + [0, 1, 2].map(col => '<w:tc><w:tcPr>'
        + (merged && col === 0 && row < 2 ? `<w:vMerge${row === 0 ? ' w:val="restart"' : ""}/>` : "")
        + '</w:tcPr><w:p><w:r><w:t>' + (merged && col === 0 && row === 1 ? "" : `R${row}C${col}`) + '</w:t></w:r></w:p></w:tc>').join("") + '</w:tr>').join("")
      + '</w:tbl><w:p><w:r><w:t>AFTER</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>'),
  });
}
