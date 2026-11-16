import { mutateDocx, body, styles } from "./regression-fixtures.js";
export function dynamicGridFixture(
  grid: string,
  span = 1,
  cellWidth = "1440",
  tableWidth = "2880",
  type = "dxa",
) {
  return mutateDocx("01-text-numbering.docx", {
    "word/styles.xml": () =>
      styles('<w:style w:type="paragraph" w:styleId="Normal" w:default="1"/>'),
    "word/document.xml": (source) =>
      body(
        source,
        '<w:p><w:r><w:t>BEFORE</w:t></w:r></w:p><w:tbl><w:tblPr><w:tblW w:type="' +
          type +
          '" w:w="' +
          tableWidth +
          '"/></w:tblPr><w:tblGrid>' +
          grid +
          '</w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:type="dxa" w:w="' +
          cellWidth +
          '"/><w:gridSpan w:val="' +
          span +
          '"/></w:tcPr><w:p><w:r><w:t>DYNAMIC</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p><w:r><w:t>AFTER</w:t></w:r></w:p>',
      ),
  });
}

export function dynamicTableLayoutFixture(count = 60): Uint8Array {
  const cell = (text: string, span = 1) =>
    "<w:tc><w:tcPr>" +
    (span > 1 ? '<w:gridSpan w:val="' + span + '"/>' : "") +
    "</w:tcPr><w:p><w:r><w:t>" +
    text +
    "</w:t></w:r></w:p></w:tc>";
  return mutateDocx("02-table-merges.docx", {
    "word/styles.xml": () =>
      styles('<w:style w:type="paragraph" w:styleId="Normal" w:default="1"/>'),
    "word/document.xml": (source) =>
      body(
        source,
        '<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/><w:tblLayout w:type="autofit"/></w:tblPr><w:tblGrid/>' +
          Array.from(
            { length: count },
            (_, row) =>
              "<w:tr>" +
              (row === 0
                ? cell("MERGED", 2)
                : row === 1
                  ? '<w:trPr><w:gridBefore w:val="1"/></w:trPr>' + cell("OMITTED")
                  : cell("A" + row) + cell("B" + row)) +
              "</w:tr>",
          ).join("") +
          "</w:tbl>",
      ),
  });
}
