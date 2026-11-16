import { mutateDocx, body, styles } from "./regression-fixtures.js";
export type GridBoundaryKind = "merged" | "vertical" | "before" | "after";
export function gridBoundaryFixture(kind: GridBoundaryKind, count = 2): Uint8Array {
  return mutateDocx("02-table-merges.docx", {
    "word/styles.xml": () =>
      styles('<w:style w:type="paragraph" w:styleId="Normal" w:default="1"/>'),
    "word/document.xml": (source) =>
      body(
        source,
        '<w:tbl><w:tblPr><w:tblW w:w="4320" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid>' +
          [800, 1200, 2320].map((w) => '<w:gridCol w:w="' + w + '"/>').join("") +
          "</w:tblGrid>" +
          Array.from({ length: count }, (_, row) => {
            const gap = kind === "before" || kind === "after" ? kind : undefined;
            const props = gap
              ? "<w:trPr><w:grid" +
                (gap === "before" ? "Before" : "After") +
                ' w:val="2"/><w:w' +
                (gap === "before" ? "Before" : "After") +
                ' w:w="25%" w:type="pct"/></w:trPr>'
              : "";
            const merge =
              kind === "vertical" ? "<w:vMerge" + (row === 0 ? ' w:val="restart"' : "") + "/>" : "";
            return (
              "<w:tr>" +
              props +
              "<w:tc><w:tcPr>" +
              (gap ? "" : '<w:gridSpan w:val="3"/>') +
              merge +
              "</w:tcPr><w:p>" +
              (kind === "vertical" && row > 0 ? "" : "<w:r><w:t>ROW" + row + "</w:t></w:r>") +
              "</w:p></w:tc></w:tr>"
            );
          }).join("") +
          "</w:tbl><w:p><w:r><w:t>AFTER</w:t></w:r></w:p>",
      ),
  });
}
