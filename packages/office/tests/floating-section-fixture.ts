import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { drawingWrapFixture } from "./drawing-wrap-fixture.js";
import { floatingTableFixture } from "./floating-table-fixture.js";
import { body } from "./regression-fixtures.js";

const cols =
  '<w:cols w:num="2" w:equalWidth="0"><w:col w:w="2400" w:space="480"/><w:col w:w="4080"/></w:cols>';
const paragraph = (text: string, contents = "", props = "") =>
  '<w:p><w:pPr><w:spacing w:after="0" w:line="320" w:lineRule="exact"/>' +
  props +
  "</w:pPr>" +
  contents +
  "<w:r><w:t>" +
  text +
  "</w:t></w:r></w:p>";
const section = (type: string, columns = cols, height = 5200, refs = "", right = 720) =>
  "<w:sectPr>" +
  refs +
  '<w:type w:val="' +
  type +
  '"/><w:pgSz w:w="8400" w:h="' +
  height +
  '"/><w:pgMar w:top="720" w:right="' +
  right +
  '" w:bottom="720" w:left="720" w:header="160" w:footer="160" w:gutter="0"/>' +
  columns +
  "</w:sectPr>";

export function floatingSectionFixture(
  kind: "bands" | "rows" | "stories" | "geometry",
): Uint8Array {
  const files = unzipSync(drawingWrapFixture({ wrap: "none" })),
    document = strFromU8(files["word/document.xml"]!),
    shapeRun = document.match(/<w:r><w:drawing[\s\S]*?<\/w:drawing><\/w:r>/)![0],
    tableBytes = unzipSync(floatingTableFixture()),
    tableXml = strFromU8(tableBytes["word/document.xml"]!),
    table = tableXml.match(/<w:tbl>[\s\S]*?<\/w:tbl>/)![0].replaceAll('w:w="2400"', 'w:w="1600"'),
    draw = (id: number, wrap = "none", vertical = "margin", y = 12) =>
      shapeRun
        .replace('wp:docPr id="1"', 'wp:docPr id="' + id + '"')
        .replaceAll("2032000", "762000")
        .replaceAll("1270000", "254000")
        .replace('relativeFrom="paragraph"', 'relativeFrom="' + vertical + '"')
        .replace("152400", String(y * 12700))
        .replace(
          "<wp:wrapNone/>",
          wrap === "topAndBottom"
            ? "<wp:wrapTopAndBottom/>"
            : wrap === "square"
              ? '<wp:wrapSquare wrapText="bothSides"/>'
              : "<wp:wrapNone/>",
        ),
    positioned = (v: string, y: number) =>
      table.replace(
        /<w:tblpPr[^>]*\/>/,
        '<w:tblpPr w:vertAnchor="' +
          v +
          '" w:horzAnchor="text" w:tblpXSpec="center" w:tblpY="' +
          (y * 20 + 1) +
          '"/><w:tblOverlap w:val="never"/>',
      );
  let content: string;
  if (kind === "bands") {
    content =
      paragraph("第一栏 FIRST", draw(1, "topAndBottom", "margin", 0)) +
      paragraph("", "", section("nextPage")) +
      paragraph("SECOND 第二栏", draw(2)) +
      positioned("margin", 12) +
      paragraph("后续文字 FOLLOW") +
      paragraph(
        "",
        "",
        section("nextColumn", cols.replace('w:equalWidth="0"', 'w:equalWidth="0" w:sep="1"')),
      ) +
      paragraph("通栏 TAIL", draw(3, "none", "paragraph", 0)) +
      section("continuous", '<w:cols w:num="1"/>');
  } else if (kind === "rows") {
    const row = table.match(/<w:tr>[\s\S]*?<\/w:tr>/)![0];
    const rows = Array.from({ length: 31 }, (_, index) => {
      let entry = row.replace(
        "CELL",
        index === 0 ? "REPEAT HEADER" : "ROW " + String(index).padStart(2, "0"),
      );
      // This case isolates row continuation across columns. Pin its cell
      // metrics as well as the surrounding paragraphs; omitted single
      // spacing is covered separately with native default-spacing fixtures.
      entry = entry.replace("<w:p>", '<w:p><w:pPr><w:spacing w:line="320" w:lineRule="exact"/></w:pPr>');
      if (index === 0) entry = entry.replace("<w:tr>", "<w:tr><w:trPr><w:tblHeader/></w:trPr>");
      if (index === 16)
        entry = entry.replace("<w:r><w:t>", draw(4, "none", "line", 0) + "<w:r><w:t>");
      return entry;
    }).join("");
    content =
      paragraph("FIRST") +
      paragraph("", "", section("nextPage")) +
      paragraph("SECOND") +
      positioned("text", 0).replace(row, rows) +
      paragraph("最后一段 AFTER ROWS", draw(7, "none", "margin", 0)) +
      section("nextColumn");
  } else if (kind === "geometry") {
    content =
      paragraph("FIRST", draw(1)) +
      paragraph("", "", section("nextPage")) +
      paragraph("NEW PAPER", draw(2)) +
      paragraph("", "", section("continuous", cols, 5800)) +
      paragraph("NEW COLUMNS", draw(3)) +
      section("nextColumn", cols.replaceAll("2400", "2040").replaceAll("4080", "4440"), 5800);
  } else {
    const ns = ' xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
      header =
        paragraph(
          "页眉 HEADER wraps the real text around the shape on each sheet",
          draw(5, "square", "page", 12),
        ) +
        positioned("page", 12).replace(
          "<w:r><w:t>CELL",
          draw(8, "none", "line", 0) + "<w:r><w:t>CELL",
        ),
      footer = paragraph("页脚 FOOTER", draw(6, "none", "page", 204)) + positioned("page", 204);
    for (const [tag, body] of [
      ["header", header],
      ["footer", footer],
    ] as const) {
      const id = "r18" + tag,
        root = tag === "header" ? "hdr" : "ftr";
      files["word/" + tag + "1.xml"] = strToU8(
        "<w:" + root + ns + ">" + body + "</w:" + root + ">",
      );
      files["word/_rels/document.xml.rels"] = strToU8(
        strFromU8(files["word/_rels/document.xml.rels"]!).replace(
          "</Relationships>",
          '<Relationship Id="' +
            id +
            '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/' +
            tag +
            '" Target="' +
            tag +
            '1.xml"/></Relationships>',
        ),
      );
      files["[Content_Types].xml"] = strToU8(
        strFromU8(files["[Content_Types].xml"]!).replace(
          "</Types>",
          '<Override PartName="/word/' +
            tag +
            '1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.' +
            tag +
            '+xml"/></Types>',
        ),
      );
    }
    files["word/settings.xml"] = strToU8(
      strFromU8(files["word/settings.xml"]!).replace(
        "</w:settings>",
        "<w:mirrorMargins/></w:settings>",
      ),
    );
    content =
      paragraph("FIRST") +
      paragraph("EVEN", "", "<w:pageBreakBefore/>") +
      paragraph("ODD", "", "<w:pageBreakBefore/>") +
      section(
        "nextPage",
        cols.replace('w:w="4080"', 'w:w="3600"'),
        5200,
        '<w:headerReference w:type="default" r:id="r18header"/><w:footerReference w:type="default" r:id="r18footer"/>',
        1200,
      );
  }
  files["word/document.xml"] = strToU8(body(document, content));
  return zipSync(files);
}
