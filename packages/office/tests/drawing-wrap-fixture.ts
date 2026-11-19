import type { WordDrawingLayout } from "@doffi/core";
import { body, mutateDocx, styles } from "./regression-fixtures.js";

export const wrapText = "中文排版 中英协作 ABCdef 𠀀\u{e0100} 😀 环绕文字保持准确选区。 ";
export const notchPolygon = [
  { x: 0, y: 0 },
  { x: 21600, y: 0 },
  { x: 21600, y: 21600 },
  { x: 16000, y: 21600 },
  { x: 16000, y: 5000 },
  { x: 5600, y: 5000 },
  { x: 5600, y: 21600 },
  { x: 0, y: 21600 },
];
export const diamondPolygon = [
  { x: 10800, y: 0 },
  { x: 21600, y: 10800 },
  { x: 10800, y: 21600 },
  { x: 0, y: 10800 },
];
/** Independent hand-authored OOXML, using only ZIP/XML fixture primitives.
 * Neither the SDK drawing contracts nor its exporter generate this source. */
export function drawingWrapFixture(
  options: {
    wrap?: WordDrawingLayout["wrap"];
    side?: WordDrawingLayout["wrapSide"];
    polygon?: "notch" | "diamond";
    inCell?: boolean;
    layoutInCell?: boolean;
    count?: number;
    allowOverlap?: boolean;
    rotation?: number;
  } = {},
): Uint8Array {
  const polygon = options.polygon === "diamond" ? diamondPolygon : notchPolygon,
    wrap = options.wrap ?? "square",
    wrapXml =
      wrap === "tight" || wrap === "through"
        ? `<wp:wrap${wrap === "tight" ? "Tight" : "Through"} wrapText="${options.side ?? "bothSides"}" distL="38100" distR="50800"><wp:wrapPolygon edited="1">` +
          `<wp:start x="${polygon[0]!.x}" y="${polygon[0]!.y}"/>` +
          polygon
            .slice(1)
            .map((p) => `<wp:lineTo x="${p.x}" y="${p.y}"/>`)
            .join("") +
          `</wp:wrapPolygon></wp:wrap${wrap === "tight" ? "Tight" : "Through"}>`
        : wrap === "square"
          ? `<wp:wrapSquare wrapText="${options.side ?? "bothSides"}" distL="38100" distR="50800"/>`
          : wrap === "none"
            ? "<wp:wrapNone/>"
            : "<wp:wrapTopAndBottom/>",
    drawing = (index: number) =>
      `<w:r><w:drawing xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wp:anchor distT="25400" distB="38100" distL="12700" distR="12700" simplePos="0" relativeHeight="${index + 1}" behindDoc="0" locked="0" layoutInCell="${options.layoutInCell === false ? 0 : 1}" allowOverlap="${options.allowOverlap === false ? 0 : 1}"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:align>center</wp:align></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>152400</wp:posOffset></wp:positionV><wp:extent cx="2032000" cy="1270000"/><wp:effectExtent l="0" t="0" r="0" b="0"/>${wrapXml}<wp:docPr id="${index + 1}" name="Wrap object ${index + 1}"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr/><wps:spPr><a:xfrm rot="${(options.rotation ?? 0) * 60000}"><a:off x="0" y="0"/><a:ext cx="2032000" cy="1270000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FACC15"/></a:solidFill><a:ln><a:noFill/></a:ln></wps:spPr><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>`,
    p = (text: string, objects = "") =>
      '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="320" w:lineRule="exact"/></w:pPr>' +
      objects +
      '<w:r><w:t xml:space="preserve">' +
      text +
      "</w:t></w:r></w:p>",
    objects = Array.from({ length: options.count ?? 1 }, (_, index) => drawing(index)).join(""),
    content = options.inCell
      ? '<w:tbl><w:tblPr><w:tblW w:w="6960" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:tblCellMar><w:top w:w="80" w:type="dxa"/><w:left w:w="80" w:type="dxa"/><w:bottom w:w="80" w:type="dxa"/><w:right w:w="80" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="3480"/><w:gridCol w:w="3480"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="3480" w:type="dxa"/></w:tcPr>' +
        p(wrapText.repeat(2), objects) +
        '</w:tc><w:tc><w:tcPr><w:tcW w:w="3480" w:type="dxa"/></w:tcPr>' +
        p("相邻单元格 " + wrapText.repeat(2)) +
        "</w:tc></w:tr></w:tbl>" +
        p("表格之后 " + wrapText.repeat(10))
      : p(wrapText.repeat(40), objects) + p("后续段落 " + wrapText.repeat(3));
  return mutateDocx("01-text-numbering.docx", {
    "word/styles.xml": () =>
      styles(
        '<w:style w:type="paragraph" w:styleId="Normal" w:default="1"><w:rPr><w:rFonts w:ascii="Liberation Serif" w:eastAsia="Noto Sans CJK SC"/><w:sz w:val="24"/></w:rPr></w:style>',
      ),
    "word/document.xml": (source) =>
      body(
        source,
        content +
          '<w:sectPr><w:pgSz w:w="8400" w:h="10000"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"/></w:sectPr>',
      ),
  });
}
