import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { rowSpacingFixture } from "./row-spacing-fixture.js";
export function rowExceptionFixture(count = 3, full = false): Uint8Array {
  const files = unzipSync(rowSpacingFixture(count));
  files["word/document.xml"] = strToU8(
    strFromU8(files["word/document.xml"]!).replaceAll(
      "<w:tr>",
      '<w:tr><w:tblPrEx><w:tblCellSpacing w:w="80" w:type="dxa"/></w:tblPrEx>',
    ),
  );
  if (full) {
    let xml = strFromU8(files["word/document.xml"]!);
    xml = xml.replaceAll(
      /<w:tcBorders>[\s\S]*?<\/w:tcBorders>|<w:shd[^>]*\/>|<w:tcMar>[\s\S]*?<\/w:tcMar>/g,
      "",
    );
    xml = xml
      .replaceAll(
        "<w:tblPrEx>",
        '<w:tblPrEx><w:tblW w:w="3200" w:type="dxa"/><w:jc w:val="center"/>',
      )
      .replaceAll(
        "</w:tblPrEx>",
        '<w:tblBorders><w:left w:val="single" w:sz="8" w:color="AA2200"/><w:right w:val="single" w:sz="8" w:color="AA2200"/><w:insideV w:val="single" w:sz="4" w:color="336699"/></w:tblBorders><w:shd w:val="clear" w:fill="FFEEDD"/><w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="160" w:type="dxa"/><w:right w:w="160" w:type="dxa"/></w:tblCellMar><w:tblLook w:val="0020"/></w:tblPrEx>',
      );
    files["word/document.xml"] = strToU8(xml);
  }
  return zipSync(files);
}
