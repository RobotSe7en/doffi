import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { drawingWrapFixture } from "./drawing-wrap-fixture.js";
import { floatingTableFixture } from "./floating-table-fixture.js";

/** Independent native anchors and table positions, with explicit producer mode
 * and fonts so that native reference positions do not depend on Normal.dotm. */
export function floatingCompatibilityFixture(mode: 14 | 15, kind: "drawing" | "cell" | "cell-outside" | "cell-none" | "table-text" | "table-center" | "table-page", preceding?: "paragraph" | "row"): Uint8Array {
  const table = kind.startsWith("table-"),
    bytes = table ? floatingTableFixture(kind === "table-center"
      ? '<w:tblpPr w:horzAnchor="margin" w:vertAnchor="text" w:tblpXSpec="center" w:tblpY="241" w:leftFromText="120" w:rightFromText="120"/>'
      : kind === "table-page" ? '<w:tblpPr w:horzAnchor="page" w:vertAnchor="page" w:tblpX="1441" w:tblpY="1441"/>' : undefined)
      : drawingWrapFixture({ inCell: kind !== "drawing", layoutInCell: kind === "cell", wrap: kind === "cell-none" ? "none" : "square" }),
    files = unzipSync(bytes);
  if (preceding) {
    const xml = strFromU8(files["word/document.xml"]!);
    files["word/document.xml"] = strToU8(preceding === "paragraph"
      ? xml.replace('<w:p><w:pPr>', '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="320" w:lineRule="exact"/></w:pPr><w:r><w:t>LEAD</w:t></w:r></w:p><w:p><w:pPr>')
      : xml.replace('<w:tr>', '<w:tr><w:trPr><w:trHeight w:val="800" w:hRule="exact"/></w:trPr><w:tc><w:p><w:r><w:t>ROWBEFORE</w:t></w:r></w:p></w:tc><w:tc><w:p/></w:tc></w:tr><w:tr>'));
  }
  files["word/settings.xml"] = strToU8(`<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="${mode}"/></w:compat></w:settings>`);
  files["word/styles.xml"] = strToU8(table
    ? '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="DengXian" w:hAnsi="DengXian" w:eastAsia="DengXian" w:cs="DengXian"/><w:sz w:val="24"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults></w:styles>'
    : strFromU8(files["word/styles.xml"]!).replaceAll("Liberation Serif", "DengXian").replaceAll("Noto Sans CJK SC", "DengXian"));
  return zipSync(files);
}

export function drawingBodyDefaultsFixture(body = ""): Uint8Array {
  const files = unzipSync(floatingCompatibilityFixture(15, "drawing"));
  files["word/document.xml"] = strToU8(strFromU8(files["word/document.xml"]!).replace(
    "<wps:bodyPr/>",
    '<wps:txbx><w:txbxContent><w:p><w:r><w:t>BODY DEFAULTS</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr ' + body + '/>',
  ));
  return zipSync(files);
}
