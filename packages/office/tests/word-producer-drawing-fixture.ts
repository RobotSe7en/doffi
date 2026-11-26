import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { floatingCompatibilityFixture } from "./floating-compatibility-fixture.js";

/** Producer properties are hand-authored, without private document bytes. */
export function wordProducerDrawingFixture(options: { hidden?: boolean; connector?: boolean; scale?: number; rotation?: number; oversized?: boolean; offset?: number } = {}): Uint8Array {
  const files = unzipSync(floatingCompatibilityFixture(14, "cell")),
    extent = options.oversized ? 2540000 : 2032000;
  let xml = strFromU8(files["word/document.xml"]!)
    .replace('<wp:cNvGraphicFramePr/>', '<wp:cNvGraphicFramePr><a:graphicFrameLocks/></wp:cNvGraphicFramePr>')
    .replaceAll('cx="2032000"', `cx="${extent}"`)
    .replace('rot="0"', `rot="${(options.rotation ?? 0) * 60000}"`)
    .replace('<wp:docPr ', `<wp:docPr hidden="${options.hidden ? 1 : 0}" `)
    .replace('<wp:effectExtent l="0" t="0" r="0" b="0"/>', '<wp:effectExtent l="12700" t="28575" r="38100" b="69850"/>')
    .replace('<wps:spPr>', '<wps:spPr bwMode="auto">');
  if (options.offset !== undefined) xml = xml.replace('<wp:align>center</wp:align>', `<wp:posOffset>${options.offset * 12700}</wp:posOffset>`);
  xml = options.connector ? xml
    .replace('<wps:cNvSpPr/>', '<wps:cNvCnPr><a:cxnSpLocks noChangeShapeType="1"/></wps:cNvCnPr>')
    .replaceAll('cy="1270000"', 'cy="0"').replace('prst="rect"', 'prst="line"')
    : xml.replace('<wps:cNvSpPr/>', '<wps:cNvSpPr txBox="1"><a:spLocks noChangeArrowheads="1"/></wps:cNvSpPr>')
      .replace('<wps:bodyPr/>', '<wps:txbx><w:txbxContent><w:p><w:r><w:rPr><w:sz w:val="40"/>'
        + (options.scale ? `<w:w w:val="${options.scale}"/>` : "")
        + '</w:rPr><w:t>合成文本框 Synthetic text</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr upright="1" lIns="0" tIns="0" rIns="0" bIns="0"/>');
  files["word/document.xml"] = strToU8(xml);
  return zipSync(files);
}

/** Identical inherited defaults with omitted versus explicit single spacing. */
export function nativeSingleSpacingFixture(explicit: boolean): Uint8Array {
  const files = unzipSync(wordProducerDrawingFixture()),
    font = '<w:rFonts w:ascii="SimSun" w:hAnsi="SimSun" w:eastAsia="SimSun"/><w:sz w:val="40"/>',
    spacing = explicit ? '<w:spacing w:line="240" w:lineRule="auto"/>' : '';
  files['word/styles.xml'] = strToU8('<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr>' + font + '</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr>' + spacing + '</w:pPr></w:pPrDefault></w:docDefaults></w:styles>');
  files['word/document.xml'] = strToU8(strFromU8(files['word/document.xml']!)
    .replace(/<w:spacing\b[^>]*\/>/g, '')
    .replace('合成文本框 Synthetic text', 'Single</w:t><w:br/><w:t>Second')
    .replace(/<w:t>[^<]*中文排版[^<]*<\/w:t>/g, '<w:t>中文排版</w:t>'));
  return zipSync(files);
}

export function nativeTextBoxGridFixture(options: { grid?: boolean; snap?: boolean; exact?: number } = {}): Uint8Array {
  const files = unzipSync(nativeSingleSpacingFixture(false));
  files['word/styles.xml'] = strToU8(strFromU8(files['word/styles.xml']!).replaceAll('w:val="40"', 'w:val="70"'));
  files['word/document.xml'] = strToU8(strFromU8(files['word/document.xml']!)
    .replaceAll('<w:sz w:val="40"/>', '<w:sz w:val="70"/>')
    .replace('<w:txbxContent><w:p>', '<w:txbxContent><w:p><w:pPr>'
      + (options.snap === false ? '<w:snapToGrid w:val="0"/>' : '')
      + (options.exact ? '<w:spacing w:line="' + options.exact * 20 + '" w:lineRule="exact"/>' : '')
      + '</w:pPr>')
    .replace('</w:sectPr>', (options.grid ? '<w:docGrid w:type="lines" w:linePitch="312"/>' : '') + '</w:sectPr>'));
  return zipSync(files);
}
