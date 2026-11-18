import { body, mutateDocx, styles, wordNamespace } from "./regression-fixtures.js";

/** Independent OOXML with the instance and level declared on different styles. */
export function styleNumberingFixture(): Uint8Array {
  const paragraph = (style: string, text: string, numbering = "") =>
    '<w:p><w:pPr><w:pStyle w:val="' + style + '"/>' + numbering +
    "</w:pPr><w:r><w:t>" + text + "</w:t></w:r></w:p>";
  const style = (id: string, base: string, outline: number, properties: string, run = "") =>
    '<w:style w:type="paragraph" w:styleId="' + id + '"><w:name w:val="' + id +
    '"/><w:basedOn w:val="' + base + '"/><w:pPr>' + properties +
    '<w:outlineLvl w:val="' + outline + '"/></w:pPr>' + run + "</w:style>";
  const level = (index: number, label: string, linkedStyle = "") =>
    '<w:lvl w:ilvl="' + index + '"><w:start w:val="1"/><w:numFmt w:val="decimal"/>' +
    (linkedStyle ? '<w:pStyle w:val="' + linkedStyle + '"/>' : "") +
    '<w:lvlText w:val="' + label + '"/><w:pPr><w:ind w:left="480" w:hanging="480"/></w:pPr></w:lvl>';
  return mutateDocx("01-text-numbering.docx", {
    "word/document.xml": (source) => body(source,
      paragraph("Chapter", "Chapter A") + paragraph("Topic", "Topic A") +
      paragraph("Detail", "Detail A") +
      paragraph("Topic", "Override", '<w:numPr><w:ilvl w:val="2"/></w:numPr>') +
      paragraph("Topic", "Unnumbered", '<w:numPr><w:numId w:val="0"/></w:numPr>') +
      paragraph("Topic", "Topic B") + paragraph("Detail", "Detail B") +
      paragraph("Chapter", "Chapter B") + paragraph("Topic", "Topic C"),
    ),
    "word/styles.xml": () => styles(
      '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/>' +
      '<w:sz w:val="28"/></w:rPr></w:rPrDefault></w:docDefaults>' +
      '<w:style w:type="paragraph" w:styleId="Normal" w:default="1"><w:name w:val="Normal"/>' +
      '<w:pPr><w:ind w:firstLine="560"/></w:pPr></w:style>' +
      style("Chapter", "Normal", 0, '<w:numPr><w:numId w:val="7"/></w:numPr>', "<w:rPr><w:b/></w:rPr>") +
      style("Topic", "Chapter", 1, '<w:numPr><w:ilvl w:val="1"/></w:numPr><w:ind w:left="0" w:firstLine="0"/>',
        '<w:rPr><w:sz w:val="30"/></w:rPr>') +
      style("Detail", "Topic", 2, ""),
    ),
    "word/settings.xml": () =>
      '<w:settings xmlns:w="' + wordNamespace + '"><w:defaultTabStop w:val="600"/></w:settings>',
    "word/numbering.xml": () =>
      '<w:numbering xmlns:w="' + wordNamespace + '"><w:abstractNum w:abstractNumId="0">' +
      '<w:multiLevelType w:val="multilevel"/>' +
      level(0, "%1") + level(1, "%1.%2") + level(2, "%1.%2.%3", "Detail") +
      '</w:abstractNum><w:num w:numId="7"><w:abstractNumId w:val="0"/></w:num></w:numbering>',
  });
}
