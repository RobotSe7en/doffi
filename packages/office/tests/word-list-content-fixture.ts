import {
  DEFAULT_DOCUMENT_READ_LIMITS,
  SourceStyleResolver,
  getWordTextDefaults,
  openDocument,
} from "@doffi/core";
import type { WordContentSemantics } from "@doffi/core";
import { importDocx } from "@doffi/office";
import { body, mutateDocx, styles } from "./regression-fixtures.js";

export function nativeListContentFixture(): Uint8Array {
  const paragraph = (text: string, id?: string, level = 0) =>
    "<w:p>" +
    (id
      ? '<w:pPr><w:numPr><w:ilvl w:val="' +
        level +
        '"/><w:numId w:val="' +
        id +
        '"/></w:numPr></w:pPr>'
      : "") +
    "<w:r><w:t>" +
    text +
    "</w:t></w:r></w:p>";
  return mutateDocx("01-text-numbering.docx", {
    "word/styles.xml": () =>
      styles('<w:style w:type="paragraph" w:styleId="Normal" w:default="1"/>'),
    "word/numbering.xml": () =>
      '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="multilevel"/>' +
      '<w:lvl w:ilvl="0"><w:start w:val="5"/><w:numFmt w:val="upperRoman"/><w:suff w:val="tab"/><w:lvlText w:val="%1)"/><w:pPr><w:ind w:left="480" w:hanging="240"/></w:pPr><w:rPr><w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif" w:eastAsia="FandolFang"/><w:color w:val="880000"/><w:sz w:val="28"/></w:rPr></w:lvl>' +
      '<w:lvl w:ilvl="1"><w:start w:val="2"/><w:numFmt w:val="lowerLetter"/><w:lvlRestart w:val="1"/><w:suff w:val="space"/><w:lvlText w:val="%1.%2)"/><w:pPr><w:ind w:left="960" w:hanging="240"/></w:pPr><w:rPr><w:rFonts w:ascii="Liberation Mono" w:hAnsi="Liberation Mono"/><w:color w:val="008800"/><w:sz w:val="24"/></w:rPr></w:lvl></w:abstractNum>' +
      '<w:abstractNum w:abstractNumId="99"><w:lvl w:ilvl="0"><w:start w:val="666"/><w:numFmt w:val="bullet"/><w:lvlText w:val="PRIVATE-MARKER"/></w:lvl></w:abstractNum>' +
      '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="0"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="10"/></w:lvlOverride></w:num><w:num w:numId="99"><w:abstractNumId w:val="99"/></w:num></w:numbering>',
    "word/document.xml": (source) =>
      body(
        source,
        paragraph("FIRST", "1") +
          paragraph("CHILD", "1", 1) +
          paragraph("SECOND", "1") +
          paragraph("RESTART", "2") +
          paragraph("OFF", "0") +
          '<w:tbl><w:tblPr><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>' +
          '<w:tr><w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr>' +
          paragraph("MERGED") +
          "</w:tc><w:tc>" +
          paragraph("R1") +
          "</w:tc></w:tr>" +
          '<w:tr><w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p><w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="2"/></w:numPr></w:pPr></w:p></w:tc><w:tc>' +
          paragraph("R2") +
          "</w:tc></w:tr></w:tbl>" +
          paragraph("PRIVATE-BODY", "99"),
      ),
  });
}

/** Native-side fixture for typed transaction/reopen tests. Browser acceptance
 * separately uses the real projection and ClipboardEvent copy path. */
export function wordListContentFixture(tableOnly = false): {
  html: string;
  semantics: WordContentSemantics;
} {
  const model = openDocument(
    importDocx(nativeListContentFixture(), {
      documentId: "list-content-source",
      unsupportedContent: "reject",
    }),
    { kind: "document" },
  );
  const snapshot = model.snapshot();
  const word = snapshot.metadata.word!;
  const all = [...snapshot.html.matchAll(/data-node-id="([^"]+)"/g)].map(
    (match) => model.getNode(match[1]!)!,
  );
  const selected = tableOnly
    ? [model.getNode(word.tables[0]!.nodeId)!]
    : all.filter(
        (node) =>
          node.tagName === "p" &&
          ["FIRST", "CHILD", "SECOND", "RESTART", "OFF"].includes(node.text),
      );
  const ids = new Set<string>();
  const classes = new Set<string>();
  const numbers = new Set<string>();
  const visit = (id: string): void => {
    ids.add(id);
    const node = model.getNode(id)!;
    if (node.attributes["data-word-numbering"]) numbers.add(node.attributes["data-word-numbering"]);
    for (const name of node.attributes.class?.split(/\s+/) ?? []) if (name) classes.add(name);
    for (const child of node.childNodeIds) visit(child);
  };
  for (const node of selected) visit(node.nodeId);
  const tables = word.tables.filter((table) => ids.has(table.nodeId));
  for (const table of tables)
    for (const entry of table.continuationFormats ?? []) {
      if (entry.cellCssClass) classes.add(entry.cellCssClass);
      for (const paragraph of entry.emptyParagraphs ?? []) {
        if (paragraph.numbering) numbers.add(paragraph.numbering.numberingId);
        if (paragraph.paragraphCssClass) classes.add(paragraph.paragraphCssClass);
        if (paragraph.runCssClass) classes.add(paragraph.runCssClass);
        for (const run of paragraph.runs) if (run.cssClass) classes.add(run.cssClass);
      }
    }
  const numbering = word.numbering.filter((definition) => numbers.has(definition.numberingId));
  for (const definition of numbering)
    for (const level of definition.levels) classes.add(level.cssClass);
  const css = new SourceStyleResolver(
    snapshot.styles.map((sheet) => sheet.css),
    DEFAULT_DOCUMENT_READ_LIMITS,
  );
  const defaults = Object.entries(getWordTextDefaults(model, selected[0]!.nodeId))
    .map(([key, value]) => key + ":" + value)
    .join(";")
    .replaceAll('"', "&quot;");
  return {
    html: '<div style="' + defaults + '">' + selected.map((node) => node.html).join("") + "</div>",
    semantics: JSON.parse(
      JSON.stringify({
        numbering,
        tables,
        extraClasses: Object.fromEntries(
          [...classes].map((name) => [name, css.specifiedClass(name)]),
        ),
      }),
    ) as WordContentSemantics,
  };
}
