import {
  openDocument,
  SourceStyleResolver,
  DEFAULT_DOCUMENT_READ_LIMITS,
  getWordTextDefaults,
} from "@doffi/core";
import type { WordContentSemantics } from "@doffi/core";
import { importDocx } from "@doffi/office";
import { continuationCellFixture } from "./continuation-cell-fixture.js";
import { body, mutateDocx, styles } from "./regression-fixtures.js";
export function nestedTableContentFixture(): Uint8Array {
  return mutateDocx("02-table-merges.docx", {
    "word/styles.xml": () =>
      styles('<w:style w:type="paragraph" w:styleId="Normal" w:default="1"/>'),
    "word/document.xml": (source) =>
      body(
        source,
        '<w:tbl><w:tblPr><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="4800"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>OUTER</w:t></w:r></w:p><w:tbl><w:tblPr><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>INNER-A</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:tc><w:p><w:r><w:t>INNER-B</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/></w:tc></w:tr><w:tr><w:tc><w:p><w:r><w:t>TAIL</w:t></w:r></w:p></w:tc></w:tr></w:tbl>',
      ),
  });
}
export function wordTableContentFixture(bytes = continuationCellFixture()): {
  html: string;
  semantics: WordContentSemantics;
} {
  const model = openDocument(
      importDocx(bytes, { documentId: "table-content-source", unsupportedContent: "reject" }),
      { kind: "document" },
    ),
    snapshot = model.snapshot(),
    word = snapshot.metadata.word!,
    table = word.tables[0]!,
    css = new SourceStyleResolver(
      snapshot.styles.map((sheet) => sheet.css),
      DEFAULT_DOCUMENT_READ_LIMITS,
    ),
    names = new Set<string>(),
    selected = new Set<string>();
  const visit = (id: string): void => {
    selected.add(id);
    const node = model.getNode(id)!;
    for (const name of node.attributes.class?.split(/\s+/) ?? []) if (name) names.add(name);
    for (const child of node.childNodeIds) visit(child);
  };
  visit(table.nodeId);
  const tables = word.tables.filter((item) => selected.has(item.nodeId));
  for (const item of tables)
    for (const entry of item.continuationFormats ?? []) {
      if (entry.cellCssClass) names.add(entry.cellCssClass);
      for (const paragraph of entry.emptyParagraphs ?? []) {
        if (paragraph.paragraphCssClass) names.add(paragraph.paragraphCssClass);
        if (paragraph.runCssClass) names.add(paragraph.runCssClass);
        for (const run of paragraph.runs) if (run.cssClass) names.add(run.cssClass);
      }
    }
  const defaults = Object.entries(getWordTextDefaults(model, table.nodeId))
    .map(([key, value]) => key + ":" + value)
    .join(";")
    .replaceAll("&", "&amp;")
    .replaceAll('"', "&quot;");
  return {
    html: '<div style="' + defaults + '">' + model.getNode(table.nodeId)!.html + "</div>",
    semantics: JSON.parse(
      JSON.stringify({
        tables,
        tabs: word.tabs.filter((item) => selected.has(item.paragraphNodeId)),
        conditionalFormats: word.conditionalFormats?.filter((item) => selected.has(item.nodeId)),
        extraClasses: Object.fromEntries(
          [...names].map((name) => [name, css.specifiedClass(name)]),
        ),
      }),
    ) as WordContentSemantics,
  };
}
