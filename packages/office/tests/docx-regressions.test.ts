import { openContentFixture as openDocument } from "../../../tests/content-fixture.js";
import { expect, test } from "vitest";
import {SourceStyleResolver,DEFAULT_DOCUMENT_READ_LIMITS} from "@doffi/core";
import type { ContentDocumentModel as DocumentModel, DocumentNode } from "@doffi/core";
import { exportDocx, importDocx as importNativeDocx, readDocumentZip, writeDocumentZip } from "@doffi/office";
// These historical rejection cases exercise the explicitly retained strict policy.
const importDocx: typeof importNativeDocx = (bytes, options) => importNativeDocx(bytes, { ...options, unsupportedContent: "reject" });
import {
  body, complexField, crossSectionRangesFixture, mutateDocx, nestedFieldFixture, styleInheritanceFixture, styles, wordNamespace,
} from "./regression-fixtures.js";

function reopen(bytes: Uint8Array): DocumentModel {
  const model = openDocument(importDocx(bytes, { documentId: "regression" }));
  expect(openDocument(readDocumentZip(writeDocumentZip(model.serialize()))).snapshot()).toEqual(model.snapshot());
  const native = exportDocx(model);
  const reopened = openDocument(importDocx(native.bytes, { documentId: "native-regression" }));
  const main = (document: DocumentModel): string => document.getNode(document.snapshot().metadata.word!.stories.find((story) => story.kind === "main")!.rootNodeId)!.text;
  expect(main(reopened)).toBe(main(model));
  expect(reopened.snapshot().annotations).toHaveLength(model.snapshot().annotations.length);
  expect(native.report.unexportedAnnotations).toEqual([]);
  return model;
}
function sourceCss(model:DocumentModel,node:DocumentNode):string{const resolver=new SourceStyleResolver(model.snapshot().styles.map(sheet=>sheet.css),DEFAULT_DOCUMENT_READ_LIMITS);const project=(node:DocumentNode):Parameters<typeof resolver.get>[0]=>({tag:node.tagName,attrs:node.attributes,parent:node.parentNodeId?project(model.getNode(node.parentNodeId)!):undefined});return Object.entries(resolver.get(project(node))).map(([key,value])=>key+':'+value).join(';');}
function descendants(model: DocumentModel, nodeId: string): DocumentNode[] {
  const node = model.getNode(nodeId)!;
  return [node, ...node.childNodeIds.flatMap((id) => descendants(model, id))];
}
function overlappingComments(count: number, runs: number, chunks = 1, quote = "x"): Uint8Array {
  return mutateDocx("08-comments-revisions.docx", {
    "word/comments.xml": () => '<w:comments xmlns:w="' + wordNamespace + '">' +
      Array.from({ length: count }, (_, id) => '<w:comment w:id="' + id + '"><w:p/></w:comment>').join("") +
      "</w:comments>",
    "word/document.xml": (source) => body(source, "<w:p>" +
      Array.from({ length: count }, (_, id) => '<w:commentRangeStart w:id="' + id + '"/>').join("") +
      ("<w:r>" + ("<w:t>" + quote + "</w:t>").repeat(chunks) + "</w:r>").repeat(runs) +
      Array.from({ length: count }, (_, id) => '<w:commentRangeEnd w:id="' + id + '"/>').join("") +
      "</w:p>"),
  });
}
test.each([
  { value: "A中".repeat(100_000), chunks: 1, limits: { maxNodes: 1000 } },
  { value: "Aé".repeat(100_000), chunks: 1, limits: { maxNodes: 1000 } },
  { value: "中é".repeat(100_000), chunks: 1, limits: { maxNodes: 1000 } },
  { value: "A中".repeat(100), chunks: 4, limits: { maxNodes: 1000 } },
  { value: "A中".repeat(2000), chunks: 1, limits: { maxFileBytes: 32_768 } },
  { value: "A中".repeat(2000), chunks: 1, limits: { maxPackageBytes: 65_536 } },
  { value: '"'.repeat(7000), chunks: 1, limits: { maxFileBytes: 32_768 } },
])("bounds generated script nodes and escaped bytes before allocation (case %#)", ({ value, chunks, limits }) => {
  const bytes = mutateDocx("06-links-tabs.docx", {
    "word/document.xml": (source) => body(source,
      '<w:p><w:r>' + ('<w:t>' + value + '</w:t>').repeat(chunks) + '</w:r></w:p>'),
  });
  expect(() => importDocx(bytes, { documentId: "budget", limits })).toThrowError(expect.objectContaining({
    code: "ResourceLimitExceeded",
    message: expect.stringContaining("Generated HTML"),
  }));
});
test("preserves bounded script segments, UTF-16 variation sequences and escaped text", () => {
  const value = 'A中\u0301𠀀\u{e0100}é&amp;&lt;&gt;"';
  const model = reopen(mutateDocx("06-links-tabs.docx", {
    "word/document.xml": (source) => body(source, '<w:p><w:r><w:t>' + value + '</w:t></w:r></w:p>'),
  }));
  const main = model.snapshot().metadata.word!.stories.find((story) => story.kind === "main")!;
  expect(model.getNode(main.rootNodeId)!.text).toBe('A中\u0301𠀀\u{e0100}é&<>"');
  const segments = descendants(model, main.rootNodeId).filter((node) => node.attributes["data-word-script"]);
  expect(segments.map((node) => [node.attributes["data-word-script"], node.text])).toEqual([
    ["eastAsia", "中\u0301𠀀\u{e0100}"], ["hAnsi", "é"],
  ]);
});

test.each([false, true])("isolates auxiliary fields, comments and bookmarks across sections (instruction open=%s)", (instructionOpen) => {
  const model = reopen(crossSectionRangesFixture(instructionOpen));
  const snapshot = model.snapshot();
  const expected = [instructionOpen ? "main after" : "main beforemain after", "header cache", "footer cache"];
  expect(snapshot.metadata.sections).toHaveLength(2);
  expect(snapshot.metadata.sections![1]!.headers).toEqual(snapshot.metadata.sections![0]!.headers);
  expect(snapshot.metadata.word!.fields).toHaveLength(3);
  for (const [index, field] of snapshot.metadata.word!.fields.entries()) {
    expect(field.endNodeId).toBeDefined();
    expect(field.instructionFields).toBeUndefined();
    expect(field.resultNodeIds!.map((id) => model.getNode(id)!.text).join("")).toBe(expected[index]);
  }
  expect(snapshot.annotations.map(({ anchor }) => anchor.kind === "text" ? anchor.quote
    : anchor.kind === "text-set" ? anchor.ranges.map((range) => range.quote).join("") : undefined)).toEqual(expected);
  expect(snapshot.metadata.word!.bookmarks.map((bookmark) => bookmark.bookmarkId).sort()).toEqual([
    "footer-1-0", "header-0-0", "main-0",
  ]);
});

test.each([
  ['<w:r><w:fldChar w:fldCharType="begin"/></w:r>', '<w:r><w:fldChar w:fldCharType="end"/></w:r>', "A field control has no matching begin."],
  ['<w:commentRangeStart w:id="0"/>', '<w:commentRangeEnd w:id="0"/>', "A comment range end is invalid."],
  ['<w:bookmarkStart w:id="0" w:name="crossing"/>', '<w:bookmarkEnd w:id="0"/>', "A bookmark end has no start."],
])("rejects an auxiliary endpoint that only matches a main-story start: %s", (start, end, message) => {
  const bytes = mutateDocx("08-comments-revisions.docx", {
    "word/document.xml": (source) => body(source,
      '<w:p><w:pPr><w:sectPr><w:headerReference w:type="default" r:id="header1"/></w:sectPr></w:pPr>' +
      start + '</w:p><w:p>' + end + '</w:p>'),
    "word/header1.xml": () => '<w:hdr xmlns:w="' + wordNamespace + '"><w:p>' + end + '</w:p></w:hdr>',
    "word/_rels/document.xml.rels": (source) => source.replace('</Relationships>',
      '<Relationship Id="header1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>'),
    "[Content_Types].xml": (source) => source.replace('</Types>',
      '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>'),
  });
  expect(() => importDocx(bytes, { documentId: "crossing" })).toThrowError(expect.objectContaining({ code: "InvalidDocument", message }));
});
test("bounds the 8,000 by 8,000 comment expansion before allocating its ranges", () => {
  expect(() => importDocx(overlappingComments(8000, 8000), { documentId: "budget" }))
    .toThrowError(expect.objectContaining({
      code: "ResourceLimitExceeded",
      message: expect.stringContaining("Generated comment ranges"),
    }));
});
test.each([
  { count: 80, runs: 80, chunks: 1, quote: "x", limits: { maxNodes: 2000 } },
  { count: 40, runs: 40, chunks: 1, quote: "x".repeat(100), limits: { maxFileBytes: 32_768 } },
  { count: 20, runs: 1, chunks: 100, quote: "x".repeat(128), limits: { maxFileBytes: 32_768 } },
  { count: 40, runs: 40, chunks: 1, quote: "x".repeat(100), limits: { maxPackageBytes: 65_536 } },
])("bounds cumulative comment range counts and new/extended quote bytes: %j", (input) => {
  expect(() => importDocx(overlappingComments(input.count, input.runs, input.chunks, input.quote), {
    documentId: "budget", limits: input.limits,
  })).toThrowError(expect.objectContaining({
    code: "ResourceLimitExceeded",
    message: expect.stringContaining("Generated comment ranges"),
  }));
});
test("keeps ordinary overlapping comment ranges correct after reopening", () => {
  const model = reopen(overlappingComments(3, 3, 2, "中😀"));
  for (const annotation of model.snapshot().annotations) {
    expect(annotation.anchor.kind).toBe("text-set");
    if (annotation.anchor.kind === "text-set")
      expect(annotation.anchor.ranges.map((range) => range.quote).join("")).toBe("中😀".repeat(6));
  }
});
test("bounds references expanded through nested cached field results", () => {
  let content = "<w:r><w:t>x</w:t></w:r>".repeat(100);
  for (let index = 0; index < 20; index++)
    content = complexField(" REF cache ", content);
  const bytes = mutateDocx("07-fields.docx", {
    "word/document.xml": (source) => body(source, "<w:p>" + content + "</w:p>"),
  });
  expect(() => importDocx(bytes, { documentId: "fields", limits: { maxNodes: 1500 } }))
    .toThrowError(expect.objectContaining({
      code: "ResourceLimitExceeded",
      message: expect.stringContaining("Generated field result references"),
    }));
});

test.each([
  ["cell property revision", "<w:tcPrChange/>", "<w:p/>"],
  ["supported cell direction", '<w:textDirection w:val="tbRl"/>', "<w:p/>"],
  ["paragraph property revision", "", "<w:p><w:pPr><w:pPrChange/></w:pPr></w:p>"],
  ["unsupported tab alignment", "", '<w:p><w:pPr><w:tabs><w:tab w:pos="100" w:val="bar"/></w:tabs></w:pPr></w:p>'],
  ["run property revision", "", "<w:p><w:r><w:rPr><w:rPrChange/></w:rPr></w:r></w:p>"],
  ["break", "", "<w:p><w:r><w:br/></w:r></w:p>"],
  ["drawing", "", "<w:p><w:r><w:drawing/></w:r></w:p>"],
  ["bookmark", "", '<w:p><w:bookmarkStart w:id="55" w:name="lost"/></w:p>'],
  ["whitespace", "", '<w:p><w:r><w:t xml:space="preserve"> </w:t></w:r></w:p>'],
])("validates merge continuations before dropping %s", (_name, properties, content) => {
  const bytes = mutateDocx("02-table-merges.docx", {
    "word/document.xml": (source) => source.replace(
      /<w:tc><w:tcPr><w:vMerge\/><\/w:tcPr>[\s\S]*?<\/w:tc>/,
      "<w:tc><w:tcPr><w:vMerge/>" + properties + "</w:tcPr>" + content + "</w:tc>",
    ),
  });
  if(_name==='supported cell direction')expect(()=>importDocx(bytes,{documentId:"merge",unsupportedContent:"reject"})).not.toThrow();else expect(() => importDocx(bytes, { documentId: "merge",unsupportedContent:"reject" })).toThrowError(expect.objectContaining({ code: "UnsupportedFeature" }));
});
test("preserves inherited table geometry, partial margin overrides and explicit highlight removal", () => {
  const model = reopen(styleInheritanceFixture());
  const snapshot = model.snapshot();
  const table = snapshot.metadata.word!.tables[0]!;
  expect(sourceCss(model,model.getNode(table.nodeId)!)).toContain("margin-left:30pt");
  for (const cell of descendants(model, table.nodeId).filter((node) => node.tagName === "td")) {
    expect(sourceCss(model,cell)).toContain("padding-left:30pt");
    expect(sourceCss(model,cell)).toContain("padding-right:40pt");
    expect(sourceCss(model,cell)).toContain("padding-top:2pt");
    expect(sourceCss(model,cell)).toContain("padding-bottom:3pt");
  }
  expect(snapshot.metadata.word!.tabs[0]!.stops.map((stop) => stop.position)).toEqual([100, 200]);
  expect(snapshot.html).toContain("background-color:var(--word-run-shading,transparent)");
  expect(snapshot.html).toContain('data-word-script="hAnsi"');
});
test("direct table and cell geometry overrides inherited geometry", () => {
  const bytes = mutateDocx("02-table-merges.docx", {
    "word/styles.xml": () => styles('<w:style w:type="table" w:styleId="T"><w:tblPr><w:tblInd w:w="600"/><w:tblCellMar><w:left w:w="600"/><w:right w:w="800"/></w:tblCellMar></w:tblPr></w:style>'),
    "word/document.xml": (source) => source
      .replace("<w:tblPr>", '<w:tblPr><w:tblStyle w:val="T"/><w:tblInd w:w="400"/><w:tblCellMar><w:left w:w="200"/></w:tblCellMar>')
      .replace("<w:tcPr>", '<w:tcPr><w:tcMar><w:left w:w="100"/></w:tcMar>'),
  });
  const model = reopen(bytes);
  const tableId = model.snapshot().metadata.word!.tables[0]!.nodeId;
  expect(sourceCss(model,model.getNode(tableId)!)).toContain("margin-left:20pt");
  const cells = descendants(model, tableId).filter((node) => node.tagName === "td");
  expect(sourceCss(model,cells[0]!)).toContain("padding-left:5pt");
  expect(sourceCss(model,cells[1]!)).toContain("padding-left:10pt");
  expect(sourceCss(model,cells[1]!)).toContain("padding-right:40pt");
});
test("preserves ignored wholeTable definitions and applies active region-specific margins", () => {
  const bytes = mutateDocx("02-table-merges.docx", {
    "word/styles.xml": () => styles(
      '<w:style w:type="table" w:styleId="T"><w:tblPr><w:tblInd w:w="400"/></w:tblPr>' +
      '<w:tblStylePr w:type="wholeTable"><w:tblPr><w:tblInd w:w="800"/><w:tblCellMar><w:left w:w="600"/><w:right w:w="800"/></w:tblCellMar></w:tblPr></w:tblStylePr>' +
      '<w:tblStylePr w:type="firstRow"><w:tblPr><w:tblCellMar><w:left w:w="1000"/></w:tblCellMar></w:tblPr></w:tblStylePr></w:style>'),
    "word/document.xml": (source) => source.replace("<w:tblPr>",
      '<w:tblPr><w:tblStyle w:val="T"/><w:tblLook w:firstRow="1" w:noHBand="1" w:noVBand="1"/>'),
  });
  const model = reopen(bytes);
  const id = model.snapshot().metadata.word!.tables[0]!.nodeId;
  expect(sourceCss(model,model.getNode(id)!)).toContain("margin-left:20pt");
  const cells = descendants(model, id).filter((node) => node.tagName === "td");
  expect(sourceCss(model,cells[0]!)).toContain("padding-left:50pt");
  expect(sourceCss(model,cells[2]!)).toContain("padding-left:5.4pt");
  expect(cells.every((node) => sourceCss(model,node).includes("padding-right:5.4pt"))).toBe(true);
});
test("rejects table indentation in a conditional cell region", () => {
  const bytes = mutateDocx("02-table-merges.docx", {
    "word/styles.xml": () => styles(
      '<w:style w:type="table" w:styleId="T"><w:tblStylePr w:type="firstRow"><w:tblPr><w:tblInd w:w="600"/></w:tblPr></w:tblStylePr></w:style>'),
  });
  expect(() => importDocx(bytes, { documentId: "conditional" }))
    .toThrowError(expect.objectContaining({ code: "UnsupportedFeature" }));
});
test("merges tab stops through defaults, base styles and direct clear/replace declarations", () => {
  const bytes = mutateDocx("06-links-tabs.docx", {
    "word/styles.xml": () => styles(
      '<w:docDefaults><w:pPrDefault><w:pPr><w:tabs><w:tab w:val="left" w:pos="1000"/></w:tabs></w:pPr></w:pPrDefault></w:docDefaults>' +
      '<w:style w:type="paragraph" w:styleId="Base"><w:pPr><w:tabs><w:tab w:val="left" w:pos="2000"/></w:tabs></w:pPr></w:style>' +
      '<w:style w:type="paragraph" w:styleId="Middle"><w:basedOn w:val="Base"/><w:pPr><w:tabs><w:tab w:val="center" w:pos="3000"/></w:tabs></w:pPr></w:style>' +
      '<w:style w:type="paragraph" w:styleId="Derived"><w:basedOn w:val="Middle"/><w:pPr><w:tabs><w:tab w:val="clear" w:pos="2000"/><w:tab w:val="left" w:pos="6000"/></w:tabs></w:pPr></w:style>',
    ),
    "word/document.xml": (source) => body(source,
      '<w:p><w:pPr><w:pStyle w:val="Middle"/><w:tabs><w:tab w:val="left" w:pos="4000"/></w:tabs></w:pPr><w:r><w:tab/></w:r></w:p>' +
      '<w:p><w:pPr><w:pStyle w:val="Derived"/><w:tabs><w:tab w:val="clear" w:pos="3000"/><w:tab w:val="right" w:pos="6000" w:leader="dot"/></w:tabs></w:pPr><w:r><w:tab/></w:r></w:p>' +
      '<w:p><w:r><w:tab/></w:r></w:p>'),
  });
  const tabs = reopen(bytes).snapshot().metadata.word!.tabs;
  expect(tabs[0]!.stops.map((stop) => stop.position)).toEqual([50, 100, 150, 200]);
  expect(tabs[1]!.stops).toEqual([
    { position: 50, alignment: "left", leader: "none" },
    { position: 300, alignment: "right", leader: "dot" },
  ]);
  expect(tabs[2]!.stops).toEqual([{ position: 50, alignment: "left", leader: "none" }]);
});
test("fails explicitly when a theme font cannot be resolved despite an explicit face", () => {
  const bytes = mutateDocx("09-chinese-fonts.docx", {
    "word/document.xml": (source) => source.replace('w:eastAsia="Noto Serif CJK SC"', 'w:eastAsia="OldFace" w:eastAsiaTheme="minorEastAsia"'),
  });
  expect(() => importDocx(bytes, { documentId: "theme" }))
    .toThrowError(expect.objectContaining({ code: "UnsupportedFeature" }));
});
test.each(["table", "row", "style", "style row", "conditional style"])("maps nonzero table and row spacing in %s properties", (location) => {
  const spacing = '<w:tblCellSpacing w:w="100" w:type="dxa"/>';
  const bytes = mutateDocx("02-table-merges.docx", {
    "word/styles.xml": (source) => location.includes("style") ? source.replace("</w:styles>",
      '<w:style w:type="table" w:styleId="Spaced">' +
      (location === "conditional style" ? '<w:tblStylePr w:type="firstRow"><w:tblPr>' + spacing + "</w:tblPr></w:tblStylePr>" :
        location === "style row" ? "<w:trPr>" + spacing + "</w:trPr>" : "<w:tblPr>" + spacing + "</w:tblPr>") +
      "</w:style></w:styles>") : source,
    "word/document.xml": (source) => location === "table" ? source.replace("<w:tblPr>", "<w:tblPr>" + spacing) :
      location === "row" ? source.replace("<w:tr>", "<w:tr><w:trPr>" + spacing + "</w:trPr>") : source,
  });
  expect(()=>importDocx(bytes,{documentId:"spacing"})).not.toThrow();
});
test("accepts explicit zero table and row spacing", () => {
  reopen(mutateDocx("02-table-merges.docx", {
    "word/document.xml": (source) => source
      .replace("<w:tblPr>", '<w:tblPr><w:tblCellSpacing w:w="0" w:type="dxa"/>')
      .replace("<w:tr>", '<w:tr><w:trPr><w:tblCellSpacing w:w="0" w:type="dxa"/></w:trPr>'),
  }));
});
test.each([
  "<a:effectDag/>", "<a:scene3d/>", "<a:sp3d/>",
  '<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>',
])("rejects unmapped picture appearance: %s", (appearance) => {
  const bytes = mutateDocx("03-images.docx", {
    "word/document.xml": (source) => source.replace("</pic:spPr>", appearance + "</pic:spPr>"),
  });
  expect(() => importDocx(bytes, { documentId: "picture" })).toThrowError(
    expect.objectContaining({ code: "UnsupportedFeature" }),
  );
});
test.each(['<a:ln w="12700"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln>', '<a:ln/>'])("maps editable native picture outlines: %s",appearance=>{
 const model=reopen(mutateDocx('03-images.docx',{'word/document.xml':source=>source.replace('</pic:spPr>',appearance+'</pic:spPr>')}));expect(model.snapshot().metadata.word!.drawings!.some(item=>item.kind==='image')).toBe(true);expect(model.snapshot().html).toContain('--office-line-');
});
test("accepts explicit no-fill picture outlines and empty effects", () => {
  reopen(mutateDocx("03-images.docx", {
    "word/document.xml": (source) => source.replace("</pic:spPr>", "<a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst/></pic:spPr>"),
  }));
});
test.each([false, true])("records every cached result in nested instruction fields (simple=%s)", (simple) => {
  const model = reopen(nestedFieldFixture(simple));
  const fields = model.snapshot().metadata.word!.fields;
  const outer = fields.find((field) => field.instruction.includes(" IF "))!;
  expect(outer.instructionFields).toHaveLength(1);
  const inner = fields.find((field) => field.nodeId === outer.instructionFields![0]!.fieldNodeId)!;
  const content = inner.resultNodeIds!.map((id) => model.getNode(id)!);
  expect(content.some((node) => node.tagName === "br")).toBe(true);
  expect(content.some((node) => node.tagName === "img")).toBe(true);
  expect(content.some((node) => node.attributes["data-word-tab"] === "true")).toBe(true);
  expect(content.some((node) => node.text === "child cache")).toBe(true);
  expect(outer.resultNodeIds!.map((id) => model.getNode(id)!.text).join("")).toBe("outer result");
});
