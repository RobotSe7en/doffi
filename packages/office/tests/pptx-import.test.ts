import { openContentFixture as openDocument } from "../../../tests/content-fixture.js";
import { readFileSync } from "node:fs";
import { expect, test } from "vitest";
import { createDocumentSession, DocumentError, readPresentationGeometry, resolvePresentationGeometry } from "@doffi/core";
import { exportPptx, importPptx, readDocumentZip, writeDocumentZip } from "@doffi/office";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";

const directory = new URL("../../../examples/pptx/", import.meta.url);
const expected = JSON.parse(readFileSync(new URL("expected.json", directory), "utf8")) as {
  cases: { file: string; text: string[]; slides: number; objects?: number }[];
  negativeCases: { file: string; error: string }[];
};
const fixture = (name = "01-rich-text.pptx") => new Uint8Array(readFileSync(new URL(name, directory)));
const load = (name = "01-rich-text.pptx") => openDocument(importPptx(fixture(name), { documentId: "ppt-test" }));
function mutate(name: string, part: string, change: (source: string) => string): Uint8Array {
  const files = unzipSync(fixture(name)); files[part] = strToU8(change(strFromU8(files[part]!))); return zipSync(files);
}
const presentationNamespace = "http://schemas.openxmlformats.org/presentationml/2006/main";
const relationshipNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const change = (files: Record<string, Uint8Array>, part: string, edit: (xml: string) => string) => {
  files[part] = strToU8(edit(strFromU8(files[part]!)));
};
const readingLater = /^(09|10|11|12|13)-/;
for (const item of expected.cases.filter(c => !readingLater.test(c.file))) for (const producer of ["", "producer/"]) {
  test("imports independent PPTX base objects and reopens the HTML package: " + producer + item.file, () => {
    const bytes = fixture(producer + item.file), before = new Uint8Array(bytes);
    const files = importPptx(bytes, { documentId: "ppt-test" }), model = openDocument(files), snapshot = model.snapshot();
    expect(bytes).toEqual(before);
    expect(snapshot.manifest).toMatchObject({ documentId: "ppt-test", revision: 0, kind: "presentation" });
    expect(snapshot.metadata.presentation!.slides).toHaveLength(item.slides);
    const slideText = snapshot.metadata.presentation!.slides.map(slide => model.getNode(slide.rootNodeId)!.text).join("\n");
    for (const text of item.text) expect(slideText).toContain(text);
    if (item.objects !== undefined) expect(snapshot.metadata.presentation!.objects).toHaveLength(item.objects);
    expect(openDocument(model.serialize()).snapshot()).toEqual(snapshot);
    expect(openDocument(readDocumentZip(writeDocumentZip(files))).snapshot()).toEqual(snapshot);
    expect(importPptx(bytes, { documentId: "ppt-test" })).toEqual(files);
    expect([...files.keys()].some(name => name.endsWith(".pptx"))).toBe(false);
  });
}

test("keeps source font roles, Unicode variation sequences, paragraphs, lists and links", () => {
  const model = load(), snapshot = model.snapshot(), p = snapshot.metadata.presentation!;
  const slide = model.getNode(p.slides[0]!.rootNodeId)!;
  expect(slide.text).toContain("扩展字𠀀与变体葛󠄀");
  expect(slide.text).toContain("换行前\n换行后");
  expect(snapshot.html).toContain('data-office-numbering="arabicPeriod"');
  expect(snapshot.html).toContain('data-office-bullet="•"');
  expect(snapshot.html).toContain("font-weight:bold");
  expect(snapshot.html).toContain("--office-font-east-asia");
  expect(snapshot.html).toContain("Noto Sans CJK SC");
  expect(p.hyperlinks).toEqual([expect.objectContaining({ target: "https://example.com/pptx" })]);
  expect(snapshot.html).not.toContain('href="https://example.com');
});

test.each([0, 7, 9999])("preserves native firstSlideNum=%i through package and native reopening", first => {
  const bytes = mutate("01-rich-text.pptx", "ppt/presentation.xml", source =>
    source.replace("<p:presentation ", `<p:presentation firstSlideNum="${first}" `));
  const model = openDocument(importPptx(bytes, { documentId: "slide-numbering" }));
  expect(model.snapshot().metadata.presentation!.firstSlideNumber).toBe(first);
  expect(openDocument(model.serialize()).snapshot()).toEqual(model.snapshot());
  const output = exportPptx(model);
  expect(strFromU8(unzipSync(output.bytes)["ppt/presentation.xml"]!)).toContain(`firstSlideNum="${first}"`);
  expect(openDocument(importPptx(output.bytes, { documentId: "number-reopen" })).snapshot().metadata.presentation!.firstSlideNumber).toBe(first);
});

test.each([-1, 10000, 1.5])("rejects a first slide number outside the Office range: %i", first => {
  const bytes = mutate("01-rich-text.pptx", "ppt/presentation.xml", source =>
    source.replace("<p:presentation ", `<p:presentation firstSlideNum="${first}" `));
  expect(() => importPptx(bytes, { documentId: "bad-number" })).toThrowError(expect.objectContaining({ code: "InvalidDocument" }));
});

test("retains an out-of-order east-Asian font as inactive while matching PowerPoint inheritance", () => {
  const bytes = mutate("01-rich-text.pptx", "ppt/slides/slide1.xml", source => source.replace(
    '<a:ea typeface="Noto Sans CJK SC"/><a:cs typeface="Liberation Sans"/>',
    '<a:cs typeface="Liberation Sans"/><a:ea typeface="Noto Serif CJK SC"/>'));
  const model = openDocument(importPptx(bytes, { documentId: "ordered-fonts" }));
  const slide = model.getNode(model.snapshot().metadata.presentation!.slides[0]!.rootNodeId)!;
  const object = model.getNode(slide.childNodeIds[0]!)!, paragraph = model.getNode(object.childNodeIds[0]!)!;
  const run = model.getNode(paragraph.childNodeIds[0]!)!;
  expect(run.attributes.style).toContain('--office-ignored-font-east-asia:"Noto Serif CJK SC"');
  expect(run.attributes.style).not.toContain('--office-font-east-asia:');
  expect(openDocument(model.serialize()).snapshot()).toEqual(model.snapshot());
  // Canonical native export expresses the effective font cascade, never an invalid child order.
  const output = exportPptx(model);
  expect(strFromU8(unzipSync(output.bytes)["ppt/slides/slide1.xml"]!)).not.toContain('typeface="Noto Serif CJK SC"');
  const reopened = openDocument(importPptx(output.bytes, { documentId: "font-reopen" }));
  expect(reopened.getNode(reopened.snapshot().metadata.presentation!.slides[0]!.rootNodeId)!.text).toBe(slide.text);
});

test("imports repeated native run formatting within default CSS limits through editing and native reopening", () => {
  const bytes = mutate("01-rich-text.pptx", "ppt/slides/slide1.xml", source => {
    const run = source.match(/<a:r>.*?<\/a:r>/s)![0];
    // Parsing these identical declarations separately exceeds 100,000 CSS nodes.
    return source.replace(run, run.repeat(6000));
  });
  const before = new Uint8Array(bytes);
  const files = importPptx(bytes, { documentId: "repeated-runs" });
  const model = openDocument(files), snapshot = model.snapshot();
  const slide = model.getNode(snapshot.metadata.presentation!.slides[0]!.rootNodeId)!;
  const object = model.getNode(slide.childNodeIds[0]!)!;
  const paragraph = model.getNode(object.childNodeIds[0]!)!;
  const run = model.getNode(paragraph.childNodeIds[0]!)!;
  expect(paragraph.text).toBe("中文与 Latin 😀，".repeat(6000) + "粗体和斜体");
  expect(run.attributes.style).toContain("--office-font-east-asia");
  expect(bytes).toEqual(before);
  expect(openDocument(readDocumentZip(writeDocumentZip(files))).snapshot()).toEqual(snapshot);
  expect(() => importPptx(bytes, { documentId: "repeated-runs", limits: { maxCssNodes: 6000 } }))
    .toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded" }));
  const session = createDocumentSession(model);
  try {
    session.commit(session.prepare({ protocolVersion: 1, documentId: "repeated-runs", baseRevision: 0,
      operations: [{ op: "replaceText", nodeId: run.nodeId, from: 0, to: 0, expectedText: "", text: "新增" }] }));
    const reopened = openDocument(importPptx(exportPptx(session).bytes, { documentId: "native-reopen" }));
    const reopenedSlide = reopened.snapshot().metadata.presentation!.slides[0]!;
    expect(reopened.getNode(reopenedSlide.rootNodeId)!.text).toBe("新增" + slide.text);
    session.undo();
    expect(session.snapshot().html).toBe(snapshot.html);
  } finally { session.dispose(); }
});

test("preserves placeholder inheritance without materializing a second transform", () => {
  const model = load("02-theme-placeholders.pptx"), p = model.snapshot().metadata.presentation!;
  expect(p.masters).toHaveLength(2);
  const slide = model.getNode(p.slides[0]!.rootNodeId)!, object = model.getNode(slide.childNodeIds[0]!)!;
  expect(readPresentationGeometry(object.attributes)).toBeUndefined();
  const layout = p.layouts.find(l => l.layoutId === p.slides[0]!.layoutId)!;
  const inherited = model.getNode(model.getNode(layout.rootNodeId)!.childNodeIds[0]!)!;
  expect(readPresentationGeometry(inherited.attributes)).toMatchObject({ x: 60, y: 40, width: 840, height: 100 });
  expect(model.snapshot().styles[0]!.css).toContain("--office-palette-accent1:#c0504d");
});

test("resolves percent-encoded part names without relying on the reference application's path support", () => {
  const files = unzipSync(fixture("17-renamed-parts.pptx"));
  for (const name of Object.keys(files)) {
    const target = name.replace("page-one.xml", "page one.xml");
    const bytes = files[name]!;
    if (target !== name) delete files[name];
    files[target] = name.endsWith(".xml") || name.endsWith(".rels") ? strToU8(strFromU8(bytes).replaceAll("page-one.xml", "page%20one.xml")) : bytes;
  }
  const model = openDocument(importPptx(zipSync(files), { documentId: "encoded" }));
  const p = model.snapshot().metadata.presentation!;
  expect(model.getNode(p.slides[0]!.rootNodeId)!.text).toBe("命名空间不依赖前缀");
});

test("keeps nested local transforms, negative coordinates, rotation and flips", () => {
  const model = load("05-groups.pptx"), p = model.snapshot().metadata.presentation!;
  const group = p.objects.find(o => o.kind === "group" && o.nativeId === "2")!;
  const node = model.getNode(group.nodeId)!;
  expect(node.attributes).toMatchObject({ "data-office-child-x": "-50", "data-office-child-y": "-25", "data-office-child-width": "900", "data-office-child-height": "450" });
  expect(readPresentationGeometry(node.attributes)).toMatchObject({ rotation: 10, flipH: true });
  expect(readPresentationGeometry(model.getNode(p.objects.find(o => o.nativeId === "5")!.nodeId)!.attributes)).toMatchObject({ x: -20, y: 0 });
});

test("resolves connector identities after all objects are mapped and retains native elbow adjustments", () => {
  for (const prefix of ["", "producer/"]) {
    const model = load(prefix + "06-connectors.pptx"), p = model.snapshot().metadata.presentation!;
    const a = p.objects.find(o => o.nativeId === "2")!, b = p.objects.find(o => o.nativeId === "3")!;
    for (const connector of p.objects.filter(o => o.kind === "connector")) expect(connector).toMatchObject({ startNodeId: a.nodeId, endNodeId: b.nodeId, startSite: 1, endSite: 3 });
    if (prefix) expect(model.snapshot().html).toContain('data-office-connector-preset="bentConnector4"');
  }
});

test("retains image bytes, crop and accessible description without fetching external resources", () => {
  const bytes = fixture("07-images.pptx"), native = unzipSync(bytes), files = importPptx(bytes, { documentId: "images" });
  const model = openDocument(files), snapshot = model.snapshot();
  expect(snapshot.manifest.assets.map(a => a.mimeType).sort()).toEqual(["image/jpeg", "image/png", "image/webp"]);
  for (const asset of snapshot.manifest.assets) expect(files.get(asset.path)).toEqual(native[asset.path.slice("assets/".length)]);
  expect(snapshot.html).toContain('data-office-crop-left="0.1"');
  expect(snapshot.html).toContain('alt="公开合成图片"');
});

test("maps native merged table cells to one HTML cell per merge origin", () => {
  for (const prefix of ["", "producer/"]) {
    const model = load(prefix + "08-table-merges.pptx"), p = model.snapshot().metadata.presentation!;
    const table = p.objects.find(o => o.kind === "table")!;
    if (table.kind !== "table") throw new Error("Expected table");
    // LibreOffice quantizes these native grid widths on resave; retain its actual
    // coordinates rather than snapping them back to the synthetic source.
    expect(table.grid).toEqual(Array(3).fill(prefix ? 239.98110236220472 : 240));
    expect(model.getNode(table.nodeId)!.html).toContain('colspan="2"');
    expect(model.getNode(table.nodeId)!.html).toContain('rowspan="2"');
    expect(model.getNode(table.nodeId)!.html.match(/<td /g)).toHaveLength(7);
  }
});

for (const item of expected.negativeCases) test("rejects malformed or unsupported native PPTX: " + item.file, () => {
  try { importPptx(fixture(item.file), { documentId: "invalid" }); throw new Error("Accepted invalid source"); }
  catch (error) { expect(error).toBeInstanceOf(DocumentError); expect((error as DocumentError).code).toBe(item.error); }
});

test("rejects wrong types, cyclic ownership, unsupported geometry and broken connection targets", () => {
  const cases: [string, string, (s: string) => string, string][] = [
    ["01-rich-text.pptx", "[Content_Types].xml", s => s.replace("presentationml.slide+xml", "presentationml.slideLayout+xml"), "InvalidDocument"],
    ["01-rich-text.pptx", "ppt/slideLayouts/_rels/slideLayout1.xml.rels", s => s.replace("../slideMasters/slideMaster1.xml", "../slideLayouts/slideLayout1.xml"), "InvalidDocument"],
    ["01-rich-text.pptx", "ppt/slides/slide1.xml", s => s.replace('prst="rect"', 'prst="cloud"'), "UnsupportedFeature"],
    ["06-connectors.pptx", "ppt/slides/slide1.xml", s => s.replace('<a:stCxn id="2"', '<a:stCxn id="999"'), "InvalidDocument"],
    ["08-table-merges.pptx", "ppt/slides/slide1.xml", s => s.replace('gridSpan="2"', 'gridSpan="99999999"'), "InvalidDocument"],
  ];
  for (const [name, part, change, code] of cases) expect(() => importPptx(mutate(name, part, change), { documentId: "bad" })).toThrowError(expect.objectContaining({ code }));
});

test("enforces byte/node/ZIP budgets, identities, format separation and opens accepted PPT sessions", () => {
  expect(() => importPptx(fixture(), { documentId: "" })).toThrowError(expect.objectContaining({ code: "InvalidDocument" }));
  expect(() => importPptx(fixture(), { documentId: "small", maxZipBytes: 10 })).toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded" }));
  expect(() => importPptx(fixture(), { documentId: "small", limits: { maxNodes: 30 } })).toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded" }));
  const docx = readFileSync(new URL("../../../examples/docx/01-text-numbering.docx", import.meta.url));
  expect(() => importPptx(docx, { documentId: "wrong" })).toThrowError(expect.objectContaining({ code: "UnsupportedFormat" }));
  expect(createDocumentSession(load()).snapshot().manifest.kind).toBe("presentation");
});

test("maps imported modern UTF-16 anchors and retains native rich threads through history and ZIP reopening", () => {
  const model = load("12-modern-comments.pptx"), session = createDocumentSession(model), original = session.getAnnotations().find(a => a.anchor.kind === "text")!;
  if (original.anchor.kind !== "text") throw Error("Expected precise imported text anchor");
  const transaction = { protocolVersion: 1 as const, documentId: "ppt-test", baseRevision: 0, operations: [{ op: "replaceText" as const, nodeId: original.anchor.nodeId, from: 0, to: 0, expectedText: "", text: "前󠄀😀" }] };
  session.commit(session.prepare(transaction));
  expect(session.getAnnotations().find(a => a.annotationId === original.annotationId)!.anchor).toMatchObject({ ...original.anchor, from: original.anchor.from + 5, to: original.anchor.to + 5 });
  expect(openDocument(readDocumentZip(writeDocumentZip(session.serialize()))).snapshot()).toEqual(session.snapshot());
  session.undo(); expect(session.snapshot().html).toBe(model.snapshot().html); expect(session.getAnnotations().map(a => a.replies)).toEqual(model.snapshot().annotations.map(a => a.replies));
});

test("ungrouping the independent native fixture keeps its explicitly calculated center, rotation and glyph flip parity",()=>{
  const model=load("05-groups.pptx"),p=model.snapshot().metadata.presentation!,group=p.objects.find(o=>o.kind==="group"&&model.getNode(o.nodeId)!.parentNodeId===p.slides[0]!.rootNodeId)!,child=model.getNode(group.nodeId)!.childNodeIds[0]!,s=createDocumentSession(model);
  expect(readPresentationGeometry(model.getNode(child)!.attributes)).toMatchObject({x:0,y:0,width:160,height:80});
  s.commit(s.prepare({protocolVersion:1,documentId:"ppt-test",baseRevision:0,operations:[{op:"ungroupPresentationObject",nodeId:group.nodeId}]}));
  const g=readPresentationGeometry(s.getNode(child)!.attributes)!,r=10*Math.PI/180;
  // Fixture: group (48,120,600,300), child origin (-50,-25), child extents
  // (900,450), rotation 10deg and horizontal flip; child center is (80,40).
  const dx=-((80+50)*2/3-300),dy=(40+25)*2/3-150;
  expect(g.x+g.width/2).toBeCloseTo(348+dx*Math.cos(r)-dy*Math.sin(r),9);
  expect(g.y+g.height/2).toBeCloseTo(270+dx*Math.sin(r)+dy*Math.cos(r),9);
  expect(g.width).toBeCloseTo(160*2/3,9);expect(g.height).toBeCloseTo(80*2/3,9);expect(g.rotation).toBeCloseTo(10,9);expect(g.flipH).toBe(true);expect(g.flipV).toBe(false);
  s.undo();expect(s.snapshot().html).toBe(model.snapshot().html);
});

test("compares native numeric identities by value, including leading-zero aliases and unsigned bounds", () => {
  const duplicate = mutate("01-rich-text.pptx", "ppt/slides/slide1.xml", source => {
    const shape = source.match(/<p:sp>.*?<\/p:sp>/s)![0];
    return source.replace("</p:spTree>", shape.replace('id="2"', 'id="02"') + "</p:spTree>");
  });
  const outside = mutate("01-rich-text.pptx", "ppt/slides/slide1.xml", source => source.replace('id="2"', 'id="4294967296"'));
  const masters = mutate("02-theme-placeholders.pptx", "ppt/presentation.xml", source => source.replace('id="2147483649"', 'id="2147483648"'));
  for (const bytes of [duplicate, outside, masters]) expect(() => importPptx(bytes, { documentId: "identity" })).toThrowError(expect.objectContaining({ code: "InvalidDocument" }));
});

test("rejects a table continuation whose direction disagrees with its merge origin", () => {
  const bytes = mutate("08-table-merges.pptx", "ppt/slides/slide1.xml", source => source.replace('vMerge="1"', 'hMerge="1"'));
  expect(() => importPptx(bytes, { documentId: "merge" })).toThrowError(expect.objectContaining({ code: "InvalidDocument" }));
});

test("preserves the reference color and transforms for a theme background", () => {
  const bytes = mutate("01-rich-text.pptx", "ppt/slides/slide1.xml", source => source.replace('<p:spTree>', '<p:bg><p:bgRef idx="1001"><a:srgbClr val="123456"><a:alpha val="50000"/></a:srgbClr></p:bgRef></p:bg><p:spTree>'));
  const snapshot = openDocument(importPptx(bytes, { documentId: "background" })).snapshot();
  expect(snapshot.html).toContain("--office-background-reference-color:#123456");
  expect(snapshot.html).toContain("--office-background-reference-alpha:0.5");
  expect(snapshot.styles[0]!.css).toContain("var(--office-background-reference-color");
});

test("compatible PPTX retains theme insertion defaults and static animation without losing the slide", () => {
  const files = unzipSync(mutate("01-rich-text.pptx", "ppt/theme/theme1.xml", xml => xml.replace(/<a:objectDefaults\s*\/>|<a:objectDefaults><\/a:objectDefaults>|(?=<\/a:theme>)/,
    '<a:objectDefaults><a:spDef><a:spPr><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></a:spPr><a:bodyPr/><a:lstStyle/></a:spDef></a:objectDefaults>')));
  const slide = "ppt/slides/slide1.xml", original = strFromU8(files[slide]!);
  const timing = '<p:timing><p:tnLst><p:par><p:cTn id="1"><p:childTnLst><p:set><p:cBhvr><p:cTn id="2"/><p:tgtEl><p:spTgt spid="2"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>';
  files[slide] = strToU8(original.replace("</p:sld>", timing + "</p:sld>"));
  const bytes = zipSync(files), options = { documentId: "static", mode: "compatible" as const };
  expect(() => importPptx(bytes, { documentId: "static" })).toThrowError(expect.objectContaining({ code: "UnsupportedFeature" }));
  const model = openDocument(importPptx(bytes, options)), metadata = model.snapshot().metadata.presentation!;
  expect(metadata.nativePreview).toEqual([
    expect.objectContaining({ feature: "theme-defaults", ownerId: metadata.themes[0]!.themeId, xml: expect.stringContaining('val="00FF00"') }),
    expect.objectContaining({ feature: "animation", ownerId: metadata.slides[0]!.slideId, xml: expect.stringContaining('spid="2"') }),
  ]);
  expect(model.snapshot().styles[0]!.css).not.toContain("#00ff00");
  expect(model.getNode(metadata.slides[0]!.rootNodeId)!.text).toBe(load().getNode(load().snapshot().metadata.presentation!.slides[0]!.rootNodeId)!.text);
  expect(openDocument(readDocumentZip(writeDocumentZip(model.serialize()))).snapshot()).toEqual(model.snapshot());
  expect(() => exportPptx(model)).toThrowError(expect.objectContaining({ code: "UnsupportedFeature" }));
  files[slide] = strToU8(original.replace("</p:sld>", timing.replace('spid="2"', 'spid="99999"') + "</p:sld>"));
  expect(() => importPptx(zipSync(files), options)).toThrowError(expect.objectContaining({ code: "InvalidDocument" }));
  files[slide] = strToU8(original.replace("</p:sld>", timing.replace("<p:set>", "<p:audio/><p:set>") + "</p:sld>"));
  expect(() => importPptx(zipSync(files), options)).toThrowError(expect.objectContaining({ code: "UnsupportedFeature" }));
});

test("compatible PPTX keeps a linked print master, static transition, and unused theme formats", () => {
  const files = unzipSync(fixture()), main = "ppt/presentation.xml", master = "ppt/slideMasters/slideMaster1.xml";
  const theme = "ppt/theme/theme1.xml", handout = "ppt/print/print.xml", slide = "ppt/slides/slide1.xml";
  const mode = { documentId: "compatible-definition", mode: "compatible" as const };
  const timing = '<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"/></p:par></p:tnLst></p:timing>';
  change(files, main, xml => xml.replace("<p:sldIdLst>", '<p:handoutMasterIdLst><p:handoutMasterId r:id="print"/></p:handoutMasterIdLst><p:sldIdLst>'));
  change(files, "ppt/_rels/presentation.xml.rels", xml => xml.replace("</Relationships>",
    '<Relationship Id="print" Type="' + relationshipNamespace + '/handoutMaster" Target="print/print.xml"/></Relationships>'));
  change(files, "[Content_Types].xml", xml => xml.replace("</Types>",
    '<Override PartName="/' + handout + '" ContentType="application/vnd.openxmlformats-officedocument.presentationml.handoutMaster+xml"/></Types>'));
  files[handout] = strToU8('<p:handoutMaster xmlns:p="' + presentationNamespace + '"><p:cSld><p:spTree/></p:cSld></p:handoutMaster>');
  files["ppt/print/_rels/print.xml.rels"] = strToU8('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="theme" Type="' + relationshipNamespace + '/theme" Target="../theme/theme1.xml"/></Relationships>');
  change(files, master, xml => xml.replace("</p:sldMaster>", timing + "</p:sldMaster>"));
  change(files, slide, xml => xml.replace("</p:sld>", '<p:transition><p:wipe dir="r"/></p:transition>' + timing + "</p:sld>"));
  change(files, theme, xml => xml.replace("<a:bgFillStyleLst>", '<a:bgFillStyleLst><a:gradFill><a:gsLst>' +
    '<a:gs pos="0"><a:srgbClr val="000000"/></a:gs><a:gs pos="100000"><a:srgbClr val="FFFFFF"/></a:gs>' +
    '</a:gsLst><a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path></a:gradFill>'));
  const bytes = zipSync(files);
  expect(() => importPptx(bytes, { documentId: mode.documentId })).toThrowError(expect.objectContaining({ code: "UnsupportedFeature" }));
  const document = openDocument(importPptx(bytes, mode)), snapshot = document.snapshot();
  expect(snapshot.metadata.presentation!.nativePreview!.map(item => item.feature)).toEqual(["theme-format", "transition", "handout-master", "handout-theme"]);
  expect(openDocument(readDocumentZip(writeDocumentZip(document.serialize()))).snapshot()).toEqual(snapshot);
  expect(() => exportPptx(document)).toThrowError(expect.objectContaining({ code: "UnsupportedFeature" }));
  const rels = files["ppt/_rels/presentation.xml.rels"]!;
  change(files, "ppt/_rels/presentation.xml.rels", xml => xml.replace('Target="print/print.xml"', 'Target="https://example.com/master.xml" TargetMode="External"'));
  expect(() => importPptx(zipSync(files), mode)).toThrowError(expect.objectContaining({ code: "UnsupportedFeature" }));
  files["ppt/_rels/presentation.xml.rels"] = rels;
  change(files, slide, xml => xml.replace("<p:spTree>", '<p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree>'));
  expect(() => importPptx(zipSync(files), mode)).toThrowError(/referenced theme style cannot be rendered/);
});

test("round trips stretched picture backgrounds and layout placeholders whose master uses a different index", () => {
  const files = unzipSync(fixture("02-theme-placeholders.pptx"));
  const layout = "ppt/slideLayouts/slideLayout2.xml", master = "ppt/slideMasters/slideMaster2.xml";
  const content = unzipSync(fixture("07-images.pptx"))["ppt/media/sample.png"]!;
  files["ppt/media/background.png"] = content;
  change(files, "[Content_Types].xml", xml => xml.replace("</Types>", '<Default Extension="png" ContentType="image/png"/></Types>'));
  change(files, "ppt/slideMasters/_rels/slideMaster2.xml.rels", xml => xml.replace("</Relationships>",
    '<Relationship Id="background" Type="' + relationshipNamespace + '/image" Target="../media/background.png"/></Relationships>'));
  change(files, master, xml => xml.replace('type="title" idx="0"', 'type="title" idx="42"').replace("<p:spTree>",
    '<p:bg><p:bgPr><a:blipFill><a:blip r:embed="background"/><a:srcRect/><a:stretch><a:fillRect/></a:stretch></a:blipFill></p:bgPr></p:bg><p:spTree>'));
  change(files, layout, xml => xml.replace(/<p:sp>.*?<\/p:sp>/s, shape =>
    shape.replace(/<a:xfrm>.*?<\/a:xfrm>/s, "").replace(/<a:bodyPr(?=[\s/>])/, '<a:bodyPr vert="eaVert"')));
  const document = openDocument(importPptx(zipSync(files), { documentId: "background-vertical" }));
  const snapshot = document.snapshot(), presentation = snapshot.metadata.presentation!;
  const slide = document.getNode(presentation.slides[0]!.rootNodeId)!;
  const layoutTitle = presentation.objects.find(object => object.placeholder?.type === "title" &&
    document.getNode(object.nodeId)?.parentNodeId === presentation.layouts.find(item => item.layoutId === presentation.slides[0]!.layoutId)!.rootNodeId)!;
  expect(readPresentationGeometry(document.getNode(layoutTitle.nodeId)!.attributes)).toBeUndefined();
  expect(resolvePresentationGeometry(document, slide.childNodeIds[0]!)).toMatchObject({ x: 48, y: 30, width: 864, height: 90 });
  expect(document.getNode(layoutTitle.nodeId)!.attributes.style).toContain("--office-text-direction:eaVert");
  const output = exportPptx(document).bytes;
  const reopened = openDocument(importPptx(output, { documentId: "background-vertical" }));
  expect(reopened.snapshot().html).toContain('data-office-background-image="true"');
  expect(reopened.snapshot().html).toContain("--office-text-direction:eaVert");
  expect(reopened.snapshot().manifest.assets[0]!.sha256).toBe(snapshot.manifest.assets[0]!.sha256);
});

test("compatible PPTX extracts a full-frame EMF DIB and retains the native bytes", () => {
  const files = unzipSync(fixture("07-images.pptx")), original = "ppt/media/preview.emf";
  const image = new Uint8Array(88 + 128 + 20), fields = new DataView(image.buffer);
  const words = [
    [0, 1], [4, 88], [16, 1], [20, 0], [40, 0x464d4520], [48, image.length], [52, 3],
    [88, 81], [92, 128], [104, 1], [128, 2], [132, 1], [136, 80], [140, 40], [144, 120], [148, 8],
    [156, 0x00cc0020], [160, 2], [164, 1], [168, 40], [172, 2], [176, 1],
    [216, 14], [220, 20], [232, 20],
  ] as const;
  for (const [offset, value] of words) fields.setUint32(offset, value, true);
  fields.setUint16(180, 1, true);
  fields.setUint16(182, 32, true);
  image.set([0, 0, 255, 0, 255, 0, 0, 0], 208);
  files[original] = image;
  change(files, "ppt/slides/_rels/slide1.xml.rels", xml => xml.replace("../media/sample.png", "../media/preview.emf"));
  change(files, "[Content_Types].xml", xml => xml.replace("</Types>", '<Default Extension="emf" ContentType="image/x-emf"/></Types>'));
  const options = { documentId: "embedded-raster", mode: "compatible" as const };
  const document = openDocument(importPptx(zipSync(files), options)), snapshot = document.snapshot();
  const preview = snapshot.metadata.presentation!.nativePreview!.find(item => item.feature === "emf-image")!;
  expect(document.serialize().get(preview.assetPath)).toEqual(image);
  const png = document.serialize().get("preview-assets/" + original + ".png")!;
  expect(new DataView(png.buffer, png.byteOffset).getUint32(16)).toBe(2);
  expect(snapshot.html).toContain("preview-assets/" + original + ".png");
  const session = createDocumentSession(document);
  const picture = document.getNode(preview.ownerId)!;
  const replacement = snapshot.manifest.assets.find(asset => asset.mimeType === "image/jpeg")!;
  session.commit(session.prepare({ protocolVersion: 1, documentId: options.documentId, baseRevision: 0, operations: [
    { op: "replaceAsset", nodeId: picture.childNodeIds[0]!, assetPath: replacement.path },
  ] }));
  expect(session.snapshot().metadata.presentation!.nativePreview).toEqual([]);
  session.undo();
  expect(session.snapshot().metadata.presentation!.nativePreview).toEqual([preview]);
  fields.setUint32(144, 128, true);
  expect(() => importPptx(zipSync(files), options)).toThrowError(expect.objectContaining({ code: "InvalidDocument" }));
});

test("compatible built-in table styles require explicit cell fills and reject unknown styles", () => {
  const main = "ppt/slides/slide1.xml";
  const preset = '<a:tableStyleId>{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}</a:tableStyleId>';
  const options = { documentId: "styled-table", mode: "compatible" as const };
  const files = unzipSync(fixture("08-table-merges.pptx"));
  change(files, main, xml => xml.replace(/<a:tblPr(?: [^>]*)?\/>|<a:tblPr(?: [^>]*)?>.*?<\/a:tblPr>/s,
    '<a:tblPr firstRow="1">' + preset + "</a:tblPr>"));
  const snapshot = openDocument(importPptx(zipSync(files), options)).snapshot();
  expect(snapshot.metadata.presentation!.nativePreview).toEqual([expect.objectContaining({ feature: "table-style" })]);
  expect(snapshot.html).toContain("font-weight:bold");
  change(files, main, xml => xml.replace(preset, preset.replace("5C22544A", "ABCDEFAB")));
  expect(() => importPptx(zipSync(files), options)).toThrowError(expect.objectContaining({ code: "UnsupportedFeature" }));
});

test("round trips native picture modulation and a custom nonrectangular mask without changing its image bytes", () => {
  const custom = '<a:custGeom><a:gdLst><a:gd name="center" fmla="*/ 1 w 2"/></a:gdLst><a:cxnLst><a:cxn ang="0"><a:pos x="center" y="b"/></a:cxn></a:cxnLst><a:pathLst><a:path w="400" h="200"><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="400" y="0"/></a:lnTo><a:lnTo><a:pt x="200" y="200"/></a:lnTo><a:close/></a:path></a:pathLst></a:custGeom>';
  const bytes = mutate("07-images.pptx", "ppt/slides/slide1.xml", xml => xml.replace(/<p:pic>[\s\S]*?<\/p:pic>/, picture =>
    picture.replace(/<a:blip ([^>]+)\/>/, '<a:blip $1><a:alphaModFix amt="45000"/></a:blip>').replace(/<a:prstGeom[^>]*>[\s\S]*?<\/a:prstGeom>/, custom)));
  const model = openDocument(importPptx(bytes, { documentId: "mask" }));
  const reopened = openDocument(importPptx(exportPptx(model).bytes, { documentId: "mask" }));
  const shape = (doc: typeof model) => doc.snapshot().metadata.presentation!.objects.find(item => item.kind === "image")!;
  expect(shape(reopened)).toMatchObject({ kind: "image", customGeometry: { connections: [{ x: "center", y: "b", angle: "0" }] } });
  expect(shape(reopened)).toMatchObject({ customGeometry: (shape(model) as Extract<ReturnType<typeof shape>, { kind: "image" }>).customGeometry });
  expect(reopened.getNode(shape(reopened).nodeId)!.attributes["data-office-image-opacity"]).toBe("0.45");
  expect(reopened.snapshot().manifest.assets.map(asset => asset.sha256)).toEqual(model.snapshot().manifest.assets.map(asset => asset.sha256));
});

test("compatible import retains opaque object tags by relationship and removes them only with their owner", () => {
  const files = unzipSync(fixture()), slidePath = "ppt/slides/slide1.xml", relsPath = "ppt/slides/_rels/slide1.xml.rels";
  const tagPath = "ppt/sample-attributes/attributes.xml";
  files[relsPath] = strToU8(strFromU8(files[relsPath]!).replace("</Relationships>", '<Relationship Id="opaque-tags" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tags" Target="../sample-attributes/attributes.xml"/></Relationships>'));
  files["[Content_Types].xml"] = strToU8(strFromU8(files["[Content_Types].xml"]!).replace("</Types>", '<Override PartName="/ppt/sample-attributes/attributes.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tags+xml"/></Types>'));
  files[tagPath] = strToU8('<p:tagLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:tag name="opaque" val="original"/></p:tagLst>');
  files[slidePath] = strToU8(strFromU8(files[slidePath]!).replace(/<p:nvSpPr>.*?<\/p:nvSpPr>/s, xml => xml.replace(/<p:nvPr\s*\/>|<p:nvPr><\/p:nvPr>/, '<p:nvPr><p:custDataLst><p:tags r:id="opaque-tags"/></p:custDataLst></p:nvPr>')));
  const model = openDocument(importPptx(zipSync(files), { documentId: "tagged", mode: "compatible" })), session = createDocumentSession(model);
  const tags = model.snapshot().metadata.presentation!.nativePreview!;
  expect(tags).toEqual([expect.objectContaining({ feature: "tags", part: tagPath, xml: expect.stringContaining('name="opaque" val="original"') })]);
  session.commit(session.prepare({ protocolVersion: 1, documentId: "tagged", baseRevision: 0, operations: [{ op: "removeNode", nodeId: tags[0]!.ownerId }] }));
  expect(session.snapshot().metadata.presentation!.nativePreview).toEqual([]);
  session.undo();
  expect(session.snapshot().metadata.presentation!.nativePreview).toEqual(tags);
});

test("keeps compound theme gradient transforms in source order through native export and reopen", () => {
  const first = '<a:tint val="67000"/><a:lumMod val="110000"/><a:satMod val="105000"/>',
    second = '<a:satMod val="250000"/><a:satMod val="50000"/><a:shade val="78000"/>';
  const files = unzipSync(mutate("01-rich-text.pptx", "ppt/theme/theme1.xml", source => source.replace(
    /(<a:bgFillStyleLst>)<a:solidFill>.*?<\/a:solidFill>/s,
    `$1<a:gradFill><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr">${first}</a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr">${second}</a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>`,
  )));
  const slide = "ppt/slides/slide1.xml";
  files[slide] = strToU8(strFromU8(files[slide]!).replace("<p:spTree>", '<p:bg><p:bgRef idx="1001"><a:srgbClr val="6699CC"/></p:bgRef></p:bg><p:spTree>'));
  const verify = (bytes: Uint8Array) => {
    const model = openDocument(importPptx(bytes, { documentId: "compound-gradient" }));
    const css = model.snapshot().styles[0]!.css;
    expect(css).toContain("--office-fill-stop-0-color-transforms:tint 67000 lumMod 110000 satMod 105000");
    expect(css).toContain("--office-fill-stop-1-color-transforms:satMod 250000 satMod 50000 shade 78000");
    return model;
  };
  const output = exportPptx(verify(zipSync(files))).bytes;
  const theme = strFromU8(unzipSync(output)["ppt/theme/theme1.xml"]!);
  expect(theme).toContain(first);
  expect(theme).toContain(second);
  verify(output);
  const themePath = "ppt/theme/theme1.xml", source = strFromU8(files[themePath]!);
  for (const val of ["-1", "2147483648", "foo"]) {
    const invalid = { ...files, [themePath]: strToU8(source.replace('<a:satMod val="105000"/>', `<a:satMod val="${val}"/>`)) };
    expect(() => importPptx(zipSync(invalid), { documentId: "invalid-transform" })).toThrowError(expect.objectContaining({ code: "InvalidDocument" }));
  }
  const repeated = { ...files, [themePath]: strToU8(source.replace(first, '<a:satMod val="100000"/>'.repeat(65))) };
  expect(() => importPptx(zipSync(repeated), { documentId: "oversized-transform" })).toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded" }));
});

test("preserves absent text-body overrides for placeholder inheritance instead of manufacturing defaults", () => {
  const input = mutate("02-theme-placeholders.pptx", "ppt/slides/slide1.xml", source => source.replace(/<a:bodyPr[^>]*>[\s\S]*?<\/a:bodyPr>/g, "<a:bodyPr/>"));
  const model = openDocument(importPptx(input, { documentId: "body-inheritance" }));
  const object = model.snapshot().metadata.presentation!.objects.find(o => o.kind === "text" && model.getNode(o.nodeId)!.text.includes("继承主题标题"))!;
  expect(object).toMatchObject({ text: {} });
  const node = model.getNode(object.nodeId)!; expect(node.attributes.style ?? "").not.toMatch(/padding-|white-space|text-anchor|tab-size/);
});

test("native classic comment points keep their slide position during visual edits and recover only when their slide is removed",()=>{
  const s=createDocumentSession(load("11-classic-comments.pptx")),before=s.getAnnotations()[0]!,p=s.snapshot().metadata.presentation!;
  expect(before.anchor).toMatchObject({kind:"region",layoutKey:"pptx-classic-point"});const original=p.slides[0]!;
  const commit=(operations:Parameters<typeof s.prepare>[0]["operations"]):void=>{s.commit(s.prepare({protocolVersion:1,documentId:s.snapshot().manifest.documentId,baseRevision:s.snapshot().manifest.revision,operations}));};
  commit([{op:"setStyle",nodeId:original.rootNodeId,property:"background-color",value:"#123456"}]);expect(s.getAnnotations()[0]!.status).toBe("open");expect(s.getAnnotations()[0]!.anchor).toEqual(before.anchor);
  commit([{op:"insertSlide",layoutId:p.layouts[0]!.layoutId}]);commit([{op:"removeSlide",slideId:original.slideId}]);expect(s.getAnnotations()[0]!.status).toBe("needs_reanchor");expect(openDocument(s.serialize()).snapshot().annotations[0]!.anchor).toEqual(before.anchor);
});

test("records DrawingML proportional and exact line-spacing rules while round-tripping native spcPct and spcPts", () => {
  let count = 0;
  const input = mutate("01-rich-text.pptx", "ppt/slides/slide1.xml", source => source.replace(/<a:pPr([^>]*)>/g, (match) => {
    count++;
    if (count === 1) return match + '<a:lnSpc><a:spcPct val="150000"/></a:lnSpc>';
    if (count === 2) return match + '<a:lnSpc><a:spcPts val="3600"/></a:lnSpc>';
    return match;
  }));
  const model = openDocument(importPptx(input, { documentId: "lnspc-rules" }));
  const html = model.snapshot().html;
  expect(html).toContain("line-height:1.5;--office-line-spacing-rule:proportional");
  expect(html).toContain("line-height:36pt;--office-line-spacing-rule:exact");
  const exported = exportPptx(model).bytes;
  const slideXml = strFromU8(unzipSync(exported)["ppt/slides/slide1.xml"]!);
  expect(slideXml).toContain('<a:spcPct val="150000"/>');
  expect(slideXml).toContain('<a:spcPts val="3600"/>');
});

test("preserves DrawingML hangingPunct/eaLnBrk/latinLnBrk on import/export and leaves empty theme fallback unquoted", () => {
  const files = unzipSync(fixture("01-rich-text.pptx"));
  const slideXml = strFromU8(files["ppt/slides/slide1.xml"]!).replace(
    "<a:pPr>",
    '<a:pPr eaLnBrk="1" latinLnBrk="0" hangingPunct="1">',
  );
  const themePath = Object.keys(files).find(name => name.startsWith("ppt/theme/theme") && name.endsWith(".xml"))!;
  const themeXml = strFromU8(files[themePath]!)
    .replace(/<a:ea typeface="[^"]*"\/>/g, '<a:ea typeface=""/>')
    .replace(/<a:font script="Hans"[^>]*\/>/g, "");
  const mutated = zipSync({ ...files, "ppt/slides/slide1.xml": strToU8(slideXml), [themePath]: strToU8(themeXml) });
  const model = openDocument(importPptx(mutated, { documentId: "hanging-and-theme-fallback" }));
  const snapshot = model.snapshot();
  expect(snapshot.html).toContain("--office-ea-line-break:true");
  expect(snapshot.html).toContain("--office-latin-line-break:false");
  expect(snapshot.html).toContain("--office-hanging-punctuation:allow-end");
  const css = snapshot.styles.map(sheet => sheet.css).join("\n");
  expect(css).toContain("--office-minor-east-asia:sans-serif");
  expect(css).not.toContain('--office-minor-east-asia:"sans-serif"');
  const exportedSlide = strFromU8(unzipSync(exportPptx(model).bytes)["ppt/slides/slide1.xml"]!);
  expect(exportedSlide).toMatch(/eaLnBrk="1"/);
  expect(exportedSlide).toMatch(/latinLnBrk="0"/);
  expect(exportedSlide).toMatch(/hangingPunct="1"/);
});
