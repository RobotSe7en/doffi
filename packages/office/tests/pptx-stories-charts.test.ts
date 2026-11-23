import { readFileSync } from "node:fs";
import { expect, test } from "vitest";
import { Ajv2020 } from "ajv/dist/2020.js";
import { DocumentError, getContractJsonSchema, openDocument, parseContract } from "@doffi/core";
import type { Annotation, PresentationDocumentMetadata } from "@doffi/core";
import { importDocx, importPptx, readDocumentZip, writeDocumentZip } from "@doffi/office";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";

const directory = new URL("../../../examples/pptx/", import.meta.url);
const expected = JSON.parse(readFileSync(new URL("expected.json", directory), "utf8")) as {
  cases: { file: string; text: string[]; slides: number; charts?: number; notes?: string[]; classicComments?: number; modernComments?: number; replies?: number }[];
};
const bytes = (name: string) => new Uint8Array(readFileSync(new URL(name, directory)));
const load = (name: string) => openDocument(importPptx(bytes(name), { documentId: "stories" }));
function mutate(name: string, part: string, change: (source: string) => string): Uint8Array {
  const files = unzipSync(bytes(name)); files[part] = strToU8(change(strFromU8(files[part]!))); return zipSync(files);
}
for (const item of expected.cases.filter(c => /^(09|10|11|12|13)-/.test(c.file))) for (const producer of ["", "producer/"]) {
  test("reads charts, notes and comments from independent sources: " + producer + item.file, () => {
    const input = bytes(producer + item.file), before = new Uint8Array(input), files = importPptx(input, { documentId: "stories" });
    const model = openDocument(files), s = model.snapshot(), p = s.metadata.presentation!;
    expect(input).toEqual(before);
    expect(p.slides).toHaveLength(item.slides);
    expect(p.charts).toHaveLength(item.charts ?? 0);
    const text = p.slides.map(slide => model.getNode(slide.rootNodeId)!.text).join("\n");
    for (const required of item.text) expect(text).toContain(required);
    for (const required of item.notes ?? []) expect(p.stories.filter(s => s.kind === "notes").map(story => model.getNode(story.rootNodeId)!.text).join("\n")).toContain(required);
    // The independent reference records LibreOffice dropping modern threads.
    const modern = producer ? 0 : item.modernComments ?? 0;
    expect(s.annotations).toHaveLength((item.classicComments ?? 0) + modern);
    expect(s.annotations.flatMap(a => a.replies ?? [])).toHaveLength(producer ? 0 : item.replies ?? 0);
    expect(openDocument(model.serialize()).snapshot()).toEqual(s);
    expect(openDocument(readDocumentZip(writeDocumentZip(files))).snapshot()).toEqual(s);
    expect(importPptx(input, { documentId: "stories" })).toEqual(files);
    expect([...files.keys()].some(name => /\.(?:xlsx|pptx)$/.test(name))).toBe(false);
  });
}

test("uses one authoritative data table per chart and preserves its native appearance references", () => {
  for (const producer of ["", "producer/"]) {
    const model = load(producer + "09-charts.pptx"), s = model.snapshot(), p = s.metadata.presentation!;
    expect(p.charts.map(c => c.kind)).toEqual(["column", "line", "pie"]);
    for (const chart of p.charts) {
      const data = p.dataSets.find(d => d.dataSetId === chart.dataSetId)!;
      expect(data.categories).toEqual(["第一季度", "第二季度", "第三季度"]);
      expect(data.series).toEqual([expect.objectContaining({ name: "交付", values: [2, 5, 8] })]);
      expect(chart.legend).toBe("bottom");
      expect(chart.showLabels).toBe(false);
      expect(model.getNode(chart.titleNodeId!)!.text).toBe("图表 " + chart.kind);
      expect(chart.seriesStyles).toHaveLength(1);
      expect(s.styles[0]!.css).toContain("." + chart.seriesStyles![0]!.cssClass + "{");
    }
    if (producer) expect(p.charts[2]!.pointStyles).toHaveLength(3);
  }
});

test("keeps native notes geometry, size, master and editable body in the same source", () => {
  for (const producer of ["", "producer/"]) {
    const model = load(producer + "10-notes.pptx"), p = model.snapshot().metadata.presentation!;
    expect(p.notesSize).toEqual({ unit: "pt", width: 540, height: 720 });
    expect(p.notesMasters).toHaveLength(1);
    const story = p.stories.find(s => s.kind === "notes")!;
    if (story.kind !== "notes") throw new Error("Expected notes");
    expect(story.notesMasterId).toBe(p.notesMasters[0]!.notesMasterId);
    expect(model.getNode(story.bodyNodeId!)!.text).toBe("演讲者备注：介绍上线、风险与培训。");
    expect(model.getNode(story.bodyNodeId!)!.html).toContain("font-weight:bold");
    expect(model.getNode(story.bodyNodeId!)!.attributes["data-office-width"]).toBeDefined();
  }
});

test("maps classic comment points from eighths of a point and preserves display authors and dates", () => {
  const model = load("11-classic-comments.pptx"), s = model.snapshot(), a = s.annotations[0]!;
  expect(a.anchor).toMatchObject({ kind: "region", rect: { x: 72, y: 72, width: 1, height: 1 }, nodeIds: [] });
  expect(a.author).toEqual({ name: "公开合成作者", initials: "SY" });
  expect(a.date).toBe("2026-10-31T08:00:00Z");
  expect(model.getNode(a.bodyNodeId!)!.text).toBe("经典批注：补充交付细节。");
  expect(s.metadata.presentation!.commentFormat).toBe("classic");
  expect(s.metadata.presentation!.commentAuthors).toEqual([expect.objectContaining({ format: "classic", nativeId: "0" })]);
});

test("preserves modern workflow, replies, opaque author IDs and exact UTF-16 text anchors", () => {
  const model = load("12-modern-comments.pptx"), s = model.snapshot(), [page, object, text] = s.annotations;
  expect(page!.anchor).toMatchObject({ kind: "node", nodeId: s.metadata.presentation!.slides[0]!.rootNodeId });
  expect(object!.status).toBe("resolved");
  expect(text!.anchor).toMatchObject({ kind: "text", from: 2, to: 6, quote: "𠀀😀" });
  expect(text!.replies![0]).toMatchObject({ author: { name: "公开作者", initials: "ZZ" }, status: "open" });
  expect(model.getNode(text!.replies![0]!.bodyNodeId!)!.text).toBe("回复：已记录。");
  expect(s.metadata.presentation!.commentAuthors).toHaveLength(2);
  expect(s.metadata.presentation!.commentAuthors![0]).toMatchObject({ format: "modern", userId: "synthetic-reviewer", providerId: "synthetic" });
});

test("maps multi-paragraph native ranges across implicit paragraph marks and explicit soft breaks", () => {
  const files = unzipSync(bytes("12-modern-comments.pptx"));
  files["ppt/slides/slide3.xml"] = strToU8(strFromU8(files["ppt/slides/slide3.xml"]!).replace(/<p:txBody>.*?<\/p:txBody>/s,
    '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>甲𠀀</a:t></a:r><a:br/><a:r><a:t>乙</a:t></a:r></a:p><a:p><a:r><a:t>丙😀丁</a:t></a:r></a:p></p:txBody>'));
  files["ppt/comments/modernComment3.xml"] = strToU8(strFromU8(files["ppt/comments/modernComment3.xml"]!).replace('cp="2" len="4"', 'cp="1" len="8"'));
  const model = openDocument(importPptx(zipSync(files), { documentId: "ranges" })), anchor = model.snapshot().annotations[2]!.anchor;
  expect(anchor.kind).toBe("text-set");
  if (anchor.kind !== "text-set") throw new Error("Expected text set");
  expect(anchor.ranges.map(r => r.quote)).toEqual(["𠀀\n乙", "丙😀"]);
  for (const range of anchor.ranges) expect(model.getNode(range.nodeId)!.text.slice(range.from, range.to)).toBe(range.quote);
});

test("verifies the optional OfficeArt UTF-16 text context hash without guessing another range", () => {
  const source = "中文𠀀😀批注锚点";
  const hash = (input: string) => { let value = 0n; for (let index = 0; index < input.length; index++) value = (value * 33n + BigInt(input.charCodeAt(index))) % 4294967296n; return String(value); };
  for (const text of [source, source + "\r"]) {
    const input = mutate("12-modern-comments.pptx", "ppt/comments/modernComment3.xml", xml => xml.replace('<ac:txMk cp="2" len="4"/>', '<ac:txMk cp="2" len="4"><ac:context len="' + text.length + '" hash="' + hash(text) + '"/></ac:txMk>'));
    expect(openDocument(importPptx(input, { documentId: "hash" })).snapshot().annotations[2]!.anchor).toMatchObject({ kind: "text", quote: "𠀀😀" });
  }
  const stale = mutate("12-modern-comments.pptx", "ppt/comments/modernComment3.xml", xml => xml.replace('<ac:txMk cp="2" len="4"/>', '<ac:txMk cp="2" len="4"><ac:context hash="0"/></ac:txMk>'));
  expect(() => importPptx(stale, { documentId: "hash" })).toThrowError(expect.objectContaining({ code: "InvalidDocument" }));
});

test("preserves closed state and explicit unknown-anchor recovery without conflating them", () => {
  const input = mutate("12-modern-comments.pptx", "ppt/comments/modernComment2.xml", xml => xml.replace('status="resolved"', 'status="closed"'));
  expect(openDocument(importPptx(input, { documentId: "closed" })).snapshot().annotations[1]!.status).toBe("closed");
  const unknown = mutate("12-modern-comments.pptx", "ppt/comments/modernComment2.xml", xml => xml.replace(/<ac:deMkLst>.*?<\/ac:deMkLst>/s, '<p188:unknownAnchor/>'));
  const s = openDocument(importPptx(unknown, { documentId: "unknown" })).snapshot();
  expect(s.annotations[1]!.status).toBe("needs_reanchor");
  expect(s.metadata.presentation!.comments[1]!.recoveryStatus).toBe("resolved");
});

test("rejects cache/workbook mismatch, advanced charts, unsafe formulas and inconsistent notes ownership", () => {
  const invalid: [string, string, (s: string) => string, string][] = [
    ["09-charts.pptx", "ppt/charts/chart1.xml", s => s.replace('<c:v>2</c:v>', '<c:v>200</c:v>'), "InvalidDocument"],
    ["09-charts.pptx", "ppt/charts/chart1.xml", s => s.replace('val="clustered"', 'val="stacked"'), "UnsupportedFeature"],
    ["09-charts.pptx", "ppt/charts/chart2.xml", s => s.replace('<c:smooth val="0"', '<c:smooth val="1"'), "UnsupportedFeature"],
    ["producer/09-charts.pptx", "ppt/charts/chart1.xml", s => s.replace('categories</c:f>', 'SUM(A1:A3)</c:f>'), "UnsupportedFeature"],
    ["10-notes.pptx", "ppt/notesSlides/_rels/notesSlide1.xml.rels", s => s.replace('../slides/slide1.xml', '../notesMasters/notesMaster1.xml'), "InvalidDocument"],
    ["12-modern-comments.pptx", "ppt/comments/modernComment3.xml", s => s.replace('cp="2" len="4"', 'cp="3" len="4"'), "InvalidDocument"],
    ["12-modern-comments.pptx", "ppt/comments/modernComment1.xml", s => s.replace('status="active"', 'status="invalid"'), "InvalidDocument"],
    ["12-modern-comments.pptx", "ppt/slides/slide1.xml", s => s.replace('r:id="comments-modern"', 'r:id="missing"'), "InvalidDocument"],
  ];
  for (const [name, part, change, code] of invalid) {
    try { importPptx(mutate(name, part, change), { documentId: "bad" }); throw new Error("Accepted invalid source: " + name); }
    catch (error) { expect(error).toBeInstanceOf(DocumentError); expect((error as DocumentError).code).toBe(code); }
  }
});

test("bounds nested XLSX expansion, cache lengths and reply work", () => {
  const input = bytes("09-charts.pptx"), native = unzipSync(input), baseBytes = Object.values(native).reduce((n, data) => n + data.length, 0);
  expect(() => importPptx(input, { documentId: "nested", limits: { maxPackageBytes: baseBytes + 20 } })).toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded" }));
  const huge = mutate("09-charts.pptx", "ppt/charts/chart1.xml", source => source.replace('ptCount val="3"', 'ptCount val="999999999"'));
  expect(() => importPptx(huge, { documentId: "huge" })).toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded" }));
});

test("preserves visible label formatting and axis options in the source styles", () => {
  const input = mutate("09-charts.pptx", "ppt/charts/chart1.xml", source => source.replace('</c:ser>',
    '<c:dLbls><c:numFmt formatCode="0.00" sourceLinked="0"/><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1200"/></a:pPr></a:p></c:txPr><c:dLblPos val="outEnd"/><c:showVal val="1"/><c:separator> / </c:separator></c:dLbls></c:ser>')
    .replace('<c:catAx>', '<c:catAx><c:majorTickMark val="out"/><c:tickLblPos val="low"/>'));
  const s = openDocument(importPptx(input, { documentId: "labels" })).snapshot(), chart = s.metadata.presentation!.charts[0]!;
  expect(chart.labelContents).toEqual(["value"]);
  expect(s.styles[0]!.css).toContain('--office-label-number-format:"0.00"');
  expect(s.styles[0]!.css).toContain("--office-axis-major-tick:out");
  expect(s.styles[0]!.css).toContain("--office-axis-label-position:low");
});

test("retains a hidden chart title's content instead of discarding it", () => {
  const input = mutate("09-charts.pptx", "ppt/charts/chart1.xml", source => source.replace('<c:plotArea>', '<c:autoTitleDeleted val="1"/><c:plotArea>'));
  const model = openDocument(importPptx(input, { documentId: "hidden-title" })), chart = model.snapshot().metadata.presentation!.charts[0]!;
  expect(model.getNode(chart.titleNodeId!)?.text).toBe("图表 column");
  expect(chart.titleVisible).toBe(false);
});

test("preserves an explicitly encoded carriage return through HTML package reopening", () => {
  const input = mutate("01-rich-text.pptx", "ppt/slides/slide1.xml", source => source.replace('中文与 Latin 😀，', '甲&#13;乙'));
  const model = openDocument(importPptx(input, { documentId: "cr" })), p = model.snapshot().metadata.presentation!;
  expect(model.getNode(p.slides[0]!.rootNodeId)!.text).toContain("甲\r乙");
  expect(openDocument(model.serialize()).getNode(p.slides[0]!.rootNodeId)!.text).toContain("甲\r乙");
});

test("rejects undeclared modern comment extensions instead of dropping them", () => {
  const input = mutate("12-modern-comments.pptx", "ppt/comments/modernComment1.xml", source => source.replace('</p188:cm>', '<p188:extLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:ext uri="unsupported-task"><p188:task/></p:ext></p188:extLst></p188:cm>'));
  expect(() => importPptx(input, { documentId: "extension" })).toThrowError(expect.objectContaining({ code: "UnsupportedFeature" }));
});

test("resolves an omitted slide placeholder type through its unique native layout index", () => {
  const input = mutate("02-theme-placeholders.pptx", "ppt/slides/slide1.xml", source => source.replace('<p:ph type="title" idx="0"/>', '<p:ph idx="0"/>'));
  const model = openDocument(importPptx(input, { documentId: "placeholder" })), p = model.snapshot().metadata.presentation!;
  const objectId = model.getNode(p.slides[0]!.rootNodeId)!.childNodeIds[0]!;
  expect(p.objects.find(o => o.nodeId === objectId)!.placeholder).toEqual({ type: "title", index: 0 });
  expect(model.getNode(objectId)!.attributes["data-office-x"]).toBeUndefined();
});

test("keeps JSON Schema contracts and cross-source references valid for the new stories and chart fields", () => {
  const files = new Map(importPptx(bytes("13-multi-slide.pptx"), { documentId: "contracts" }));
  const metadata = JSON.parse(strFromU8(files.get("metadata.json")!)) as PresentationDocumentMetadata;
  const annotations = JSON.parse(strFromU8(files.get("annotations.json")!)) as Annotation[];
  const validator = new Ajv2020({ strict: false });
  expect(validator.compile(getContractJsonSchema("document-metadata"))(metadata)).toBe(true);
  expect(validator.compile(getContractJsonSchema("annotations"))(annotations)).toBe(true);
  expect(parseContract("document-metadata", metadata)).toEqual(metadata);
  const invalid = [
    (m: PresentationDocumentMetadata) => { m.presentation.charts[0]!.seriesStyles![0]!.cssClass = "missing"; },
    (m: PresentationDocumentMetadata) => { m.presentation.comments[0]!.authorId = "missing"; },
    (m: PresentationDocumentMetadata) => { const note = m.presentation.stories.find(s => s.kind === "notes")!; if (note.kind === "notes") note.bodyNodeId = m.presentation.slides[0]!.rootNodeId; },
    (m: PresentationDocumentMetadata) => { m.presentation.comments[0]!.recoveryStatus = "resolved"; },
  ];
  for (const change of invalid) {
    const changed = structuredClone(metadata); change(changed);
    const candidate = new Map(files); candidate.set("metadata.json", strToU8(JSON.stringify(changed)));
    expect(() => openDocument(candidate)).toThrowError(expect.objectContaining({ code: "InvalidDocument" }));
  }
});

test("preserves the closed annotation state in current Word source packages", () => {
  const source = readFileSync(new URL("../docx/08-comments-revisions.docx", directory));
  const files = new Map(importDocx(source, { documentId: "word-closed" }));
  const annotations = JSON.parse(strFromU8(files.get("annotations.json")!)) as Annotation[];
  annotations[0]!.status = "closed"; files.set("annotations.json", strToU8(JSON.stringify(annotations)));
  const reopened = openDocument(files);
  expect(openDocument(reopened.serialize()).snapshot().annotations[0]!.status).toBe("closed");
});
