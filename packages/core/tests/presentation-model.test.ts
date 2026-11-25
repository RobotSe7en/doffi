import { openContentFixture as openDocument } from "../../../tests/content-fixture.js";
import { Ajv2020 } from "ajv/dist/2020.js";
import { readFileSync, readdirSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, test } from "vitest";
import { createDocumentSession, createPresentation, DocumentError, getContractJsonSchema, parseContract, readPresentationGeometry } from "@doffi/core";
import type { Annotation, DocumentErrorCode, PresentationDocumentMetadata } from "@doffi/core";
import { repository } from "./fixtures.js";

const encoder = new TextEncoder(), decoder = new TextDecoder();
const initial = () => new Map(createPresentation({ documentId: "deck" }));
const json = (files: Map<string, Uint8Array>, path: string, data: unknown): void => { files.set(path, encoder.encode(JSON.stringify(data))); };
const source = (files: Map<string, Uint8Array>): string => decoder.decode(files.get("content.html"));
const metadata = (files: Map<string, Uint8Array>): PresentationDocumentMetadata => JSON.parse(decoder.decode(files.get("metadata.json"))) as PresentationDocumentMetadata;
function editMetadata(files: Map<string, Uint8Array>, edit: (data: PresentationDocumentMetadata) => void): void { const data = metadata(files); edit(data); json(files, "metadata.json", data); }
function expectFailure(action: () => unknown, code: DocumentErrorCode = "InvalidDocument"): void {
  try { action(); } catch (error) { expect(error).toBeInstanceOf(DocumentError); expect((error as DocumentError).code).toBe(code); return; }
  throw new Error("Expected failure.");
}
function addText(files: Map<string, Uint8Array>, id: string, text = "中文𠀀与emoji😀"): void {
  const html = '<div data-node-id="' + id + '" data-office-kind="text" data-office-x="-10" data-office-y="20" data-office-width="120" data-office-height="60"><p data-node-id="' + id + '-p">' + text + '</p></div>';
  files.set("content.html", encoder.encode(source(files).replace(/(<section[^>]*data-node-id="slide-root-1"[^>]*>)/, "$1" + html)));
  editMetadata(files, (data) => { data.presentation.objects.push({ nodeId: id, kind: "text" }); });
}

describe("Presentation package model", () => {
  test("opens the complete public example with authorized resources and rich replies", () => {
    const directory = join(repository, "examples/presentation");
    const paths = [...readdirSync(directory).filter((name) => name.endsWith(".json") || name.endsWith(".html") || name.endsWith(".css")), "assets/sample.png"];
    const model = openDocument(new Map(paths.map((path) => [path, new Uint8Array(readFileSync(join(directory, path)))])));
    expect(model.snapshot().metadata.presentation!.objects.map((object) => object.kind)).toEqual(expect.arrayContaining(["text", "geometry", "group", "connector", "table", "chart", "image"]));
    expect(model.getNode("speaker-notes-p")!.text).toContain("演讲者提示");
    expect(model.snapshot().annotations[0]!.replies![0]!.bodyNodeId).toBe("reply-body");
    expect(openDocument(model.serialize()).snapshot()).toEqual(model.snapshot());
  });
  test("creates a neutral deck with native source definitions and no synthetic Word metadata", () => {
    const files = initial(), model = openDocument(files), snapshot = model.snapshot();
    expect(snapshot.manifest).toMatchObject({ kind: "presentation", documentId: "deck", revision: 0 });
    expect(snapshot.metadata).not.toHaveProperty("pageSetup");
    expect(snapshot.metadata).not.toHaveProperty("sections");
    expect(snapshot.metadata.presentation).toMatchObject({ size: { unit: "pt", width: 960, height: 540 }, commentFormat: "modern" });
    expect(snapshot.metadata.presentation!.slides).toHaveLength(1);
    expect(snapshot.metadata.presentation!.layouts).toHaveLength(3);
    expect(snapshot.metadata.presentation!.objects).toHaveLength(4);
    expect(model.getNode("slide-root-1")!.childNodeIds).toEqual([]);
    expect(openDocument(model.serialize()).snapshot()).toEqual(snapshot);
    expect(openDocument(model.serialize()).serialize()).toEqual(model.serialize());
  });

  test("supports explicit title, locale and page size without HTML injection", () => {
    const title = '中文<&"标题😀';
    const model = openDocument(createPresentation({ documentId: "custom", title, locale: "zh-Hans", size: { unit: "pt", width: 720, height: 540 } }));
    expect(model.snapshot().metadata).toMatchObject({ title, locale: "zh-Hans", presentation: { size: { width: 720, height: 540 } } });
    expect(model.snapshot().html).toContain("&lt;&amp;");
    expect(model.getNode("layout-body")!.attributes["data-office-width"]).toBe("648");
  });

  test("shares identical PPT inline declarations while budgeting every use and validating new styles", () => {
    const files = initial();
    files.set("styles.css", encoder.encode(".theme-default{}"));
    addText(files, "repeated", Array.from({ length: 20 }, (_, index) =>
      `<span data-node-id="run-${index}" style="color: red">字</span>`).join(""));
    // Six stylesheet nodes, four for the first declaration list, nineteen reused lists.
    const limits = { maxCssNodes: 29 };
    const model = openDocument(files, { limits });
    expect(model.getNode("repeated")!.text).toBe("字".repeat(20));
    expect(model.getNode("run-19")!.attributes.style).toBe("color:red");
    expect(openDocument(model.serialize(), { limits }).snapshot()).toEqual(model.snapshot());
    expect(() => openDocument(files, { limits: { maxCssNodes: 28 } })).toThrowError(
      expect.objectContaining({ code: "ResourceLimitExceeded", message: "Document input exceeds maxCssNodes (29 > 28) in inline CSS." }));
    files.set("content.html", encoder.encode(source(files).replace('style="color: red"', 'style="color: red;background-color:url(https://example.com/image)"')));
    expectFailure(() => openDocument(files), "InvalidDocument");
  });

  test("isolates bytes, source snapshots and different document identities", () => {
    const files = initial(), first = openDocument(files), second = openDocument(createPresentation({ documentId: "other" }));
    const snapshot = first.snapshot();
    files.get("metadata.json")!.fill(0); files.clear();
    expect(Reflect.set(snapshot.metadata.presentation!.size, "width", 1)).toBe(false);
    const serialized = first.serialize(); serialized.get("content.html")!.fill(0);
    expect(openDocument(first.serialize()).snapshot()).toEqual(snapshot);
    expect(second.snapshot().manifest.documentId).toBe("other");
  });

  test("validates both structural branches with independent JSON Schema", () => {
    const data = metadata(initial()), validate = new Ajv2020({ strict: false }).compile(getContractJsonSchema("document-metadata"));
    expect(validate(data)).toBe(true);
    expect(parseContract("document-metadata", data)).toEqual(data);
    for (const changed of [{ ...data, sections: [] }, { ...data, unknown: true }, { ...data, presentation: { ...data.presentation, unknown: true } }]) {
      expect(validate(changed)).toBe(false); expectFailure(() => parseContract("document-metadata", changed));
    }
  });

  test.each([
    (data: PresentationDocumentMetadata) => { data.presentation.slides[0]!.layoutId = "missing"; },
    (data: PresentationDocumentMetadata) => { data.presentation.layouts[0]!.masterId = "missing"; },
    (data: PresentationDocumentMetadata) => { data.presentation.masters[0]!.themeId = "missing"; },
    (data: PresentationDocumentMetadata) => { data.presentation.themes[0]!.cssClass = "missing"; },
    (data: PresentationDocumentMetadata) => { data.presentation.slides.push({ ...data.presentation.slides[0]! }); },
    (data: PresentationDocumentMetadata) => { data.presentation.layouts[0]!.rootNodeId = data.presentation.slides[0]!.rootNodeId; },
    (data: PresentationDocumentMetadata) => { data.presentation.objects[0]!.nodeId = "missing"; },
  ])("rejects dangling, duplicate and wrong-role definitions", (edit) => {
    const files = initial(); editMetadata(files, edit); expectFailure(() => openDocument(files));
  });

  test("rejects mixed manifest/metadata kinds, including workbook kind without chunks", () => {
    for (const kind of ["document", "workbook"]) {
      const files = initial(), manifest = JSON.parse(decoder.decode(files.get("manifest.json"))) as { kind: string };
      manifest.kind = kind; json(files, "manifest.json", manifest);
      expectFailure(() => openDocument(files), "InvalidDocument");
    }
  });

  test("keeps negative local coordinates, source order and UTF-16 text", () => {
    const files = initial(); addText(files, "text-a"); addText(files, "text-b", "第二对象");
    editMetadata(files, (data) => { data.presentation.objects.reverse(); });
    const model = openDocument(files);
    expect(model.getNode("slide-root-1")!.childNodeIds).toEqual(["text-b", "text-a"]);
    expect(model.getNode("text-a-p")!.text).toBe("中文𠀀与emoji😀");
    expect(readPresentationGeometry(model.getNode("text-a")!.attributes)).toEqual({ x: -10, y: 20, width: 120, height: 60, rotation: 0, flipH: false, flipV: false });
  });

  test.each(["NaN", "Infinity", "1px", "", "-1"])("rejects invalid geometry width %j", (width) => {
    const files = initial(); files.set("content.html", encoder.encode(source(files).replace('data-office-width="816"', 'data-office-width="' + width + '"')));
    expectFailure(() => openDocument(files));
  });

  test("rejects missing transforms and unregistered source objects", () => {
    const files = initial(); addText(files, "text-a");
    files.set("content.html", encoder.encode(source(files).replace(' data-office-width="120"', "")));
    expectFailure(() => openDocument(files));
    const unknown = initial(); addText(unknown, "text-a");
    editMetadata(unknown, (data) => { data.presentation.objects = data.presentation.objects.filter((item) => item.nodeId !== "text-a"); });
    expectFailure(() => openDocument(unknown));
  });

  test("resolves an inherited placeholder by its layout identity and rejects unresolved ones", () => {
    const files = initial(); addText(files, "text-a");
    files.set("content.html", encoder.encode(source(files).replace(' data-office-x="-10" data-office-y="20" data-office-width="120" data-office-height="60"', "")));
    editMetadata(files, (data) => {
      data.presentation.slides[0]!.layoutId = "layout-body";
      const object = data.presentation.objects.find((item) => item.nodeId === "text-a")!;
      object.placeholder = { type: "title", index: 0 };
    });
    expect(openDocument(files).getNode("text-a")!.text).toBe("中文𠀀与emoji😀");
    editMetadata(files, (data) => { data.presentation.objects.find((item) => item.nodeId === "text-a")!.placeholder!.index = 9; });
    expectFailure(() => openDocument(files));
  });

  test("does not interpret inherited slide geometry as a group's local coordinates", () => {
    const files = initial(); addText(files, "text-a");
    const child = openDocument(files).getNode("text-a")!.html;
    const inherited = child.replace(' data-office-x="-10" data-office-y="20" data-office-width="120" data-office-height="60"', "");
    files.set("content.html", encoder.encode(source(files).replace(child, '<div data-node-id="group" data-office-kind="group" data-office-x="100" data-office-y="100" data-office-width="400" data-office-height="200" data-office-child-x="0" data-office-child-y="0" data-office-child-width="400" data-office-child-height="200">' + inherited + '</div>')));
    editMetadata(files, (data) => {
      data.presentation.slides[0]!.layoutId = "layout-body";
      data.presentation.objects.push({ nodeId: "group", kind: "group" });
      data.presentation.objects.find((item) => item.nodeId === "text-a")!.placeholder = { type: "title", index: 0 };
    });
    expectFailure(() => openDocument(files));
  });

  test("validates native table grids and row spans before allocating expanded cells", () => {
    const files = initial(); addText(files, "table-object", "");
    const table = '<table data-node-id="table"><tr data-node-id="row-1"><td data-node-id="cell-a" rowspan="2"><p data-node-id="p-a">跨行</p></td><td data-node-id="cell-b"><p data-node-id="p-b">乙</p></td></tr><tr data-node-id="row-2"><td data-node-id="cell-c"><p data-node-id="p-c">丙</p></td></tr></table>';
    files.set("content.html", encoder.encode(source(files).replace('data-node-id="table-object" data-office-kind="text"', 'data-node-id="table-object" data-office-kind="table"').replace('<p data-node-id="table-object-p"></p>', table)));
    editMetadata(files, (data) => {
      data.presentation.objects = data.presentation.objects.filter((item) => item.nodeId !== "table-object");
      data.presentation.objects.push({ nodeId: "table-object", kind: "table", grid: [60, 60] });
    });
    expect(openDocument(files).getNode("cell-a")!.text).toBe("跨行");
    const grouped = new Map(files);
    grouped.set("content.html", encoder.encode(source(grouped).replace('<table data-node-id="table">', '<table data-node-id="table"><tbody data-node-id="table-body">').replace('</table>', '</tbody></table>')));
    expect(openDocument(grouped).getNode("cell-a")!.text).toBe("跨行");
    const wide = new Map(files);
    editMetadata(wide, (data) => { const tableObject = data.presentation.objects.find((item) => item.kind === "table")!; if (tableObject.kind === "table") tableObject.grid = new Array<number>(1000).fill(1); });
    expectFailure(() => openDocument(wide, { limits: { maxNodes: 200 } }), "ResourceLimitExceeded");
    files.set("content.html", encoder.encode(source(files).replace('rowspan="2"', 'rowspan="9007199254740991"')));
    expectFailure(() => openDocument(files));
  });

  test("rejects direct unregistered text and forged nested story boundaries", () => {
    const files = initial(); addText(files, "text-a");
    files.set("content.html", encoder.encode(source(files).replace('<section data-node-id="slide-root-1"', 'unregistered<section data-node-id="slide-root-1"')));
    expectFailure(() => openDocument(files));
    const forged = initial(); addText(forged, "text-a");
    forged.set("content.html", encoder.encode(source(forged).replace('data-node-id="text-a-p"', 'data-node-id="text-a-p" data-office-story="comment"')));
    expectFailure(() => openDocument(forged));
  });

  test("supports registered rich comment bodies and bounded reply identities", () => {
    const files = initial(); addText(files, "text-a");
    files.set("content.html", encoder.encode(source(files).replace("</article>", '<div data-node-id="comment-root" data-office-kind="comment" data-office-story="comment"><p data-node-id="comment-p">富文本批注</p></div></article>')));
    editMetadata(files, (data) => { data.presentation.stories.push({ storyId: "comment", kind: "comment", rootNodeId: "comment-root", slideId: "slide-1" });
      data.presentation.comments.push({ annotationId: "a", format: "modern" }); });
    const annotations: Annotation[] = [{ annotationId: "a", documentId: "deck", baseRevision: 0, status: "resolved", anchor: { kind: "node", nodeId: "text-a" }, bodyNodeId: "comment-root", replies: [{ replyId: "reply", instruction: "回复", author: { name: "作者" } }] }];
    json(files, "annotations.json", annotations);
    expect(openDocument(files).snapshot().annotations).toEqual(annotations);
    annotations[0]!.replies![0]!.replyId = "a"; json(files, "annotations.json", annotations);
    expectFailure(() => openDocument(files));
  });

  test("validates chart data and object references without creating a second source", () => {
    const files = initial(); addText(files, "chart-object", "");
    files.set("content.html", encoder.encode(source(files).replace('data-node-id="chart-object" data-office-kind="text"', 'data-node-id="chart-object" data-office-kind="chart"')));
    editMetadata(files, (data) => {
      data.presentation.objects = data.presentation.objects.filter((item) => item.nodeId !== "chart-object");
      data.presentation.objects.push({ nodeId: "chart-object", kind: "chart", chartId: "chart" });
      data.presentation.charts.push({ chartId: "chart", nodeId: "chart-object", dataSetId: "data", kind: "column", legend: "right", showLabels: false });
      data.presentation.dataSets.push({ dataSetId: "data", categories: ["甲", "乙"], series: [{ seriesId: "s", name: "销售额", values: [2, null] }] });
    });
    expect(openDocument(files).snapshot().metadata.presentation!.charts).toHaveLength(1);
    const extra = new Map(files);
    extra.set("content.html", encoder.encode(source(extra).replace('<p data-node-id="chart-object-p"></p>', '<p data-node-id="chart-object-p">不能丢失的内容</p>')));
    expectFailure(() => openDocument(extra));
    editMetadata(files, (data) => { data.presentation.dataSets[0]!.series[0]!.values.pop(); });
    expectFailure(() => openDocument(files));
  });

  test("checks security and budgets before opening the document", () => {
    const files = initial(); addText(files, "text-a");
    files.set("content.html", encoder.encode(source(files).replace('data-node-id="text-a"', 'onclick="alert(1)" data-node-id="text-a"')));
    expectFailure(() => openDocument(files));
    expectFailure(() => createPresentation({ documentId: "d" }, { limits: { maxNodes: 5 } }), "ResourceLimitExceeded");
    expectFailure(() => createPresentation({ documentId: "d", size: { unit: "pt", width: -1, height: 1 } }));
    expectFailure(() => createPresentation({ documentId: "" }));
    expectFailure(() => createPresentation({ documentId: "d", title: "\ud800" }));
    expectFailure(() => createPresentation({ documentId: "d", title: "\u0000" }));
  });

  test("opens presentation sessions from revalidated serialized source after reading acceptance", () => {
    expect(createDocumentSession(openDocument(initial())).snapshot().manifest.kind).toBe("presentation");
    const actual = openDocument(initial());
    const fake = { ...actual, snapshot: () => ({ ...actual.snapshot(), metadata: { pageSetup: { unit: "pt" as const, width: 10, height: 10, margins: { top: 0, right: 0, bottom: 0, left: 0 } }, sections: [{ sectionId: "fake", rootNodeId: "fake" }] } }) };
    expect(createDocumentSession(fake).snapshot().manifest.kind).toBe("presentation");
  });
});
