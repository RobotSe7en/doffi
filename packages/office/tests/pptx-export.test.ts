import { openContentFixture as openDocument } from "../../../tests/content-fixture.js";
import { readFileSync, readdirSync } from "node:fs";
import { expect, test } from "vitest";
import { createDocumentSession, createPresentation, DocumentError, resolvePresentationGeometry } from "@doffi/core";
import type { ContentDocumentModel as DocumentModel } from "@doffi/core";
import { exportDocx, exportPptx, importPptx } from "@doffi/office";
import { strFromU8, strToU8, unzipSync } from "fflate";
const directory = new URL("../../../examples/pptx/", import.meta.url);
const model = (name: string) => openDocument(importPptx(new Uint8Array(readFileSync(new URL(name, directory))), { documentId: "native-pptx" }));
const reopen = (source: DocumentModel) => openDocument(importPptx(exportPptx(source).bytes, { documentId: "reopened" }));
const texts = (source: DocumentModel) => source.snapshot().metadata.presentation!.slides.map(slide => source.getNode(slide.rootNodeId)!.text);
const data = (source: DocumentModel) => source.snapshot().metadata.presentation!.dataSets.map(data => ({ categories: data.categories, series: data.series.map(series => ({ name: series.name, values: series.values })) }));
for (const producer of ["", "producer/"])
    for (const name of readdirSync(new URL(producer || ".", directory)).filter(name => name.endsWith(".pptx")))
        test("native current-source PPTX export/reopen: " + producer + name, () => {
            const source = model(producer + name), before = source.serialize(), snapshot = source.snapshot(), output = exportPptx(source), native = unzipSync(output.bytes), target = openDocument(importPptx(output.bytes, { documentId: "reopened" })), after = target.snapshot();
            expect(source.serialize()).toEqual(before);
            expect(Object.keys(native)).not.toContain("manifest.json");
            expect(texts(target)).toEqual(texts(source));
            expect(data(target)).toEqual(data(source));
            expect(after.metadata.presentation!.objects.map(object => ({ kind: object.kind, name: object.name, placeholder: object.placeholder }))).toEqual(snapshot.metadata.presentation!.objects.map(object => ({ kind: object.kind, name: object.name ?? "Object " + nativeId(source, object.nodeId), placeholder: object.placeholder })));
            expect(after.annotations.map(annotation => ({ status: annotation.status, text: annotation.bodyNodeId ? target.getNode(annotation.bodyNodeId)!.text : annotation.instruction, replies: annotation.replies?.map(reply => ({ status: reply.status, text: reply.bodyNodeId ? target.getNode(reply.bodyNodeId)!.text : reply.instruction })) }))).toEqual(snapshot.annotations.map(annotation => ({ status: annotation.status, text: annotation.bodyNodeId ? source.getNode(annotation.bodyNodeId)!.text : annotation.instruction, replies: annotation.replies?.map(reply => ({ status: reply.status, text: reply.bodyNodeId ? source.getNode(reply.bodyNodeId)!.text : reply.instruction })) })));
            expect(output.report.unexportedAnnotations).toEqual([]);
            expect(output.report.annotationLosses).toEqual([]);
            const beforeObjects = snapshot.metadata.presentation!.objects, afterObjects = after.metadata.presentation!.objects;
            for (const [index, object] of beforeObjects.entries()) {
                const a = resolvePresentationGeometry(source, object.nodeId), b = resolvePresentationGeometry(target, afterObjects[index]!.nodeId);
                for (const key of ["x", "y", "width", "height", "rotation"] as const)
                    expect(b[key]).toBeCloseTo(a[key], 3);
            }
        });
function nativeId(source: DocumentModel, id: string): number { const p = source.snapshot().metadata.presentation!, object = p.objects.find(object => object.nodeId === id)!; return Number(object.nativeId ?? 2); }
test.each([undefined, false, true].flatMap(scaled => [false, true].map(rotates => ({ scaled, rotates }))))("current-source gradient scaling $scaled and rotation $rotates survive native export and reopen", ({ scaled, rotates }) => {
    const session = createDocumentSession(openDocument(createPresentation({ documentId: "gradient-export" })));
    const style = '--office-fill-kind:linear;--office-fill-angle:45;--office-fill-stop-count:2;--office-fill-stop-0-position:0;--office-fill-stop-0-color:#ff0000;--office-fill-stop-1-position:1;--office-fill-stop-1-color:#0000ff;--office-fill-rotate-with-shape:' + rotates + (scaled === undefined ? '' : ';--office-fill-scaled:' + scaled);
    try {
        session.commit(session.prepare({ protocolVersion: 1, documentId: "gradient-export", baseRevision: 0, operations: [{ op: "insertPresentationObject", parentNodeId: "slide-root-1", geometry: { x: 20, y: 20, width: 200, height: 100, rotation: 45, flipH: true, flipV: false }, object: { kind: "geometry", preset: "rect", style } }] }));
        const output = exportPptx(session), native = unzipSync(output.bytes);
        expect(strFromU8(native['ppt/slides/slide1.xml']!)).toContain(`<a:lin ang="2700000" scaled="${scaled === true ? 1 : 0}"`);
        expect(strFromU8(native['ppt/slides/slide1.xml']!)).toContain(`<a:gradFill rotWithShape="${rotates ? 1 : 0}"`);
        const target = openDocument(importPptx(output.bytes, { documentId: "gradient-reopen" }));
        const object = target.snapshot().metadata.presentation!.objects.find(object => object.kind === 'geometry')!, appearance = target.getNode(object.nodeId)!.attributes.style;
        expect(appearance).toContain('--office-fill-scaled:' + (scaled === true));
        expect(appearance).toContain('--office-fill-angle:45');
        expect(appearance).toContain('--office-fill-rotate-with-shape:' + rotates);
        expect(target.getNode(object.nodeId)!.attributes).toMatchObject({ 'data-office-rotation': '45', 'data-office-flip-h': 'true' });
    } finally { session.dispose(); }
});
test("unspecified and explicit-zero outline widths survive native export and reopen", () => {
    for (const width of [undefined, 0]) {
        const session = createDocumentSession(openDocument(createPresentation({ documentId: "line-export" })));
        const style = '--office-line-kind:solid;--office-line-color:#123456' + (width === undefined ? '' : ';--office-line-width:0pt');
        try {
            session.commit(session.prepare({ protocolVersion: 1, documentId: "line-export", baseRevision: 0, operations: [{ op: "insertPresentationObject", parentNodeId: "slide-root-1", geometry: { x: 20, y: 20, width: 200, height: 100, rotation: 0, flipH: false, flipV: false }, object: { kind: "geometry", preset: "rect", style } }] }));
            const output = exportPptx(session), native = unzipSync(output.bytes);
            expect(strFromU8(native['ppt/slides/slide1.xml']!)).toContain(`<a:ln w="${width === 0 ? 0 : 9525}"`);
            const target = openDocument(importPptx(output.bytes, { documentId: "line-reopen" })), object = target.snapshot().metadata.presentation!.objects.find(object => object.kind === 'geometry')!;
            expect(target.getNode(object.nodeId)!.attributes.style).toContain('--office-line-width:' + (width ?? .75) + 'pt');
        } finally { session.dispose(); }
    }
});
test("new source, rich notes, UTF-16/variation sequences and current edits do not depend on original PPTX bytes", () => {
    const session = createDocumentSession(openDocument(createPresentation({ documentId: "new-presentation" }))), slide = session.snapshot().metadata.presentation!.slides[0]!;
    session.commit(session.prepare({ protocolVersion: 1, documentId: "new-presentation", baseRevision: 0, operations: [{ op: "insertPresentationObject", parentNodeId: slide.rootNodeId, geometry: { x: 30, y: 30, width: 400, height: 120, rotation: 0, flipH: false, flipV: false }, object: { kind: "text", text: "𠀀󠄀 新增内容" } }, { op: "insertPresentationNotes", slideId: slide.slideId, html: '<p><em>演讲者备注</em></p>' }] }));
    const target = reopen(session);
    expect(texts(target).join("")).toContain("𠀀󠄀 新增内容");
    const story = target.snapshot().metadata.presentation!.stories.find(story => story.kind === "notes")!;
    expect(target.getNode(story.rootNodeId)!.text).toBe("演讲者备注");
    expect(exportPptx(session).report.revision).toBe(1);
    expect(() => exportDocx(session)).toThrowError(expect.objectContaining({ code: "UnsupportedFormat" }));
    session.dispose();
});
test("changed chart constants are synchronized with a native editable embedded workbook", () => {
    const session = createDocumentSession(model("09-charts.pptx")), chart = session.snapshot().metadata.presentation!.charts[0]!, dataset = session.snapshot().metadata.presentation!.dataSets.find(data => data.dataSetId === chart.dataSetId)!;
    const files = new Map(session.serialize()), metadata = JSON.parse(strFromU8(files.get("metadata.json")!));
    metadata.presentation.dataSets.find((data: {
        dataSetId: string;
    }) => data.dataSetId === dataset.dataSetId).series[0].values = [123, null, -4];
    files.set("metadata.json", strToU8(JSON.stringify(metadata)));
    const source = openDocument(files), native = unzipSync(exportPptx(source).bytes), target = reopen(source);
    expect(data(target)[0]!.series[0]!.values).toEqual([123, null, -4]);
    const books = Object.entries(native).filter(([path]) => path.endsWith(".xlsx"));
    expect(books.length).toBeGreaterThan(0);
    const sheet = strFromU8(unzipSync(books[0]![1])["xl/worksheets/sheet1.xml"]!);
    expect(sheet).toContain("<v>123</v>");
    expect(sheet).not.toContain("<f>");
    session.dispose();
});
test("classic compatibility export reports thread, status, rich format and precise anchor losses", () => {
    const source = model("12-modern-comments.pptx"), output = exportPptx(source, { commentFormat: "classic" }), target = openDocument(importPptx(output.bytes, { documentId: "classic" }));
    expect(target.snapshot().annotations).toHaveLength(source.snapshot().annotations.length);
    expect(output.report.annotationLosses.map(loss => loss.kind)).toEqual(expect.arrayContaining(["anchor", "thread", "status", "format"]));
    expect(texts(target)).toEqual(texts(source));
});
test("unrepresentable region anchors retain full modern threads and report recovery", () => {
    const source = model("11-classic-comments.pptx"), output = exportPptx(source, { commentFormat: "modern" }), target = openDocument(importPptx(output.bytes, { documentId: "modern" }));
    expect(target.snapshot().annotations[0]!.status).toBe("needs_reanchor");
    expect(output.report.annotationLosses).toEqual(expect.arrayContaining([expect.objectContaining({ kind: "anchor" })]));
});
test("output quotas include embedded workbook parts and leave the source immutable", () => {
    const source = model("09-charts.pptx"), before = source.serialize();
    const outerFiles = Object.keys(unzipSync(exportPptx(source).bytes)).length;
    for (const options of [{ limits: { maxNodes: 30 } }, { limits: { maxFiles: outerFiles } }, { maxZipBytes: 30 }])
        expect(() => exportPptx(source, options)).toThrow(DocumentError);
    expect(source.serialize()).toEqual(before);
});
test("classic export reports formatting applied to a paragraph without inline wrappers", () => {
    const session = createDocumentSession(model("11-classic-comments.pptx")), annotation = session.snapshot().annotations[0]!, body = session.getNode(annotation.bodyNodeId!)!, paragraph = body.childNodeIds.map(id => session.getNode(id)!).find(node => node.tagName === "p")!;
    session.commit(session.prepare({ protocolVersion: 1, documentId: "native-pptx", baseRevision: 0, operations: [{ op: "setStyle", nodeId: paragraph.nodeId, property: "font-weight", value: "700" }] }));
    const output = exportPptx(session);
    expect(output.report.annotationLosses).toEqual(expect.arrayContaining([expect.objectContaining({ annotationId: annotation.annotationId, kind: "format" })]));
    session.dispose();
});
test("native output rejects independent underline color instead of dropping body formatting", () => {
    const session = createDocumentSession(model("01-rich-text.pptx")), node = [...session.snapshot().html.matchAll(/data-node-id="([^"]+)"/g)].map(match => session.getNode(match[1]!)!).find(node => node.tagName === "span" && node.text)!;
    session.commit(session.prepare({ protocolVersion: 1, documentId: "native-pptx", baseRevision: 0, operations: [{ op: "setStyle", nodeId: node.nodeId, property: "text-decoration-line", value: "underline" }, { op: "setStyle", nodeId: node.nodeId, property: "text-decoration-color", value: "#fedcba" }] }));
    expect(() => exportPptx(session)).toThrowError(expect.objectContaining({ code: "UnsupportedFeature" }));
    session.dispose();
});
test("native bullet percentages cannot silently round away fractional source formatting", () => {
    const session = createDocumentSession(model("01-rich-text.pptx")), node = [...session.snapshot().html.matchAll(/data-node-id="([^"]+)"/g)].map(match => session.getNode(match[1]!)!).find(node => node.tagName === "p" && node.attributes["data-office-bullet"])!;
    session.commit(session.prepare({ protocolVersion: 1, documentId: "native-pptx", baseRevision: 0, operations: [{ op: "setStyle", nodeId: node.nodeId, property: "--office-bullet-size", value: ".455em" }] }));
    expect(() => exportPptx(session)).toThrowError(expect.objectContaining({ code: "UnsupportedFeature" }));
    session.dispose();
});
test.each(["margin-top", "margin-left", "padding-left"])("percentage %s is rejected instead of being exported as zero", property => {
    const session = createDocumentSession(model("01-rich-text.pptx")), node = [...session.snapshot().html.matchAll(/data-node-id="([^"]+)"/g)].map(match => session.getNode(match[1]!)!).find(node => node.tagName === "p")!;
    session.commit(session.prepare({ protocolVersion: 1, documentId: "native-pptx", baseRevision: 0, operations: [{ op: "setStyle", nodeId: node.nodeId, property, value: "10%" }] }));
    const before = session.serialize();
    expect(() => exportPptx(session)).toThrowError(expect.objectContaining({ code: "UnsupportedFeature" }));
    expect(session.serialize()).toEqual(before);
    session.dispose();
});
test("new table cell insets are counted once and nonwrapping text survives native reopen", () => {
    const session = createDocumentSession(openDocument(createPresentation({ documentId: "table-export" })));
    session.commit(session.prepare({ protocolVersion: 1, documentId: "table-export", baseRevision: 0, operations: [{ op: "insertPresentationObject", parentNodeId: "slide-root-1", geometry: { x: 30, y: 30, width: 400, height: 150, rotation: 0, flipH: false, flipV: false }, object: { kind: "table", rows: 2, columns: 2 } }] }));
    const table = session.snapshot().metadata.presentation!.objects.find(object => object.kind === "table")!, cell = [...session.snapshot().html.matchAll(/data-node-id="([^"]+)"/g)].map(match => session.getNode(match[1]!)!).find(node => node.tagName === "td")!;
    session.commit(session.prepare({ protocolVersion: 1, documentId: "table-export", baseRevision: 1, operations: [{ op: "editPresentationTable", nodeId: table.nodeId, change: { kind: "setCells", row: 0, column: 0, values: [["长文本不换行 ABCDEFGHIJKLMNOPQRSTUVWXYZ"]] } }, { op: "setStyle", nodeId: cell.nodeId, property: "padding-left", value: "8pt" }, { op: "setStyle", nodeId: cell.nodeId, property: "white-space", value: "nowrap" }] }));
    const output = exportPptx(session), native = strFromU8(unzipSync(output.bytes)["ppt/slides/slide1.xml"]!), block = /<a:tc(?:\s[^>]*)?>([\s\S]*?)<\/a:tc>/.exec(native)![1]!, body = /<a:bodyPr\b([^>]*)>/.exec(block)![1]!, properties = /<a:tcPr\b([^>]*)>/.exec(block)![1]!;
    expect(Number(/lIns="(\d+)"/.exec(body)?.[1] ?? 0) + Number(/marL="(\d+)"/.exec(properties)?.[1] ?? 0)).toBe(8 * 12700);
    const target = openDocument(importPptx(output.bytes, { documentId: "table-reopen" }));
    expect(texts(target)).toEqual(texts(session));
    expect(target.snapshot().html + target.snapshot().styles.map(style => style.css).join("\n")).toContain("white-space:nowrap");
    session.dispose();
});
test.each([['p', 'border-left', '2pt solid #123456'], ['p', 'width', '80pt'], ['span', 'padding-left', '3pt']])("review: unsupported %s %s fails instead of losing current source formatting", (tag, property, value) => {
    const session = createDocumentSession(model("01-rich-text.pptx"));
    const node = [...session.snapshot().html.matchAll(/data-node-id="([^"]+)"/g)].map(match => session.getNode(match[1]!)!).find(node => node.tagName === tag && node.text)!;
    session.commit(session.prepare({ protocolVersion: 1, documentId: "native-pptx", baseRevision: 0, operations: [{ op: "setStyle", nodeId: node.nodeId, property, value }] }));
    const before = session.serialize();
    expect(() => exportPptx(session)).toThrowError(expect.objectContaining({ code: "UnsupportedFeature" }));
    expect(session.serialize()).toEqual(before);
    session.dispose();
});
test("review: edited CSS underline style takes precedence over the imported native hint", () => {
    const session = createDocumentSession(model("01-rich-text.pptx"));
    const paragraph = [...session.snapshot().html.matchAll(/data-node-id="([^"]+)"/g)].map(match => session.getNode(match[1]!)!).find(node => node.tagName === "p")!;
    session.commit(session.prepare({ protocolVersion: 1, documentId: "native-pptx", baseRevision: 0, operations: [{ op: "setStyle", nodeId: paragraph.nodeId, property: "text-decoration-line", value: "underline" }, { op: "setStyle", nodeId: paragraph.nodeId, property: "--office-underline", value: "sng" }, { op: "setStyle", nodeId: paragraph.nodeId, property: "text-decoration-style", value: "double" }] }));
    const native = strFromU8(unzipSync(exportPptx(session).bytes)["ppt/slides/slide1.xml"]!);
    expect(/<a:pPr\b[\s\S]*?<a:defRPr\b([^>]*)>/.exec(native)![1]).toContain('u="dbl"');
    session.dispose();
});
test("review: HTML rowspan zero covers the remaining table rows in native output", () => {
    const session = createDocumentSession(openDocument(createPresentation({ documentId: "rowspan-zero" })));
    session.commit(session.prepare({ protocolVersion: 1, documentId: "rowspan-zero", baseRevision: 0, operations: [{ op: "insertPresentationObject", parentNodeId: "slide-root-1", geometry: { x: 30, y: 30, width: 400, height: 150, rotation: 0, flipH: false, flipV: false }, object: { kind: "table", rows: 3, columns: 2 } }] }));
    const files = new Map(session.serialize()), entry = session.snapshot().manifest.entrypoint;
    const table = '<table data-node-id="review-table"><tbody data-node-id="review-body"><tr data-node-id="review-row-1"><td data-node-id="review-cell-1" rowspan="0">A</td><td data-node-id="review-cell-2">B</td></tr><tr data-node-id="review-row-2"><td data-node-id="review-cell-3">C</td></tr><tr data-node-id="review-row-3"><td data-node-id="review-cell-4">D</td></tr></tbody></table>';
    files.set(entry, strToU8(strFromU8(files.get(entry)!).replace(/<table\b[\s\S]*?<\/table>/, table)));
    const source = openDocument(files), output = exportPptx(source), native = strFromU8(unzipSync(output.bytes)["ppt/slides/slide1.xml"]!);
    expect(native).toContain('rowSpan="3"');
    expect(texts(openDocument(importPptx(output.bytes, { documentId: "rowspan-reopen" })))).toEqual(texts(source));
    session.dispose();
});

test("chart label hierarchy exports point/series overrides independently of later general CSS order", () => {
    const session = createDocumentSession(openDocument(createPresentation({ documentId: "label-layer" })));
    try {
        session.commit(session.prepare({ protocolVersion: 1, documentId: "label-layer", baseRevision: 0, operations: [{ op: "insertPresentationChart", parentNodeId: "slide-root-1", geometry: { x: 20, y: 20, width: 400, height: 240, rotation: 0, flipH: false, flipV: false }, kind: "column", data: { categories: ["甲", "乙"], series: [{ name: "系列", values: [2, 4] }] } }] }));
        const files = new Map(session.serialize()), metadata = JSON.parse(strFromU8(files.get("metadata.json")!)), chart = metadata.presentation.charts[0], series = metadata.presentation.dataSets[0].series[0];
        chart.showLabels = true; chart.labelCssClass = "all-labels"; chart.seriesStyles = [{ seriesId: series.seriesId, cssClass: "series-paint", labelCssClass: "series-label" }]; chart.pointStyles = [{ seriesId: series.seriesId, index: 0, cssClass: "point-paint", labelCssClass: "point-label" }];
        files.set("metadata.json", strToU8(JSON.stringify(metadata)));
        const path = session.snapshot().manifest.styles[0]!;
        files.set(path, strToU8(strFromU8(files.get(path)!) + '.series-paint{}.point-paint{}.point-label{color:#703090;font-size:16pt;text-decoration-line:none}.series-label{color:#20a040;font-size:11pt}.all-labels{color:#a05020;font-size:9pt;--decor:underline double currentColor;text-decoration:var(--decor)}'));
        const native = unzipSync(exportPptx(openDocument(files)).bytes), xml = strFromU8(native["ppt/charts/chart1.xml"]!);
        const labels = xml.match(/<c:dLbls>[\s\S]*?<\/c:dLbls>/g)!;
        const seriesLabels = labels.find(value => value.includes("<c:dLbl>"))!, point = seriesLabels.match(/<c:dLbl>[\s\S]*?<\/c:dLbl>/)![0];
        expect(point).toContain('sz="1600"'); expect(point).toContain('u="none"'); expect(point).toContain('val="703090"'); expect(point).not.toContain('val="a05020"');
        expect(seriesLabels.slice(point.length)).toContain('val="20a040"'); expect(seriesLabels.slice(point.length)).toContain('sz="1100"'); expect(seriesLabels.slice(point.length)).toContain('u="dbl"');
    } finally { session.dispose(); }
});
