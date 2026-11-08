import { encodeTestPackage } from './package-transport.js';
import { openSdkHost } from './sdk-host.js';
import { openContentFixture as openDocument } from "../content-fixture.js";
import { readFileSync, readdirSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { expect, test } from "@playwright/test";
import type { Page } from "@playwright/test";
import { build } from "vite";
import { importPptx, readDocumentZip } from "@doffi/office";
import { writeOpcZip } from "../../packages/office/src/document-zip.js";
import { createPresentation, readPresentationGeometry } from "@doffi/core";
import type { DocumentPackage, DocumentMetadata } from "@doffi/core";
import type { DocumentPreviewOptions } from "@doffi/render";
import type { PresentationHost } from "./fixtures/presentation-host.js";
declare global { interface Window { presentationHost: PresentationHost } }
let code: string;
test.beforeAll(async () => {
  const result = await build({ configFile: false, logLevel: "silent", build: { write: false, target: "es2022", lib: { entry: fileURLToPath(new URL("./fixtures/presentation-host.ts", import.meta.url)), formats: ["es"] } } });
  const output = Array.isArray(result) ? result[0]!.output : "output" in result ? result.output : [];
  code = output.flatMap(item => item.type === "chunk" ? [item.code] : []).join("\n");
});
const entries = encodeTestPackage;
async function mount(page: Page, files: DocumentPackage, options: DocumentPreviewOptions = {}): Promise<void> {
  await openSdkHost(page); await page.addScriptTag({ type: "module", content: code }); await page.waitForFunction(() => !!window.presentationHost);
  await page.evaluate(args => window.presentationHost.mount(args.files, args.options), { files: entries(files), options });
}
const fixture = (name: string, producer = false): DocumentPackage => importPptx(new Uint8Array(readFileSync(`examples/pptx/${producer ? "producer/" : ""}${name}.pptx`)), { documentId: "ppt-preview" });
test("new decks render fixed pages without an editing session or changes to the canonical source", async ({ page }) => {
  const files = createPresentation({ documentId: "new-deck" }); await mount(page, files);
  const report = await page.evaluate(() => window.presentationHost.view!.layout());
  expect(report.pages).toHaveLength(1); expect(report.pages[0]!.setup.width).toBe(960); expect(report.pages[0]!.setup.height).toBe(540);
  const source = await page.evaluate(() => window.presentationHost.model!.snapshot().html); expect(source).toBe(openDocument(files).snapshot().html);
  await expect(page.frameLocator("iframe").locator("script,input,textarea,[contenteditable=true]")).toHaveCount(0);
  const projected = await page.evaluate(() => window.presentationHost.projected());
  expect(projected).not.toMatch(/svg|data-sdk|sdk-ppt|position:|transform:/);
});

test("Word UI adapters reject PPT without retaining a partial mount", async ({ page }) => {
  await mount(page, createPresentation({ documentId: "format-gate" }));
  expect(await page.evaluate(() => window.presentationHost.wordAdapterGates())).toEqual(Array.from({ length: 3 }, () => ({ code: "UnsupportedFeature", children: 0 })));
});

test("PPT core edits update the same preview and source with one undo/redo revision", async ({ page }) => {
  const files = fixture("01-rich-text"), model = openDocument(files), p = model.snapshot().metadata.presentation!, slide = model.getNode(p.slides[0]!.rootNodeId)!, object = p.objects.find(o => o.kind === "text" && slide.childNodeIds.includes(o.nodeId))!, node = model.getNode(object.nodeId)!, geometry = readPresentationGeometry(node.attributes)!;
  const paragraph = model.getNode(node.childNodeIds[0]!)!, text = paragraph.text;
  await mount(page, files);
  const box = page.frameLocator("iframe").locator('[data-node-id="' + object.nodeId + '"]'), before = await box.boundingBox();
  await page.evaluate(async args => window.presentationHost.edit([{ op: "setPresentationGeometry", nodeId: args.id, geometry: args.geometry }, { op: "replaceText", nodeId: args.paragraph, from: 0, to: 0, expectedText: "", text: "新增𠀀😀" }]), { id: object.nodeId, geometry: { ...geometry, x: geometry.x + 27 }, paragraph: paragraph.nodeId });
  expect((await box.boundingBox())!.x - before!.x).toBeCloseTo(36, 1);
  await expect(box).toContainText("新增𠀀😀" + text);
  expect(await page.evaluate(() => window.presentationHost.model!.snapshot().manifest.revision)).toBe(1);
  expect(await page.evaluate(() => window.presentationHost.projected())).not.toMatch(/data-sdk-|sdk-ppt|<svg|blob:/);
  await page.evaluate(() => window.presentationHost.history("undo")); expect((await box.boundingBox())!.x).toBeCloseTo(before!.x, 1); await expect(box.locator('[data-node-id="' + paragraph.nodeId + '"]')).toHaveText(text);
  await page.evaluate(() => window.presentationHost.history("redo")); expect((await box.boundingBox())!.x - before!.x).toBeCloseTo(36, 1);
  expect(await page.evaluate(() => window.presentationHost.model!.snapshot().manifest.revision)).toBe(3);
});
for (const producer of [false, true]) for (const file of readdirSync("examples/pptx").filter(name => name.endsWith(".pptx"))) {
  test(`${producer ? "producer" : "source"} ${file} retains its package while composing the complete declared read-only view`, async ({ page }, testInfo) => {
    const files = fixture(file.slice(0, -5), producer), model = openDocument(files), snapshot = model.snapshot();
    await mount(page, files);
    const state = await page.evaluate(() => ({ report: window.presentationHost.view!.layout(), html: window.presentationHost.model!.snapshot().html, projected: window.presentationHost.projected(), print: window.presentationHost.view!.printDocument().body.innerHTML }));
    expect(state.html).toBe(snapshot.html); expect(state.report.pages).toHaveLength(snapshot.metadata.presentation!.slides.length);
    expect(state.report.fields.every(f => f.kind === "PAGE")).toBe(true);
    expect(state.projected).not.toMatch(/data-sdk-|sdk-ppt|<svg|blob:/);
    expect(state.print).not.toContain("演讲者备注"); expect(state.print).not.toContain('class="sdk-ppt-comments"');
    const slides = page.frameLocator("iframe").locator('article > [data-office-kind="slide"]');
    for (let i = 0; i < state.report.pages.length; i++) {
      const clip = await slides.nth(i).boundingBox(); expect(clip!.width).toBeCloseTo(snapshot.metadata.presentation!.size.width * 4 / 3, 0); expect(clip!.height).toBeCloseTo(snapshot.metadata.presentation!.size.height * 4 / 3, 0);
      await testInfo.attach(`slide-${i + 1}`, { body: await slides.nth(i).screenshot({ path: `.cache/presentation-preview/${testInfo.project.name}/${producer ? "producer-" : "source-"}${file}-${i + 1}.png` }), contentType: "image/png" });
      if (!producer && testInfo.project.name === "chromium" && !process.env.DOFFI_BROWSER_EXECUTABLE && ["04-shapes.pptx", "05-groups.pptx", "09-charts.pptx"].includes(file))
        await expect(slides.nth(i)).toHaveScreenshot(`${file}-${i + 1}.png`, { maxDiffPixelRatio: 0.005, threshold: 0.2 });
    }
  });
}

test("shape text rectangles wrap long words and group transforms preserve glyph orientation and point sizes", async ({ page }) => {
  await mount(page, fixture("04-shapes"));
  const shapes = await page.frameLocator("iframe").locator('article > [data-office-kind="slide"] [data-office-kind="geometry"]').evaluateAll(nodes => nodes.map(node => {
    const text = node.querySelector<HTMLElement>(".sdk-ppt-text")!, path = node.querySelector("svg > path")!;
    return { preset: node.querySelector("svg")!.getAttribute("data-sdk-preset"), width: text.clientWidth, scrollWidth: text.scrollWidth, display: getComputedStyle(text).display, path: path.getAttribute("d") };
  }));
  expect(shapes).toHaveLength(16); expect(shapes.find(s => s.preset === "line")!.display).toBe("none");
  expect(shapes.find(s => s.preset === "downArrow")!.scrollWidth).toBeLessThanOrEqual(shapes.find(s => s.preset === "downArrow")!.width + 1);
  await mount(page, fixture("05-groups"));
  const glyphs = await page.frameLocator("iframe").locator('article > [data-office-kind="slide"] .sdk-ppt-text').evaluateAll(nodes => nodes.map(node => {
    let matrix = new DOMMatrix();
    for (let current: Element | null = node; current && current.getAttribute("data-office-kind") !== "slide"; current = current.parentElement) {
      const transform = getComputedStyle(current).transform; if (transform !== "none") matrix = new DOMMatrix(transform).multiply(matrix);
    }
    return { scale: Math.hypot(matrix.a, matrix.b), determinant: matrix.a * matrix.d - matrix.b * matrix.c, fontSize: getComputedStyle(node.querySelector("[data-office-script]")!).fontSize, text: node.textContent };
  }));
  expect(glyphs).toHaveLength(2); for (const glyph of glyphs) { expect(glyph.scale).toBeCloseTo(1, 4); expect(glyph.determinant).toBeCloseTo(1, 4); expect(parseFloat(glyph.fontSize)).toBeCloseTo(32, 1); }
});

test("bound connectors meet native connection sites without changing arrow definitions or the source transform", async ({ page }) => {
  await mount(page, fixture("06-connectors"));
  const paths = await page.frameLocator("iframe").locator('article > [data-office-kind="slide"] [data-office-kind="connector"] svg > path').evaluateAll(nodes => nodes.map(node => {
    const path = node as SVGPathElement, ctm = path.getScreenCTM()!, start = path.getPointAtLength(0).matrixTransform(ctm), end = path.getPointAtLength(path.getTotalLength()).matrixTransform(ctm);
    return { start: { x: start.x, y: start.y }, end: { x: end.x, y: end.y }, marker: path.parentElement!.querySelector("marker path")!.getAttribute("d") };
  }));
  expect(paths).toHaveLength(2);
  for (const path of paths) { expect(path.start.x).toBeCloseTo(paths[0]!.start.x, 1); expect(path.start.y).toBeCloseTo(paths[0]!.start.y, 1); expect(path.end.x).toBeCloseTo(paths[0]!.end.x, 1); expect(path.end.y).toBeCloseTo(paths[0]!.end.y, 1); expect(path.marker!.length).toBeLessThan(60); }
  const start = await page.frameLocator("iframe").locator('article > [data-office-kind="slide"] [data-office-kind="geometry"]').first().boundingBox();
  expect(paths[0]!.start.x).toBeCloseTo(start!.x, 1); expect(paths[0]!.start.y).toBeCloseTo(start!.y + start!.height / 2, 1);
});

test("chart decorations use final substituted fonts and survive print without entering the source", async ({ page }) => {
  const files = new Map(fixture("09-charts")), model = openDocument(files), snapshot = model.snapshot(), metadata = JSON.parse(new TextDecoder().decode(files.get("metadata.json")!)) as DocumentMetadata, chart = metadata.presentation!.charts[0]!;
  chart.showLabels = true; chart.labelContents = ["value"]; chart.labelCssClass = "decor-label";
  chart.pointStyles = [{ seriesId: metadata.presentation!.dataSets.find(data => data.dataSetId === chart.dataSetId)!.series[0]!.seriesId, index: 0, cssClass: "decor-point", labelCssClass: "decor-none" }];
  files.set("metadata.json", new TextEncoder().encode(JSON.stringify(metadata)));
  const path = snapshot.manifest.styles[0]!;
  files.set(path, new TextEncoder().encode(new TextDecoder().decode(files.get(path)!) + '.decor-label{--office-font-ascii:"DOFFI Unavailable Chart Font";font-size:18pt;text-decoration:underline double currentColor}.decor-point{}.decor-none{text-decoration-line:none}'));
  await mount(page, files);
  const result = await page.evaluate(() => {
    const view = window.presentationHost.view!, document = view.element.contentDocument!;
    return { substituted: view.fonts().some(font => font.requestedFamily === "DOFFI Unavailable Chart Font" && font.substituted), projected: window.presentationHost.projected(),
      print: view.printDocument().querySelectorAll('[data-sdk-chart-decoration-paint] line').length,
      labels: [...document.querySelectorAll<SVGTextElement>('svg[data-sdk-chart=column] [data-sdk-chart-label=data-label]')].map(label => {
        const bounds = label.getBBox(), carrier = label.closest('[data-sdk-chart-text-decoration]');
        return { accessible: !carrier?.hasAttribute('aria-hidden'), strokes: [...(carrier?.querySelectorAll<SVGLineElement>('[data-sdk-chart-decoration-paint] line') ?? [])].map(line => ({ x: line.x1.baseVal.value, width: line.x2.baseVal.value - line.x1.baseVal.value })), bounds: { x: bounds.x, width: bounds.width } };
      }) };
  });
  expect(result.substituted).toBe(true); expect(result.print).toBeGreaterThan(0); expect(result.projected).not.toMatch(/data-sdk-|sdk-ppt|<svg/);
  expect(result.labels[0]!.strokes).toHaveLength(0);
  for (const label of result.labels.slice(1)) { expect(label.accessible).toBe(true); expect(label.strokes).toHaveLength(2); for (const stroke of label.strokes) { expect(stroke.x).toBeCloseTo(label.bounds.x, 4); expect(stroke.width).toBeCloseTo(label.bounds.width, 4); } }
});

test("charts plot current constants at physical point sizes and rich UTF-16 selection maps to canonical text", async ({ page }) => {
  await mount(page, fixture("09-charts"));
  const state = await page.frameLocator("iframe").locator('svg[data-sdk-chart="column"]').evaluate(node => ({ texts: [...node.querySelectorAll("text")].map(n => n.textContent), sizes: [...node.querySelectorAll("text")].map(n => parseFloat(getComputedStyle(n).fontSize)), bars: [...node.querySelectorAll("rect")].filter(n => Number(n.getAttribute("height")) < 200 && Number(n.getAttribute("width")) > 30).map(n => Number(n.getAttribute("height"))) }));
  expect(state.texts).toEqual(expect.arrayContaining(["第一季度", "第二季度", "第三季度", "交付", "0", "9"]));
  expect(state.sizes.every(n => Math.abs(n - 10) < 0.1)).toBe(true); expect(state.bars).toHaveLength(3); expect(state.bars[2]! / state.bars[0]!).toBeCloseTo(4, 4);
  await mount(page, fixture("12-modern-comments"));
  const anchor = await page.evaluate(() => {
    const host = window.presentationHost, source = host.view!.source()!, p = [...source.root.querySelectorAll<HTMLElement>('section[data-office-kind="slide"] p')].find(p => p.textContent!.includes("𠀀😀"))!;
    const range = p.ownerDocument.createRange(); range.selectNodeContents(p); return source.textAnchor(range, host.model!);
  });
  expect(anchor).toMatchObject({ kind: "text", from: 0, to: "中文𠀀😀批注锚点".length, quote: "中文𠀀😀批注锚点" });
  await expect(page.frameLocator("iframe").locator(".sdk-ppt-comments")).toContainText("现代批注");
});

test("updates revoke private image URLs and failed updates restore the previous view", async ({ page }) => {
  await page.addInitScript(() => {
    const created: string[] = [], revoked: string[] = [], create = URL.createObjectURL.bind(URL), revoke = URL.revokeObjectURL.bind(URL);
    URL.createObjectURL = blob => { const url = create(blob); created.push(url); return url; }; URL.revokeObjectURL = url => { revoked.push(url); revoke(url); };
    Object.assign(window, { pptUrls: { created, revoked } });
  });
  await mount(page, fixture("07-images"), { limits: { maxPages: 1 } });
  await expect(page.frameLocator("iframe").locator("img").first()).toBeVisible();
  await expect(page.evaluate(files => window.presentationHost.update(files), entries(fixture("09-charts")))).rejects.toThrow("page budget");
  expect(await page.evaluate(() => window.presentationHost.view!.layout().pages.length)).toBe(1);
  await expect(page.frameLocator("iframe").locator("img")).toHaveCount(3);
  const pending = await page.evaluate(async files => {
    const promise = window.presentationHost.view!.settled(); window.presentationHost.update(files);
    try { await promise; return "unexpected success"; } catch (error) { return String(error); }
  }, entries(createPresentation({ documentId: "replacement" })));
  expect(pending).toContain("replaced");
  await page.evaluate(() => window.presentationHost.view!.dispose());
  const urls = await page.evaluate(() => (window as unknown as { pptUrls: { created: string[]; revoked: string[] } }).pptUrls);
  expect(urls.created).toHaveLength(3); expect(urls.revoked.sort()).toEqual(urls.created.sort()); await expect(page.locator("iframe")).toHaveCount(0);
});

test("render and print budgets fail explicitly and failed mounts release their resources", async ({ page }) => {
  await expect(mount(page, fixture("07-images"), { limits: { maxNodes: 3 } })).rejects.toThrow("node budget"); await expect(page.locator("iframe")).toHaveCount(0);
  await mount(page, fixture("09-charts"), { limits: { maxPrintNodes: 10 } });
  await expect(page.evaluate(() => window.presentationHost.view!.printDocument())).rejects.toThrow("node budget");
  await expect(mount(page, createPresentation({ documentId: "continuous" }), { layout: "continuous" })).rejects.toThrow("fixed pages");
});

test("normal and shape auto-fit remain derived and print retains resolved inheritance CSS", async ({ page }) => {
  const files = new Map(fixture("01-rich-text")), encoder = new TextEncoder(), decoder = new TextDecoder();
  const metadata = JSON.parse(decoder.decode(files.get("metadata.json"))) as { presentation: { objects: { kind: string; text?: { autoFit: string; wrap: boolean } }[] } };
  const object = metadata.presentation.objects.find(o => o.kind === "text")!; object.text = { autoFit: "normal", wrap: true };
  files.set("metadata.json", encoder.encode(JSON.stringify(metadata)));
  files.set("content.html", encoder.encode(decoder.decode(files.get("content.html")).replace('data-office-height="360"', 'data-office-height="70"').replace(/data-office-height="[\d.]+"/, 'data-office-height="70"').replace('--office-fill-kind:none', '--office-fill-kind:linear')));
  files.set("styles.css", encoder.encode(decoder.decode(files.get("styles.css")) + '\n[data-office-kind="text"]{--office-fill-kind:linear;--office-fill-angle:90;--office-fill-stop-count:2;--office-fill-stop-0-position:0;--office-fill-stop-0-color:#ff0000;--office-fill-stop-1-position:1;--office-fill-stop-1-color:#0000ff}'));
  await mount(page, files);
  const normal = await page.frameLocator("iframe").locator('section[data-office-kind="slide"] .sdk-ppt-text').evaluate(node => ({ height: node.clientHeight, scroll: node.scrollHeight, font: parseFloat(getComputedStyle(node.querySelector("[data-office-script]")!).fontSize) }));
  expect(normal.scroll).toBeLessThanOrEqual(normal.height + 1); expect(normal.font).toBeLessThan(32);
  object.text.autoFit = "shape"; files.set("metadata.json", encoder.encode(JSON.stringify(metadata))); await mount(page, files);
  const shape = await page.frameLocator("iframe").locator('section[data-office-kind="slide"] .sdk-ppt-text').evaluate(node => ({ height: node.clientHeight, scroll: node.scrollHeight, font: parseFloat(getComputedStyle(node.querySelector("[data-office-script]")!).fontSize) }));
  expect(shape.height).toBeGreaterThan(100); expect(shape.scroll).toBeLessThanOrEqual(shape.height + 1); expect(shape.font).toBeCloseTo(32, 1);
  const gradient = await page.frameLocator("iframe").locator('section[data-office-kind="slide"] [data-office-kind="text"] > svg').first().evaluate(node => ({ height: (node as SVGSVGElement).viewBox.baseVal.height, y1: Number(node.querySelector('linearGradient')!.getAttribute('y1')), y2: Number(node.querySelector('linearGradient')!.getAttribute('y2')) }));
  expect(gradient.height).toBeGreaterThan(70); expect(gradient.y1).toBeCloseTo(0, 8); expect(gradient.y2).toBeCloseTo(gradient.height, 8);
  expect(await page.evaluate(() => window.presentationHost.model!.snapshot().html)).toContain('data-office-height="70"');
  await mount(page, fixture("02-theme-placeholders"));
  const print = await page.evaluate(() => {
    const host = window.presentationHost, doc = host.view!.printDocument(), iframe = document.createElement("iframe");
    document.body.append(iframe); iframe.contentDocument!.open(); iframe.contentDocument!.write(doc.documentElement.outerHTML); iframe.contentDocument!.close();
    const source = host.view!.source()!.root.querySelector<HTMLElement>('section[data-office-kind="slide"] .sdk-ppt-text p')!, copy = iframe.contentDocument!.querySelector<HTMLElement>('.sdk-ppt-text p')!;
    const actual = host.view!.element.contentWindow!.getComputedStyle(source), printed = iframe.contentWindow!.getComputedStyle(copy);
    const result = { source: [actual.color, actual.fontFamily, actual.fontSize], print: [printed.color, printed.fontFamily, printed.fontSize] }; iframe.remove(); return result;
  });
  expect(print.print).toEqual(print.source);
});

test("review: inherited relative leading follows a local paragraph font size", async ({ page }) => {
  const files = new Map(fixture("01-rich-text")), decoder = new TextDecoder(), encoder = new TextEncoder();
  const model = openDocument(files), master = model.snapshot().metadata.presentation!.masters[0]!;
  files.set("styles.css", encoder.encode(decoder.decode(files.get("styles.css")) + `\n.${master.otherStyle}-level-0{line-height:1.5}`));
  files.set("content.html", encoder.encode(decoder.decode(files.get("content.html")).replace(/<p\b[^>]*>/, tag => tag.slice(0, -1) + ' style="font-size:48pt">')));
  await mount(page, files);
  const values = await page.frameLocator("iframe").locator('section[data-office-kind="slide"] .sdk-ppt-text p').first().evaluate(node => ({ size: parseFloat(getComputedStyle(node).fontSize), leading: parseFloat(getComputedStyle(node).lineHeight) }));
  expect(values.leading).toBeCloseTo(values.size * 1.5, 1);
});

test("review: HTML text uses its declared color transforms as well as its base CSS color", async ({ page }) => {
  const files = new Map(fixture("01-rich-text")), decoder = new TextDecoder(), encoder = new TextEncoder();
  files.set("content.html", encoder.encode(decoder.decode(files.get("content.html")).replace(/<p\b[^>]*>/, tag => tag.slice(0, -1) + ' style="color:#c00000;--office-text-color:#c00000;--office-text-alpha:0.5">')));
  await mount(page, files);
  const color = await page.frameLocator("iframe").locator('section[data-office-kind="slide"] .sdk-ppt-text p').first().locator("[data-office-script]").first().evaluate(node => getComputedStyle(node).color);
  expect(color).toBe("rgba(192, 0, 0, 0.5)");
});

test("review: a native theme style composes reference-color alpha with the format fill", async ({ page }) => {
  const source = new Map(readDocumentZip(readFileSync("examples/pptx/04-shapes.pptx"))), decoder = new TextDecoder(), encoder = new TextEncoder();
  const reference = '<p:style><a:lnRef idx="0"><a:srgbClr val="000000"/></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"><a:alpha val="50000"/></a:schemeClr></a:fillRef><a:effectRef idx="0"><a:srgbClr val="000000"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style>';
  source.set("ppt/slides/slide1.xml", encoder.encode(decoder.decode(source.get("ppt/slides/slide1.xml")).replace(/<p:sp>[\s\S]*?<\/p:sp>/, shape => shape.replace(/<a:solidFill>[\s\S]*?<\/a:solidFill>/, "").replace("<p:txBody>", reference + "<p:txBody>"))));
  await mount(page, importPptx(writeOpcZip(source), { documentId: "reference-alpha" }));
  const fill = await page.frameLocator("iframe").locator('article > [data-office-kind="slide"] [data-office-kind="geometry"] svg > path').first().getAttribute("fill");
  expect(fill).toBe("rgba(68,114,196,0.5)");
});

test("a platform glyph-raster failure is reported as incomplete rather than rejecting a readable deck", async ({ page }) => {
  await page.addInitScript(() => { CanvasRenderingContext2D.prototype.getImageData = () => { throw new DOMException("Raster unavailable", "InvalidStateError"); }; });
  await mount(page, fixture("09-charts"));
  const report = await page.evaluate(() => window.presentationHost.view!.layout()); expect(report.pages).toHaveLength(3); expect(report.fonts.some(font => !font.glyphAuditComplete)).toBe(true);
  await expect(page.frameLocator("iframe").locator('[role="note"]')).toContainText("部分字形未核验");
});

test("expanding notes updates later slide coordinates and iframe height without changing the source", async ({ page }) => {
  await mount(page, fixture("13-multi-slide"));
  const before = await page.evaluate(() => ({ html: window.presentationHost.model!.snapshot().html, y: window.presentationHost.view!.layout().pages[1]!.rect.y, height: parseFloat(window.presentationHost.view!.element.style.height) }));
  await page.frameLocator("iframe").locator(".sdk-ppt-notes summary").click();
  await expect.poll(() => page.evaluate(() => window.presentationHost.view!.layout().pages[1]!.rect.y)).toBeGreaterThan(before.y);
  expect(await page.evaluate(() => parseFloat(window.presentationHost.view!.element.style.height))).toBeGreaterThan(before.height);
  expect(await page.evaluate(() => window.presentationHost.model!.snapshot().html)).toBe(before.html);
  const prevented = await page.frameLocator("iframe").locator(".sdk-ppt-notes p").first().evaluate(node => {
    const range = document.createRange(); range.selectNodeContents(node); const selection = document.getSelection()!; selection.removeAllRanges(); selection.addRange(range);
    const event = new ClipboardEvent("copy", { bubbles: true, cancelable: true, clipboardData: new DataTransfer() }); document.dispatchEvent(event); return event.defaultPrevented;
  });
  expect(prevented).toBe(false);
});

test("notes body resolves its notes-master typography and background in the read-only panel", async ({ page }) => {
  const files = new Map(fixture("10-notes")), decoder = new TextDecoder(), encoder = new TextEncoder(), metadata = openDocument(files).snapshot().metadata.presentation!, master = metadata.notesMasters[0]!;
  files.set("styles.css", encoder.encode(decoder.decode(files.get("styles.css")) + `\n.${master.textStyle}-level-0{font-style:italic;--office-font-east-asia:"Noto Serif CJK SC";color:#990000}`));
  files.set("content.html", encoder.encode(decoder.decode(files.get("content.html")).replace(new RegExp(`<section data-node-id="${master.rootNodeId}"[^>]*>`), tag => tag.replace('style="', 'style="--office-fill-kind:solid;--office-fill-color:#ffff00;'))));
  await mount(page, files); await page.frameLocator("iframe").locator(".sdk-ppt-notes summary").click();
  const span = page.frameLocator("iframe").locator('.sdk-ppt-notes [data-node-id] [data-office-script="east-asia"]').first();
  const style = await span.evaluate(node => { const s = getComputedStyle(node); return { font: s.fontFamily, italic: s.fontStyle, color: s.color }; });
  expect(style.font).toContain("Noto Serif CJK SC"); expect(style.italic).toBe("italic"); expect(style.color).toBe("rgb(153, 0, 0)");
  expect(await page.frameLocator("iframe").locator(".sdk-ppt-note-content").evaluate(node => getComputedStyle(node).backgroundColor)).toBe("rgb(255, 255, 0)");
});

test("modern comment bodies resolve their slide theme font references and preserve rich lists", async ({ page }) => {
  const files = new Map(fixture("12-modern-comments")), decoder = new TextDecoder(), encoder = new TextEncoder(), model = openDocument(files), metadata = model.snapshot().metadata.presentation!, body = model.getNode(model.snapshot().annotations[0]!.bodyNodeId!)!;
  const theme = metadata.themes[0]!, bodyClass = body.attributes.class!.split(" ")[0]!;
  files.set("styles.css", encoder.encode(decoder.decode(files.get("styles.css")) + `\n.${theme.cssClass}{--office-minor-east-asia:"Noto Serif CJK SC"}\n.${bodyClass}{--office-font-east-asia:var(--office-minor-east-asia);--office-list-kind:bullet;--office-list-character:"•"}`));
  await mount(page, files);
  const bodyView = page.frameLocator("iframe").locator(`.sdk-ppt-comments [data-node-id="${body.nodeId}"]`);
  expect(await bodyView.locator('[data-office-script="east-asia"]').first().evaluate(node => getComputedStyle(node).fontFamily)).toContain("Noto Serif CJK SC");
  await expect(bodyView.locator(".sdk-ppt-marker")).toHaveText("•");
  await expect(page.frameLocator("iframe").locator(".sdk-ppt-comments")).toContainText(model.snapshot().annotations[0]!.date!);
});
