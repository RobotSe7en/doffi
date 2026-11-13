import { openSdkHost } from './sdk-host.js';
import { fileURLToPath } from "node:url";
import { expect, test } from "@playwright/test";
import type { Page } from "@playwright/test";
import { build } from "vite";
import type { DocumentPackage } from "@doffi/core";
import type { DocumentPreviewOptions } from "@doffi/editor";
import { fieldsFiles, longTable, pagedFiles } from "./fixtures/pagination-files.js";
import type { PaginationHost } from "./fixtures/pagination-host.js";
declare global { interface Window { paginationHost: PaginationHost } }
let code: string;
test.beforeAll(async () => {
  const result = await build({ configFile: false, logLevel: "silent", build: { write: false, target: "es2022", lib: {
    entry: fileURLToPath(new URL("./fixtures/pagination-host.ts", import.meta.url)), formats: ["es"],
  } } });
  const output = Array.isArray(result) ? result[0]!.output : "output" in result ? result.output : [];
  code = output.flatMap((item) => item.type === "chunk" ? [item.code] : []).join("\n");
});
async function mount(page: Page, files: DocumentPackage, mode: "preview" | "edit" | "interaction" = "preview", options: DocumentPreviewOptions = {}): Promise<void> {
  await openSdkHost(page); await page.addScriptTag({ type: "module", content: code }); await page.waitForFunction(() => !!window.paginationHost);
  await page.evaluate(({ entries, mode, options }) => window.paginationHost.mount(entries, mode, options), {
    entries: [...files].map(([path, bytes]) => [path, [...bytes]] as [string, number[]]), mode, options,
  });
}
test('source projection strips late owned wrappers while preserving bound roots and unowned lookalikes', async ({ page }) => {
  await mount(page, pagedFiles('<p data-node-id="text">文档</p>'));
  const result = await page.evaluate(() => window.paginationHost.lateProjection());
  expect(result.html).toBe(result.original); expect(result.instance).toBe(result.original); expect(result.shellMissing).toBe(true);
  expect(result.copy).toBe('<span data-node-id="run" style="color:#123456">原文</span>');
  expect(result.unowned).toContain('<span class="sdk-inline-flow" style="display: block;">源内容</span>');
});
test("read-only pagination requires no editing session and preserves the complete package source", async ({ page }) => {
  const files = pagedFiles(Array.from({ length: 40 }, (_, index) => '<p data-node-id="p-' + index + '">LINE-' + index + "</p>").join(""));
  await mount(page, files);
  const state = await page.evaluate(() => ({ report: window.paginationHost.layout(), session: !!window.paginationHost.session, html: window.paginationHost.model!.snapshot().html }));
  expect(state.session).toBe(false); expect(state.report.pages).toHaveLength(4);
  const frame = page.frameLocator("iframe"); await expect(frame.locator('[contenteditable="true"]')).toHaveCount(0);
  await expect(frame.locator("script,input,textarea")).toHaveCount(0);
  expect(state.html).not.toMatch(/sdk-page|column-height|contenteditable/);
  expect(state.html).toContain("LINE-39");
  const geometry = await frame.locator('[data-node-id="p-12"]').evaluate((node) => ({ y: node.getBoundingClientRect().top, text: node.textContent }));
  expect(geometry.y).toBeGreaterThan(300); expect(state.report.nodePages["p-12"]).toEqual(["page-2"]);
});
test("one paragraph spans pages and maps a cross-page selection to one canonical UTF-16 range", async ({ page }) => {
  await mount(page, pagedFiles('<p data-node-id="long">' + Array.from({ length: 40 }, (_, index) => "LINE-" + index + " 😀").join("<br>") + "</p>"), "interaction");
  const frame = page.frameLocator("iframe");
  await frame.locator('[data-node-id="long"]').evaluate((node) => {
    const range = document.createRange(); range.selectNodeContents(node); const selected = document.getSelection()!; selected.removeAllRanges(); selected.addRange(range); document.dispatchEvent(new Event("selectionchange"));
  });
  await expect.poll(() => page.evaluate(() => window.paginationHost.interaction!.selection()?.anchor)).toEqual(await page.evaluate(() => ({ kind: "text", nodeId: "long", from: 0,
    to: window.paginationHost.session!.getNode("long")!.text.length, quote: window.paginationHost.session!.getNode("long")!.text })));
  expect((await page.evaluate(() => window.paginationHost.layout().nodePages.long))!.length).toBeGreaterThan(1);
});
test("hiding a preview preserves its accepted pages and revealing it remeasures the visible canvas", async ({ page }) => {
  const errors: string[] = []; page.on('pageerror', error => errors.push(error.message));
  await mount(page, pagedFiles(Array.from({ length: 40 }, (_, index) => `<p data-node-id="p-${index}">LINE-${index}</p>`).join('')));
  const original = await page.evaluate(() => window.paginationHost.layout()); expect(original.pages).toHaveLength(4);
  const hidden = await page.evaluate(async () => { (document.body.firstElementChild as HTMLElement).hidden = true; return window.paginationHost.view!.settled(); });
  expect(hidden.pages).toEqual(original.pages); expect(hidden.nodePages).toEqual(original.nodePages);
  await page.evaluate(() => { (document.body.firstElementChild as HTMLElement).hidden = false; });
  await expect.poll(() => page.evaluate(async () => (await window.paginationHost.view!.settled()).pages)).toEqual(original.pages);
  await expect.poll(() => page.frameLocator('iframe').locator('[data-node-id="p-12"]').evaluate(node => node.getBoundingClientRect().top)).toBeGreaterThan(300);
  expect(errors).toEqual([]);
});
test("editing reflows pages immediately and undo restores layout without persisting page wrappers", async ({ page }) => {
  await mount(page, pagedFiles('<p data-node-id="a">Start</p>'), "edit");
  const before = await page.evaluate(() => window.paginationHost.layout().pages.length);
  await page.frameLocator("iframe").locator('[data-node-id="a"]').evaluate((node) => {
    const range = document.createRange(); range.selectNodeContents(node); range.collapse(false); const selection = document.getSelection()!; selection.removeAllRanges(); selection.addRange(range); (document.querySelector("[data-office-editor]") as HTMLElement).focus();
  });
  await page.keyboard.insertText(" long text".repeat(300));
  expect(await page.evaluate(() => window.paginationHost.layout().pages.length)).toBeGreaterThan(before);
  expect(await page.evaluate(() => window.paginationHost.session!.snapshot().html)).not.toMatch(/sdk-page|column-height|contenteditable/);
  await page.keyboard.press("Control+z");
  expect(await page.evaluate(() => window.paginationHost.layout().pages.length)).toBe(before);
  expect(await page.evaluate(() => window.paginationHost.session!.getNode("a")!.text)).toBe("Start");
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
});
test("repeated table headers remain a display projection while body cells can be edited", async ({ page }) => {
  await mount(page, longTable(), "edit");
  const pages = await page.evaluate(() => window.paginationHost.layout().pages.length);
  expect(pages).toBeGreaterThan(1);
  await expect(page.frameLocator("iframe").locator(".sdk-repeat-header")).toHaveCount(pages - 1);
  await page.frameLocator("iframe").locator('[data-node-id="cell-20"]').first().evaluate((node) => {
    const range = document.createRange(); range.selectNodeContents(node); range.collapse(false); const selected = document.getSelection()!; selected.removeAllRanges(); selected.addRange(range); (document.querySelector("[data-office-editor]") as HTMLElement).focus();
  });
  await page.keyboard.insertText(" EDITED");
  expect(await page.evaluate(() => window.paginationHost.session!.getNode("cell-20")!.text)).toBe("UNIQUE-20 EDITED");
  const html = await page.evaluate(() => window.paginationHost.session!.snapshot().html);
  expect(html.match(/REPEATED HEADER/g)).toHaveLength(1); expect(html).not.toMatch(/sdk-repeat|box-decoration-break|contenteditable/);
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
});
test("page fields and TOC are derived; explicit cache update is one core undo step", async ({ page }) => {
  await mount(page, fieldsFiles(), "interaction");
  const state = await page.evaluate(() => ({ report: window.paginationHost.layout(), html: window.paginationHost.session!.snapshot().html, revision: window.paginationHost.session!.snapshot().manifest.revision }));
  expect(state.report.pages).toHaveLength(2);
  expect(state.report.fields.find((field) => field.kind === "NUMPAGES")?.value).toBe("2");
  expect(state.report.fields.find((field) => field.kind === "TOC")?.value).toBe("Alpha\t1\nBeta\t2");
  expect(state.html).not.toContain("sdk-field");
  await page.evaluate(() => window.paginationHost.updateFields());
  expect(await page.evaluate(() => window.paginationHost.session!.snapshot().manifest.revision)).toBe(state.revision + 1);
  expect(await page.evaluate(() => window.paginationHost.session!.history().undoSteps)).toBe(1);
  await page.evaluate(() => window.paginationHost.session!.undo());
  expect(await page.evaluate(() => window.paginationHost.session!.snapshot().html)).toBe(state.html);
});
test("continuous preview remains available and page budgets fail explicitly", async ({ page }) => {
  const files = pagedFiles(Array.from({ length: 40 }, (_, index) => '<p data-node-id="p-' + index + '">LINE-' + index + "</p>").join(""));
  await mount(page, files, "preview", { layout: "continuous" });
  expect(await page.evaluate(() => window.paginationHost.layout().layout)).toBe("continuous");
  await expect(page.frameLocator("iframe").locator(".sdk-paper")).toHaveCount(0);
  await expect(mount(page, files, "preview", { limits: { maxPages: 1 } })).rejects.toThrow("page budget");
});

const prefix = (count: number): string => Array.from({ length: count }, (_, index) => `<p data-node-id="prefix-${index}">LINE</p>`).join("");
async function atStart(page: Page, id: string): Promise<void> {
  await page.frameLocator("iframe").locator(`[data-node-id="${id}"]`).first().evaluate((node) => {
    const walker = document.createTreeWalker(node, NodeFilter.SHOW_TEXT), text = walker.nextNode()!;
    const range = document.createRange(); range.setStart(text, 0); range.collapse(true);
    const selected = document.getSelection()!; selected.removeAllRanges(); selected.addRange(range); (document.querySelector("[data-office-editor]") as HTMLElement).focus();
  });
}
test("headings stay with following text and short paragraphs respect orphan and widow limits", async ({ page }) => {
  await mount(page, pagedFiles(prefix(11) + '<h1 data-node-id="heading" style="font-size:12pt;line-height:16pt;break-after:avoid">HEADING</h1><p data-node-id="after">FOLLOWING</p>'));
  expect(await page.evaluate(() => window.paginationHost.layout().nodePages.heading)).toEqual(["page-2"]);
  expect(await page.evaluate(() => window.paginationHost.layout().nodePages.after)).toEqual(["page-2"]);
  await mount(page, pagedFiles(prefix(11) + '<p data-node-id="short" style="orphans:2;widows:2">FIRST<br>SECOND<br>THIRD</p>'));
  expect(await page.evaluate(() => window.paginationHost.layout().nodePages.short)).toEqual(["page-2"]);
  await mount(page, pagedFiles(prefix(9) + '<p data-node-id="four" style="orphans:2;widows:2">FIRST<br>SECOND<br>THIRD<br>FOURTH</p>'));
  const y = await page.frameLocator("iframe").locator('[data-node-id="four"]').evaluate((node) => [...node.childNodes].filter((item) => item.nodeType === 3 && item.textContent).map((item) => {
    const range = document.createRange(); range.selectNodeContents(item); return range.getBoundingClientRect().top;
  }));
  expect(y[1]! - y[0]!).toBeLessThan(30); expect(y[2]! - y[1]!).toBeGreaterThan(60); expect(y[3]! - y[2]!).toBeLessThan(30);
});
test("unidentified hard page breaks stay canonical after edit, save and reopen", async ({ page }) => {
  await mount(page, pagedFiles('<p data-node-id="a">AAA<br data-word-break="page">BBB</p>'), "edit");
  expect(await page.evaluate(() => window.paginationHost.layout().pages.length)).toBe(2);
  await atStart(page, "a"); await page.keyboard.insertText("X");
  expect(await page.evaluate(() => window.paginationHost.session!.getNode("a")!.text)).toBe("XAAA\nBBB");
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
  const entries = await page.evaluate(() => [...window.paginationHost.session!.serialize()].map(([path, bytes]) => [path, [...bytes]] as [string, number[]]));
  await page.evaluate((files) => window.paginationHost.mount(files, "preview"), entries);
  expect(await page.evaluate(() => window.paginationHost.layout().pages.length)).toBe(2);
  expect(await page.evaluate(() => window.paginationHost.model!.snapshot().html)).not.toMatch(/display:\s*none|sdk-page-gap/);
});
test("direct container text is paginated without losing its source identity", async ({ page }) => {
  await mount(page, pagedFiles('<div data-node-id="long">' + Array.from({ length: 40 }, (_, index) => "LINE-" + index).join("<br>") + "</div>"));
  expect(await page.evaluate(() => window.paginationHost.layout().pages.length)).toBeGreaterThan(1);
  expect(await page.evaluate(() => window.paginationHost.layout().nodePages.long!.length)).toBeGreaterThan(1);
  expect(await page.frameLocator("iframe").locator('[data-node-id="long"]').count()).toBe(1);
});
test("section metadata order never reorders source or saved edits", async ({ page }) => {
  const files = pagedFiles('<section data-node-id="first"><p data-node-id="a">A</p></section><section data-node-id="second"><p data-node-id="b">B</p></section>');
  const metadata = JSON.parse(new TextDecoder().decode(files.get("metadata.json"))); metadata.sections = [{ sectionId: "second", rootNodeId: "second", breakType: "nextPage" }, { sectionId: "first", rootNodeId: "first" }];
  files.set("metadata.json", new TextEncoder().encode(JSON.stringify(metadata)));
  await mount(page, files, "edit"); await atStart(page, "a"); await page.keyboard.insertText("X");
  const state = await page.evaluate(() => ({ text: window.paginationHost.session!.getNode("root")!.text, html: window.paginationHost.session!.snapshot().html, errors: window.paginationHost.errors }));
  expect(state.text).toBe("XAB"); expect(state.html.indexOf('data-node-id="first"')).toBeLessThan(state.html.indexOf('data-node-id="second"')); expect(state.errors).toEqual([]);
});
test("deletion without target-range support crosses view gaps and preserves graphemes", async ({ page }) => {
  await mount(page, pagedFiles('<p data-node-id="a">AB漢󠄀😀</p><p data-node-id="b">NEXT</p>'), "edit");
  await page.frameLocator("iframe").locator('[data-node-id="a"]').evaluate((node) => {
    Object.defineProperty(InputEvent.prototype, "getTargetRanges", { configurable: true, value: undefined });
    const range = document.createRange(); range.selectNodeContents(node); range.collapse(false); const selected = document.getSelection()!; selected.removeAllRanges(); selected.addRange(range); (document.querySelector("[data-office-editor]") as HTMLElement).focus();
  });
  await page.keyboard.press("Backspace"); await page.keyboard.press("Backspace");
  expect(await page.evaluate(() => window.paginationHost.session!.getNode("a")!.text)).toBe("AB");
  await atStart(page, "b"); await page.keyboard.press("Backspace");
  expect(await page.evaluate(() => window.paginationHost.session!.getNode("a")!.text)).toBe("ABNEXT");
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
});
test("long table rows fragment while continued pages reserve repeated-header space", async ({ page }) => {
  const lines = Array.from({ length: 30 }, (_, index) => "CELL-" + index).join("<br>");
  await mount(page, pagedFiles('<table data-node-id="table"><tr data-node-id="header" data-word-repeat-header="true"><th data-node-id="h1">HEAD</th><th data-node-id="h2">HEAD2</th></tr><tr data-node-id="long-row"><td data-node-id="left">' + lines + '</td><td data-node-id="right">' + lines + '</td></tr><tr data-node-id="tail"><td data-node-id="tail-left">END</td><td data-node-id="tail-right">END2</td></tr></table>'));
  const report = await page.evaluate(() => window.paginationHost.layout()); expect(report.pages.length).toBeGreaterThan(2);
  await expect(page.frameLocator("iframe").locator(".sdk-repeat-header")).toHaveCount(report.pages.length - 1);
  const overlaps = await page.frameLocator("iframe").locator('[data-node-id="left"]').evaluate((node) => {
    const headerRects = [...document.querySelectorAll(".sdk-repeat-header")].map((header) => header.getBoundingClientRect()), walker = document.createTreeWalker(node, NodeFilter.SHOW_TEXT); let item; const bad: string[] = [];
    while ((item = walker.nextNode())) { if (!item.textContent) continue; const range = document.createRange(); range.selectNodeContents(item); for (const rect of range.getClientRects()) if (headerRects.some((header) => rect.top < header.bottom - 0.5 && rect.bottom > header.top + 0.5)) bad.push(item.textContent); }
    return bad;
  });
  expect(overlaps).toEqual([]);
});
test("mixed container inline text remains visible and authored copy excludes page furniture", async ({ page }) => {
  await mount(page, pagedFiles('<div data-node-id="mixed">' + Array.from({ length: 30 }, (_, index) => "DIRECT-" + index).join("<br>") + '<p data-node-id="end">END</p></div>'));
  expect(await page.evaluate(() => window.paginationHost.layout().pages.length)).toBeGreaterThan(2);
  const copy = await page.frameLocator("iframe").locator('[data-node-id="mixed"]').evaluate((node) => {
    const range = document.createRange(); range.selectNodeContents(node); const selected = document.getSelection()!; selected.removeAllRanges(); selected.addRange(range);
    const data = new Map<string, string>(), event = new Event("copy", { bubbles: true, cancelable: true });
    Object.defineProperty(event, "clipboardData", { value: { setData: (type: string, value: string) => data.set(type, value) } }); document.dispatchEvent(event);
    return { html: data.get("text/html") ?? "", text: data.get("text/plain") ?? "", prevented: event.defaultPrevented };
  });
  expect(copy.prevented).toBe(true); expect(copy.html).not.toMatch(/sdk-page|sdk-inline-flow|contenteditable|clip-path/); expect(copy.text).toContain("DIRECT-29"); expect(copy.text).toContain("END");
});
test("host interaction events work while CSP blocks executable document content", async ({ page }) => {
  await mount(page, pagedFiles('<p data-node-id="a">SAFE</p>'));
  await page.evaluate(() => {
    (window as unknown as { sourceExecuted: number; sourceViolations: number }).sourceExecuted = 0;
    (window as unknown as { sourceViolations: number }).sourceViolations = 0;
    const frame = document.querySelector("iframe")!.contentDocument!;
    frame.addEventListener("securitypolicyviolation", (event) => { if (event.violatedDirective.startsWith("script-src")) (window as unknown as { sourceViolations: number }).sourceViolations++; });
    const script = frame.createElement("script"); script.textContent = "parent.sourceExecuted++"; frame.body.append(script);
    const image = frame.createElement("img"); image.setAttribute("onerror", "parent.sourceExecuted++"); image.src = "invalid-source:"; frame.body.append(image);
  });
  await expect.poll(() => page.evaluate(() => (window as unknown as { sourceViolations: number }).sourceViolations)).toBeGreaterThanOrEqual(2);
  expect(await page.evaluate(() => (window as unknown as { sourceExecuted: number }).sourceExecuted)).toBe(0);
  expect(await page.frameLocator("iframe").locator('meta[http-equiv="Content-Security-Policy"]').getAttribute("content")).toContain("default-src 'none'");
});
test("overheight headers stay outside body text and report their clipping", async ({ page }) => {
  const files = pagedFiles('<section data-node-id="main" data-office-story="main"><section data-node-id="sec"><p data-node-id="body-text">BODY FIRST LINE</p></section></section><section data-node-id="head" data-office-story="header"><p data-node-id="header-a">HEADER 1</p><p data-node-id="header-b">HEADER 2</p><p data-node-id="header-c">HEADER 3</p></section>');
  const metadata = JSON.parse(new TextDecoder().decode(files.get("metadata.json"))); metadata.pageSetup = { ...metadata.pageSetup, height: 300, margins: { top: 72, bottom: 72, left: 24, right: 24 } };
  metadata.sections = [{ sectionId: "s", rootNodeId: "sec", headers: { default: "hdr" }, headerDistance: 36 }];
  metadata.word = { stories: [{ storyId: "main", kind: "main", rootNodeId: "main" }, { storyId: "hdr", kind: "header", rootNodeId: "head" }], styles: [], numbering: [], hyperlinks: [], bookmarks: [], fields: [], tabs: [], tables: [], revisions: [] };
  files.set("metadata.json", new TextEncoder().encode(JSON.stringify(metadata))); await mount(page, files);
  expect((await page.evaluate(() => window.paginationHost.layout().diagnostics)).some((item) => item.code === "ContentOverflow" && item.nodeId === "head")).toBe(true);
  const bounds = await page.frameLocator("iframe").locator("body").evaluate(() => ({ header: document.querySelector(".sdk-paper .sdk-page-story")!.getBoundingClientRect().bottom, body: document.querySelector('[data-node-id="body-text"]')!.getBoundingClientRect().top }));
  expect(bounds.header).toBeLessThanOrEqual(bounds.body + 0.5);
});
test("copy keeps nested block boundaries and source attributes cannot impersonate the editor", async ({ page }) => {
  const files = pagedFiles('<section data-node-id="nested"><p data-node-id="first">FIRST</p><p data-node-id="second">SECOND</p></section><table data-node-id="grid"><tr data-node-id="r"><td data-node-id="c">C</td><td data-node-id="d">D</td></tr></table>');
  files.set("content.html", new TextEncoder().encode(new TextDecoder().decode(files.get("content.html")).replace('data-node-id="root"', 'data-node-id="root" data-office-editor="source-data"')));
  await mount(page, files);
  const copied = await page.frameLocator("iframe").locator('[data-node-id="root"]').evaluate((node) => {
    const range = document.createRange(); range.selectNodeContents(node); document.getSelection()!.removeAllRanges(); document.getSelection()!.addRange(range);
    const values: Record<string, string> = {}, event = new Event("copy", { bubbles: true, cancelable: true }); Object.defineProperty(event, "clipboardData", { value: { setData: (type: string, value: string) => { values[type] = value; } } }); document.dispatchEvent(event);
    return { text: values["text/plain"], html: values["text/html"], prevented: event.defaultPrevented, editable: (node as HTMLElement).isContentEditable };
  });
  expect(copied.editable).toBe(false); expect(copied.prevented).toBe(true); expect(copied.text).toBe("FIRST\nSECOND\nC\tD"); expect(copied.html).not.toContain("sdk-page");
});
for (const mode of ["preview", "edit"] as const) test("partial field display copy stays within the selected range: " + mode, async ({ page }) => {
  await mount(page, fieldsFiles(), mode);
  const copied = await page.frameLocator("iframe").locator(".sdk-field-display").filter({ hasText: "Alpha" }).first().evaluate((node) => {
    const walker = document.createTreeWalker(node, NodeFilter.SHOW_TEXT), text = walker.nextNode()!;
    const range = document.createRange(); range.setStart(text, 0); range.setEnd(text, 3); document.getSelection()!.removeAllRanges(); document.getSelection()!.addRange(range);
    const values: Record<string, string> = {}, event = new Event("copy", { bubbles: true, cancelable: true }); Object.defineProperty(event, "clipboardData", { value: { setData: (type: string, value: string) => { values[type] = value; } } });
    node.dispatchEvent(event); return values;
  });
  expect(copied["text/plain"]).toBe("Alp"); expect(copied["text/html"]).not.toContain("Beta"); expect(copied["text/html"]).not.toMatch(/sdk-field|data-word-field|data-node-id/);
});
test("disposing a preview during settled invalidates its pending report", async ({ page }) => {
  await mount(page, pagedFiles('<p data-node-id="a">TEXT</p>'));
  const result = await page.evaluate(async () => {
    const view = window.paginationHost.view as import("@doffi/editor").DocumentPreview;
    const pending = view.settled(); await Promise.resolve(); view.dispose();
    return pending.then(() => "success", (error: { code: string }) => error.code);
  });
  expect(result).toBe("InvalidOperation");
});
test("odd-page section breaks create a blank sheet using the new geometry and restart printed numbering", async ({ page }) => {
  const files = pagedFiles('<section data-node-id="first"><p data-node-id="a">FIRST</p></section><section data-node-id="second"><p data-node-id="b">SECOND</p></section>');
  const metadata = JSON.parse(new TextDecoder().decode(files.get("metadata.json")));
  metadata.sections = [{ sectionId: "first", rootNodeId: "first" }, { sectionId: "second", rootNodeId: "second", breakType: "oddPage", pageNumberStart: 8, pageNumberFormat: "lowerRoman", pageSetup: { ...metadata.pageSetup, width: 240, height: 300 } }];
  files.set("metadata.json", new TextEncoder().encode(JSON.stringify(metadata))); await mount(page, files);
  const report = await page.evaluate(() => window.paginationHost.layout());
  expect(report.pages.map((item) => [item.blank, item.setup.width, item.setup.height, item.label])).toEqual([[false, 300, 240, "1"], [true, 240, 300, "2"], [false, 240, 300, "viii"]]);
  expect(report.nodePages.a).toEqual(["page-1"]); expect(report.nodePages.b).toEqual(["page-3"]);
});
test("unitless inherited Word line spacing works without Typed OM and preserves absolute overrides", async ({ page }) => {
  await page.addInitScript(() => { Object.defineProperty(HTMLElement.prototype, "computedStyleMap", { configurable: true, value: undefined }); });
  const files = pagedFiles('<p data-node-id="auto" data-word-line-rule="auto">AUTO</p><p data-node-id="absolute" data-word-line-rule="auto" style="line-height:30px">ABSOLUTE</p>');
  const metadata = JSON.parse(new TextDecoder().decode(files.get("metadata.json")));
  metadata.word = { stories: [{ storyId: "main", kind: "main", rootNodeId: "root" }], styles: [], numbering: [], hyperlinks: [], bookmarks: [], fields: [], tabs: [], tables: [], revisions: [] };
  files.set("metadata.json", new TextEncoder().encode(JSON.stringify(metadata)));
  files.set("content.html", new TextEncoder().encode(new TextDecoder().decode(files.get("content.html")).replace('data-office-kind="document"', 'data-office-kind="document" data-office-story="main"')));
  files.set("styles.css", new TextEncoder().encode('article{font-family:"Liberation Serif",serif;font-size:12pt}p{line-height:1.5;margin:0}'));
  await mount(page, files);
  const measured = await page.frameLocator("iframe").locator('[data-node-id="auto"]').evaluate((node) => {
    const probe = document.createElement("span"); probe.style.cssText = "position:absolute;display:block;line-height:normal;white-space:pre;font:inherit"; probe.style.lineHeight = "normal"; probe.textContent = "汉M"; node.append(probe);
    const normal = probe.getBoundingClientRect().height; probe.remove(); return { line: parseFloat(getComputedStyle(node).lineHeight), normal };
  });
  expect(measured.line).toBeCloseTo(measured.normal * 1.5, 1);
  await expect(page.frameLocator("iframe").locator('[data-node-id="absolute"]')).toHaveCSS("line-height", "30px");
});
for (const layout of ["pages", "continuous"] as const) test("editing preserves source content before, between and after sections: " + layout, async ({ page }) => {
  const files = pagedFiles('<p data-node-id="lead">L</p><section data-node-id="first"><p data-node-id="a">A</p></section><p data-node-id="middle">M</p><section data-node-id="second"><p data-node-id="b">B</p></section><p data-node-id="tail">T</p>');
  const metadata = JSON.parse(new TextDecoder().decode(files.get("metadata.json"))); metadata.sections = [{ sectionId: "first", rootNodeId: "first" }, { sectionId: "second", rootNodeId: "second", breakType: "nextPage" }];
  files.set("metadata.json", new TextEncoder().encode(JSON.stringify(metadata))); await mount(page, files, "edit", { layout });
  await atStart(page, "a"); await page.keyboard.insertText("X");
  const state = await page.evaluate(() => ({ text: window.paginationHost.session!.getNode("root")!.text, html: window.paginationHost.session!.snapshot().html, errors: window.paginationHost.errors }));
  expect(state.text).toBe("LXAMBT"); expect(state.errors).toEqual([]);
  expect(state.html).not.toMatch(/sdk-page|edit-\d|<div/);
  const structure = await page.evaluate(() => window.paginationHost.session!.getNode("root")!.childNodeIds);
  expect(structure).toEqual(["lead", "first", "middle", "second", "tail"]);
  await page.keyboard.press("Control+z"); expect(await page.evaluate(() => window.paginationHost.session!.getNode("root")!.text)).toBe("LAMBT");
});
test("formatted table whitespace does not become extra plain-text rows or cells", async ({ page }) => {
  await mount(page, pagedFiles('<table data-node-id="table">\n <tbody>\n <tr data-node-id="row">\n  <td data-node-id="c">C</td>\n  <td data-node-id="d">D</td>\n </tr>\n </tbody>\n</table>'));
  const copied = await page.frameLocator("iframe").locator('[data-node-id="table"]').evaluate((node) => {
    const range = document.createRange(); range.selectNode(node); document.getSelection()!.removeAllRanges(); document.getSelection()!.addRange(range);
    const values: Record<string, string> = {}, event = new Event("copy", { bubbles: true, cancelable: true }); Object.defineProperty(event, "clipboardData", { value: { setData: (type: string, value: string) => { values[type] = value; } } }); node.dispatchEvent(event); return values;
  });
  expect(copied["text/plain"]).toBe("C\tD");
});
