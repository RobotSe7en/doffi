import { openSdkHost } from './sdk-host.js';
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { expect, test } from "@playwright/test";
import type { Page } from "@playwright/test";
import { build } from "vite";
import { importDocx } from "@doffi/office";
import { createAssetDescriptor } from "@doffi/core";
import type { DocumentPackage } from "@doffi/core";
import type { TextEditorHost } from "./fixtures/text-editor-host.js";
declare global { interface Window { textEditorHost: TextEditorHost } }
let code: string;
const report = new Map(["manifest.json", "content.html", "styles.css", "metadata.json", "annotations.json"].map((path) => [path,
  new Uint8Array(readFileSync(new URL("../../examples/report/" + path, import.meta.url))),
]));
function fixture(html = '<p data-node-id="a">甲甲甲</p><p data-node-id="b">第二段</p>'): DocumentPackage {
  const files = new Map(report); files.set("annotations.json", new TextEncoder().encode("[]"));
  files.set("content.html", new TextEncoder().encode(new TextDecoder().decode(files.get("content.html")).replace(/<article[\s\S]*<\/article>/,
    '<article data-node-id="root" data-office-kind="document">' + html + "</article>"))); return files;
}
async function mount(page: Page, files = fixture(), beforeReadyText?: string): Promise<void> {
  await openSdkHost(page); await page.addScriptTag({ type: "module", content: code });
  await page.waitForFunction(() => !!window.textEditorHost);
  await page.evaluate(({ entries, text }) => window.textEditorHost.mount(entries, text), {
    entries: [...files].map(([path, value]) => [path, [...value]] as [string, number[]]), text: beforeReadyText,
  });
}
async function selection(page: Page, fromId: string, from: number, toId = fromId, to = from): Promise<void> {
  await page.frameLocator('iframe[title="正文编辑"]').locator("body").evaluate((_, value) => {
    const find = (id: string, offset: number): [Node, number] => {
      const element = document.querySelector('[data-node-id="' + CSS.escape(id) + '"]')!, walker = document.createTreeWalker(element, NodeFilter.SHOW_TEXT);
      let node: Node | null;
      while ((node = walker.nextNode())) { if (offset <= node.textContent!.length) return [node, offset]; offset -= node.textContent!.length; }
      return [element, element.childNodes.length];
    };
    const range = document.createRange(); range.setStart(...find(value.fromId, value.from)); range.setEnd(...find(value.toId, value.to));
    const selected = document.getSelection()!; selected.removeAllRanges(); selected.addRange(range);
    (document.querySelector('[data-office-editor]') as HTMLElement).focus(); document.dispatchEvent(new Event("selectionchange"));
  }, { fromId, from, toId, to });
}
test.beforeAll(async () => {
  const result = await build({ configFile: false, logLevel: "silent", build: { write: false, target: "es2022", lib: {
    entry: fileURLToPath(new URL("./fixtures/text-editor-host.ts", import.meta.url)), formats: ["es"],
  } } });
  const output = Array.isArray(result) ? result[0]!.output : "output" in result ? result.output : [];
  code = output.flatMap((item) => item.type === "chunk" ? [item.code] : []).join("\n");
});

test("direct typing preserves source node identity and uses core history", async ({ page }) => {
  await mount(page); await selection(page, "a", 1); await page.keyboard.insertText("新😀");
  await expect.poll(() => page.evaluate(() => window.textEditorHost.session!.getNode("a")!.text)).toBe("甲新😀甲甲");
  expect(await page.evaluate(() => window.textEditorHost.session!.history().undoSteps)).toBe(1);
  await page.keyboard.press("Control+z");
  await expect.poll(() => page.evaluate(() => window.textEditorHost.session!.getNode("a")!.text)).toBe("甲甲甲");
  await page.keyboard.press("Control+Shift+z");
  await expect.poll(() => page.evaluate(() => window.textEditorHost.session!.getNode("a")!.text)).toBe("甲新😀甲甲");
  expect(await page.evaluate(() => window.textEditorHost.errors)).toEqual([]);
});

test("an empty document offers a derived paragraph and commits it only after input", async ({ page }) => {
  await mount(page, fixture(""));
  expect(await page.evaluate(() => window.textEditorHost.session!.getNode("root")!.childNodeIds)).toEqual([]);
  await page.frameLocator('iframe[title="正文编辑"]').getByRole("textbox", { name: "文档正文" }).focus();
  await page.keyboard.insertText("空文档输入");
  await expect.poll(() => page.evaluate(() => window.textEditorHost.session!.getNode("root")!.text)).toBe("空文档输入");
  expect(await page.evaluate(() => window.textEditorHost.session!.history().undoSteps)).toBe(1);
});

test("a host commit before ready is visible and accepts input at the current revision", async ({ page }) => {
  await mount(page, fixture(), "宿主先更新");
  await expect(page.frameLocator('iframe[title="正文编辑"]').locator('[data-node-id="a"]')).toHaveText("宿主先更新甲甲甲");
  await selection(page, "a", 0); await page.keyboard.insertText("人工后输入");
  expect(await page.evaluate(() => window.textEditorHost.session!.getNode("a")!.text)).toBe("人工后输入宿主先更新甲甲甲");
  expect(await page.evaluate(() => window.textEditorHost.errors)).toEqual([]);
});

test("an authorized image replacement creates a current blob URL without changing package paths", async ({ page }) => {
  const files = new Map(fixture('<p data-node-id="a"><img data-node-id="image" src="assets/a.png" width="10" height="20"></p>'));
  const bytes = new Uint8Array(readFileSync(new URL("../../examples/fidelity/assets/sample.png", import.meta.url)));
  const manifest = JSON.parse(new TextDecoder().decode(files.get("manifest.json")!));
  manifest.assets = ["assets/a.png", "assets/b.png"].map((path) => createAssetDescriptor(path, "image/png", bytes));
  files.set("manifest.json", new TextEncoder().encode(JSON.stringify(manifest))); files.set("assets/a.png", bytes); files.set("assets/b.png", bytes);
  await mount(page, files);
  const image = page.frameLocator('iframe[title="正文编辑"]').locator('[data-node-id="image"]');
  const previous = await image.getAttribute("src");
  await page.evaluate(() => {
    const s = window.textEditorHost.session!;
    s.commit(s.prepare({ protocolVersion: 1, documentId: s.snapshot().manifest.documentId, baseRevision: 0,
      operations: [{ op: "replaceAsset", nodeId: "image", assetPath: "assets/b.png" }] }));
  });
  await expect(image).toHaveAttribute("src", /^blob:/);
  expect(await image.getAttribute("src")).not.toBe(previous);
  await expect.poll(() => image.evaluate((node) => (node as HTMLImageElement).naturalWidth)).toBeGreaterThan(0);
  expect(await page.evaluate(() => window.textEditorHost.session!.getNode("image")!.attributes.src)).toBe("assets/b.png");
  expect(await page.evaluate(() => window.textEditorHost.errors)).toEqual([]);
});

test("repeated text anchors follow the actual insertion point rather than a guessed string diff", async ({ page }) => {
  const files = new Map(fixture()); files.set("annotations.json", new TextEncoder().encode(JSON.stringify([{
    annotationId: "note", documentId: "report-demo", baseRevision: 0, status: "open", instruction: "middle",
    anchor: { kind: "text", nodeId: "a", from: 1, to: 2, quote: "甲" },
  }])));
  await mount(page, files); await selection(page, "a", 1); await page.keyboard.insertText("甲");
  await expect.poll(() => page.evaluate(() => window.textEditorHost.session!.getAnnotations()[0]!.anchor)).toEqual({ kind: "text", nodeId: "a", from: 2, to: 3, quote: "甲" });
});

test("a known paragraph split retains annotations on the untouched prefix", async ({ page }) => {
  const files = new Map(fixture('<p data-node-id="a">甲乙丙</p>'));
  files.set("annotations.json", new TextEncoder().encode(JSON.stringify([{
    annotationId: "note", documentId: "report-demo", baseRevision: 0, status: "open", instruction: "prefix",
    anchor: { kind: "text", nodeId: "a", from: 0, to: 1, quote: "甲" },
  }])));
  await mount(page, files); await selection(page, "a", 2); await page.keyboard.press("Enter");
  expect(await page.evaluate(() => window.textEditorHost.session!.getAnnotations()[0]!.status)).toBe("open");
  expect(await page.evaluate(() => window.textEditorHost.session!.getAnnotations()[0]!.anchor)).toEqual({ kind: "text", nodeId: "a", from: 0, to: 1, quote: "甲" });
});

test("native deletion uses its target range and joins paragraphs through core history", async ({ page }) => {
  await mount(page); await selection(page, "a", 2); await page.keyboard.press("Backspace");
  expect(await page.evaluate(() => window.textEditorHost.session!.getNode("a")!.text)).toBe("甲甲");
  await page.keyboard.press("Delete");
  expect(await page.evaluate(() => window.textEditorHost.session!.getNode("a")!.text)).toBe("甲");
  await selection(page, "b", 0); await page.keyboard.press("Backspace");
  expect(await page.evaluate(() => window.textEditorHost.session!.getNode("a")!.text)).toBe("甲第二段");
  await page.keyboard.press("Control+z");
  expect(await page.evaluate(() => window.textEditorHost.session!.getNode("b")!.text)).toBe("第二段");
  expect(await page.evaluate(() => window.textEditorHost.errors)).toEqual([]);
});

test("splitting a font-substituted run preserves authored fonts rather than measured presentation", async ({ page }) => {
  await mount(page, fixture('<p data-node-id="a"><span data-node-id="s" style="font-family:UnavailableSourceFace;color:#123456">甲乙丙</span></p>'));
  await selection(page, "s", 1); await page.keyboard.press("Enter"); await page.keyboard.insertText("新");
  const html = await page.evaluate(() => window.textEditorHost.session!.snapshot().html);
  expect(html.match(/font-family:UnavailableSourceFace/g)).toHaveLength(2);
  expect(html).not.toMatch(/Noto|sdk-number-marker|contenteditable|data-office-editor|role="textbox"/);
  expect(await page.evaluate(() => window.textEditorHost.errors)).toEqual([]);
});

test("editing preserves imported numbering, tabs, grid and untouched source attributes", async ({ page }) => {
  for (const name of ["01-text-numbering.docx", "02-table-merges.docx", "06-links-tabs.docx"]) {
    const path = new URL("../../examples/docx/" + name, import.meta.url);
    const files = importDocx(readFileSync(path), { documentId: "projection" });
    await mount(page, files);
    const before = await page.evaluate(() => {
      const s = window.textEditorHost.session!, source = new DOMParser().parseFromString(s.snapshot().html, "text/html");
      const paragraph = [...source.querySelectorAll("p")].find((node) => node.textContent?.trim())!;
      return { id: paragraph.getAttribute("data-node-id")!, nodes: [...source.querySelectorAll("[data-node-id]")].map((node) => ({ id: node.getAttribute("data-node-id")!, html: node.outerHTML })) };
    });
    await selection(page, before.id, 0); await page.keyboard.insertText("新增");
    const after = await page.evaluate(() => window.textEditorHost.session!.snapshot().html);
    expect(after).toContain("新增"); expect(after).not.toMatch(/sdk-number-marker|contenteditable|data-office-editor|blob:/);
    const untouched = before.nodes.filter((node) => !node.html.includes('data-node-id="' + before.id + '"') && node.id !== before.id && !before.nodes.find((item) => item.id === before.id)!.html.includes('data-node-id="' + node.id + '"'));
    // Core may assign an identity to a previously implicit table wrapper.
    const structural = (html: string): string => html.replace(/<(tbody|thead|tfoot) data-node-id="edit-[^"]+"/g, "<$1");
    for (const node of untouched) expect(structural((await page.evaluate((id) => window.textEditorHost.session!.getNode(id)?.html, node.id))!)).toBe(node.html);
    expect(await page.evaluate(() => window.textEditorHost.errors)).toEqual([]);
  }
});

test("Enter splits paragraphs, cross-paragraph replacement and table input remain serializable", async ({ page }) => {
  await mount(page); await selection(page, "a", 1); await page.keyboard.press("Enter");
  await expect.poll(() => page.evaluate(() => window.textEditorHost.session!.getNode("root")!.childNodeIds.length)).toBe(3);
  expect(await page.evaluate(() => window.textEditorHost.errors)).toEqual([]);
  await mount(page); await selection(page, "a", 1, "b", 1); await page.keyboard.insertText("替换");
  await expect.poll(() => page.evaluate(() => window.textEditorHost.session!.getNode("a")!.text)).toBe("甲替换二段");
  await mount(page, fixture('<table data-node-id="t"><tr data-node-id="r"><td data-node-id="c">单元格</td><td data-node-id="d">另一格</td></tr></table>'));
  await selection(page, "c", 1); await page.keyboard.insertText("中文");
  await expect.poll(() => page.evaluate(() => window.textEditorHost.session!.getNode("c")!.text)).toBe("单中文元格");
  expect(await page.evaluate(() => window.textEditorHost.errors)).toEqual([]);
});

for (const offset of [1, 3]) {
  test("typing after Enter stays in the new paragraph: split at " + offset, async ({ page }) => {
    await mount(page); await selection(page, "a", offset); await page.keyboard.press("Enter"); await page.keyboard.insertText("新段");
    const paragraphs = await page.evaluate(() => {
      const s = window.textEditorHost.session!; return s.getNode("root")!.childNodeIds.map((id) => s.getNode(id)!.text);
    });
    expect(paragraphs).toEqual(["甲".repeat(offset), "新段" + "甲".repeat(3 - offset), "第二段"]);
    expect(await page.evaluate(() => window.textEditorHost.errors)).toEqual([]);
  });
}

test("formatting preserves plain text and identities while sharing one undo step", async ({ page }) => {
  await mount(page); await selection(page, "a", 1, "a", 2); await page.getByRole("button", { name: "加粗", exact: true }).click();
  await expect.poll(() => page.evaluate(() => window.textEditorHost.session!.getNode("a")!.html)).toContain("font-weight:700");
  expect(await page.evaluate(() => window.textEditorHost.session!.getNode("a")!.text)).toBe("甲甲甲");
  expect(await page.evaluate(() => window.textEditorHost.session!.history().undoSteps)).toBe(1);
  await page.getByRole("button", { name: "撤销输入" }).click();
  expect(await page.evaluate(() => window.textEditorHost.session!.getNode("a")!.html)).not.toContain("font-weight");
});

test("real Chromium composition commits one Chinese input batch", async ({ page }) => {
  await mount(page); await selection(page, "a", 1);
  const cdp = await page.context().newCDPSession(page);
  await cdp.send("Input.imeSetComposition", { text: "中", selectionStart: 1, selectionEnd: 1 });
  expect(await page.evaluate(() => window.textEditorHost.session!.snapshot().manifest.revision)).toBe(0);
  await cdp.send("Input.imeSetComposition", { text: "中文", selectionStart: 2, selectionEnd: 2 });
  expect(await page.evaluate(() => window.textEditorHost.session!.snapshot().manifest.revision)).toBe(0);
  await cdp.send("Input.insertText", { text: "中文" });
  await expect.poll(() => page.evaluate(() => window.textEditorHost.session!.getNode("a")!.text)).toBe("甲中文甲甲");
  expect(await page.evaluate(() => window.textEditorHost.session!.history().undoSteps)).toBe(1);
  expect(await page.evaluate(() => window.textEditorHost.errors)).toEqual([]);
});

test("a collapsed character format applies to IME text in its first and only undo step", async ({ page }) => {
  await mount(page); await selection(page, "a", 1);
  await page.getByRole("button", { name: "加粗", exact: true }).click();
  expect(await page.evaluate(() => window.textEditorHost.session!.snapshot().manifest.revision)).toBe(0);
  const cdp = await page.context().newCDPSession(page);
  await cdp.send("Input.imeSetComposition", { text: "中文", selectionStart: 2, selectionEnd: 2 });
  await cdp.send("Input.insertText", { text: "中文" });
  expect(await page.evaluate(() => window.textEditorHost.session!.getNode("a")!.text)).toBe("甲中文甲甲");
  expect(await page.evaluate(() => window.textEditorHost.session!.getNode("a")!.html)).toMatch(/font-weight:700[^>]*>中文/);
  expect(await page.evaluate(() => window.textEditorHost.session!.history().undoSteps)).toBe(1);
  await page.keyboard.press("Control+z");
  expect(await page.evaluate(() => window.textEditorHost.session!.getNode("a")!.text)).toBe("甲甲甲");
  expect(await page.evaluate(() => window.textEditorHost.errors)).toEqual([]);
});

test("external changes during composition reject the pending input and restore the authoritative source", async ({ page }) => {
  await mount(page); await selection(page, "a", 1);
  const cdp = await page.context().newCDPSession(page);
  await cdp.send("Input.imeSetComposition", { text: "中文", selectionStart: 2, selectionEnd: 2 });
  await page.evaluate(() => {
    const s = window.textEditorHost.session!;
    s.commit(s.prepare({ protocolVersion: 1, documentId: "report-demo", baseRevision: 0,
      operations: [{ op: "setStyle", nodeId: "b", property: "color", value: "blue" }] }));
  });
  await cdp.send("Input.insertText", { text: "中文" });
  await expect(page.getByRole("alert")).toContainText("输入期间文档已变化");
  expect(await page.evaluate(() => window.textEditorHost.session!.getNode("a")!.text)).toBe("甲甲甲");
  expect(await page.evaluate(() => window.textEditorHost.session!.snapshot().manifest.revision)).toBe(1);
  await expect(page.frameLocator('iframe[title="正文编辑"]').locator('[data-node-id="a"]')).toHaveText("甲甲甲");
});

test("supported HTML paste is validated atomically and executable or external content never enters the editor", async ({ page }) => {
  await mount(page); await selection(page, "a", 1);
  const external: string[] = []; page.on("request", (request) => { if (request.url().startsWith("https://example.com")) external.push(request.url()); });
  const paste = async (html: string): Promise<void> => {
    await page.frameLocator('iframe[title="正文编辑"]').getByRole("textbox", { name: "文档正文" }).evaluate((element, html) => {
      const data = new DataTransfer(); data.setData("text/html", html);
      element.dispatchEvent(new ClipboardEvent("paste", { clipboardData: data, bubbles: true, cancelable: true }));
    }, html);
  };
  await paste('<span style="color:#123456">粘贴中文</span>');
  await expect.poll(() => page.evaluate(() => window.textEditorHost.session!.getNode("root")!.text)).toContain("粘贴中文");
  const before = await page.evaluate(() => window.textEditorHost.session!.snapshot());
  for (const html of ['<script>bad()</script>', '<img src="https://example.com/private" onerror="bad()">']) {
    await paste(html); await expect(page.getByRole("alert")).toBeVisible();
    expect(await page.evaluate(() => window.textEditorHost.session!.snapshot())).toEqual(before);
  }
  expect(external).toEqual([]);
});

test("native copy, paste and cut transfer source formatting rather than browser computed styles", async ({ page }) => {
  await page.context().grantPermissions(["clipboard-read", "clipboard-write"]);
  await mount(page, fixture('<p data-node-id="a"><span data-node-id="s" style="font-family:UnavailableSourceFace;color:#123456">甲乙丙</span></p><p data-node-id="b">另一段</p>'));
  await selection(page, "s", 0, "s", 2); await page.keyboard.press("Control+c");
  const html = await page.evaluate(async () => {
    const items = await navigator.clipboard.read(); return (await items[0]!.getType("text/html")).text();
  });
  expect(html).toContain("UnavailableSourceFace"); expect(html).not.toMatch(/Noto|data-node-id|webkit/);
  await selection(page, "b", 3); await page.keyboard.press("Control+v");
  await expect.poll(() => page.evaluate(() => window.textEditorHost.session!.getNode("root")!.text)).toBe("甲乙丙另一段甲乙");
  expect(await page.evaluate(() => window.textEditorHost.session!.getNode("b")!.text)).toBe("另一段甲乙");
  expect(await page.evaluate(() => window.textEditorHost.session!.getNode("root")!.childNodeIds)).toHaveLength(2);
  expect(await page.evaluate(() => window.textEditorHost.errors)).toEqual([]);
  await selection(page, "s", 0, "s", 1); await page.keyboard.press("Control+x");
  expect(await page.evaluate(() => window.textEditorHost.session!.getNode("a")!.text)).toBe("乙丙");
  await page.keyboard.press("Control+z");
  expect(await page.evaluate(() => window.textEditorHost.session!.getNode("a")!.text)).toBe("甲乙丙");
});

for (const name of ["01-text-numbering.docx", "02-table-merges.docx"]) {
  test("copy and paste omit derived numbering and table grid children: " + name, async ({ page }) => {
    await page.context().grantPermissions(["clipboard-read", "clipboard-write"]);
    await mount(page, importDocx(readFileSync(new URL("../../examples/docx/" + name, import.meta.url)), { documentId: "copy-derived" }));
    const target = await page.evaluate(() => {
      const s = window.textEditorHost.session!, root = s.snapshot().metadata.word!.stories.find((story) => story.kind === "main")!.rootNodeId;
      s.commit(s.prepare({ protocolVersion: 1, documentId: s.snapshot().manifest.documentId, baseRevision: 0,
        operations: [{ op: "insertFragment", nodeId: root, position: "end", html: "<p>粘贴到此</p>" }] }));
      return s.getNode(root)!.childNodeIds.at(-1)!;
    });
    const canvas = page.frameLocator('iframe[title="正文编辑"]');
    const selector = name.startsWith("01") ? "[data-word-numbering]" : "table";
    const copied = await canvas.locator(selector).first().evaluate((node) => {
      (document.querySelector("[data-office-editor]") as HTMLElement).focus();
      const range = document.createRange(); range.selectNode(node);
      const selected = document.getSelection()!; selected.removeAllRanges(); selected.addRange(range);
      return node.getAttribute("data-node-id")!;
    });
    const expected = await page.evaluate((id) => window.textEditorHost.session!.getNode(id)!.text, copied);
    await page.keyboard.press("Control+c");
    const clipboard = await page.evaluate(async () => {
      const items = await navigator.clipboard.read(); return {
        html: await (await items[0]!.getType("text/html")).text(), text: await navigator.clipboard.readText(),
      };
    });
    expect(clipboard.html).not.toMatch(/sdk-number-marker|colgroup|contenteditable|data-node-id/);
    expect(name.startsWith("02") ? clipboard.text.replace(/[\t\n]/g, "") : clipboard.text).toBe(expected);
    await selection(page, target, 4); await page.keyboard.press("Control+v");
    expect(await page.evaluate(() => window.textEditorHost.errors)).toEqual([]);
    const text = await page.evaluate(() => window.textEditorHost.session!.snapshot().html);
    expect(text).not.toMatch(/sdk-number-marker|colgroup|contenteditable/);
    expect(await canvas.locator(selector).count()).toBeGreaterThan(1);
  });
}

test("table Tab navigation and paragraph/character formats use the same session", async ({ page }) => {
  await mount(page, fixture('<table data-node-id="t"><tr data-node-id="r"><td data-node-id="c">第一格</td><td data-node-id="d">第二格</td></tr></table>'));
  await selection(page, "c", 0); await page.keyboard.press("Tab"); await page.keyboard.insertText("新增");
  await expect.poll(() => page.evaluate(() => window.textEditorHost.session!.getNode("d")!.text)).toBe("新增第二格");
  await mount(page); await selection(page, "a", 0, "a", 2); await page.getByLabel("字号").selectOption("16pt");
  await expect.poll(() => page.evaluate(() => window.textEditorHost.session!.getNode("a")!.html)).toContain("font-size:16pt");
  await page.getByRole("button", { name: "居中", exact: true }).click();
  await expect.poll(() => page.evaluate(() => window.textEditorHost.session!.getNode("a")!.attributes.style)).toContain("text-align:center");
});

test("disposing during composition cancels the temporary view and leaves the host session usable", async ({ page }) => {
  await mount(page); await selection(page, "a", 1);
  const cdp = await page.context().newCDPSession(page);
  await cdp.send("Input.imeSetComposition", { text: "中文", selectionStart: 2, selectionEnd: 2 });
  await page.evaluate(() => window.textEditorHost.view!.dispose());
  await expect(page.locator("iframe")).toHaveCount(0);
  expect(await page.evaluate(() => ({ revision: window.textEditorHost.session!.snapshot().manifest.revision, steps: window.textEditorHost.session!.history().undoSteps }))).toEqual({ revision: 0, steps: 0 });
});

test("input cannot split a UTF-16 surrogate pair or detach a variation selector", async ({ page }) => {
  for (const offset of [2, 5]) {
    await mount(page, fixture('<p data-node-id="a">甲😀𠀀\u{e0100}乙</p>'));
    await selection(page, "a", offset); await page.keyboard.insertText("新");
    const text = await page.evaluate(() => window.textEditorHost.session!.getNode("a")!.text);
    expect(text).toContain("😀"); expect(text).toContain("𠀀\u{e0100}");
    expect([...text].some((character) => character.length === 1 && character.charCodeAt(0) >= 0xd800 && character.charCodeAt(0) <= 0xdfff)).toBe(false);
  }
});

test("an incrementally typed variation selector stays in its base character's source font scope", async ({ page }) => {
  await mount(page, fixture('<p data-node-id="a"><span data-node-id="s" data-word-run="true" data-word-script="eastAsia" style="font-family:CJKFace;--word-font-eastAsia:CJKFace;--word-font-hAnsi:OtherFace">漢</span></p>'));
  await selection(page, "s", 1); await page.keyboard.insertText("\u{e0100}");
  await expect.poll(() => page.evaluate(() => window.textEditorHost.session!.getNode("s")!.text)).toBe("漢\u{e0100}");
  expect(await page.evaluate(() => window.textEditorHost.session!.getNode("s")!.childNodeIds)).toEqual([]);
  expect(await page.evaluate(() => window.textEditorHost.errors)).toEqual([]);
});

test("replacing the complete editable story retains document and section identities", async ({ page }) => {
  for (const files of [fixture(), importDocx(readFileSync(new URL("../../examples/docx/04-styles-sections.docx", import.meta.url)), { documentId: "whole-story" })]) {
    await mount(page, files);
    const sections = await page.evaluate(() => window.textEditorHost.session!.snapshot().metadata.sections!.map((section) => section.rootNodeId));
    await page.frameLocator('iframe[title="正文编辑"]').getByRole("textbox", { name: "文档正文" }).focus();
    await page.keyboard.press("Control+a"); await page.keyboard.insertText("全篇替换");
    await expect.poll(() => page.evaluate(() => window.textEditorHost.session!.snapshot().html)).toContain("全篇替换");
    expect(await page.evaluate((sections) => sections.every((id) => !!window.textEditorHost.session!.getNode(id)), sections)).toBe(true);
    expect(await page.evaluate(() => window.textEditorHost.errors)).toEqual([]);
  }
});

test("React workbench edits source, returns to preview and saves a reopenable package", async ({ page }, info) => {
  await page.addInitScript(() => localStorage.setItem("doffi:editing", "preview"));
  await page.goto("/"); await page.getByRole("button", { name: "打开报告协作样例" }).click();
  await page.getByRole("button", { name: "编辑正文", exact: true }).click();
  await expect(page.frameLocator('iframe[title="正文编辑"]').getByRole("textbox", { name: "文档正文" })).toBeVisible();
  await selection(page, "n-summary", 0); await page.keyboard.insertText("人工编辑：");
  await page.getByRole("button", { name: "返回预览与批注", exact: true }).click();
  await expect(page.frameLocator('iframe[title="文档只读预览"]').locator('[data-node-id="n-summary"]')).toContainText("人工编辑：");
  const download = page.waitForEvent("download"); await page.getByRole("button", { name: "下载文档包" }).click();
  const path = info.outputPath("edited.zip"); await (await download).saveAs(path);
  await page.getByLabel("打开文件").setInputFiles(path);
  await expect(page.frameLocator('iframe[title="文档只读预览"]').locator('[data-node-id="n-summary"]')).toContainText("人工编辑：");
});

for (const file of ["07-fields.docx", "08-comments-revisions.docx", "09-chinese-fonts.docx"]) {
  test("editing imported DOCX preserves native semantics and font roles: " + file, async ({ page }) => {
    const files = importDocx(readFileSync(new URL("../../examples/docx/" + file, import.meta.url)), { documentId: "rich-docx" });
    await mount(page, files);
    const before = await page.evaluate(() => window.textEditorHost.session!.snapshot().metadata);
    const id = await page.evaluate(() => {
      const s = window.textEditorHost.session!;
      return [...s.snapshot().html.matchAll(/data-node-id="([^"]+)"/g)].map((match) => match[1]!).find((id) => s.getNode(id)!.tagName === "p" && s.getNode(id)!.text.trim())!;
    });
    await selection(page, id, 0); await page.keyboard.insertText("中文ABC𠀀\u{e0100}");
    await expect.poll(() => page.evaluate((id) => window.textEditorHost.session!.getNode(id)!.text, id)).toContain("中文ABC𠀀\u{e0100}");
    const after = await page.evaluate(() => window.textEditorHost.session!.snapshot().metadata);
    expect(after.word!.stories).toEqual(before.word!.stories); expect(after.word!.revisions).toEqual(before.word!.revisions);
    expect(after.word!.styles).toEqual(before.word!.styles);
    expect(await page.evaluate(() => window.textEditorHost.errors)).toEqual([]);
  });
}
