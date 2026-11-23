import { encodeTestPackage } from "./package-transport.js";
import { openSdkHost } from "./sdk-host.js";
import { openContentFixture as openDocument } from "../content-fixture.js";
import { expect, test } from "@playwright/test";
import type { Page } from "@playwright/test";
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { build } from "vite";
import {} from "@doffi/core";
import type { DocumentPackage, Rectangle } from "@doffi/core";
import { importDocx, readDocumentZip } from "@doffi/office";
import type {} from "./fixtures/interaction-host.js";
import { nestedFieldFixture } from "../../packages/office/tests/regression-fixtures.js";

let hostCode: string;
const report = new Map(
  ["manifest.json", "content.html", "styles.css", "metadata.json", "annotations.json"].map(
    (path) => [
      path,
      new Uint8Array(readFileSync(new URL("../../examples/report/" + path, import.meta.url))),
    ],
  ),
);
function fixture(html?: string, rootTag = "article"): DocumentPackage {
  const files = new Map(report);
  files.set("annotations.json", new TextEncoder().encode("[]"));
  if (html)
    files.set(
      "content.html",
      new TextEncoder().encode(
        new TextDecoder()
          .decode(files.get("content.html"))
          .replace(
            /<article[\s\S]*<\/article>/,
            "<" +
              rootTag +
              ' data-node-id="root" data-office-kind="document">' +
              html +
              "</" +
              rootTag +
              ">",
          ),
      ),
    );
  return files;
}
const entries = encodeTestPackage;
async function mount(page: Page, files = fixture(), readOnly = false): Promise<void> {
  await openSdkHost(page);
  await page.addScriptTag({ type: "module", content: hostCode });
  await page.waitForFunction(() => !!window.interactionHost);
  await page.evaluate(({ input, readOnly }) => window.interactionHost.mount(input, readOnly), {
    input: entries(files),
    readOnly,
  });
}
async function select(
  page: Page,
  startId: string,
  from: number,
  endId = startId,
  to?: number,
): Promise<void> {
  await page
    .frameLocator("iframe")
    .locator("body")
    .evaluate(
      (_, input) => {
        const start = document.querySelector('[data-node-id="' + input.startId + '"]')!;
        const end = document.querySelector('[data-node-id="' + input.endId + '"]')!;
        const range = document.createRange();
        range.setStart(start.firstChild!, input.from);
        range.setEnd(end.firstChild!, input.to ?? end.firstChild!.textContent!.length);
        const selected = window.getSelection()!;
        selected.removeAllRanges();
        selected.addRange(range);
        document.dispatchEvent(new Event("selectionchange"));
      },
      { startId, from, endId, to },
    );
}
async function drag(page: Page, rect: Rectangle): Promise<void> {
  const frame = page.locator("iframe"),
    box = (await frame.boundingBox())!;
  const displayScale = await frame.evaluate(
    (node) => node.getBoundingClientRect().width / (node as HTMLElement).offsetWidth,
  );
  const surface = await page
    .frameLocator("iframe")
    .locator('[data-node-id="root"]')
    .evaluate((node) => {
      const box = node.getBoundingClientRect();
      return { x: box.left, y: box.top };
    });
  const x = box.x + (surface.x + (rect.x * 96) / 72) * displayScale;
  const y = box.y + (surface.y + (rect.y * 96) / 72) * displayScale;
  await page.mouse.move(x, y);
  await page.mouse.down();
  await page.mouse.move(
    x + ((rect.width * 96) / 72) * displayScale,
    y + ((rect.height * 96) / 72) * displayScale,
    { steps: 5 },
  );
  await page.mouse.up();
}

test.beforeAll(async () => {
  const result = await build({
    configFile: false,
    logLevel: "silent",
    build: {
      write: false,
      target: "es2022",
      lib: {
        entry: fileURLToPath(new URL("./fixtures/interaction-host.ts", import.meta.url)),
        formats: ["es"],
      },
    },
  });
  const output = Array.isArray(result)
    ? result[0]!.output
    : "output" in result
      ? result.output
      : [];
  hostCode = output.flatMap((item) => (item.type === "chunk" ? [item.code] : [])).join("\n");
});

for (const rootTag of ["article", "section", "div"]) {
  for (const readOnly of [false, true]) {
    test(`${readOnly ? "read-only" : "interactive"} preview accepts a ${rootTag} document root`, async ({
      page,
    }) => {
      await mount(page, fixture('<p data-node-id="a">合法正文</p>', rootTag), readOnly);
      const frame = page.frameLocator("iframe");
      await expect(frame.locator('body > [data-office-kind="document"]')).toHaveJSProperty(
        "tagName",
        rootTag.toUpperCase(),
      );
      await expect(frame.locator('[data-node-id="a"]')).toBeVisible();
      if (!readOnly) {
        await select(page, "a", 0);
        expect(await page.evaluate(() => window.interactionHost.view!.selection()?.anchor)).toEqual(
          {
            kind: "text",
            nodeId: "a",
            from: 0,
            to: 4,
            quote: "合法正文",
          },
        );
      }
      expect(await page.evaluate(() => window.interactionHost.errors)).toEqual([]);
    });
  }
}

test("text selection, annotation editing, resolution and history use one session", async ({
  page,
}) => {
  await mount(page);
  await select(page, "n-summary", 2, "n-summary", 8);
  await expect(page.getByLabel("当前选择")).toHaveText("计划于第三季");
  expect(
    await page.evaluate(() => window.interactionHost.session!.snapshot().manifest.revision),
  ).toBe(0);
  await page.getByLabel("新批注内容").fill("改为第四季度");
  await page.getByRole("button", { name: "添加批注", exact: true }).click();
  await expect(page.getByLabel("会话状态")).toHaveText("修订 1 · 批注 1");
  await page.getByRole("button", { name: "定位", exact: true }).click();
  await expect(page.frameLocator("iframe").locator("[data-sdk-highlight]")).not.toHaveCount(0);
  await page.getByRole("button", { name: "修改批注", exact: true }).click();
  await page.getByLabel("修改批注内容").fill("改为第四季度交付");
  await page.getByRole("button", { name: "保存批注", exact: true }).click();
  await page.getByRole("button", { name: "标为已处理", exact: true }).click();
  await expect(page.getByLabel("会话状态")).toHaveText("修订 3 · 批注 1");
  await page.getByRole("button", { name: "撤销", exact: true }).click();
  await expect(page.getByText("待处理", { exact: true })).toBeVisible();
  await page.getByRole("button", { name: "重做", exact: true }).click();
  await expect(page.getByText("已处理", { exact: true })).toBeVisible();
  const saved = await page.evaluate(() =>
    Array.from(
      window.interactionHost.session!.serialize(),
      ([path, bytes]) => [path, [...bytes]] as [string, number[]],
    ),
  );
  const reopened = openDocument(
    new Map(saved.map(([path, bytes]) => [path, new Uint8Array(bytes)])),
  );
  expect(reopened.snapshot().annotations[0]).toMatchObject({
    instruction: "改为第四季度交付",
    status: "resolved",
    anchor: { kind: "text", nodeId: "n-summary", from: 2, to: 8, quote: "计划于第三季" },
  });
  expect(reopened.snapshot().html).not.toContain("data-sdk-highlight");
  await page.getByRole("button", { name: "删除批注", exact: true }).click();
  await expect(page.getByLabel("会话状态")).toHaveText("修订 6 · 批注 0");
});

test("saving an unchanged annotation closes its editor without changing revision or history", async ({
  page,
}) => {
  await mount(page);
  await select(page, "n-summary", 2, "n-summary", 8);
  await page.getByLabel("新批注内容").fill("原说明");
  await page.getByRole("button", { name: "添加批注", exact: true }).click();
  await page.getByRole("button", { name: "标为已处理", exact: true }).click();
  await page.getByRole("button", { name: "撤销", exact: true }).click();
  const before = await page.evaluate(() => ({
    snapshot: window.interactionHost.session!.snapshot(),
    history: window.interactionHost.session!.history(),
  }));
  await page.getByRole("button", { name: "修改批注", exact: true }).click();
  await expect(page.getByLabel("修改批注内容")).toHaveValue("原说明");
  await page.getByRole("button", { name: "保存批注", exact: true }).click();
  await expect(page.getByLabel("修改批注内容")).toHaveCount(0);
  await expect(page.getByRole("button", { name: "修改批注", exact: true })).toBeVisible();
  expect(
    await page.evaluate(() => ({
      snapshot: window.interactionHost.session!.snapshot(),
      history: window.interactionHost.session!.history(),
    })),
  ).toEqual(before);
  await page.getByRole("button", { name: "重做", exact: true }).click();
  await expect(page.getByText("已处理", { exact: true })).toBeVisible();
  expect(await page.evaluate(() => window.interactionHost.errors)).toEqual([]);
});

test("cross-paragraph and table ranges retain source order and UTF-16 quotes", async ({ page }) => {
  await mount(
    page,
    fixture(
      '<p data-node-id="a">中文😀𠀀\u{e0100}结尾</p><p data-node-id="b">第二段</p><table data-node-id="t"><tr data-node-id="tr"><td data-node-id="c">单元格</td></tr></table>',
    ),
  );
  await select(page, "a", 2, "c", 2);
  const selection = await page.evaluate(() => window.interactionHost.view!.selection());
  expect(selection?.anchor).toEqual({
    kind: "text-set",
    ranges: [
      { nodeId: "a", from: 2, to: 10, quote: "😀𠀀\u{e0100}结尾" },
      { nodeId: "b", from: 0, to: 3, quote: "第二段" },
      { nodeId: "c", from: 0, to: 2, quote: "单元" },
    ],
  });
  await page.getByLabel("新批注内容").fill("跨节点批注");
  await page.getByRole("button", { name: "添加批注", exact: true }).click();
  await expect(page.getByLabel("会话状态")).toHaveText("修订 1 · 批注 1");
  await select(page, "a", 3, "a", 4);
  await expect(page.getByRole("alert")).toContainText("代理对");
  expect(
    await page.evaluate(() => window.interactionHost.session!.snapshot().manifest.revision),
  ).toBe(1);
});

test("inline styles, line breaks and embedded objects map to canonical text", async ({ page }) => {
  const bytes = readFileSync(new URL("../../examples/docx/03-images.docx", import.meta.url));
  const files = new Map(importDocx(bytes, { documentId: "inline-objects" }));
  // Importer output provides a real declared image and its package bytes.
  const model = openDocument(files);
  const firstImage = /<img[^>]*>/.exec(model.snapshot().html)![0];
  const plain = new Map(
    fixture(
      '<p data-node-id="a">前<strong data-node-id="s">加粗😀</strong><br data-node-id="br">' +
        firstImage.replace(/data-node-id="[^"]*"/, 'data-node-id="image"') +
        "后</p>",
    ),
  );
  const manifest = JSON.parse(new TextDecoder().decode(plain.get("manifest.json")));
  manifest.assets = model.snapshot().manifest.assets;
  plain.set("manifest.json", new TextEncoder().encode(JSON.stringify(manifest)));
  for (const asset of manifest.assets) plain.set(asset.path, files.get(asset.path)!);
  await mount(page, plain);
  await page
    .frameLocator("iframe")
    .locator('[data-node-id="a"]')
    .evaluate((node) => {
      const range = document.createRange();
      range.selectNodeContents(node);
      window.getSelection()!.removeAllRanges();
      window.getSelection()!.addRange(range);
      document.dispatchEvent(new Event("selectionchange"));
    });
  expect((await page.evaluate(() => window.interactionHost.view!.selection()))?.anchor).toEqual({
    kind: "text",
    nodeId: "a",
    from: 0,
    to: 8,
    quote: "前加粗😀\n\ufffc后",
  });
});

test("selection mapping reads each owner once across thousands of inline runs", async ({
  page,
}, info) => {
  const runs = Array.from(
    { length: 2000 },
    (_, index) => '<span data-node-id="run-' + index + '">中😀</span>',
  ).join("");
  await mount(page, fixture('<p data-node-id="a">' + runs + '</p><p data-node-id="b">末尾</p>'));
  const measured = await page.evaluate(() => {
    const host = window.interactionHost,
      frame =
        host.view!.element.shadowRoot!.querySelector<HTMLIFrameElement>("iframe")!.contentDocument!;
    const range = frame.createRange();
    range.setStart(frame.querySelector('[data-node-id="a"]')!, 0);
    range.setEnd(frame.querySelector('[data-node-id="b"]')!.firstChild!, 2);
    frame.getSelection()!.removeAllRanges();
    frame.getSelection()!.addRange(range);
    const before = host.nodeReads.length,
      started = performance.now();
    frame.dispatchEvent(new Event("selectionchange"));
    return {
      reads: host.nodeReads.slice(before),
      durationMs: performance.now() - started,
      anchor: host.view!.selection()?.anchor,
    };
  });
  await info.attach("selection-mapping", {
    body: JSON.stringify({ reads: measured.reads, durationMs: measured.durationMs }),
    contentType: "application/json",
  });
  expect(measured.reads).toEqual(["a", "b"]);
  expect(measured.anchor).toEqual({
    kind: "text-set",
    ranges: [
      { nodeId: "a", from: 0, to: 6000, quote: "中😀".repeat(2000) },
      { nodeId: "b", from: 0, to: 2, quote: "末尾" },
    ],
  });
  expect(await page.evaluate(() => window.interactionHost.session!.history().undoSteps)).toBe(0);
});

test("DOCX numbering and tab leaders never become source text", async ({ page }) => {
  for (const file of ["01-text-numbering.docx", "06-links-tabs.docx"]) {
    await mount(
      page,
      importDocx(readFileSync(new URL("../../examples/docx/" + file, import.meta.url)), {
        documentId: file,
      }),
    );
    const frame = page.frameLocator("iframe");
    const paragraph = frame
      .locator(file.startsWith("01") ? "[data-word-numbering]" : "p:has([data-word-tab])")
      .first();
    const id = await paragraph.getAttribute("data-node-id");
    await paragraph.evaluate((node) => {
      const range = document.createRange();
      range.selectNodeContents(node);
      window.getSelection()!.removeAllRanges();
      window.getSelection()!.addRange(range);
      document.dispatchEvent(new Event("selectionchange"));
    });
    expect(await page.evaluate(() => window.interactionHost.view!.selection()?.anchor)).toEqual({
      kind: "text",
      nodeId: id,
      from: 0,
      to: openDocument(
        importDocx(readFileSync(new URL("../../examples/docx/" + file, import.meta.url)), {
          documentId: file,
        }),
      ).getNode(id!)!.text.length,
      quote: openDocument(
        importDocx(readFileSync(new URL("../../examples/docx/" + file, import.meta.url)), {
          documentId: file,
        }),
      ).getNode(id!)!.text,
    });
  }
});

test("hidden instruction caches stay outside a visible field selection", async ({ page }) => {
  await mount(page, importDocx(nestedFieldFixture(false), { documentId: "hidden-cache" }));
  await page
    .frameLocator("iframe")
    .locator("p")
    .first()
    .evaluate((node) => {
      const range = document.createRange();
      range.selectNodeContents(node);
      window.getSelection()!.removeAllRanges();
      window.getSelection()!.addRange(range);
      document.dispatchEvent(new Event("selectionchange"));
    });
  const anchor = (await page.evaluate(() => window.interactionHost.view!.selection()))!.anchor;
  const quote =
    anchor.kind === "text"
      ? anchor.quote
      : anchor.kind === "text-set"
        ? anchor.ranges.map((range) => range.quote).join("")
        : "";
  expect(quote).toContain("outer result");
  expect(quote).not.toContain("cache");
  await page.getByLabel("新批注内容").fill("字段可见结果");
  await page.getByRole("button", { name: "添加批注", exact: true }).click();
  await expect(page.getByLabel("会话状态")).toHaveText("修订 1 · 批注 1");
});

test("object selection supports ordered node sets and leaves history unchanged", async ({
  page,
}) => {
  await mount(page);
  await page.getByRole("button", { name: "对象选择", exact: true }).click();
  const frame = page.frameLocator("iframe");
  await frame.locator('[data-node-id="n-summary"]').click();
  await frame.locator('[data-node-id="n-title"]').click({ modifiers: ["Control"] });
  expect((await page.evaluate(() => window.interactionHost.view!.selection()))?.anchor).toEqual({
    kind: "node-set",
    nodeIds: ["n-title", "n-summary"],
  });
  expect(await page.evaluate(() => window.interactionHost.session!.history().undoSteps)).toBe(0);
  await page.getByLabel("新批注内容").fill("对象批注");
  await page.getByRole("button", { name: "添加批注", exact: true }).click();
  await expect(page.getByLabel("会话状态")).toHaveText("修订 1 · 批注 1");
});

test("native mouse selection and document keyboard shortcuts use the core history", async ({
  page,
}) => {
  await mount(page);
  const box = (await page.locator("iframe").boundingBox())!;
  const text = await page
    .frameLocator("iframe")
    .locator('[data-node-id="n-summary"]')
    .evaluate((node) => {
      const range = document.createRange();
      range.setStart(node.firstChild!, 2);
      range.setEnd(node.firstChild!, 8);
      const rect = range.getBoundingClientRect();
      return { x: rect.left, y: rect.top + rect.height / 2, width: rect.width };
    });
  await page.mouse.move(box.x + text.x + 0.1, box.y + text.y);
  await page.mouse.down();
  // Playwright's Chromium drag helper waits for a timer inside the scriptless frame.
  // Raw trusted mouse movement exercises native selection without that injected timer.
  const client = await page.context().newCDPSession(page);
  for (let step = 1; step <= 8; step++)
    await client.send("Input.dispatchMouseEvent", {
      type: "mouseMoved",
      x: box.x + text.x + (text.width * step) / 8 - 0.1,
      y: box.y + text.y,
      button: "left",
      buttons: 1,
    });
  await page.mouse.up();
  await client.detach();
  await expect(page.getByLabel("当前选择")).toHaveText("计划于第三季");
  await page.getByLabel("新批注内容").fill("鼠标划选");
  await page.getByRole("button", { name: "添加批注", exact: true }).click();
  await page.frameLocator("iframe").locator('[data-node-id="n-summary"]').click();
  await page.keyboard.press("Control+z");
  await expect(page.getByLabel("会话状态")).toHaveText("修订 2 · 批注 0");
  await page.keyboard.press("Control+Shift+z");
  await expect(page.getByLabel("会话状态")).toHaveText("修订 3 · 批注 1");
  await page.keyboard.press("Escape");
  await expect(page.getByLabel("当前选择")).toHaveText("尚未选择");
});

for (const action of ["clear button", "Escape in document", "Escape in toolbar", "mode change"]) {
  test(`${action} clears native ranges and prevents an old selection from returning`, async ({
    page,
  }) => {
    await mount(page);
    await select(page, "n-summary", 2, "n-summary", 8);
    await page.getByLabel("新批注内容").fill("焦点移走后保留选择");
    await expect(page.getByLabel("当前选择")).toHaveText("计划于第三季");
    await expect(page.getByRole("button", { name: "添加批注", exact: true })).toBeEnabled();
    expect(
      await page
        .frameLocator("iframe")
        .locator("body")
        .evaluate(() => window.getSelection()!.isCollapsed),
    ).toBe(false);
    if (action === "clear button")
      await page.getByRole("button", { name: "清除选择", exact: true }).click();
    else if (action === "mode change") {
      await page.getByRole("button", { name: "对象选择", exact: true }).click();
      await page.getByRole("button", { name: "文字划选", exact: true }).click();
    } else {
      if (action === "Escape in document")
        await page
          .locator("iframe")
          .evaluate((node) => (node as HTMLIFrameElement).contentWindow!.focus());
      else await page.getByRole("button", { name: "文字划选", exact: true }).focus();
      await page.keyboard.press("Escape");
    }
    expect(
      await page
        .frameLocator("iframe")
        .locator("body")
        .evaluate(() => {
          const selected = window.getSelection()!;
          document.dispatchEvent(new Event("selectionchange"));
          return { rangeCount: selected.rangeCount, isCollapsed: selected.isCollapsed };
        }),
    ).toEqual({ rangeCount: 0, isCollapsed: true });
    await page.evaluate(() => window.interactionHost.settle());
    await expect(page.getByLabel("当前选择")).toHaveText("尚未选择");
    await expect(page.getByRole("button", { name: "添加批注", exact: true })).toBeDisabled();
    expect(await page.evaluate(() => window.interactionHost.view!.selection())).toBeUndefined();
    expect(
      await page.evaluate(() => window.interactionHost.session!.snapshot().manifest.revision),
    ).toBe(0);
    expect(await page.evaluate(() => window.interactionHost.session!.history().undoSteps)).toBe(0);
  });
}

test("native rich comments retain their body and auxiliary stories cannot enter a main selection", async ({
  page,
}) => {
  const files = importDocx(
    readFileSync(new URL("../../examples/docx/08-comments-revisions.docx", import.meta.url)),
    { documentId: "comments" },
  );
  await mount(page, files);
  await page.getByRole("button", { name: "标为已处理", exact: true }).click();
  expect(
    await page.evaluate(() => window.interactionHost.session!.getAnnotations()[0]!.bodyNodeId),
  ).toBe(openDocument(files).snapshot().annotations[0]!.bodyNodeId);
  await page.getByRole("button", { name: "定位", exact: true }).click();
  await page
    .frameLocator("iframe")
    .locator("article")
    .evaluate((node) => {
      const range = document.createRange();
      range.selectNodeContents(node);
      window.getSelection()!.removeAllRanges();
      window.getSelection()!.addRange(range);
      document.dispatchEvent(new Event("selectionchange"));
    });
  await expect(page.getByRole("alert")).toContainText("同一正文内容域");
  expect(await page.evaluate(() => window.interactionHost.view!.selection())).toBeUndefined();
  expect(
    await page.evaluate(() => window.interactionHost.session!.snapshot().manifest.revision),
  ).toBe(1);
});

test("region coordinates exclude zoom and scrolling and preserve empty regions", async ({
  page,
}) => {
  const errors: string[] = [];
  page.on("pageerror", (error) => errors.push(error.message));
  await mount(page);
  await page.getByRole("button", { name: "矩形框选", exact: true }).click();
  let key: string | undefined;
  for (const zoom of [1, 0.5, 2]) {
    await page.getByLabel("预览缩放").selectOption(String(zoom));
    await page.evaluate(() => window.interactionHost.settle());
    await drag(page, { x: 6, y: 6, width: 18, height: 12 });
    const anchor = (await page.evaluate(() => window.interactionHost.view!.selection()))?.anchor;
    expect(anchor?.kind).toBe("region");
    if (anchor?.kind !== "region") throw new Error("Expected region");
    expect(anchor.rect.x).toBeCloseTo(6, 0);
    expect(anchor.rect.y).toBeCloseTo(6, 0);
    expect(anchor.rect.width).toBeCloseTo(18, 0);
    expect(anchor.rect.height).toBeCloseTo(12, 0);
    expect(anchor.nodeIds).toEqual([]);
    if (key) expect(anchor.layoutKey).toBe(key);
    key = anchor.layoutKey;
  }
  expect(
    await page.evaluate(() => window.interactionHost.session!.snapshot().manifest.revision),
  ).toBe(0);
  await page.getByLabel("新批注内容").fill("空白区域添加内容");
  await page.getByRole("button", { name: "添加批注", exact: true }).click();
  await expect(page.getByLabel("会话状态")).toHaveText("修订 1 · 批注 1");
  await page.getByLabel("预览缩放").selectOption("1");
  await page.evaluate(() => window.interactionHost.settle());
  expect(
    await page.evaluate(() => window.interactionHost.session!.getAnnotations()[0]!.status),
  ).toBe("open");
  await page.evaluate(async () => {
    window.interactionHost.view!.setZoom(2);
    window.interactionHost.view!.dispose();
    await new Promise<void>((resolve) =>
      requestAnimationFrame(() => requestAnimationFrame(() => resolve())),
    );
  });
  expect(errors).toEqual([]);
});

test("regions hit multiple nodes and layout changes require explicit reanchoring", async ({
  page,
}) => {
  await mount(page);
  await page.getByRole("button", { name: "矩形框选", exact: true }).click();
  await drag(page, { x: 60, y: 60, width: 200, height: 70 });
  const before = (await page.evaluate(() => window.interactionHost.view!.selection()))?.anchor;
  expect(before?.kind === "region" && before.nodeIds.length).toBeGreaterThan(1);
  await page.getByLabel("新批注内容").fill("区域批注");
  await page.getByRole("button", { name: "添加批注", exact: true }).click();
  await page.evaluate(async () => {
    const session = window.interactionHost.session!,
      manifest = session.snapshot().manifest;
    session.commit(
      session.prepare({
        protocolVersion: 1,
        documentId: manifest.documentId,
        baseRevision: manifest.revision,
        operations: [{ op: "setStyle", nodeId: "n-title", property: "font-size", value: "40pt" }],
      }),
    );
    await window.interactionHost.settle();
  });
  await expect(page.getByText("需要重新定位", { exact: true })).toBeVisible();
  expect(
    await page.evaluate(() => window.interactionHost.session!.getAnnotations()[0]!.anchor),
  ).toEqual(before);
  await drag(page, { x: 5, y: 5, width: 12, height: 12 });
  await page.getByRole("button", { name: "用当前选择重新定位", exact: true }).click();
  await expect(page.getByText("待处理", { exact: true })).toBeVisible();
  await page.getByRole("button", { name: "撤销", exact: true }).click();
  await expect(page.getByText("需要重新定位", { exact: true })).toBeVisible();
});

test("Escape cancels a captured region drag", async ({ page }) => {
  await mount(page);
  await page.getByRole("button", { name: "矩形框选", exact: true }).click();
  const box = (await page.locator("iframe").boundingBox())!;
  const point = await page
    .frameLocator("iframe")
    .locator('[data-node-id="root"]')
    .evaluate((node) => {
      const box = node.getBoundingClientRect();
      return { x: box.left + 10, y: box.top + 10 };
    });
  await page.mouse.move(box.x + point.x, box.y + point.y);
  await page.mouse.down();
  await page.mouse.move(box.x + point.x + 20, box.y + point.y + 20);
  await page.keyboard.press("Escape");
  await page.mouse.move(box.x + point.x + 40, box.y + point.y + 40);
  await page.mouse.up();
  await expect(page.getByLabel("当前选择")).toHaveText("尚未选择");
  expect(await page.evaluate(() => window.interactionHost.session!.history().undoSteps)).toBe(0);
});

test("scrolling and reverse drags keep the same logical rectangle", async ({ page }) => {
  await mount(
    page,
    fixture('<p data-node-id="a">正文</p><div data-node-id="space" style="height:600pt"></div>'),
  );
  await page.getByRole("button", { name: "矩形框选", exact: true }).click();
  const rect = { x: 160, y: 100, width: 24, height: 24 };
  await drag(page, rect);
  const original = (await page.evaluate(() => window.interactionHost.view!.selection()))!.anchor;
  await page.getByLabel("预览缩放").selectOption("2");
  await page.evaluate(async () => {
    const stage = window.interactionHost.view!.element.shadowRoot!.querySelector(".stage")!;
    stage.scrollTop = 80;
    stage.scrollLeft = 120;
    await window.interactionHost.settle();
  });
  await drag(page, {
    x: rect.x + rect.width,
    y: rect.y + rect.height,
    width: -rect.width,
    height: -rect.height,
  });
  const current = (await page.evaluate(() => window.interactionHost.view!.selection()))!.anchor;
  expect(current.kind).toBe("region");
  if (original.kind !== "region" || current.kind !== "region") throw new Error("Expected region");
  for (const key of ["x", "y", "width", "height"] as const)
    expect(current.rect[key]).toBeCloseTo(original.rect[key], 0);
  expect(current.layoutKey).toBe(original.layoutKey);
});

test("environment layout changes invalidate regions without serializing preview DOM", async ({
  page,
}) => {
  await mount(page);
  await page.getByRole("button", { name: "矩形框选", exact: true }).click();
  await drag(page, { x: 5, y: 5, width: 15, height: 15 });
  await page.getByLabel("新批注内容").fill("布局失效");
  await page.getByRole("button", { name: "添加批注", exact: true }).click();
  const original = await page.evaluate(() => window.interactionHost.session!.getAnnotations()[0]!);
  await page
    .frameLocator("iframe")
    .locator('[data-node-id="n-summary"]')
    .evaluate((node) => {
      (node as HTMLElement).style.fontFamily = "monospace";
      document.fonts.dispatchEvent(new Event("loadingdone"));
    });
  await expect(page.getByText("需要重新定位", { exact: true })).toBeVisible();
  const changed = await page.evaluate(() => window.interactionHost.session!.snapshot());
  expect(changed.manifest.revision).toBe(2);
  expect(changed.annotations[0]!.anchor).toEqual(original.anchor);
  expect(changed.html).not.toContain("monospace");
  await page.getByRole("button", { name: "定位", exact: true }).click();
  await expect(page.getByRole("alert")).toContainText("需要重新定位");
});

test("external commits clear stale selections and retain mapped annotations", async ({ page }) => {
  await mount(page);
  await select(page, "n-summary", 2, "n-summary", 8);
  await page.getByLabel("新批注内容").fill("已有批注");
  await page.getByRole("button", { name: "添加批注", exact: true }).click();
  await select(page, "n-summary", 2, "n-summary", 8);
  await page.evaluate(async () => {
    const session = window.interactionHost.session!,
      manifest = session.snapshot().manifest;
    session.commit(
      session.prepare({
        protocolVersion: 1,
        documentId: manifest.documentId,
        baseRevision: manifest.revision,
        operations: [
          {
            op: "replaceText",
            nodeId: "n-summary",
            from: 0,
            to: 0,
            expectedText: "",
            text: "新：",
          },
        ],
      }),
    );
    await window.interactionHost.settle();
  });
  await expect(page.getByLabel("当前选择")).toHaveText("尚未选择");
  await expect(page.getByRole("button", { name: "添加批注", exact: true })).toBeDisabled();
  expect(
    await page.evaluate(() => window.interactionHost.session!.getAnnotations()[0]!.anchor),
  ).toMatchObject({ from: 4, to: 10, quote: "计划于第三季" });
  await page.getByRole("button", { name: "定位", exact: true }).click();
  await expect(page.frameLocator("iframe").locator("[data-sdk-highlight]")).not.toHaveCount(0);
});

test("an external commit preserves an annotation draft and rejects saving its old revision", async ({
  page,
}) => {
  await mount(page);
  await select(page, "n-summary", 2, "n-summary", 8);
  await page.getByLabel("新批注内容").fill("原说明");
  await page.getByRole("button", { name: "添加批注", exact: true }).click();
  await page.getByRole("button", { name: "修改批注", exact: true }).click();
  await page.getByLabel("修改批注内容").fill("尚未提交的草稿");
  await page.evaluate(async () => {
    const session = window.interactionHost.session!,
      manifest = session.snapshot().manifest;
    session.commit(
      session.prepare({
        protocolVersion: 1,
        documentId: manifest.documentId,
        baseRevision: manifest.revision,
        operations: [{ op: "setStyle", nodeId: "n-title", property: "color", value: "blue" }],
      }),
    );
    await window.interactionHost.settle();
  });
  await expect(page.getByLabel("修改批注内容")).toHaveValue("尚未提交的草稿");
  await page.getByRole("button", { name: "保存批注", exact: true }).click();
  await expect(page.getByRole("alert")).toContainText("different document revision");
  await expect(page.getByLabel("修改批注内容")).toHaveValue("尚未提交的草稿");
  expect(
    await page.evaluate(() => window.interactionHost.session!.snapshot().manifest.revision),
  ).toBe(2);
  expect(
    await page.evaluate(() => window.interactionHost.session!.getAnnotations()[0]!.instruction),
  ).toBe("原说明");
  await page.getByRole("button", { name: "取消修改", exact: true }).click();
  await expect(page.getByLabel("修改批注内容")).toHaveCount(0);
});

test("host styles stay isolated and unmount releases listeners and image URLs without disposing the session", async ({
  page,
}) => {
  await page.addInitScript(() => {
    const counts = { created: 0, revoked: 0 };
    Object.assign(window, { urlCounts: counts });
    const create = URL.createObjectURL.bind(URL),
      revoke = URL.revokeObjectURL.bind(URL);
    URL.createObjectURL = (value) => {
      counts.created++;
      return create(value);
    };
    URL.revokeObjectURL = (value) => {
      counts.revoked++;
      revoke(value);
    };
  });
  await mount(
    page,
    importDocx(readFileSync(new URL("../../examples/docx/03-images.docx", import.meta.url)), {
      documentId: "cleanup",
    }),
  );
  await page.addStyleTag({
    content:
      "p,button,textarea{color:rgb(255,0,0)!important;font-size:50px!important}iframe{color:red}",
  });
  await expect(page.getByRole("button", { name: "文字划选", exact: true })).not.toHaveCSS(
    "color",
    "rgb(255, 0, 0)",
  );
  await expect(page.frameLocator("iframe").locator("p").first()).not.toHaveCSS("font-size", "50px");
  await page.evaluate(() => {
    window.interactionHost.view!.dispose();
    window.interactionHost.view!.dispose();
  });
  await expect(page.locator("iframe")).toHaveCount(0);
  const counts = await page.evaluate(
    () => (window as unknown as { urlCounts: { created: number; revoked: number } }).urlCounts,
  );
  expect(counts.created).toBe(3);
  expect(counts.revoked).toBe(3);
  expect(
    await page.evaluate(() => window.interactionHost.session!.snapshot().manifest.documentId),
  ).toBe("cleanup");
  await page.evaluate(() => window.interactionHost.session!.undo());
  expect(await page.evaluate(() => window.interactionHost.errors)).toEqual([]);
});

test("disposing during initial rendering prevents a late view from reconnecting", async ({
  page,
}) => {
  await mount(page);
  await page.evaluate(async (files) => {
    const pending = window.interactionHost.mount(files);
    window.interactionHost.view!.dispose();
    await pending;
  }, entries(fixture()));
  await expect(page.locator("iframe")).toHaveCount(0);
  expect(await page.evaluate(() => window.interactionHost.session!.history().undoSteps)).toBe(0);
  expect(await page.evaluate(() => window.interactionHost.errors)).toEqual([]);
});

test("Playground downloads annotation transactions and reopens them", async ({ page }, info) => {
  await page.goto("/");
  await page.getByRole("button", { name: "打开 DOCX 样例" }).click();
  await page.getByRole("button", { name: "返回预览与批注", exact: true }).click();
  await page
    .frameLocator('iframe[title="文档只读预览"]')
    .getByRole("heading", { name: "实施报告", exact: true })
    .evaluate((node) => {
      const range = document.createRange();
      range.selectNodeContents(node);
      window.getSelection()!.removeAllRanges();
      window.getSelection()!.addRange(range);
      document.dispatchEvent(new Event("selectionchange"));
    });
  await page.getByRole("tab", { name: "审阅", exact: true }).click();
  await page.getByRole("button", { name: "新建批注", exact: true }).click();
  const dialog = page.getByRole("dialog", { name: "新建批注", exact: true });
  await dialog.getByLabel("新批注内容").fill("保存并重开");
  await dialog.getByRole("button", { name: "添加批注", exact: true }).click();
  await expect(page.locator(".office-app")).toHaveAttribute("data-revision", "1");
  const download = page.waitForEvent("download");
  await page.getByRole("button", { name: "下载文档包", exact: true }).click();
  const path = info.outputPath("annotations.zip");
  await (await download).saveAs(path);
  const saved = openDocument(readDocumentZip(readFileSync(path))).snapshot();
  expect(saved.annotations).toHaveLength(1);
  expect(saved.annotations[0]).toMatchObject({
    instruction: "保存并重开", anchor: { kind: "text", from: 0, to: 4, quote: "实施报告" },
  });
  await page.getByLabel("打开文件", { exact: true }).filter({ visible: false }).setInputFiles(path);
  await expect(page.getByRole("heading", { name: "annotations.zip", exact: true })).toBeVisible();
  await expect(page.locator(".office-app")).toHaveAttribute("data-revision", "1");
  await expect(
    page
      .getByRole("region", { name: "文档批注", exact: true })
      .getByText("保存并重开", { exact: true }),
  ).toBeVisible();
});

test("Word comment reply drafts, edits, deletion and undo use the same session", async ({
  page,
}) => {
  await mount(page, fixture('<p data-node-id="a">讨论正文😀</p>'));
  await select(page, "a", 0, "a", 2);
  await page.getByLabel("新批注内容").fill("原意见");
  await page.getByRole("button", { name: "添加批注", exact: true }).click();
  await page.getByRole("button", { name: "回复批注", exact: true }).click();
  await page.getByLabel("回复内容", { exact: true }).fill("回复漢\u{e0100}");
  await page.getByLabel("回复作者", { exact: true }).fill("乙");
  await page.getByRole("button", { name: "发送回复", exact: true }).click();
  await expect(page.locator("[data-reply-id]")).toContainText("回复漢\u{e0100}");
  expect(
    await page.evaluate(
      () => window.interactionHost.session!.getAnnotations()[0]!.replies![0]!.author?.name,
    ),
  ).toBe("乙");
  await page.getByRole("button", { name: "修改回复", exact: true }).click();
  await page.getByLabel("批注消息内容", { exact: true }).fill("新回复");
  await page.getByRole("button", { name: "保存消息", exact: true }).click();
  await expect(page.locator("[data-reply-id]")).toContainText("新回复");
  await page.getByRole("button", { name: "删除回复", exact: true }).click();
  await expect(page.locator("[data-reply-id]")).toHaveCount(0);
  await page.getByRole("button", { name: "撤销", exact: true }).click();
  await expect(page.locator("[data-reply-id]")).toContainText("新回复");
  expect(await page.evaluate(() => window.interactionHost.errors)).toEqual([]);
});
