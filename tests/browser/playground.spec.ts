import { expect, test } from "@playwright/test";
import type { Page } from "@playwright/test";
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import {
  editedChineseFontFixture, indentedTabsFixture, mutateDocx, nestedFieldFixture, sectionPackageFixture, styleInheritanceFixture, stylesheetPackageFixture,
} from "../../packages/office/tests/regression-fixtures.js";

const samples = fileURLToPath(new URL("../../examples/docx/", import.meta.url));
test.beforeEach(async ({ page }) => {
  // Keep R1-10's continuous reading baselines; pagination has its own fixed
  // fixtures and all three browser-engine gates in pagination.spec.ts.
  await page.addInitScript(() => { if (window === window.top) { localStorage.setItem("doffi:layout", "continuous"); localStorage.setItem("doffi:editing", "preview"); } });
});
test("fixed-font continuous rich-text input fixture", async ({ page }) => {
  await page.goto("/"); await page.getByRole("button", { name: "打开报告协作样例" }).click();
  await page.getByRole("button", { name: "编辑正文", exact: true }).click();
  const paragraph = page.frameLocator('iframe[title="正文编辑"]').locator('[data-node-id="n-summary"]');
  await paragraph.click(); await page.keyboard.press("End"); await page.keyboard.insertText(" 中文输入 ABC");
  await paragraph.evaluate(async () => document.fonts.ready);
  await expect(paragraph).toContainText("中文输入 ABC");
  await expect(paragraph).toHaveScreenshot("continuous-rich-text.png", { animations: "disabled", caret: "hide", maxDiffPixelRatio: 0.005, threshold: 0.2 });
});
async function expandDocumentScreenshot(page: Page): Promise<void> {
  // Measure the complete source rendering, independent of the interaction scroll viewport.
  await page.locator("iframe").evaluate((frame) => {
    const stage = (frame.getRootNode() as ShadowRoot).querySelector<HTMLElement>(".stage");
    if (stage) stage.style.maxHeight = "none";
  });
}
const cases = JSON.parse(
  readFileSync(
    new URL("../../examples/docx/expected.json", import.meta.url),
    "utf8",
  ),
) as { file: string; requiredText: string[] }[];

for (const [terminator, escapedLessThan] of [
  ["</style>", false], ["</StYlE>", false], ["</STYLE >", false], ["</style>", true],
] as const) {
  test("keeps stylesheet raw-text terminators inside CSS: " + terminator + " escaped=" + escapedLessThan, async ({ page }) => {
    const payload = terminator + '<input id=escaped><a href=https://example.com/injected>link</a><style>';
    await page.goto("/");
    await page.getByLabel("打开文件").setInputFiles({
      name: "stylesheet.zip", mimeType: "application/zip",
      buffer: Buffer.from(stylesheetPackageFixture('p{--payload:"' + (escapedLessThan ? '\\' : '') + payload + '";color:#123456}')),
    });
    await expect(page.locator(".office-app")).toHaveAttribute("data-loading", "false"); await expect(page.locator(".office-app")).toHaveAttribute("data-revision", "0");
    const frame = page.frameLocator("iframe");
    await expect(frame.locator("input,a[href]")).toHaveCount(0);
    await expect(frame.locator("style")).toHaveCount(2);
    const paragraph = frame.locator("p").first();
    await expect(paragraph).toHaveCSS("color", "rgb(18, 52, 86)");
    const literal = await paragraph.evaluate((node) => {
      const probe = document.createElement("span");
      probe.style.fontFamily = "var(--payload)";
      node.append(probe);
      const result = getComputedStyle(probe).fontFamily;
      probe.remove();
      return JSON.parse(result) as string;
    });
    expect(literal).toBe(payload);
  });
}

test("positions custom and default tabs from the unindented text origin", async ({ page }) => {
  await page.goto("/");
  await page.getByLabel("打开文件").setInputFiles({
    name: "indented-tabs.docx",
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    buffer: Buffer.from(indentedTabsFixture()),
  });
  await expect(page.locator(".office-app")).toHaveAttribute("data-loading", "false"); await expect(page.locator(".office-app")).toHaveAttribute("data-revision", "0");
  const frame = page.frameLocator("iframe");
  await expect(frame.locator("[data-word-tab]").last()).toHaveCSS("display", "inline-block");
  const positions = await frame.locator("[data-word-tab]").evaluateAll((tabs) => tabs.map((tab) => {
    const section = tab.closest("section")!;
    const origin = section.getBoundingClientRect().left + parseFloat(getComputedStyle(section).paddingLeft);
    return (tab.getBoundingClientRect().right - origin) * 72 / 96;
  }));
  for (const [index, expected] of [72, 72, 144, 72, 72, 72].entries())
    expect(positions[index]).toBeCloseTo(expected, 1);
  await expect(frame.locator("[data-word-run]").first()).toHaveCSS("font-family", /Liberation Serif/);
  await expandDocumentScreenshot(page);
  await expect(frame.locator("article")).toHaveScreenshot("indented-tabs.png", {
    animations: "disabled", maxDiffPixelRatio: 0.005, threshold: 0.2,
  });
});

test("opens a DOCX sample through the public APIs and releases its preview on replacement", async ({
  page,
}) => {
  const errors: string[] = [];
  page.on("pageerror", (error) => errors.push(error.message));
  await page.goto("/");
  await expect(
    page.getByRole("link", { name: "doffi Playground", exact: true }),
  ).toBeVisible();
  await page.getByRole("button", { name: "打开 DOCX 样例" }).click();
  await expect(page.locator(".office-app")).toHaveAttribute("data-loading", "false"); await expect(page.locator(".office-app")).toHaveAttribute("data-revision", "0");
  const frame = page.frameLocator('iframe[title="文档只读预览"]');
  await expect(frame.getByRole("heading", { name: "实施报告" })).toBeVisible();
  await expect(frame.locator(".sdk-number-marker")).toHaveText([
    "1.",
    "1.a)",
    "2.",
    "4.",
  ]);
  const color = await frame
    .getByRole("heading", { name: "实施报告" })
    .evaluate((node) => getComputedStyle(node).color);
  expect(color).toBe("rgb(31, 75, 153)");
  const inherited = await frame
    .getByRole("heading", { name: "实施报告" })
    .evaluate((node) => {
      (node as HTMLElement).style.color = "#884422";
      return getComputedStyle(node.querySelector("[data-word-run]")!).color;
    });
  expect(inherited).toBe("rgb(136, 68, 34)");
  await page
    .getByLabel("打开文件")
    .setInputFiles(samples + "08-comments-revisions.docx");
  await expect(frame.locator("ins")).toHaveText("新安排");
  await expect(frame.locator("del")).toHaveText("旧安排");
  await expect(frame.getByText("请修改", { exact: false })).toBeVisible();
  await expect(page.locator("iframe")).toHaveCount(1);
  expect(errors).toEqual([]);
});
for (const item of cases) {
  test("read-only visual fixture: " + item.file, async ({ page }) => {
    const external: string[] = [];
    page.on("request", (request) => {
      if (request.url().startsWith("https://example.com"))
        external.push(request.url());
    });
    await page.goto("/");
    await page.getByLabel("打开文件").setInputFiles(samples + item.file);
    await expect(page.locator(".office-app")).toHaveAttribute("data-loading", "false"); await expect(page.locator(".office-app")).toHaveAttribute("data-revision", "0");
    const frame = page.frameLocator("iframe");
    await expect(frame.locator("article")).toBeVisible();
    await frame.locator("body").evaluate(async () => {
      await document.fonts.ready;
      await Promise.all([...document.images].map((image) => image.decode()));
    });
    if (item.file === "03-images.docx")
      await expect(frame.locator("img")).toHaveCount(3);
    if (item.file === "06-links-tabs.docx") {
      const positions = await frame
        .locator("[data-word-tab]")
        .evaluate((tab) => ({
          left: tab.getBoundingClientRect().left,
          right: tab.getBoundingClientRect().right,
        }));
      expect(positions.right).toBeGreaterThan(positions.left);
      await expect(frame.locator("a[href]")).toHaveCount(0);
    }
    await expandDocumentScreenshot(page);
    await expect(frame.locator("article")).toHaveScreenshot(
      item.file.replace(".docx", ".png"),
      { animations: "disabled", maxDiffPixelRatio: 0.005, threshold: 0.2 },
    );
    expect(external).toEqual([]);
  });
}
test("downloads its document package and reopens it", async ({
  page,
}, info) => {
  await page.goto("/");
  await page
    .getByLabel("打开文件")
    .setInputFiles(samples + "08-comments-revisions.docx");
  await expect(page.locator(".office-app")).toHaveAttribute("data-loading", "false"); await expect(page.locator(".office-app")).toHaveAttribute("data-revision", "0");
  const download = page.waitForEvent("download");
  await page.getByRole("button", { name: "下载文档包" }).click();
  const path = info.outputPath("document.zip");
  await (await download).saveAs(path);
  await page.getByLabel("打开文件").setInputFiles(path);
  await expect(page.getByRole("status")).toContainText("批注 1");
  await expect(page.frameLocator("iframe").locator("ins")).toHaveText("新安排");
});

test("keeps distinct Latin and Chinese faces and reports missing Chinese fonts", async ({
  page,
}) => {
  await page.goto("/");
  await page
    .getByLabel("打开文件")
    .setInputFiles(samples + "09-chinese-fonts.docx");
  await expect(page.locator(".office-app")).toHaveAttribute("data-loading", "false"); await expect(page.locator(".office-app")).toHaveAttribute("data-revision", "0");
  const frame = page.frameLocator("iframe");
  const source = frame.locator('[data-word-script="eastAsia"]').first();
  await expect(source).toBeVisible();
  expect(
    await source.evaluate((node) => getComputedStyle(node).fontFamily),
  ).toContain("Noto Serif CJK SC");
  await expect(frame.getByRole("note")).toContainText(
    "Dagent Office Missing CJK Test Font",
  );
  await expect(frame.getByRole("note")).toContainText("U+20000");
  const sourceNames = await frame
    .locator("article")
    .evaluate((node) => node.outerHTML);
  expect(sourceNames).toContain("--word-font-eastAsia");
});

for (const [family, selected] of [
  ["FandolSong", "FandolSong"], ["FandolFang", "FandolFang"], ["FandolKai", "FandolKai"], ["LXGW WenKai", "LXGW WenKai"],
  ["仿宋_GB2312", "FandolFang"], ["STFangsong", "FandolFang"], ["楷体_GB2312", "FandolKai"], ["STKaiti", "FandolKai"],
] as const) {
  test("uses installed Chinese font or reports its explicit substitution: " + family, async ({ page }) => {
    const bytes = mutateDocx("09-chinese-fonts.docx", {
      "word/styles.xml": (source) => source.replace(/w:eastAsia="[^"]*"/g, 'w:eastAsia="' + family + '"'),
      "word/document.xml": (source) => source.replace(/w:eastAsia="[^"]*"/g, 'w:eastAsia="' + family + '"'),
    });
    await page.goto("/");
    await page.getByLabel("打开文件").setInputFiles({ name: "chinese-font.docx", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document", buffer: Buffer.from(bytes) });
    await expect(page.locator(".office-app")).toHaveAttribute("data-loading", "false"); await expect(page.locator(".office-app")).toHaveAttribute("data-revision", "0");
    const frame = page.frameLocator("iframe");
    const eastAsian = frame.locator('[data-word-script="eastAsia"]').first();
    await expect(eastAsian).toBeVisible();
    expect(await eastAsian.evaluate((node) => getComputedStyle(node).fontFamily)).toContain(selected);
    if (family !== selected) await expect(frame.getByRole("note")).toContainText(family);
    const latin = frame.locator('[data-word-run="true"]').filter({ hasText: "ABC" }).first();
    await expect(latin).toBeVisible();
    expect(await latin.evaluate((node) => getComputedStyle(node).fontFamily)).toContain("Liberation Serif");
  });
}

test("previews a committed/reopened Chinese font edit with independent Latin and east-Asian roles", async ({ page }) => {
  await page.goto("/");
  await page.getByLabel("打开文件").setInputFiles({ name: "edited-fonts.zip", mimeType: "application/zip", buffer: Buffer.from(editedChineseFontFixture()) });
  await expect(page.locator(".office-app")).toHaveAttribute("data-loading", "false"); await expect(page.locator(".office-app")).toHaveAttribute("data-revision", "0");
  const frame = page.frameLocator("iframe");
  const inserted = frame.locator('[data-word-script="eastAsia"]').filter({ hasText: "新增仿宋" }).first();
  expect(await inserted.evaluate((node) => getComputedStyle(node).fontFamily)).toContain("FandolFang");
  const run = frame.locator('[data-word-run="true"]').filter({ hasText: "新增仿宋" }).first();
  await expect(run).toContainText("新增仿宋ABC𠀀\u{e0100}");
  expect(await run.evaluate((node) => getComputedStyle(node).fontFamily)).toContain("Liberation Serif");
});

test("releases image URLs when switching documents", async ({ page }) => {
  await page.addInitScript(() => {
    const counts = { created: 0, revoked: 0 };
    (window as unknown as { urlCounts: typeof counts }).urlCounts = counts;
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
  await page.goto("/");
  await page.getByLabel("打开文件").setInputFiles(samples + "03-images.docx");
  await expect(page.frameLocator("iframe").locator("img")).toHaveCount(3);
  await page
    .getByLabel("打开文件")
    .setInputFiles(samples + "08-comments-revisions.docx");
  await expect(page.frameLocator("iframe").locator("ins")).toHaveText("新安排");
  expect(
    await page.evaluate(
      () =>
        (
          window as unknown as {
            urlCounts: { created: number; revoked: number };
          }
        ).urlCounts,
    ),
  ).toEqual({ created: 3, revoked: 3 });
});

test("previews partial font overrides, cleared highlighting, inherited tabs and table margins", async ({ page }) => {
  await page.goto("/");
  await page.getByLabel("打开文件").setInputFiles({
    name: "style-inheritance.docx",
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    buffer: Buffer.from(styleInheritanceFixture()),
  });
  await expect(page.locator(".office-app")).toHaveAttribute("data-loading", "false"); await expect(page.locator(".office-app")).toHaveAttribute("data-revision", "0");
  const frame = page.frameLocator("iframe");
  const run = frame.locator("[data-word-run]").filter({ hasText: "ABCé中文" });
  const font = await run.evaluate((node) => ({
    latin: getComputedStyle(node).fontFamily,
    highAnsi: getComputedStyle(node.querySelector('[data-word-script="hAnsi"]')!).fontFamily,
    eastAsia: getComputedStyle(node.querySelector('[data-word-script="eastAsia"]')!).fontFamily,
    background: getComputedStyle(node).backgroundColor,
  }));
  expect(font.latin).toContain("Liberation Serif");
  expect(font.highAnsi).toContain("Noto Sans CJK SC");
  expect(font.eastAsia).toContain("Noto Serif CJK SC");
  expect(font.background).toBe("rgba(0, 0, 0, 0)");
  await expect(frame.locator("[data-word-run]").filter({ hasText: "cleared" }))
    .toHaveCSS("background-color", "rgba(0, 0, 0, 0)");
  for (const text of ["shaded direct", "shaded style"])
    await expect(frame.locator("[data-word-run]").filter({ hasText: text }))
      .toHaveCSS("background-color", "rgb(221, 238, 255)");
  const tabPosition = await frame.locator("[data-word-tab]").evaluate((tab) => {
    const paragraph = tab.closest("p")!;
    return (tab.getBoundingClientRect().right - paragraph.getBoundingClientRect().left) * 72 / 96;
  });
  expect(tabPosition).toBeCloseTo(100, 1);
  const cell = frame.locator("td").first();
  await expect(cell).toHaveCSS("padding-left", "40px");
  await expect(cell).toHaveCSS("padding-right", /53\.33/);
  await expect(cell).toHaveCSS("padding-top", /2\.66/);
  await expect(cell).toHaveCSS("padding-bottom", "4px");
  await expect(frame.locator("table")).toHaveCSS("margin-left", "40px");
  await expandDocumentScreenshot(page);
  await expect(frame.locator("article")).toHaveScreenshot("style-inheritance.png", {
    animations: "disabled", maxDiffPixelRatio: 0.005, threshold: 0.2,
  });
});

for (const simple of [false, true]) {
  test("hides the complete nested instruction cache, including breaks and images: simple=" + simple, async ({ page }) => {
    await page.goto("/");
    await page.getByLabel("打开文件").setInputFiles({
      name: "nested-field.docx",
      mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      buffer: Buffer.from(nestedFieldFixture(simple)),
    });
    await expect(page.locator(".office-app")).toHaveAttribute("data-loading", "false"); await expect(page.locator(".office-app")).toHaveAttribute("data-revision", "0");
    const frame = page.frameLocator("iframe");
    await expect(frame.getByText("outer result", { exact: true })).toBeVisible();
    await expect(frame.getByText("cache", { exact: true })).toBeHidden();
    await expect(frame.getByText("child cache", { exact: true })).toBeHidden();
    await expect(frame.locator("br")).toBeHidden();
    await expect(frame.locator("img")).toBeHidden();
    await expect(frame.locator("[data-word-tab]")).toBeHidden();
    expect(await frame.locator("p").evaluate((node) => node.getBoundingClientRect().height))
      .toBeLessThan(40);
  });
}

test("uses each section's width and margins in continuous view", async ({ page }) => {
  const bytes = mutateDocx("04-styles-sections.docx", {
    "word/document.xml": (source) => source
      .replace('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>',
        '<w:sectPr><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>')
      .replace('<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"',
        '<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/><w:pgMar w:top="720" w:right="960" w:bottom="1080" w:left="960"')
      .replace("第二节正文", "第二节正文 ".repeat(25)),
  });
  await page.goto("/");
  await page.locator("main").evaluate((node) => {
    node.style.width = "calc(100% - 2rem)";
  });
  await page.getByLabel("打开文件").setInputFiles({
    name: "section-geometry.docx",
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    buffer: Buffer.from(bytes),
  });
  await expect(page.locator(".office-app")).toHaveAttribute("data-loading", "false"); await expect(page.locator(".office-app")).toHaveAttribute("data-revision", "0");
  const frame = page.frameLocator("iframe");
  const geometry = await frame.locator("section").evaluateAll((sections) => sections.map((node) => {
    const computed = getComputedStyle(node);
    return {
      width: parseFloat(computed.width) * 72 / 96,
      padding: [computed.paddingTop, computed.paddingRight, computed.paddingBottom, computed.paddingLeft]
        .map((value) => parseFloat(value) * 72 / 96),
      contentWidth: node.querySelector("p:last-child")!.getBoundingClientRect().width * 72 / 96,
    };
  }));
  expect(geometry).toHaveLength(2);
  expect(geometry[0]!.width).toBeCloseTo(595.3, 1);
  expect(geometry[0]!.padding).toEqual([72, 72, 72, 72]);
  expect(geometry[0]!.contentWidth).toBeCloseTo(451.3, 1);
  expect(geometry[1]!.width).toBeCloseTo(841.9, 1);
  expect(geometry[1]!.padding).toEqual([36, 48, 54, 48]);
  expect(geometry[1]!.contentWidth).toBeCloseTo(745.9, 1);
  await expandDocumentScreenshot(page);
  await expect(frame.locator("article")).toHaveScreenshot("section-geometry.png", {
    animations: "disabled", maxDiffPixelRatio: 0.005, threshold: 0.2,
  });
});

for (const articleRoot of [false, true]) {
  test("applies package section geometry and document defaults: article root=" + articleRoot, async ({ page }) => {
    await page.goto("/");
    await page.getByLabel("打开文件").setInputFiles({
      name: "section-package.zip",
      mimeType: "application/zip",
      buffer: Buffer.from(sectionPackageFixture(articleRoot)),
    });
    await expect(page.locator(".office-app")).toHaveAttribute("data-loading", "false"); await expect(page.locator(".office-app")).toHaveAttribute("data-revision", "0");
    const frame = page.frameLocator("iframe");
    const widths = await frame.locator(articleRoot ? "article" : "section").evaluateAll(
      (nodes) => nodes.map((node) => node.getBoundingClientRect().width * 72 / 96),
    );
    expect(widths[0]).toBeCloseTo(841.9, 1);
    if (!articleRoot) expect(widths[1]).toBeCloseTo(595.3, 1);
  });
}
