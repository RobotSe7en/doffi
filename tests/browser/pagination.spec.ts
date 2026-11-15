import { registerWordListClipboardTests } from "./word-list-clipboard-cases.js";
import { registerWordCrossClipboardTests } from "./word-cross-clipboard-cases.js";
import { registerWordMoveClipboardTests } from "./word-move-clipboard-cases.js";
import { registerWordObjectClipboardTests } from "./word-object-clipboard-cases.js";
import { registerWordNoteClipboardTests } from "./word-note-clipboard-cases.js";
import {
  floatingTableFixture,
  floatingHeaderFixture,
} from "../../packages/office/tests/floating-table-fixture.js";
import { nativeCommentFixture } from "../../packages/office/tests/native-comment-fixture.js";
import { rowExceptionFixture } from "../../packages/office/tests/row-exception-fixture.js";
import { rowSpacingFixture } from "../../packages/office/tests/row-spacing-fixture.js";
import { tableSpacingFixture } from "../../packages/office/tests/table-spacing-fixture.js";
import { tableBidiFixture } from "../../packages/office/tests/table-bidi-fixture.js";
import { continuationCellFixture } from "../../packages/office/tests/continuation-cell-fixture.js";
import { readFileSync, writeFileSync } from "node:fs";
import {
  tableStylesFixture,
  rowStylesFixture,
  conditionalFormatFixture,
} from "../../packages/office/tests/table-style-fixture.js";
import { importDocx, exportDocx } from "@doffi/office";
import { openSdkHost } from "./sdk-host.js";
import { rowGapFixture } from "../../packages/office/tests/row-gap-fixture.js";
import {
  mutateDocx,
  body as nativeBody,
  styles as nativeStyles,
} from "../../packages/office/tests/regression-fixtures.js";
import { fileURLToPath } from "node:url";
import { expect, test } from "@playwright/test";
import type { Page } from "@playwright/test";
import { build } from "vite";
import type { DocumentPackage } from "@doffi/core";
import type { DocumentPreviewOptions } from "@doffi/editor";
import { fieldsFiles, longTable, pagedFiles } from "./fixtures/pagination-files.js";
import type { PaginationHost } from "./fixtures/pagination-host.js";
import {
  createDocumentSession,
  createWordDocument,
  openDocument as openWordDocument,
} from "@doffi/core";
declare global {
  interface Window {
    paginationHost: PaginationHost;
  }
}
let code: string;
function continuousColumnsFiles(lines = 6, noBalance = false, hardBreak = false): DocumentPackage {
  const session = createDocumentSession(
      openWordDocument(
        createWordDocument({
          documentId: "continuous-columns",
          pageSetup: {
            unit: "pt",
            width: 300,
            height: 300,
            margins: { top: 24, right: 24, bottom: 24, left: 24 },
          },
        }),
        { kind: "document" },
      ),
    ),
    apply = (operations: import("@doffi/core").DocumentOperation[]) =>
      session.commit(
        session.prepare({
          protocolVersion: 1,
          documentId: "continuous-columns",
          baseRevision: session.snapshot().manifest.revision,
          operations,
        }),
      );
  apply([
    {
      op: "replaceFragment",
      nodeId: "section-root-1",
      html:
        '<p class="band-title" style="line-height:16pt;margin:0">TITLE</p><p class="band-columns" style="line-height:16pt;margin:0' +
        (hardBreak ? ";break-before:page" : "") +
        '">' +
        Array.from({ length: lines }, (_, index) => "L" + index).join("<br>") +
        '</p><p class="band-tail" style="line-height:16pt;margin:0">TAIL</p>',
    },
  ]);
  const [titleId, columnsId] = session.getNode("section-root-1")!.childNodeIds;
  apply([{ op: "splitWordSection", nodeId: titleId!, breakType: "continuous" }]);
  const middle = session.snapshot().metadata.sections![1]!.rootNodeId;
  apply([
    { op: "splitWordSection", nodeId: columnsId!, breakType: "continuous" },
    {
      op: "setWordSection",
      nodeId: middle,
      changes: { columns: { count: 2, gap: 24, separator: true } },
    },
    { op: "setWordOptions", nodeId: "document-root", noColumnBalance: noBalance },
  ]);
  return session.serialize();
}
async function bandTextBounds(
  page: Page,
  token: string,
): Promise<{ x: number; y: number; bottom: number }> {
  return page
    .frameLocator("iframe")
    .locator('[data-office-kind="document"]')
    .evaluate((root, token) => {
      const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
      for (let node = walker.nextNode(); node; node = walker.nextNode()) {
        const from = node.textContent?.indexOf(token) ?? -1;
        if (from < 0) continue;
        const range = document.createRange();
        range.setStart(node, from);
        range.setEnd(node, from + token.length);
        const rect = range.getBoundingClientRect();
        return { x: rect.x, y: rect.y, bottom: rect.bottom };
      }
      throw new Error("Missing source token: " + token);
    }, token);
}
function nextColumnFiles(count: number): DocumentPackage {
  const session = createDocumentSession(
      openWordDocument(
        createWordDocument({
          documentId: "next-column",
          pageSetup: {
            unit: "pt",
            width: 300,
            height: 300,
            margins: { top: 24, right: 24, bottom: 24, left: 24 },
          },
        }),
        { kind: "document" },
      ),
    ),
    apply = (operations: import("@doffi/core").DocumentOperation[]) =>
      session.commit(
        session.prepare({
          protocolVersion: 1,
          documentId: "next-column",
          baseRevision: session.snapshot().manifest.revision,
          operations,
        }),
      );
  apply([
    {
      op: "replaceFragment",
      nodeId: "section-root-1",
      html: '<p class="column-first" style="line-height:16pt;margin:0">FIRST0<br>FIRST1</p><p class="column-next" style="line-height:16pt;margin:0">NEXT0<br>NEXT1<br>NEXT2</p><p class="column-last" style="line-height:16pt;margin:0">LAST</p>',
    },
    { op: "setWordSection", nodeId: "section-root-1", changes: { columns: { count, gap: 24 } } },
  ]);
  const [first, next] = session.getNode("section-root-1")!.childNodeIds;
  apply([{ op: "splitWordSection", nodeId: first!, breakType: "nextColumn" }]);
  apply([{ op: "splitWordSection", nodeId: next!, breakType: "nextColumn" }]);
  return session.serialize();
}
for (const count of [1, 2])
  test(
    "next-column section breaks use remaining columns, advance full pages and expose editable section starts " +
      count,
    async ({ page }) => {
      await mount(page, nextColumnFiles(count), "edit");
      const pages = count === 1 ? 3 : 2;
      expect((await page.evaluate(() => window.paginationHost.layout())).pages).toHaveLength(pages);
      const first = await bandTextBounds(page, "FIRST0"),
        next = await bandTextBounds(page, "NEXT0");
      if (count === 2) {
        expect(next.x - first.x).toBeGreaterThan(100);
        expect(Math.abs(next.y - first.y)).toBeLessThan(2);
      } else expect(next.y - first.y).toBeGreaterThan(350);
      await page.frameLocator("iframe").locator(".column-last").click();
      await page.evaluate(() =>
        (window.paginationHost.view as import("@doffi/editor").DocumentTextEditor).execute({
          type: "word",
          action: "pageSettings",
        }),
      );
      await expect(page.getByLabel("节起始位置", { exact: true })).toHaveValue("nextColumn");
      await page.getByLabel("节起始位置", { exact: true }).selectOption("continuous");
      await page.getByRole("button", { name: "应用节起始位置", exact: true }).click();
      expect((await page.evaluate(() => window.paginationHost.layout())).pages).toHaveLength(
        pages - 1,
      );
      await page.evaluate(() => window.paginationHost.session!.undo());
      expect((await page.evaluate(() => window.paginationHost.layout())).pages).toHaveLength(pages);
      expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
    },
  );
test("continuous single and multi-column sections share one page, balance their end and keep second-column input in the same source", async ({
  page,
}) => {
  await mount(page, continuousColumnsFiles(), "edit");
  expect((await page.evaluate(() => window.paginationHost.layout())).pages).toHaveLength(1);
  const first = await bandTextBounds(page, "L0"),
    second = await bandTextBounds(page, "L3"),
    tail = await bandTextBounds(page, "TAIL"),
    last = await bandTextBounds(page, "L5");
  expect(second.x - first.x).toBeGreaterThan(100);
  expect(Math.abs(second.y - first.y)).toBeLessThan(2);
  expect(tail.y - last.y).toBeGreaterThan((16 * 4) / 3 - 0.5);
  expect(Math.abs(tail.x - first.x)).toBeLessThan(2);
  await page
    .frameLocator("iframe")
    .locator(".band-columns")
    .evaluate((root) => {
      const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
      for (let node = walker.nextNode(); node; node = walker.nextNode()) {
        const at = node.textContent?.indexOf("L4") ?? -1;
        if (at < 0) continue;
        (root.closest("[data-office-editor]") as HTMLElement).focus();
        const range = document.createRange();
        range.setStart(node, at + 2);
        range.collapse(true);
        const selection = document.getSelection()!;
        selection.removeAllRanges();
        selection.addRange(range);
        return;
      }
      throw new Error("Missing L4");
    });
  await page.evaluate(() => window.paginationHost.view!.settled());
  await page.keyboard.type("输入");
  const columnsId = await page
    .frameLocator("iframe")
    .locator(".band-columns")
    .getAttribute("data-node-id");
  await expect
    .poll(() => page.evaluate((id) => window.paginationHost.session!.getNode(id!)!.text, columnsId))
    .toContain("L4输入");
  expect(await page.evaluate(() => window.paginationHost.session!.snapshot().html)).not.toMatch(
    /sdk-section-band|sdk-column-exclusion|position:absolute/,
  );
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
});
test("continuous column tables repeat their header once per used column and leave the next section below the rows", async ({
  page,
}) => {
  const session = createDocumentSession(
      openWordDocument(continuousColumnsFiles(), { kind: "document" }),
    ),
    section = session.snapshot().metadata.sections![1]!.rootNodeId;
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: session.snapshot().manifest.documentId,
      baseRevision: session.snapshot().manifest.revision,
      operations: [
        {
          op: "replaceFragment",
          nodeId: section,
          html:
            '<table style="width:96pt;border-collapse:collapse"><tr data-word-repeat-header="true"><th><p style="line-height:16pt;margin:0">HEADER</p></th></tr>' +
            Array.from(
              { length: 20 },
              (_, index) =>
                '<tr><td><p style="line-height:16pt;margin:0">ROW-' + index + "</p></td></tr>",
            ).join("") +
            "</table>",
        },
      ],
    }),
  );
  await mount(page, session.serialize());
  expect((await page.evaluate(() => window.paginationHost.layout())).pages).toHaveLength(1);
  await expect(page.frameLocator("iframe").locator(".sdk-repeat-header")).toHaveCount(1);
  const last = await bandTextBounds(page, "ROW-19"),
    tail = await bandTextBounds(page, "TAIL");
  expect(tail.y - last.y).toBeGreaterThan((16 * 4) / 3 - 0.5);
  expect(await page.evaluate(() => window.paginationHost.model!.snapshot().html)).not.toMatch(
    /sdk-repeat-header|sdk-section-band/,
  );
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
});
test("column balance controls retain unbalanced flow and undo shared-page geometry", async ({
  page,
}) => {
  await mount(page, continuousColumnsFiles(6, true), "edit");
  const first = await bandTextBounds(page, "L0"),
    fourth = await bandTextBounds(page, "L3"),
    tail = await bandTextBounds(page, "TAIL");
  expect(Math.abs(fourth.x - first.x)).toBeLessThan(2);
  await page.frameLocator("iframe").locator(".band-columns").click();
  await page.evaluate(() =>
    (window.paginationHost.view as import("@doffi/editor").DocumentTextEditor).execute({
      type: "word",
      action: "pageSettings",
    }),
  );
  await page.getByLabel("节末分栏自动平衡", { exact: true }).check();
  await page.getByRole("button", { name: "应用分栏平衡", exact: true }).click();
  expect((await bandTextBounds(page, "L3")).x - first.x).toBeGreaterThan(100);
  expect((await bandTextBounds(page, "TAIL")).y).toBeLessThan(tail.y - 20);
  await page.evaluate(() => window.paginationHost.session!.undo());
  expect(Math.abs((await bandTextBounds(page, "L3")).x - first.x)).toBeLessThan(2);
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
});
for (const hardBreak of [false, true])
  test(
    "continuous column sections preserve multi-page flow and explicit page breaks " + hardBreak,
    async ({ page }) => {
      await mount(page, continuousColumnsFiles(hardBreak ? 6 : 40, false, hardBreak), "edit");
      const layout = await page.evaluate(() => window.paginationHost.layout());
      expect(layout.pages).toHaveLength(2);
      const nodePages = layout.nodePages,
        tailId = await page
          .frameLocator("iframe")
          .locator(".band-tail")
          .getAttribute("data-node-id"),
        columnsId = await page
          .frameLocator("iframe")
          .locator(".band-columns")
          .getAttribute("data-node-id");
      expect(nodePages[tailId!]).toEqual(["page-2"]);
      if (hardBreak) expect(nodePages[columnsId!]).toEqual(["page-2"]);
      expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
    },
  );
function characterGridFiles(content: string, columns = 1): DocumentPackage {
  const session = createDocumentSession(
    openWordDocument(
      createWordDocument({
        documentId: "fixed-grid",
        pageSetup: {
          unit: "pt",
          width: 300,
          height: 240,
          margins: { top: 24, right: 24, bottom: 24, left: 24 },
        },
      }),
      { kind: "document" },
    ),
  );
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "fixed-grid",
      baseRevision: 0,
      operations: [
        { op: "replaceFragment", nodeId: "section-root-1", html: content },
        {
          op: "setWordSection",
          nodeId: "section-root-1",
          changes: {
            grid: { type: "snapToChars", linePitch: 24, characterSpacing: 8192 },
            ...(columns > 1 ? { columns: { count: columns, gap: 24 } } : {}),
          },
        },
      ],
    }),
  );
  return session.serialize();
}
test("fixed character grids use Normal pitch across font sizes and preserve supplementary ideographs and variation sequences", async ({
  page,
}) => {
  await mount(
    page,
    characterGridFiles('<p class="fixed-cjk" style="font-size:18pt;margin:0">汉𠀀\u{e0100}文</p>'),
    "edit",
  );
  const paragraph = page.frameLocator("iframe").locator(".fixed-cjk"),
    points = await paragraph.evaluate((root) => {
      const textNodes: Text[] = [],
        walker = document.createTreeWalker(root, 4);
      for (let node = walker.nextNode(); node; node = walker.nextNode())
        textNodes.push(node as Text);
      return [0, 1, 5].map((offset) => {
        for (const text of textNodes) {
          if (offset < text.length) {
            const range = document.createRange();
            range.setStart(text, offset);
            range.setEnd(text, offset + (text.data.codePointAt(offset)! > 65535 ? 2 : 1));
            return range.getBoundingClientRect().x;
          }
          offset -= text.length;
        }
        throw new Error("Missing grid character");
      });
    });
  expect(Math.abs(points[1]! - points[0]! - (14 * 4) / 3)).toBeLessThan(1);
  expect(Math.abs(points[2]! - points[1]! - (14 * 4) / 3)).toBeLessThan(1);
  expect(
    await paragraph
      .locator(".sdk-grid-run")
      .evaluateAll((nodes) =>
        nodes.reduce((sum, node) => sum + node.getBoundingClientRect().width, 0),
      ),
  ).toBeCloseTo((3 * 14 * 4) / 3, 1);
  const nodeId = await paragraph.getAttribute("data-node-id");
  expect(
    await page.evaluate((id) => window.paginationHost.session!.getNode(id!)!.text, nodeId),
  ).toBe("汉𠀀\u{e0100}文");
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
});
test("Latin blocks span character runs in whole grid units while complex-script blocks stay left aligned", async ({
  page,
}) => {
  await mount(
    page,
    characterGridFiles(
      '<p class="fixed-latin" style="margin:0">汉AB<strong>cd</strong>EF汉</p><p class="fixed-cs" style="margin:0">汉مرحبا汉</p><p class="fixed-indic" style="margin:0">汉नमस्ते汉</p>',
    ),
  );
  for (const selector of [".fixed-latin", ".fixed-cs", ".fixed-indic"]) {
    const geometry = await page
      .frameLocator("iframe")
      .locator(selector)
      .evaluate((root) => {
        const walker = document.createTreeWalker(root, 4),
          latin: DOMRect[] = [],
          cjk: DOMRect[] = [];
        for (let node = walker.nextNode(); node; node = walker.nextNode()) {
          const range = document.createRange();
          range.selectNodeContents(node);
          const rect = range.getBoundingClientRect();
          if (node.textContent === "汉") cjk.push(rect);
          else latin.push(rect);
        }
        return {
          left: root.getBoundingClientRect().left,
          cjk: cjk.map((rect) => rect.x),
          first: Math.min(...latin.map((rect) => rect.left)),
          width: latin.reduce((sum, rect) => sum + rect.width, 0),
        };
      });
    const pitch = (14 * 4) / 3,
      units = Math.ceil((geometry.width - 0.02) / pitch);
    expect(geometry.cjk[1]! - geometry.cjk[0]!).toBeCloseTo((units + 1) * pitch, 0);
    expect(geometry.first - geometry.left - pitch).toBeCloseTo(
      selector !== ".fixed-latin" ? 0 : (units * pitch - geometry.width) / 2,
      0,
    );
  }
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
});
test("mixed font sizes reserve line-grid multiples only on the lines containing large runs", async ({
  page,
}) => {
  await mount(
    page,
    characterGridFiles(
      '<p class="mixed-grid" style="line-height:1.2;margin:0"><span>A</span><br><span style="font-size:30pt">大A</span><br><span>A</span></p>',
    ),
  );
  const geometry = await page
    .frameLocator("iframe")
    .locator(".mixed-grid")
    .evaluate((root) => {
      const nodes = [...root.querySelectorAll("span")].filter(
          (node) =>
            node.textContent === "A" &&
            node.childNodes.length === 1 &&
            node.firstChild?.nodeType === 3,
        ),
        rects = nodes.map((node) => {
          const range = document.createRange();
          range.selectNodeContents(node);
          return range.getBoundingClientRect();
        });
      return {
        first: rects[0]!.top,
        last: rects.at(-1)!.top,
        paragraph: parseFloat(getComputedStyle(root).lineHeight),
        large: [...root.querySelectorAll("span")]
          .filter((node) => node.style.fontSize === "30pt")
          .map((node) => parseFloat(getComputedStyle(node).lineHeight)),
      };
    });
  expect(geometry.paragraph).toBeCloseTo(32);
  expect(geometry.large[0]).toBeCloseTo(64);
  expect(geometry.last - geometry.first).toBeCloseTo(96, 0);
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
});
test("literal newlines terminate grid blocks without moving end padding onto the next line", async ({
  page,
}) => {
  await mount(
    page,
    characterGridFiles(
      '<p class="literal-grid" style="margin:0;white-space:pre-wrap">ABC\nDEFG</p>',
    ),
  );
  const result = await page
    .frameLocator("iframe")
    .locator(".literal-grid")
    .evaluate((root) => {
      const blocks = [...root.querySelectorAll<HTMLElement>(".sdk-grid-run")].filter(
        (node) => node.textContent === "ABC" || node.textContent === "DEFG",
      );
      return {
        height: root.getBoundingClientRect().height,
        blocks: blocks.map((node) => {
          const range = document.createRange();
          range.selectNodeContents(node);
          return {
            width: node.getBoundingClientRect().width,
            natural: range.getBoundingClientRect().width,
          };
        }),
      };
    });
  expect(result.height).toBeCloseTo(64, 0);
  expect(result.blocks).toHaveLength(2);
  for (const block of result.blocks)
    expect(block.width).toBeCloseTo(
      (Math.ceil((block.natural - 0.02) / ((14 * 4) / 3)) * 14 * 4) / 3,
      0,
    );
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
});
for (const direction of ["rtl", "tbRl", "btLr"])
  test(
    "character grid units follow the inline axis of bidirectional and vertical text " + direction,
    async ({ page }) => {
      const text = "汉AB<strong>cd</strong>汉",
        content =
          direction === "rtl"
            ? '<p class="direction-grid" dir="rtl" style="margin:0">' + text + "</p>"
            : '<table style="width:80pt;height:120pt;border-collapse:collapse"><tr><td data-word-text-direction="' +
              direction +
              '" style="height:120pt"><p class="direction-grid" style="margin:0">' +
              text +
              "</p></td></tr></table>";
      await mount(page, characterGridFiles(content));
      const result = await page
        .frameLocator("iframe")
        .locator(".direction-grid")
        .evaluate((root, vertical) => {
          const cjk = [...root.querySelectorAll<HTMLElement>(".sdk-grid-run")].filter(
              (node) => node.textContent === "汉",
            ),
            walker = document.createTreeWalker(root, 4);
          let natural = 0;
          for (let node = walker.nextNode(); node; node = walker.nextNode()) {
            if (node.textContent === "汉") continue;
            const range = document.createRange();
            range.selectNodeContents(node);
            const box = range.getBoundingClientRect();
            natural += vertical ? box.height : box.width;
          }
          const a = cjk[0]!.getBoundingClientRect(),
            b = cjk[1]!.getBoundingClientRect();
          return { gap: Math.abs(vertical ? b.top - a.top : b.left - a.left), natural };
        }, direction !== "rtl");
      const pitch = (14 * 4) / 3;
      expect(result.gap).toBeCloseTo((1 + Math.ceil((result.natural - 0.02) / pitch)) * pitch, 0);
      expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
    },
  );
test("fixed grid measurement follows regenerated field values while their source cache stays unchanged", async ({
  page,
}) => {
  const session = createDocumentSession(
      openWordDocument(
        characterGridFiles(
          '<p class="fixed-field" style="margin:0;font-family:&#34;Liberation Mono&#34;;--word-font-ascii:&#34;Liberation Mono&#34;;--word-font-eastAsia:&#34;Noto Serif CJK SC&#34;">汉汉</p><p style="margin:0">' +
            Array.from({ length: 80 }, (_, index) => "L" + index).join("<br>") +
            "</p>",
        ),
        { kind: "document" },
      ),
    ),
    paragraph = session.getNode("section-root-1")!.childNodeIds[0]!,
    snapshot = session.snapshot();
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: snapshot.manifest.documentId,
      baseRevision: snapshot.manifest.revision,
      operations: [
        {
          op: "insertWordField",
          nodeId: paragraph,
          offset: 1,
          instruction: "NUMPAGES",
          cachedText: "0",
        },
      ],
    }),
  );
  await mount(page, session.serialize());
  const frame = page.frameLocator("iframe"),
    field = frame.locator(".fixed-field .sdk-field-display");
  await expect(field).toHaveText(
    String((await page.evaluate(() => window.paginationHost.layout())).pages.length),
  );
  expect((await field.textContent())!.length).toBeGreaterThan(1);
  const geometry = await frame.locator(".fixed-field").evaluate((root) => {
    const cjk = [...root.querySelectorAll<HTMLElement>(".sdk-grid-run")].filter(
        (node) => node.textContent === "汉",
      ),
      display = root.querySelector(".sdk-field-display")!,
      walker = document.createTreeWalker(display, 4);
    let width = 0;
    for (let node = walker.nextNode(); node; node = walker.nextNode()) {
      const range = document.createRange();
      range.selectNodeContents(node);
      width += range.getBoundingClientRect().width;
    }
    return {
      gap: cjk[1]!.getBoundingClientRect().left - cjk[0]!.getBoundingClientRect().left,
      width,
    };
  });
  expect(geometry.gap).toBeCloseTo(
    ((1 + Math.ceil((geometry.width - 0.02) / ((14 * 4) / 3))) * 14 * 4) / 3,
    0,
  );
  expect(
    await page.evaluate((id) => window.paginationHost.model!.getNode(id)!.text, paragraph),
  ).toBe("汉0汉");
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
});
test("fixed grid blocks wrap through columns and retain exact input and undo after repeated measurement", async ({
  page,
}) => {
  const original = "latin ".repeat(200) + "汉𠀀\u{e0100}";
  await mount(
    page,
    characterGridFiles('<p class="long-grid" style="margin:0">' + original + "</p>", 2),
    "edit",
  );
  const frame = page.frameLocator("iframe"),
    paragraph = frame.locator(".long-grid");
  expect((await page.evaluate(() => window.paginationHost.layout())).pages.length).toBeGreaterThan(
    1,
  );
  await paragraph.evaluate((root) => {
    (root.closest("[data-office-editor]") as HTMLElement).focus();
    const walker = document.createTreeWalker(root, 4);
    let offset = 200;
    for (let node = walker.nextNode(); node; node = walker.nextNode()) {
      if (offset < node.textContent!.length) {
        const range = document.createRange();
        range.setStart(node, offset);
        range.collapse(true);
        document.getSelection()!.removeAllRanges();
        document.getSelection()!.addRange(range);
        return;
      }
      offset -= node.textContent!.length;
    }
  });
  await page.evaluate(() => window.paginationHost.view!.settled());
  await page.keyboard.insertText("EDIT");
  const nodeId = await paragraph.getAttribute("data-node-id");
  expect(
    await page.evaluate((id) => window.paginationHost.session!.getNode(id!)!.text, nodeId),
  ).toBe(original.slice(0, 200) + "EDIT" + original.slice(200));
  await page.keyboard.type("X");
  await page.evaluate(() => window.paginationHost.session!.undo());
  await page.evaluate(() => window.paginationHost.session!.undo());
  expect(
    await page.evaluate((id) => window.paginationHost.session!.getNode(id!)!.text, nodeId),
  ).toBe(original);
  expect(await page.evaluate(() => window.paginationHost.session!.snapshot().html)).not.toMatch(
    /sdk-grid|letter-spacing|padding-left/,
  );
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
  expect(
    await page.evaluate(() =>
      window.paginationHost.layout().diagnostics.filter((item) => item.code === "NativeFormatting"),
    ),
  ).toEqual([]);
});
test("document grid keeps paragraph line and run character snapping independent and applies table lines only when enabled", async ({
  page,
}) => {
  const session = createDocumentSession(
    openWordDocument(createWordDocument({ documentId: "grid-rules" }), { kind: "document" }),
  );
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "grid-rules",
      baseRevision: 0,
      operations: [
        {
          op: "replaceFragment",
          nodeId: "section-root-1",
          html: '<p class="grid-off" style="--word-paragraph-snap-to-grid:0;line-height:12pt">中文<span style="--word-run-snap-to-grid:0">关闭字符</span></p><p class="grid-exact" style="--word-line-rule:exact;line-height:10pt">固定行距</p><table style="border-collapse:collapse"><tr><td><p class="grid-cell" style="line-height:12pt">表格</p></td></tr></table>',
        },
        {
          op: "setWordSection",
          nodeId: "section-root-1",
          changes: { grid: { type: "snapToChars", linePitch: 24, characterSpacing: 8192 } },
        },
      ],
    }),
  );
  await mount(page, session.serialize(), "edit");
  const frame = page.frameLocator("iframe");
  expect(
    await frame
      .locator(".grid-off")
      .evaluate((node) => parseFloat(getComputedStyle(node).lineHeight)),
  ).toBeCloseTo(16);
  expect(
    await frame
      .locator('.grid-off span[style*="letter-spacing"]')
      .first()
      .evaluate((node) => parseFloat(getComputedStyle(node).letterSpacing)),
  ).toBeCloseTo(8 / 3);
  await expect(frame.locator(".grid-off > span").last()).toContainText("关闭字符");
  expect(
    await frame
      .locator(".grid-exact")
      .evaluate((node) => parseFloat(getComputedStyle(node).lineHeight)),
  ).toBeCloseTo(40 / 3, 1);
  expect(
    await frame
      .locator(".grid-cell")
      .evaluate((node) => parseFloat(getComputedStyle(node).lineHeight)),
  ).toBeCloseTo(16);
  await frame.locator(".grid-off").click();
  await page.evaluate(() =>
    (window.paginationHost.view as import("@doffi/editor").DocumentTextEditor).execute({
      type: "word",
      action: "pageSettings",
    }),
  );
  await page.getByLabel("表格文字对齐行网格", { exact: true }).check();
  await page.getByRole("button", { name: "应用表格行网格", exact: true }).click();
  expect(
    await frame
      .locator(".grid-cell")
      .evaluate((node) => parseFloat(getComputedStyle(node).lineHeight)),
  ).toBeCloseTo(32);
  await page.evaluate(() => window.paginationHost.session!.undo());
  expect(
    await frame
      .locator(".grid-cell")
      .evaluate((node) => parseFloat(getComputedStyle(node).lineHeight)),
  ).toBeCloseTo(16);
  expect(await page.evaluate(() => window.paginationHost.session!.snapshot().html)).not.toMatch(
    /letter-spacing|sdk-grid/,
  );
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
});
test("find and replace follow the revision view and preserve hidden deletion source through one undo", async ({
  page,
}) => {
  const session = createDocumentSession(
    openWordDocument(createWordDocument({ documentId: "visible-replace" }), { kind: "document" }),
  );
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "visible-replace",
      baseRevision: 0,
      operations: [
        {
          op: "replaceText",
          nodeId: "paragraph-1",
          from: 0,
          to: 0,
          expectedText: "",
          text: "abXcd",
        },
        {
          op: "trackWordText",
          nodeId: "paragraph-1",
          from: 2,
          to: 3,
          expectedText: "X",
          text: "",
          author: "old",
        },
        { op: "setWordReview", nodeId: "document-root", trackChanges: true, author: "new" },
      ],
    }),
  );
  await mount(page, session.serialize(), "edit");
  await page.evaluate(() =>
    (window.paginationHost.view as import("@doffi/editor").DocumentTextEditor).execute({
      type: "word",
      action: "review",
    }),
  );
  await page.getByLabel("修订显示", { exact: true }).selectOption("final");
  await page.evaluate(() =>
    (window.paginationHost.view as import("@doffi/editor").DocumentTextEditor).execute({
      type: "word",
      action: "replace",
    }),
  );
  await page.getByLabel("查找内容", { exact: true }).fill("abcd");
  await expect(page.getByRole("region", { name: "查找和替换" })).toContainText("1 处匹配");
  await page.getByRole("button", { name: "1 · abcd", exact: true }).click();
  expect(
    await page
      .frameLocator("iframe")
      .locator("[data-office-editor]")
      .evaluate(() => document.getSelection()!.toString()),
  ).toBe("abcd");
  await page.getByLabel("替换为", { exact: true }).fill("Z");
  await page.getByRole("button", { name: "全部替换", exact: true }).click();
  expect(
    await page.evaluate(() =>
      window.paginationHost
        .session!.snapshot()
        .metadata.word!.revisions.filter((item) => item.author === "old")
        .map((item) => window.paginationHost.session!.getNode(item.nodeId)!.text),
    ),
  ).toEqual(["X"]);
  expect(
    await page
      .frameLocator("iframe")
      .locator('[data-node-id="paragraph-1"]')
      .evaluate((node) => (node as HTMLElement).innerText),
  ).toBe("Z");
  await page.evaluate(() => window.paginationHost.session!.undo());
  await page.getByLabel("查找内容", { exact: true }).fill("abcd");
  await expect(page.getByRole("region", { name: "查找和替换" })).toContainText("1 处匹配");
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
});
test("inserting a table from a column paragraph uses its section width in the actual editor", async ({
  page,
}) => {
  await mount(page, continuousColumnsFiles(), "edit");
  await page
    .frameLocator("iframe")
    .locator(".band-columns")
    .click({ position: { x: 5, y: 5 } });
  await page.evaluate(() =>
    (window.paginationHost.view as import("@doffi/editor").DocumentTextEditor).execute({
      type: "word",
      action: "insertTable",
    }),
  );
  await page.getByRole("button", { name: "插入表格", exact: true }).click();
  const grid = await page.evaluate(
    () => window.paginationHost.session!.snapshot().metadata.word!.tables[0]!.grid,
  );
  expect(grid).toEqual([38, 38, 38]);
  expect(
    await page
      .frameLocator("iframe")
      .locator("[data-office-editor] table")
      .evaluate((table) => Math.abs(table.getBoundingClientRect().width * 0.75 - 114)),
  ).toBeLessThan(1);
  await page.evaluate(() => window.paginationHost.session!.undo());
  await expect(page.frameLocator("iframe").locator("[data-office-editor] table")).toHaveCount(0);
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
});
test("native omitted grid edges share columns, skip keyboard navigation and survive source input", async ({
  page,
}) => {
  const files = importDocx(rowGapFixture(), {
    documentId: "row-gap-ui",
    unsupportedContent: "reject",
  });
  await mount(page, files, "edit");
  const frame = page.frameLocator("iframe"),
    table = frame.locator("table").first();
  expect(await table.evaluate((node) => node.getBoundingClientRect().width * 0.75)).toBeCloseTo(
    216,
    0,
  );
  const cells = await table.locator("tr").evaluateAll((rows) =>
    rows.map((row) =>
      [...row.querySelectorAll("td")].map((cell) => ({
        x: cell.getBoundingClientRect().x,
        gap: cell.getAttribute("data-word-grid-gap"),
        visibility: getComputedStyle(cell).visibility,
      })),
    ),
  );
  expect(cells[1]![1]!.x).toBeCloseTo(cells[0]![1]!.x, 0);
  expect(cells[1]![0]!.visibility).toBe("hidden");
  expect(cells[1]![2]!.visibility).toBe("hidden");
  await table.locator("td").filter({ hasText: "C" }).click();
  await page.keyboard.press("Tab");
  await page.keyboard.insertText("EDIT");
  const snapshot = await page.evaluate(() => window.paginationHost.session!.snapshot());
  expect(snapshot.html).toContain("EDITX");
  expect(snapshot.html).toContain('data-word-grid-gap="before"');
  expect(snapshot.html).toContain('data-word-grid-gap="after"');
  expect(snapshot.html).not.toMatch(/visibility:hidden|contenteditable/);
  await page.evaluate(() => window.paginationHost.session!.undo());
  expect(await page.evaluate(() => window.paginationHost.session!.snapshot().html)).not.toContain(
    "EDIT",
  );
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
});
test("original row-gap formatting projects old widths while typing retains the current table source", async ({
  page,
}) => {
  const session = createDocumentSession(
      openWordDocument(
        importDocx(rowGapFixture(), { documentId: "gap-format-ui", unsupportedContent: "reject" }),
        { kind: "document" },
      ),
    ),
    tableId = session.snapshot().metadata.word!.tables[0]!.nodeId,
    root = session.getNode(session.snapshot().metadata.word!.stories[0]!.rootNodeId)!.parentNodeId!;
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "gap-format-ui",
      baseRevision: 0,
      operations: [
        { op: "setWordReview", nodeId: root, trackChanges: true, author: "A" },
        { op: "editWordTable", nodeId: tableId, change: { kind: "resize", grid: [60, 80, 76] } },
      ],
    }),
  );
  await mount(page, session.serialize(), "edit");
  const frame = page.frameLocator("iframe"),
    gap = frame.locator('[data-word-grid-gap="before"]');
  expect(await gap.evaluate((node) => parseFloat((node as HTMLElement).style.width))).toBe(60);
  await frame.locator("td").filter({ hasText: "X" }).click();
  await page.evaluate(() =>
    (window.paginationHost.view as import("@doffi/editor").DocumentTextEditor).execute({
      type: "word",
      action: "review",
    }),
  );
  await page.getByLabel("修订显示", { exact: true }).selectOption("original");
  await expect
    .poll(() => gap.evaluate((node) => parseFloat((node as HTMLElement).style.width)))
    .toBe(72);
  await frame.locator("td").filter({ hasText: "X" }).click();
  await page.keyboard.press("End");
  await page.keyboard.insertText("EDIT");
  const snapshot = await page.evaluate(() => window.paginationHost.session!.snapshot());
  expect(snapshot.metadata.word!.tables[0]!.grid).toEqual([60, 80, 76]);
  expect(snapshot.html).toContain("EDIT");
  expect(snapshot.html).not.toMatch(/visibility:hidden|contenteditable/);
  const markup = snapshot.html.match(/<td[^>]*data-word-grid-gap="before"[^>]*>/)![0];
  expect(markup).toContain("width:60pt");
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
});
test("unmerging a native continuation transfers its paragraph formats into an editable cell with shared undo", async ({
  page,
}) => {
  const bytes = mutateDocx("01-text-numbering.docx", {
      "word/styles.xml": () =>
        nativeStyles(
          '<w:style w:type="paragraph" w:styleId="Normal" w:default="1"/><w:style w:type="paragraph" w:styleId="Empty"><w:basedOn w:val="Normal"/><w:rPr><w:sz w:val="32"/></w:rPr></w:style>',
        ),
      "word/document.xml": (source) =>
        nativeBody(
          source,
          '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>TOP</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p><w:pPr><w:pStyle w:val="Empty"/><w:numPr><w:numId w:val="0"/></w:numPr><w:rPr><w:sz w:val="36"/></w:rPr></w:pPr><w:r><w:rPr><w:sz w:val="28"/></w:rPr></w:r></w:p></w:tc></w:tr></w:tbl>',
        ),
    }),
    session = createDocumentSession(
      openWordDocument(
        importDocx(bytes, { documentId: "unmerge-ui", unsupportedContent: "reject" }),
        { kind: "document" },
      ),
    );
  await mount(page, session.serialize(), "edit");
  const frame = page.frameLocator("iframe");
  await frame.locator("td").first().click();
  await page.evaluate(() =>
    (window.paginationHost.view as import("@doffi/editor").DocumentTextEditor).execute({
      type: "word",
      action: "tableSettings",
    }),
  );
  await page.getByRole("button", { name: "拆分单元格", exact: true }).click();
  await expect(frame.locator("td")).toHaveCount(2);
  const paragraph = frame.locator("td").last().locator('[data-word-style="Empty"]');
  await expect(paragraph).toHaveCount(1);
  expect(await paragraph.getAttribute("data-word-numbering-off")).toBe("true");
  const id = (await paragraph.getAttribute("data-node-id"))!;
  await paragraph.click();
  await page.keyboard.insertText("EDIT");
  await expect
    .poll(() => page.evaluate((id) => window.paginationHost.session!.getNode(id)!.text, id))
    .toBe("EDIT");
  const files = await page.evaluate(() =>
      [...window.paginationHost.session!.serialize()].map(
        ([path, bytes]) => [path, Array.from(bytes)] as const,
      ),
    ),
    model = openWordDocument(new Map(files.map(([path, bytes]) => [path, new Uint8Array(bytes)])), {
      kind: "document",
    }),
    reopened = openWordDocument(
      importDocx(exportDocx(model).bytes, {
        documentId: "unmerge-native",
        unsupportedContent: "reject",
      }),
      { kind: "document" },
    );
  expect(reopened.snapshot().html).toContain("EDIT");
  expect(reopened.snapshot().html).toContain('data-word-numbering-off="true"');
  await page.evaluate(() => window.paginationHost.session!.undo());
  await expect
    .poll(() => page.evaluate((id) => window.paginationHost.session!.getNode(id)!.text, id))
    .toBe("");
  await page.evaluate(() => window.paginationHost.session!.undo());
  await expect(frame.locator("td")).toHaveCount(1);
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
});
test("new table grid defaults switch to a plain named style without retaining direct borders", async ({
  page,
}) => {
  const session = createDocumentSession(
    openWordDocument(createWordDocument({ documentId: "new-grid-ui" }), { kind: "document" }),
  );
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "new-grid-ui",
      baseRevision: 0,
      operations: [
        { op: "insertWordTable", nodeId: "paragraph-1", position: "after", rows: 2, columns: 2 },
      ],
    }),
  );
  await mount(page, session.serialize(), "edit");
  const frame = page.frameLocator("iframe"),
    table = frame.locator("[data-office-editor] table").first();
  await expect(table).toHaveCSS("border-top-style", "solid");
  await table.locator("td").first().click();
  await page.evaluate(() =>
    (window.paginationHost.view as import("@doffi/editor").DocumentTextEditor).execute({
      type: "word",
      action: "tableSettings",
    }),
  );
  await page.getByLabel("表格样式", { exact: true }).selectOption("TableNormal");
  await page.getByRole("button", { name: "应用表格样式", exact: true }).click();
  await expect(table).toHaveCSS("border-top-style", "none");
  await page.evaluate(() => window.paginationHost.session!.undo());
  await expect(table).toHaveCSS("border-top-style", "solid");
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
});
test("conditional native markers stay isolated when editing a merged table in an auxiliary story", async ({
  page,
}) => {
  const session = createDocumentSession(
      openWordDocument(
        importDocx(conditionalFormatFixture(true), {
          documentId: "cnf-story-ui",
          unsupportedContent: "reject",
        }),
        { kind: "document" },
      ),
    ),
    header = session.snapshot().metadata.word!.stories.find((story) => story.kind === "header")!;
  await mount(page, session.serialize(), "edit");
  await page.evaluate(
    (id) =>
      (window.paginationHost.view as import("@doffi/editor").DocumentTextEditor).editStory(id),
    header.storyId,
  );
  const frame = page.frameLocator("iframe"),
    text = frame.locator("[data-office-editor] td p").first();
  await text.click();
  await page.keyboard.press("End");
  await page.keyboard.type("EDIT");
  await expect
    .poll(() => page.evaluate(() => window.paginationHost.session!.snapshot().html))
    .toContain("AEDIT");
  expect(
    await page.evaluate(
      () => window.paginationHost.session!.snapshot().metadata.word!.conditionalFormats!.length,
    ),
  ).toBe(8);
  await page.evaluate(() =>
    (window.paginationHost.view as import("@doffi/editor").DocumentTextEditor).editStory(),
  );
  await expect(frame.locator("[data-office-editor] td p").first()).toHaveText("A");
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
});
test("table conditional style definitions edit and remove the chosen region through shared undo", async ({
  page,
}) => {
  const session = createDocumentSession(
    openWordDocument(
      importDocx(tableStylesFixture(), {
        documentId: "condition-ui",
        unsupportedContent: "reject",
      }),
      { kind: "document" },
    ),
  );
  await mount(page, session.serialize(), "edit");
  const frame = page.frameLocator("iframe"),
    first = frame.locator("[data-office-editor] td").first();
  await first.locator("p").click();
  await page.evaluate(() =>
    (window.paginationHost.view as import("@doffi/editor").DocumentTextEditor).execute({
      type: "word",
      action: "styleSettings",
    }),
  );
  await page.getByLabel("已有样式", { exact: true }).selectOption("Red");
  await page.getByLabel("表格样式区域", { exact: true }).selectOption("firstRow");
  await page.getByLabel("表格样式单元格底纹", { exact: true }).fill("#663399");
  await page.getByLabel("样式重复表头", { exact: true }).selectOption("1");
  await page.getByRole("button", { name: "保存样式", exact: true }).click();
  await expect(first).toHaveCSS("background-color", "rgb(102, 51, 153)");
  await page.getByRole("button", { name: "删除当前条件格式", exact: true }).click();
  await expect(first).toHaveCSS("background-color", "rgb(204, 0, 0)");
  await page.evaluate(() => window.paginationHost.session!.undo());
  await expect(first).toHaveCSS("background-color", "rgb(102, 51, 153)");
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
});
test("conditional table row styles repeat source headers and explicit disabling shares undo and native source", async ({
  page,
}) => {
  const session = createDocumentSession(
      openWordDocument(
        importDocx(rowStylesFixture(), {
          documentId: "styled-headers",
          unsupportedContent: "reject",
        }),
        { kind: "document" },
      ),
    ),
    section = session.snapshot().metadata.sections![0]!;
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "styled-headers",
      baseRevision: 0,
      operations: [
        {
          op: "setWordSection",
          nodeId: section.rootNodeId,
          changes: {
            pageSetup: {
              unit: "pt",
              width: 300,
              height: 240,
              margins: { top: 24, right: 24, bottom: 24, left: 24 },
            },
          },
        },
      ],
    }),
  );
  await mount(page, session.serialize(), "edit");
  const frame = page.frameLocator("iframe"),
    headers = frame.locator(".sdk-repeat-header");
  await expect.poll(() => headers.count()).toBeGreaterThan(0);
  const count = await headers.count();
  expect(await page.evaluate(() => window.paginationHost.session!.snapshot().html)).not.toContain(
    "sdk-repeat-header",
  );
  await frame.locator("[data-office-editor] td p").first().click();
  await page.evaluate(() =>
    (window.paginationHost.view as import("@doffi/editor").DocumentTextEditor).execute({
      type: "word",
      action: "tableSettings",
    }),
  );
  await expect(page.getByLabel("重复表头行数", { exact: true })).toHaveValue("1");
  await page.getByLabel("重复表头行数", { exact: true }).fill("0");
  await page.getByRole("button", { name: "设置重复表头", exact: true }).click();
  await expect(headers).toHaveCount(0);
  await page.evaluate(() => window.paginationHost.session!.undo());
  await expect(headers).toHaveCount(count);
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
});
test("table styles switch without stale paints, conditional options and original review preserve direct cell fills", async ({
  page,
}) => {
  const session = createDocumentSession(
      openWordDocument(
        importDocx(tableStylesFixture(), { documentId: "style-ui", unsupportedContent: "reject" }),
        { kind: "document" },
      ),
    ),
    root = session.getNode(
      session.snapshot().metadata.word!.stories.find((story) => story.kind === "main")!.rootNodeId,
    )!.parentNodeId!;
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "style-ui",
      baseRevision: 0,
      operations: [{ op: "setWordReview", nodeId: root, trackChanges: true, author: "甲" }],
    }),
  );
  await mount(page, session.serialize(), "edit");
  const frame = page.frameLocator("iframe"),
    cells = frame.locator("[data-office-editor] td");
  await expect(cells.first()).toHaveCSS("background-color", "rgb(170, 17, 17)");
  await cells.first().locator("p").click();
  await page.evaluate(() =>
    (window.paginationHost.view as import("@doffi/editor").DocumentTextEditor).execute({
      type: "word",
      action: "tableSettings",
    }),
  );
  await page.getByLabel("表格样式", { exact: true }).selectOption("Blue");
  await page.getByRole("button", { name: "应用表格样式", exact: true }).click();
  await expect(cells.first()).toHaveCSS("background-color", "rgb(17, 17, 170)");
  await expect(cells.nth(3)).toHaveCSS("background-color", "rgb(0, 255, 0)");
  await page.getByLabel("首行样式", { exact: true }).uncheck();
  await page.getByRole("button", { name: "应用表格样式选项", exact: true }).click();
  await expect(cells.first()).toHaveCSS("background-color", "rgb(0, 0, 204)");
  await page.evaluate(() =>
    (window.paginationHost.view as import("@doffi/editor").DocumentTextEditor).execute({
      type: "word",
      action: "review",
    }),
  );
  await page.getByLabel("修订显示", { exact: true }).selectOption("original");
  await expect(cells.first()).toHaveCSS("background-color", "rgb(170, 17, 17)");
  expect(
    await page.evaluate(
      () => window.paginationHost.session!.snapshot().metadata.word!.tables[0]!.lookMask,
    ),
  ).toBe(1536);
  await page.getByLabel("修订显示", { exact: true }).selectOption("all");
  await expect(cells.first()).toHaveCSS("background-color", "rgb(0, 0, 204)");
  await page.getByRole("button", { name: "拒绝全部显示的修订", exact: true }).click();
  await expect(cells.first()).toHaveCSS("background-color", "rgb(170, 17, 17)");
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
});
test("watermark controls repeat editable text behind every page, preserve source and remove with one undo", async ({
  page,
}) => {
  const session = createDocumentSession(
    openWordDocument(
      createWordDocument({
        documentId: "watermark-ui",
        pageSetup: {
          unit: "pt",
          width: 300,
          height: 240,
          margins: { top: 24, right: 24, bottom: 24, left: 24 },
        },
      }),
      { kind: "document" },
    ),
  );
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "watermark-ui",
      baseRevision: 0,
      operations: [
        {
          op: "replaceFragment",
          nodeId: "section-root-1",
          html:
            '<p style="line-height:16pt;margin:0">' +
            Array.from({ length: 40 }, (_, index) => "LINE-" + index).join("<br>") +
            "</p>",
        },
      ],
    }),
  );
  await mount(page, session.serialize(), "edit");
  const frame = page.frameLocator("iframe");
  await frame.locator("[data-office-editor] p").first().click();
  await page.evaluate(() =>
    (window.paginationHost.view as import("@doffi/editor").DocumentTextEditor).execute({
      type: "word",
      action: "watermark",
    }),
  );
  await page.getByLabel("水印文字", { exact: true }).fill("CONFIDENTIAL");
  await page.getByLabel("水印字号（磅）", { exact: true }).fill("24");
  await page.getByRole("button", { name: "应用文字水印", exact: true }).click();
  const pages = (await page.evaluate(() => window.paginationHost.layout())).pages.length,
    watermarks = frame.locator(".sdk-paper > [data-sdk-word-story-anchor]");
  await expect(watermarks).toHaveCount(pages);
  await expect(watermarks.first()).toContainText("CONFIDENTIAL");
  const positions = await frame.locator("body").evaluate((body) =>
    [...body.querySelectorAll(".sdk-paper > [data-sdk-word-story-anchor]")].map((node) => {
      const parent = node.parentElement!.getBoundingClientRect(),
        rect = node.getBoundingClientRect();
      return {
        x: (rect.left + rect.width / 2 - parent.left) * 0.75,
        y: (rect.top + rect.height / 2 - parent.top) * 0.75,
      };
    }),
  );
  for (const position of positions) {
    expect(position.x).toBeCloseTo(150, 0);
    expect(position.y).toBeCloseTo(120, 0);
  }
  expect(await page.evaluate(() => window.paginationHost.session!.snapshot().html)).not.toMatch(
    /sdk-word-story-anchor|sdk-word-drawing-text|position:absolute/,
  );
  await page.getByRole("button", { name: "移除水印", exact: true }).click();
  await expect(watermarks).toHaveCount(0);
  await page.evaluate(() => window.paginationHost.session!.undo());
  await expect(watermarks).toHaveCount(pages);
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
  const files = await page.evaluate(() =>
    [...window.paginationHost.session!.serialize()].map(
      ([path, bytes]) => [path, [...bytes]] as [string, number[]],
    ),
  );
  await page.evaluate((entries) => window.paginationHost.mount(entries, "preview"), files);
  expect(
    await page.evaluate(
      () =>
        (window.paginationHost.view as import("@doffi/editor").DocumentPreview)
          .printDocument()
          .querySelectorAll(".sdk-print-page [data-sdk-word-story-anchor]").length,
    ),
  ).toBe(pages);
});
test("fixed table row heights clip long source text, avoid phantom pages and edit height rules with shared undo", async ({
  page,
}) => {
  const session = createDocumentSession(
    openWordDocument(
      createWordDocument({
        documentId: "fixed-rows",
        pageSetup: {
          unit: "pt",
          width: 300,
          height: 240,
          margins: { top: 24, right: 24, bottom: 24, left: 24 },
        },
      }),
      { kind: "document" },
    ),
  );
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "fixed-rows",
      baseRevision: 0,
      operations: [
        {
          op: "insertWordTable",
          nodeId: "paragraph-1",
          position: "after",
          rows: 2,
          columns: 2,
          width: 240,
        },
      ],
    }),
  );
  const table = session.snapshot().metadata.word!.tables[0]!.nodeId,
    text = Array.from({ length: 40 }, (_, index) => "LONG-" + index).join("\n");
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "fixed-rows",
      baseRevision: 1,
      operations: [
        {
          op: "editWordTable",
          nodeId: table,
          change: {
            kind: "setCells",
            row: 0,
            column: 0,
            values: [
              [text, "A"],
              ["B", "C"],
            ],
          },
        },
        { op: "editWordTable", nodeId: table, change: { kind: "resize", rowHeights: [24, 24] } },
        { op: "formatWordTable", nodeId: table, rowHeightRule: "exact" },
      ],
    }),
  );
  await mount(page, session.serialize(), "edit");
  const frame = page.frameLocator("iframe");
  expect(
    await frame
      .locator("tr")
      .first()
      .evaluate((row) => row.getBoundingClientRect().height * 0.75),
  ).toBeCloseTo(24, 0);
  expect((await page.evaluate(() => window.paginationHost.layout())).pages).toHaveLength(1);
  expect(await frame.locator(".sdk-word-row-clip").count()).toBe(4);
  expect(
    await page.evaluate(
      () =>
        window.paginationHost.session!.getNode(
          window.paginationHost.session!.snapshot().metadata.word!.tables[0]!.nodeId,
        )!.text,
    ),
  ).toContain(text);
  await frame.locator("td").first().click();
  await page.evaluate(() =>
    (window.paginationHost.view as import("@doffi/editor").DocumentTextEditor).execute({
      type: "word",
      action: "tableSettings",
    }),
  );
  await page.getByLabel("行高规则", { exact: true }).selectOption("atLeast");
  await page.getByRole("button", { name: "设置行高", exact: true }).click();
  await expect(frame.locator(".sdk-word-row-clip")).toHaveCount(0);
  expect((await page.evaluate(() => window.paginationHost.layout())).pages.length).toBeGreaterThan(
    1,
  );
  await page.evaluate(() => window.paginationHost.session!.undo());
  await expect(frame.locator(".sdk-word-row-clip")).toHaveCount(4);
  expect((await page.evaluate(() => window.paginationHost.layout())).pages).toHaveLength(1);
  expect(await page.evaluate(() => window.paginationHost.session!.snapshot().html)).not.toMatch(
    /sdk-word-row-clip|overflow:hidden|flex-direction/,
  );
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
});
test("table border guides drag columns and one row through tracked atomic undo while Escape cancels the gesture", async ({
  page,
}) => {
  const session = createDocumentSession(
    openWordDocument(createWordDocument({ documentId: "table-drag" }), { kind: "document" }),
  );
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "table-drag",
      baseRevision: 0,
      operations: [
        {
          op: "insertWordTable",
          nodeId: "paragraph-1",
          position: "after",
          rows: 2,
          columns: 2,
          width: 240,
        },
        { op: "setWordReview", nodeId: "document-root", trackChanges: true, author: "甲" },
      ],
    }),
  );
  const tableId = session.snapshot().metadata.word!.tables[0]!.nodeId;
  await mount(page, session.serialize(), "edit");
  const frame = page.frameLocator("iframe"),
    cell = frame.locator("td").first();
  await cell.hover();
  const column = frame.getByRole("button", { name: "调整第 1 列宽度", exact: true }),
    handle = await column.boundingBox(),
    width = await frame.locator("table").evaluate((table) => table.getBoundingClientRect().width);
  await page.mouse.move(handle!.x + 4, handle!.y + 12);
  await page.mouse.down();
  await page.mouse.move(handle!.x + 44, handle!.y + 12, { steps: 4 });
  await page.mouse.up();
  const resized = await page.evaluate(
    (id) =>
      window.paginationHost
        .session!.snapshot()
        .metadata.word!.tables.find((table) => table.nodeId === id)!.grid,
    tableId,
  );
  expect(resized[0]).toBeCloseTo(120 + Math.round(((40 * 240) / width) * 4) / 4);
  expect(resized[0]! + resized[1]!).toBe(240);
  expect(
    await page.evaluate(() =>
      window.paginationHost
        .session!.snapshot()
        .metadata.word!.revisions.some((item) => item.property === "tableGrid"),
    ),
  ).toBe(true);
  await page.evaluate(() => window.paginationHost.session!.undo());
  expect(
    await page.evaluate(
      () => window.paginationHost.session!.snapshot().metadata.word!.tables[0]!.grid,
    ),
  ).toEqual([120, 120]);
  await cell.hover();
  const row = frame.getByRole("button", { name: "调整第 1 行高度", exact: true }),
    rowHandle = await row.boundingBox();
  await page.mouse.move(rowHandle!.x + 12, rowHandle!.y + 4);
  await page.mouse.down();
  await page.mouse.move(rowHandle!.x + 12, rowHandle!.y + 24, { steps: 3 });
  await page.mouse.up();
  const result = await page.evaluate(() => {
    const session = window.paginationHost.session!,
      snapshot = session.snapshot(),
      table = session.getNode(snapshot.metadata.word!.tables[0]!.nodeId)!,
      rows = table.childNodeIds.flatMap((id) => {
        const node = session.getNode(id)!;
        return node.tagName === "tr" ? [node] : node.childNodeIds.map((id) => session.getNode(id)!);
      });
    return {
      snapshot,
      rows: rows.map((row) => row.attributes.style),
      errors: window.paginationHost.errors,
    };
  });
  expect(
    result.snapshot.metadata.word!.revisions.filter((item) => item.property === "row"),
  ).toHaveLength(1);
  expect(result.rows[0]).toContain("height:");
  expect(result.rows[1]).toBeUndefined();
  expect(result.snapshot.html).not.toMatch(/sdk-table-resize|position:fixed/);
  expect(result.errors).toEqual([]);
  await page.evaluate(() => window.paginationHost.session!.undo());
  await cell.hover();
  const cancelHandle = await frame
    .getByRole("button", { name: "调整第 1 列宽度", exact: true })
    .boundingBox();
  await page.mouse.move(cancelHandle!.x + 4, cancelHandle!.y + 12);
  await page.mouse.down();
  await page.mouse.move(cancelHandle!.x + 24, cancelHandle!.y + 12);
  await page.keyboard.press("Escape");
  await page.mouse.up();
  expect(
    await page.evaluate(
      () => window.paginationHost.session!.snapshot().metadata.word!.tables[0]!.grid,
    ),
  ).toEqual([120, 120]);
  expect(
    await page.evaluate(() => window.paginationHost.session!.snapshot().metadata.word!.revisions),
  ).toEqual([]);
});
test("unequal text columns reflow to individual widths, keep exact second-column typing and expose reversible page controls", async ({
  page,
}) => {
  const text = "abcdefghij ".repeat(160),
    session = createDocumentSession(
      openWordDocument(
        createWordDocument({
          documentId: "unequal-columns",
          pageSetup: {
            unit: "pt",
            width: 300,
            height: 240,
            margins: { top: 24, right: 24, bottom: 24, left: 24 },
          },
        }),
        { kind: "document" },
      ),
    );
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "unequal-columns",
      baseRevision: 0,
      operations: [
        {
          op: "replaceFragment",
          nodeId: "section-root-1",
          html:
            '<p style="font-family:&quot;Liberation Mono&quot;;font-size:10pt;line-height:12pt;margin:0;background-color:#eeeeee;border:.5pt solid black">' +
            text +
            "</p>",
        },
        {
          op: "setWordSection",
          nodeId: "section-root-1",
          changes: {
            columns: { count: 2, gap: 24, widths: [72, 156], spaces: [24, 0], separator: true },
          },
        },
      ],
    }),
  );
  await mount(page, session.serialize(), "edit");
  const frame = page.frameLocator("iframe"),
    paragraph = frame.locator("[data-office-editor] p").first(),
    lines = await paragraph.evaluate((element) => {
      const walker = document.createTreeWalker(element, 4),
        base = element.closest(".sdk-page-flow")!.getBoundingClientRect(),
        lines: { x: number; y: number; width: number }[] = [];
      for (let text = walker.nextNode(); text; text = walker.nextNode()) {
        const range = document.createRange();
        range.selectNodeContents(text);
        for (const rect of range.getClientRects())
          if (rect.width > 0.1 && rect.height > 0.1)
            lines.push({ x: rect.left - base.left, y: rect.top - base.top, width: rect.width });
      }
      return lines;
    });
  expect(
    Math.max(...lines.filter((line) => line.x < 100 && line.y < 250).map((line) => line.width)),
  ).toBeLessThan(98);
  expect(
    Math.max(...lines.filter((line) => line.x > 120 && line.y < 250).map((line) => line.width)),
  ).toBeGreaterThan(150);
  await paragraph.evaluate((element) => {
    const walker = document.createTreeWalker(element, 4),
      base = element.closest(".sdk-page-flow")!.getBoundingClientRect();
    for (let text = walker.nextNode(); text; text = walker.nextNode()) {
      const range = document.createRange();
      range.selectNodeContents(text);
      const rect = range.getBoundingClientRect();
      if (rect.left - base.left > 120 && rect.top - base.top < 250) {
        range.setStart(text, 0);
        range.collapse(true);
        const selection = document.getSelection()!;
        selection.removeAllRanges();
        selection.addRange(range);
        (element.closest("[data-office-editor]") as HTMLElement).focus();
        return;
      }
    }
    throw new Error("Missing wide-column text.");
  });
  await page.keyboard.insertText("EDIT");
  const current = await page.evaluate(
    () =>
      window.paginationHost.session!.getNode(
        window.paginationHost.session!.snapshot().metadata.sections![0]!.rootNodeId,
      )!.text,
  );
  expect(current.replace("EDIT", "")).toBe(text);
  expect(await page.evaluate(() => window.paginationHost.session!.snapshot().html)).not.toMatch(
    /sdk-column-exclusion|shape-outside|position:relative/,
  );
  await page.evaluate(() =>
    (window.paginationHost.view as import("@doffi/editor").DocumentTextEditor).execute({
      type: "word",
      action: "pageSettings",
    }),
  );
  await page.getByLabel("等宽分栏", { exact: true }).uncheck();
  await page.getByLabel("各栏宽度（磅，以逗号分隔）", { exact: true }).fill("100,128");
  await page.getByLabel("各栏后间距（磅，以逗号分隔）", { exact: true }).fill("24,0");
  await page.getByRole("button", { name: "应用分栏", exact: true }).click();
  expect(
    await page.evaluate(
      () => window.paginationHost.session!.snapshot().metadata.sections![0]!.columns!.widths,
    ),
  ).toEqual([100, 128]);
  await page.evaluate(() => window.paginationHost.session!.undo());
  expect(
    await page.evaluate(
      () => window.paginationHost.session!.snapshot().metadata.sections![0]!.columns!.widths,
    ),
  ).toEqual([72, 156]);
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
});
test("original review projects earliest run and section formatting while new typing keeps current source formatting and history", async ({
  page,
}) => {
  const session = createDocumentSession(
    openWordDocument(createWordDocument({ documentId: "original-format" }), { kind: "document" }),
  );
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "original-format",
      baseRevision: 0,
      operations: [
        {
          op: "replaceText",
          nodeId: "paragraph-1",
          from: 0,
          to: 0,
          expectedText: "",
          text: "ABCDE",
        },
        { op: "setWordReview", nodeId: "document-root", trackChanges: true, author: "甲" },
        {
          op: "formatWordText",
          nodeId: "paragraph-1",
          from: 0,
          to: 5,
          expectedText: "ABCDE",
          properties: { color: "#cc0000", "font-size": "20pt" },
        },
        { op: "setWordSection", nodeId: "section-root-1", changes: { gutter: 18 } },
      ],
    }),
  );
  await mount(page, session.serialize(), "edit");
  const frame = page.frameLocator("iframe"),
    paragraph = frame.locator("[data-node-id=paragraph-1]");
  await paragraph.click();
  await page.evaluate(() =>
    (window.paginationHost.view as import("@doffi/editor").DocumentTextEditor).execute({
      type: "word",
      action: "review",
    }),
  );
  await page.getByLabel("修订显示", { exact: true }).selectOption("original");
  expect(
    await paragraph
      .locator("[data-node-id]")
      .first()
      .evaluate((node) => getComputedStyle(node).color),
  ).toBe("rgb(0, 0, 0)");
  expect(
    (await page.evaluate(() => window.paginationHost.layout())).pages[0]!.setup.margins.left,
  ).toBe(72);
  await page.getByRole("button", { name: "关闭修订审阅", exact: true }).click();
  await paragraph.click();
  await page.keyboard.press("End");
  await page.keyboard.insertText("X");
  const snapshot = await page.evaluate(() => window.paginationHost.session!.snapshot());
  expect(snapshot.html).toContain("color:#cc0000");
  expect(snapshot.metadata.sections![0]!.gutter).toBe(18);
  expect(snapshot.metadata.word!.revisions.some((item) => item.property === "section")).toBe(true);
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
  await page.evaluate(() =>
    (window.paginationHost.view as import("@doffi/editor").DocumentTextEditor).execute({
      type: "word",
      action: "review",
    }),
  );
  await page.getByLabel("修订显示", { exact: true }).selectOption("final");
  expect(
    (await page.evaluate(() => window.paginationHost.layout())).pages[0]!.setup.margins.left,
  ).toBe(90);
  await page.getByRole("button", { name: "拒绝全部显示的修订", exact: true }).click();
  expect(
    await page.evaluate(() => window.paginationHost.session!.getNode("paragraph-1")!.text),
  ).toBe("ABCDE");
  expect(await page.evaluate(() => window.paginationHost.session!.snapshot().html)).not.toContain(
    "color:#cc0000",
  );
});
test("page borders respect section-first pages, header inclusion and front print layers while page controls share undo", async ({
  page,
}) => {
  const session = createDocumentSession(
      openWordDocument(
        createWordDocument({
          documentId: "page-borders",
          pageSetup: {
            unit: "pt",
            width: 300,
            height: 240,
            margins: { top: 36, right: 24, bottom: 36, left: 24 },
          },
        }),
        { kind: "document" },
      ),
    ),
    border = { style: "double" as const, width: 1, space: 12, color: "224466" };
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "page-borders",
      baseRevision: 0,
      operations: [
        {
          op: "replaceFragment",
          nodeId: "section-root-1",
          html:
            '<p style="line-height:16pt;margin:0">' +
            Array.from({ length: 32 }, (_, index) => "LINE-" + index).join("<br>") +
            "</p>",
        },
        {
          op: "insertWordStory",
          nodeId: "section-root-1",
          kind: "header",
          variant: "default",
          html: "<p>页眉</p>",
        },
        {
          op: "setWordSection",
          nodeId: "section-root-1",
          changes: {
            headerDistance: 18,
            pageBorders: {
              offsetFrom: "text",
              display: "firstPage",
              zOrder: "front",
              top: border,
              left: border,
              bottom: border,
              right: border,
            },
          },
        },
      ],
    }),
  );
  await mount(page, session.serialize(), "edit");
  const frame = page.frameLocator("iframe"),
    front = frame.locator(".sdk-page-border-layer");
  await expect(front.locator("[data-sdk-page-border]")).toHaveCount(4);
  await expect(front.locator("[data-sdk-page-border=top]")).toHaveAttribute(
    "data-page-id",
    "page-1",
  );
  expect(
    await front
      .locator("[data-sdk-page-border=top]")
      .evaluate((node) => parseFloat((node as HTMLElement).style.top)),
  ).toBe(5);
  await frame.locator("[data-office-editor] p").first().click();
  await page.evaluate(() =>
    (window.paginationHost.view as import("@doffi/editor").DocumentTextEditor).execute({
      type: "word",
      action: "pageSettings",
    }),
  );
  await page.getByLabel("页面边框页面", { exact: true }).selectOption("notFirstPage");
  await page.getByLabel("页面边框相对于", { exact: true }).selectOption("page");
  await page.getByLabel("页面边框距离（磅）", { exact: true }).fill("18");
  await page.getByRole("button", { name: "应用页面边框", exact: true }).click();
  const pages = await page.evaluate(() => window.paginationHost.layout().pages.length);
  await expect(front.locator("[data-sdk-page-border]")).toHaveCount((pages - 1) * 4);
  await expect(front.locator("[data-page-id=page-1]")).toHaveCount(0);
  await page.evaluate(() => window.paginationHost.session!.undo());
  await expect(front.locator("[data-sdk-page-border]")).toHaveCount(4);
  expect(await page.evaluate(() => window.paginationHost.session!.snapshot().html)).not.toContain(
    "sdk-page-border",
  );
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
  const files = await page.evaluate(() =>
    [...window.paginationHost.session!.serialize()].map(
      ([path, bytes]) => [path, [...bytes]] as [string, number[]],
    ),
  );
  await page.evaluate((entries) => window.paginationHost.mount(entries, "preview"), files);
  const printed = await page.evaluate(() => {
    const output = (
      window.paginationHost.view as import("@doffi/editor").DocumentPreview
    ).printDocument();
    return {
      borders: output.querySelectorAll("[data-sdk-page-border]").length,
      first: output
        .querySelector(".sdk-print-page [data-sdk-page-border=top]")
        ?.getAttribute("style"),
    };
  });
  expect(printed.borders).toBe(4);
  expect(printed.first).toContain("top: 5pt");
});
test.beforeAll(async () => {
  const result = await build({
    configFile: false,
    logLevel: "silent",
    build: {
      write: false,
      target: "es2022",
      lib: {
        entry: fileURLToPath(new URL("./fixtures/pagination-host.ts", import.meta.url)),
        formats: ["es"],
      },
    },
  });
  const output = Array.isArray(result)
    ? result[0]!.output
    : "output" in result
      ? result.output
      : [];
  code = output.flatMap((item) => (item.type === "chunk" ? [item.code] : [])).join("\n");
});
async function mount(
  page: Page,
  files: DocumentPackage,
  mode: "preview" | "edit" | "interaction" = "preview",
  options: DocumentPreviewOptions = {},
): Promise<void> {
  await openSdkHost(page);
  await page.addScriptTag({ type: "module", content: code });
  await page.waitForFunction(() => !!window.paginationHost);
  await page.evaluate(
    ({ entries, mode, options }) => window.paginationHost.mount(entries, mode, options),
    {
      entries: [...files].map(([path, bytes]) => [path, [...bytes]] as [string, number[]]),
      mode,
      options,
    },
  );
}
function columnFiles(table = false): DocumentPackage {
  const session = createDocumentSession(
    openWordDocument(
      createWordDocument({
        documentId: "columns",
        pageSetup: {
          unit: "pt",
          width: 300,
          height: 240,
          margins: { top: 24, right: 24, bottom: 24, left: 24 },
        },
      }),
      { kind: "document" },
    ),
  );
  const content = table
    ? '<table style="width:100pt;border-collapse:collapse"><tr data-word-repeat-header="true"><th style="border:1pt solid black"><p>HEADER</p></th></tr>' +
      Array.from(
        { length: 32 },
        (_, n) => '<tr><td style="border:1pt solid black;padding:0"><p>ROW-' + n + "</p></td></tr>",
      ).join("") +
      "</table>"
    : '<p style="line-height:16pt;margin:0">' +
      Array.from({ length: 40 }, (_, n) => "LINE-" + n + " 😀").join("<br>") +
      "</p>";
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "columns",
      baseRevision: 0,
      operations: [
        { op: "replaceFragment", nodeId: "section-root-1", html: content },
        {
          op: "setWordSection",
          nodeId: "section-root-1",
          changes: { columns: { count: 2, gap: 24, separator: true } },
        },
      ],
    }),
  );
  return session.serialize();
}
test("equation input and component editing use one source atom and reversible MathML views", async ({
  page,
}) => {
  const session = createDocumentSession(
    openWordDocument(createWordDocument({ documentId: "equation-input" }), { kind: "document" }),
  );
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "equation-input",
      baseRevision: 0,
      operations: [
        {
          op: "replaceText",
          nodeId: "paragraph-1",
          from: 0,
          to: 0,
          expectedText: "",
          text: "前后",
        },
      ],
    }),
  );
  await mount(page, session.serialize(), "edit");
  const frame = page.frameLocator("iframe"),
    paragraph = frame.locator('[data-node-id="paragraph-1"]');
  await paragraph.click();
  await page.keyboard.press("End");
  await page.evaluate(() =>
    (window.paginationHost.view as import("@doffi/editor").DocumentTextEditor).execute({
      type: "word",
      action: "insertEquation",
    }),
  );
  await page.getByLabel("线性公式", { exact: true }).fill("\\frac{a_1^2}{\\sqrt[3]{b}}");
  await page.getByRole("button", { name: "插入公式", exact: true }).click();
  await expect(frame.locator("math")).toHaveCount(1);
  await expect(frame.locator("math mfrac")).toHaveCount(1);
  await expect(frame.locator("math mroot")).toHaveCount(1);
  await page.getByLabel("公式组件文字", { exact: true }).fill("z");
  await page.getByRole("button", { name: "更新公式文字组件", exact: true }).click();
  await expect(frame.locator("math")).toContainText("z");
  await page.getByRole("button", { name: "关闭公式编辑", exact: true }).click();
  await paragraph.click({ position: { x: 5, y: 5 } });
  await page.keyboard.press("Home");
  await page.keyboard.insertText("EDIT");
  const result = await page.evaluate(() => ({
    snapshot: window.paginationHost.session!.snapshot(),
    text: window.paginationHost.session!.getNode("paragraph-1")!.text,
    errors: window.paginationHost.errors,
  }));
  expect(result.text).toBe("EDIT前后\ufffc");
  expect(result.snapshot.metadata.word!.equations).toHaveLength(1);
  expect(result.snapshot.html).not.toMatch(/<math|mfrac|sdk-word-equation|contenteditable/);
  expect(result.errors).toEqual([]);
  await page.evaluate(() => window.paginationHost.session!.undo());
  expect(
    await page.evaluate(() => window.paginationHost.session!.getNode("paragraph-1")!.text),
  ).toBe("前后\ufffc");
});
test("bookmark and hyperlink panels apply typed source changes and retain their text selection", async ({
  page,
}) => {
  const session = createDocumentSession(
    openWordDocument(createWordDocument({ documentId: "link-controls" }), { kind: "document" }),
  );
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "link-controls",
      baseRevision: 0,
      operations: [
        {
          op: "replaceText",
          nodeId: "paragraph-1",
          from: 0,
          to: 0,
          expectedText: "",
          text: "链接文字",
        },
      ],
    }),
  );
  await mount(page, session.serialize(), "edit");
  const frame = page.frameLocator("iframe"),
    paragraph = frame.locator('[data-node-id="paragraph-1"]');
  await paragraph.evaluate((node) => {
    (node.closest("[data-office-editor]") as HTMLElement).focus();
    const range = document.createRange();
    range.selectNodeContents(node);
    const selection = document.getSelection()!;
    selection.removeAllRanges();
    selection.addRange(range);
  });
  await page.evaluate(() =>
    (window.paginationHost.view as import("@doffi/editor").DocumentTextEditor).execute({
      type: "word",
      action: "bookmarks",
    }),
  );
  await page.getByLabel("书签名称", { exact: true }).fill("LinkedText");
  await page.getByRole("button", { name: "添加书签", exact: true }).click();
  await page.evaluate(() =>
    (window.paginationHost.view as import("@doffi/editor").DocumentTextEditor).execute({
      type: "word",
      action: "insertLink",
    }),
  );
  await page.getByLabel("链接地址", { exact: true }).fill("https://example.com");
  await page.getByRole("button", { name: "插入链接", exact: true }).click();
  const result = await page.evaluate(() => ({
    snapshot: window.paginationHost.session!.snapshot(),
    text: window.paginationHost.session!.getNode("paragraph-1")!.text,
    errors: window.paginationHost.errors,
  }));
  expect(result.snapshot.metadata.word!.bookmarks[0]!.name).toBe("LinkedText");
  expect(result.snapshot.metadata.word!.hyperlinks[0]!.target).toBe("https://example.com");
  expect(result.text).toBe("链接文字");
  expect(result.snapshot.html).not.toMatch(/contenteditable|sdk-page/);
  expect(result.errors).toEqual([]);
});
test("bookmark references and caption fields derive current values and explicit cache updates share undo history", async ({
  page,
}) => {
  const session = createDocumentSession(
    openWordDocument(createWordDocument({ documentId: "reference-fields" }), { kind: "document" }),
  );
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "reference-fields",
      baseRevision: 0,
      operations: [
        {
          op: "replaceText",
          nodeId: "paragraph-1",
          from: 0,
          to: 0,
          expectedText: "",
          text: "引用文字",
        },
        {
          op: "insertWordBookmark",
          nodeId: "paragraph-1",
          from: 0,
          to: 4,
          expectedText: "引用文字",
          name: "Target",
        },
        { op: "insertFragment", nodeId: "paragraph-1", position: "after", html: "<p>引用：</p>" },
      ],
    }),
  );
  const paragraph = session.getNode("section-root-1")!.childNodeIds[1]!;
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "reference-fields",
      baseRevision: 1,
      operations: [
        {
          op: "insertWordField",
          nodeId: paragraph,
          instruction: "REF Target \\h",
          cachedText: "cache",
        },
        {
          op: "insertWordField",
          nodeId: paragraph,
          instruction: "PAGEREF Target \\h",
          cachedText: "cache",
        },
        {
          op: "insertWordField",
          nodeId: paragraph,
          instruction: "SEQ Figure",
          cachedText: "cache",
        },
      ],
    }),
  );
  await mount(page, session.serialize(), "edit");
  const frame = page.frameLocator("iframe");
  await expect(frame.locator(".sdk-field-display").nth(0)).toHaveText("引用文字");
  await expect(frame.locator(".sdk-field-display").nth(1)).toHaveText("1");
  await expect(frame.locator(".sdk-field-display").nth(2)).toHaveText("1");
  const before = await page.evaluate(() => window.paginationHost.session!.snapshot());
  expect(before.html).toContain("cache");
  await page.evaluate(() => window.paginationHost.updateFields());
  const after = await page.evaluate(() => window.paginationHost.session!.snapshot());
  expect(after.html).not.toContain("cache");
  expect(after.html).not.toContain("sdk-field-display");
  await page.evaluate(() => window.paginationHost.session!.undo());
  expect((await page.evaluate(() => window.paginationHost.session!.snapshot())).html).toBe(
    before.html,
  );
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
});
function notesFiles(long = false): DocumentPackage {
  const session = createDocumentSession(
    openWordDocument(
      createWordDocument({
        documentId: "word-notes",
        pageSetup: {
          unit: "pt",
          width: 300,
          height: 240,
          margins: { top: 24, right: 24, bottom: 24, left: 24 },
        },
      }),
      { kind: "document" },
    ),
  );
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "word-notes",
      baseRevision: 0,
      operations: [
        {
          op: "replaceText",
          nodeId: "paragraph-1",
          from: 0,
          to: 0,
          expectedText: "",
          text: "正文前后",
        },
        {
          op: "insertWordNote",
          nodeId: "paragraph-1",
          offset: 2,
          kind: "footnote",
          text: long
            ? Array.from({ length: 60 }, (_, index) => "长脚注 " + index).join("\n")
            : "脚注文字",
        },
        {
          op: "insertWordNote",
          nodeId: "paragraph-1",
          offset: 0,
          kind: "endnote",
          text: "尾注文字",
          customMark: "*",
        },
      ],
    }),
  );
  return session.serialize();
}
test("footnotes move beneath text through section controls and double-click opens their shared editable story", async ({
  page,
}) => {
  await mount(page, notesFiles(), "edit");
  const frame = page.frameLocator("iframe"),
    notes = frame.locator(".sdk-page-notes");
  const noteTop = () =>
    frame.locator("body").evaluate((root) => {
      const note = root.querySelector(".sdk-page-notes"),
        paper = root.querySelector(".sdk-paper");
      return note && paper
        ? note.getBoundingClientRect().top - paper.getBoundingClientRect().top
        : null;
    });
  await expect.poll(noteTop).not.toBeNull();
  const bottom = (await noteTop())!;
  await frame.locator('[data-node-id="paragraph-1"]').click();
  await page.evaluate(() =>
    (window.paginationHost.view as import("@doffi/editor").DocumentTextEditor).execute({
      type: "word",
      action: "pageSettings",
    }),
  );
  await page.getByLabel("本节脚注位置", { exact: true }).selectOption("beneathText");
  await page.getByRole("button", { name: "应用本节脚注设置", exact: true }).click();
  await expect
    .poll(async () => {
      const top = await noteTop();
      return top === null ? Infinity : top;
    })
    .toBeLessThan(bottom - 50);
  await page.evaluate(() => window.paginationHost.session!.undo());
  await expect.poll(async () => Math.abs(((await noteTop()) ?? Infinity) - bottom)).toBeLessThan(2);
  await page.evaluate(() => window.paginationHost.session!.redo());
  await notes.locator("p").dblclick();
  await expect(frame.locator("[data-office-editor]")).toContainText("脚注文字");
  await frame.locator("[data-office-editor] p").first().click();
  await page.keyboard.press("End");
  await page.keyboard.insertText("编辑");
  await page.evaluate(() =>
    (window.paginationHost.view as import("@doffi/editor").DocumentTextEditor).execute({
      type: "word",
      action: "mainStory",
    }),
  );
  await expect(notes).toContainText("脚注文字编辑");
  expect(await page.evaluate(() => window.paginationHost.session!.snapshot().html)).not.toMatch(
    /sdk-word-endnotes|sdk-word-note-story|sdk-page-notes/,
  );
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
});
function sectionNotesFiles(suppressed: boolean): DocumentPackage {
  const session = createDocumentSession(
      openWordDocument(
        createWordDocument({
          documentId: "section-notes-ui",
          pageSetup: {
            unit: "pt",
            width: 300,
            height: 300,
            margins: { top: 24, right: 24, bottom: 24, left: 24 },
          },
        }),
        { kind: "document" },
      ),
    ),
    apply = (operations: import("@doffi/core").DocumentOperation[]) =>
      session.commit(
        session.prepare({
          protocolVersion: 1,
          documentId: "section-notes-ui",
          baseRevision: session.snapshot().manifest.revision,
          operations,
        }),
      );
  apply([
    {
      op: "replaceFragment",
      nodeId: "section-root-1",
      html: '<p class="note-a">BODY-A</p><p class="note-b">BODY-B</p><p class="note-c">BODY-C</p>',
    },
  ]);
  const [a, b, c] = session.getNode("section-root-1")!.childNodeIds;
  apply([{ op: "splitWordSection", nodeId: a!, breakType: "continuous" }]);
  apply([{ op: "splitWordSection", nodeId: b!, breakType: "nextPage" }]);
  const sections = session.snapshot().metadata.sections!;
  apply([
    {
      op: "setWordNoteOptions",
      nodeId: "document-root",
      kind: "endnote",
      options: { position: "sectEnd" },
    },
    {
      op: "setWordSection",
      nodeId: sections[0]!.rootNodeId,
      changes: {
        footnoteOptions: {
          format: "lowerRoman",
          start: 3,
          restart: "eachSect",
          position: "beneathText",
        },
        noEndnote: suppressed,
      },
    },
    {
      op: "setWordSection",
      nodeId: sections[1]!.rootNodeId,
      changes: {
        footnoteOptions: {
          format: "decimal",
          start: 7,
          restart: "eachSect",
          position: "pageBottom",
        },
        endnoteOptions: { format: "upperLetter", start: 2, restart: "eachSect" },
      },
    },
    {
      op: "setWordSection",
      nodeId: sections[2]!.rootNodeId,
      changes: { footnoteOptions: { format: "decimal", restart: "continuous" } },
    },
    ...[a!, b!, c!].flatMap((nodeId, index): import("@doffi/core").DocumentOperation[] => [
      { op: "insertWordNote", nodeId, offset: 0, kind: "footnote", text: "FOOT-" + index },
      {
        op: "insertWordNote",
        nodeId,
        offset: 1,
        kind: "endnote",
        text:
          index === 2
            ? Array.from({ length: 40 }, (_, line) => "END-C-" + line).join("\n")
            : index === 0
              ? "END-A"
              : "END-B",
      },
    ]),
  ]);
  return session.serialize();
}
for (const suppressed of [false, true])
  test(
    "continuous sections keep local note numbering, section endnotes and deferred endnotes in reading order " +
      suppressed,
    async ({ page }) => {
      await mount(page, sectionNotesFiles(suppressed));
      const frame = page.frameLocator("iframe"),
        a = await bandTextBounds(page, "END-A"),
        b = await bandTextBounds(page, "BODY-B"),
        endB = await bandTextBounds(page, "END-B");
      expect(suppressed ? a.y > b.y : a.y < b.y).toBe(true);
      expect(a.y).toBeLessThan(endB.y);
      expect(
        await frame
          .locator('[data-word-note-reference="footnote"]')
          .evaluateAll((nodes) => nodes.map((node) => node.lastElementChild!.textContent)),
      ).toEqual(["iii", "7", "8"]);
      expect(
        (await page.evaluate(() => window.paginationHost.layout())).pages.length,
      ).toBeGreaterThan(2);
      await expect(frame.locator(".sdk-word-endnotes")).toHaveCount(suppressed ? 2 : 3);
      const after = await frame
          .locator('.sdk-page-notes[data-note-placement="after"]')
          .first()
          .boundingBox(),
        bottom = await frame
          .locator('.sdk-page-notes[data-note-placement="bottom"]')
          .first()
          .boundingBox();
      expect(after!.y).toBeLessThan(bottom!.y);
      expect(await page.evaluate(() => window.paginationHost.model!.snapshot().html)).not.toMatch(
        /sdk-word-endnotes|sdk-note/,
      );
      expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
    },
  );
test("a text annotation in a continued footnote maps only to its visible clipped copy", async ({
  page,
}) => {
  const session = createDocumentSession(openWordDocument(notesFiles(true), { kind: "document" })),
    word = session.snapshot().metadata.word!,
    story = word.stories.find((story) => story.kind === "footnote")!,
    paragraph = session.getNode(story.rootNodeId)!.childNodeIds[0]!,
    quote = "长脚注 59",
    from = session.getNode(paragraph)!.text.indexOf(quote);
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "word-notes",
      baseRevision: session.snapshot().manifest.revision,
      operations: [
        {
          op: "addAnnotation",
          annotation: {
            annotationId: "continued-note",
            documentId: "word-notes",
            baseRevision: session.snapshot().manifest.revision,
            status: "open",
            anchor: { kind: "text", nodeId: paragraph, from, to: from + quote.length, quote },
            instruction: "检查脚注最后一行",
          },
        },
      ],
    }),
  );
  await mount(page, session.serialize(), "interaction");
  await page.evaluate(() => window.paginationHost.interaction!.focusAnnotation("continued-note"));
  const highlights = page.frameLocator("iframe").locator("[data-sdk-highlight]");
  expect(await highlights.count()).toBeGreaterThan(0);
  const visible = await page.frameLocator("iframe").locator(".sdk-page-notes").last().boundingBox(),
    boxes = await highlights.evaluateAll((nodes) =>
      nodes.map((node) => {
        const box = node.getBoundingClientRect();
        return { top: box.top, bottom: box.bottom };
      }),
    ),
    frame = await page.locator("iframe").boundingBox();
  for (const box of boxes) {
    expect(box.top + frame!.y).toBeGreaterThanOrEqual(visible!.y - 1);
    expect(box.bottom + frame!.y).toBeLessThanOrEqual(visible!.y + visible!.height + 1);
  }
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
});
test("footnotes reserve page space, display editable auxiliary stories and retain native reference atoms", async ({
  page,
}) => {
  await mount(page, notesFiles(), "edit");
  await page.evaluate(() => window.paginationHost.view!.settled());
  const frame = page.frameLocator("iframe");
  await expect(frame.locator("[data-word-note-reference]")).toHaveCount(2);
  await expect(frame.locator(".sdk-page-notes")).toHaveCount(1);
  await expect(frame.locator(".sdk-page-notes")).toContainText("脚注文字");
  const before = await page.evaluate(() => window.paginationHost.session!.snapshot());
  await frame.locator('[data-node-id="paragraph-1"]').click();
  await page.evaluate(() =>
    (window.paginationHost.view as import("@doffi/editor").DocumentTextEditor).execute({
      type: "word",
      action: "referenceSettings",
    }),
  );
  await page.getByRole("button", { name: "编辑 ", exact: false }).first().click();
  const input = frame.locator("[data-office-editor] p").first();
  await input.click();
  await page.keyboard.press("End");
  await page.keyboard.insertText("补充");
  await page.evaluate(() =>
    (window.paginationHost.view as import("@doffi/editor").DocumentTextEditor).execute({
      type: "word",
      action: "mainStory",
    }),
  );
  await expect(frame.locator(".sdk-page-notes")).toContainText("补充");
  const result = await page.evaluate(() => ({
    snapshot: window.paginationHost.session!.snapshot(),
    text: window.paginationHost.session!.getNode("paragraph-1")!.text,
    errors: window.paginationHost.errors,
  }));
  expect(result.snapshot.metadata.word!.noteReferences).toEqual(
    before.metadata.word!.noteReferences,
  );
  expect(result.snapshot.html).not.toMatch(/sdk-page-notes|sdk-word-notes|contenteditable/);
  expect(result.errors).toEqual([]);
});
test("footnote-only continuation pages share native page labels, header fields and physical paper surfaces", async ({
  page,
}) => {
  const session = createDocumentSession(openWordDocument(notesFiles(true), { kind: "document" })),
    apply = (operations: import("@doffi/core").DocumentOperation[]) =>
      session.commit(
        session.prepare({
          protocolVersion: 1,
          documentId: "word-notes",
          baseRevision: session.snapshot().manifest.revision,
          operations,
        }),
      );
  apply([
    {
      op: "setWordSection",
      nodeId: "section-root-1",
      changes: { headerDistance: 12, pageNumberStart: 4, pageNumberFormat: "lowerRoman" },
    },
    {
      op: "insertWordStory",
      nodeId: "section-root-1",
      kind: "header",
      variant: "default",
      html: '<p style="font-size:8pt;line-height:10pt;margin:0">HEADER </p>',
    },
  ]);
  const header = session
      .snapshot()
      .metadata.word!.stories.find((story) => story.kind === "header")!,
    paragraph = session.getNode(header.rootNodeId)!.childNodeIds[0]!;
  apply([
    { op: "insertWordField", nodeId: paragraph, instruction: "PAGE", cachedText: "0" },
    { op: "insertWordField", nodeId: paragraph, instruction: "NUMPAGES", cachedText: "0" },
  ]);
  await mount(page, session.serialize());
  const report = await page.evaluate(() => window.paginationHost.layout()),
    frame = page.frameLocator("iframe");
  expect(report.pages.length).toBeGreaterThan(2);
  await expect(frame.locator(".sdk-word-notes .sdk-paper")).toHaveCount(0);
  await expect(frame.locator(".sdk-page-story")).toHaveCount(report.pages.length);
  for (const [index, definition] of report.pages.entries()) {
    const paper = frame.locator('.sdk-paper[data-page-id="' + definition.pageId + '"]');
    await expect(paper.locator(".sdk-field-display").nth(0)).toHaveText(definition.label);
    await expect(paper.locator(".sdk-field-display").nth(1)).toHaveText(
      String(report.pages.length),
    );
    expect(definition.number).toBe(index + 4);
  }
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
});
test("long footnotes continue across measured pages without truncating the saved source", async ({
  page,
}) => {
  await mount(page, notesFiles(true));
  const frame = page.frameLocator("iframe");
  expect(await frame.locator(".sdk-page-notes").count()).toBeGreaterThan(2);
  const report = await page.evaluate(() => window.paginationHost.layout());
  expect(report.pages.length).toBeGreaterThan(2);
  await expect(frame.locator(".sdk-word-notes")).toContainText("长脚注 59");
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
});
test("custom footnote separators and notices reserve their own continuation space and open a shared formatting story", async ({
  page,
}) => {
  await mount(page, notesFiles(true), "edit");
  await page.evaluate(() =>
    (window.paginationHost.view as import("@doffi/editor").DocumentTextEditor).execute({
      type: "word",
      action: "referenceSettings",
    }),
  );
  await page.getByLabel("注释分隔内容类型", { exact: true }).selectOption("continuationNotice");
  await page.getByLabel("注释分隔内容", { exact: true }).click();
  await page.evaluate(() => window.paginationHost.view!.settled());
  await page.getByLabel("注释分隔内容", { exact: true }).fill("下页接续");
  await page.getByRole("button", { name: "应用注释分隔内容", exact: true }).click();
  const frame = page.frameLocator("iframe"),
    notes = frame.locator(".sdk-page-notes");
  expect(
    await page.evaluate(() => ({
      errors: window.paginationHost.errors,
      note: window.paginationHost
        .session!.snapshot()
        .metadata.word!.notes!.find((note) => note.type === "continuationNotice"),
    })),
  ).toMatchObject({ errors: [], note: { kind: "footnote" } });
  expect(
    await page.evaluate(() => {
      const session = window.paginationHost.session!,
        word = session.snapshot().metadata.word!,
        note = word.notes!.find((note) => note.type === "continuationNotice")!,
        story = word.stories.find((story) => story.storyId === note.storyId)!;
      return {
        text: session.getNode(story.rootNodeId)!.text,
        main: session.getNode("paragraph-1")!.text,
      };
    }),
  ).toMatchObject({ text: "下页接续", main: "\ufffc正文\ufffc前后" });
  await expect(notes.first()).toContainText("下页接续");
  await expect(notes.last()).not.toContainText("下页接续");
  await page.getByLabel("注释分隔内容类型", { exact: true }).selectOption("continuationSeparator");
  await page.getByLabel("注释分隔内容", { exact: true }).click();
  await page.getByLabel("注释分隔内容", { exact: true }).fill("续页注释");
  await page.getByRole("button", { name: "应用注释分隔内容", exact: true }).click();
  await expect(notes.nth(1)).toContainText("续页注释");
  await expect(notes.first()).not.toContainText("续页注释");
  const boxes = await notes.evaluateAll((nodes) =>
    nodes.map((node) => {
      const box = node.getBoundingClientRect(),
        children = [...node.children].map((child) => child.getBoundingClientRect());
      return {
        top: box.top,
        bottom: box.bottom,
        children: children.map((child) => ({ top: child.top, bottom: child.bottom })),
      };
    }),
  );
  for (const box of boxes)
    for (const child of box.children) {
      expect(child.top).toBeGreaterThanOrEqual(box.top - 0.5);
      expect(child.bottom).toBeLessThanOrEqual(box.bottom + 1);
    }
  await page.getByRole("button", { name: "编辑分隔内容格式", exact: true }).click();
  await expect(frame.locator("[data-office-editor]")).toContainText("续页注释");
  await page.evaluate(() =>
    (window.paginationHost.view as import("@doffi/editor").DocumentTextEditor).execute({
      type: "word",
      action: "mainStory",
    }),
  );
  expect(await page.evaluate(() => window.paginationHost.session!.snapshot().html)).not.toMatch(
    /sdk-word-note-special|sdk-note-header/,
  );
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
});
test("endnotes use a distinct continuation separator and reserve a notice below text on continued pages", async ({
  page,
}) => {
  const session = createDocumentSession(
      openWordDocument(sectionNotesFiles(false), { kind: "document" }),
    ),
    snapshot = session.snapshot();
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: snapshot.manifest.documentId,
      baseRevision: snapshot.manifest.revision,
      operations: [
        {
          op: "setWordNoteSeparator",
          nodeId: "document-root",
          kind: "endnote",
          type: "separator",
          text: "尾注开始",
        },
        {
          op: "setWordNoteSeparator",
          nodeId: "document-root",
          kind: "endnote",
          type: "continuationSeparator",
          text: "尾注接续",
        },
        {
          op: "setWordNoteSeparator",
          nodeId: "document-root",
          kind: "endnote",
          type: "continuationNotice",
          text: "尾注转下页",
        },
      ],
    }),
  );
  await mount(page, session.serialize());
  const frame = page.frameLocator("iframe");
  await expect(
    frame.locator(".sdk-repeat-header").filter({ hasText: "尾注接续" }).first(),
  ).toBeVisible();
  await expect(
    frame.locator('.sdk-word-notes > [data-sdk-word-note-special="continuationNotice"]').first(),
  ).toHaveText("尾注转下页");
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
  expect(await page.evaluate(() => window.paginationHost.model!.snapshot().html)).not.toMatch(
    /sdk-word-endnotes|sdk-note-header/,
  );
});
test("bookmark controls create and redefine ranges across source paragraphs", async ({ page }) => {
  const session = createDocumentSession(
    openWordDocument(createWordDocument({ documentId: "range-bookmark-ui" }), { kind: "document" }),
  );
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "range-bookmark-ui",
      baseRevision: 0,
      operations: [
        {
          op: "replaceFragment",
          nodeId: "section-root-1",
          html: '<p class="range-a">前😀甲</p><p class="range-b">中间</p><p class="range-c">乙后</p>',
        },
      ],
    }),
  );
  await mount(page, session.serialize(), "edit");
  const frame = page.frameLocator("iframe");
  await frame.locator("[data-office-editor]").evaluate((root) => {
    (root as HTMLElement).focus();
    const a = root.querySelector(".range-a")!,
      c = root.querySelector(".range-c")!,
      range = document.createRange();
    range.setStart(a.firstChild!, 1);
    range.setEnd(c.firstChild!, 1);
    document.getSelection()!.removeAllRanges();
    document.getSelection()!.addRange(range);
  });
  await page.evaluate(() =>
    (window.paginationHost.view as import("@doffi/editor").DocumentTextEditor).execute({
      type: "word",
      action: "bookmarks",
    }),
  );
  await page.getByLabel("书签名称", { exact: true }).fill("Across");
  await page.getByRole("button", { name: "添加书签", exact: true }).click();
  await frame.locator("[data-office-editor]").evaluate((root) => {
    (root as HTMLElement).focus();
    const b = root.querySelector(".range-b")!,
      c = root.querySelector(".range-c")!,
      range = document.createRange();
    range.setStart(b.firstChild!, 0);
    range.setEnd(c.firstChild!, 1);
    document.getSelection()!.removeAllRanges();
    document.getSelection()!.addRange(range);
  });
  await page.evaluate(() =>
    (window.paginationHost.view as import("@doffi/editor").DocumentTextEditor).execute({
      type: "word",
      action: "bookmarks",
    }),
  );
  await page.getByLabel("书签名称", { exact: true }).fill("Renamed");
  await page.getByRole("button", { name: "重定义书签 Across", exact: true }).click();
  expect(
    await page.evaluate(
      () => window.paginationHost.session!.snapshot().metadata.word!.bookmarks[0]!.name,
    ),
  ).toBe("Renamed");
  await page.evaluate(() => window.paginationHost.session!.undo());
  expect(
    await page.evaluate(
      () => window.paginationHost.session!.snapshot().metadata.word!.bookmarks[0]!.name,
    ),
  ).toBe("Across");
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
  await page.evaluate(() => {
    const session = window.paginationHost.session!;
    session.redo();
    const snapshot = session.snapshot(),
      nodeId = session.getNode(snapshot.metadata.sections![0]!.rootNodeId)!.childNodeIds.at(-1)!;
    session.commit(
      session.prepare({
        protocolVersion: 1,
        documentId: snapshot.manifest.documentId,
        baseRevision: snapshot.manifest.revision,
        operations: [
          {
            op: "insertWordField",
            nodeId,
            offset: session.getNode(nodeId)!.text.length,
            instruction: "REF Renamed",
            cachedText: "缓存",
          },
        ],
      }),
    );
  });
  expect(
    await page.evaluate(
      () => window.paginationHost.layout().fields.find((field) => field.kind === "REF")!.value,
    ),
  ).toBe("中间\n乙");
});
test("equation controls create multiple groups and apply whole or component font roles without defeating script sizing", async ({
  page,
}) => {
  await mount(page, createWordDocument({ documentId: "group-math-ui" }), "edit");
  const frame = page.frameLocator("iframe");
  await frame.locator('[data-node-id="paragraph-1"]').click();
  await page.evaluate(() =>
    (window.paginationHost.view as import("@doffi/editor").DocumentTextEditor).execute({
      type: "word",
      action: "insertEquation",
    }),
  );
  await page.getByLabel("多行公式组", { exact: true }).check();
  await page.getByLabel("线性公式", { exact: true }).fill("x_1+漢\ny^2");
  await page.getByRole("button", { name: "插入公式", exact: true }).click();
  await expect(frame.locator("math > mtable > mtr")).toHaveCount(2);
  await page.getByLabel("公式字号", { exact: true }).fill("18");
  await page.getByLabel("公式颜色", { exact: true }).fill("#224466");
  await page.getByLabel("公式字体 eastAsia", { exact: true }).fill("serif");
  await page.getByRole("button", { name: "应用公式字符格式", exact: true }).click();
  await expect(frame.locator('math [data-word-script="eastAsia"]')).toHaveCSS(
    "font-family",
    /serif/,
  );
  const sizes = await frame.locator("math").evaluate((root) => ({
    base: parseFloat(
      getComputedStyle([...root.querySelectorAll("mi")].find((node) => node.textContent === "x")!)
        .fontSize,
    ),
    script: parseFloat(
      getComputedStyle([...root.querySelectorAll("mn")].find((node) => node.textContent === "1")!)
        .fontSize,
    ),
  }));
  expect(sizes.base).toBeCloseTo(24, 0);
  expect(sizes.script).toBeLessThan(sizes.base * 0.9);
  await page.getByLabel("公式字符格式目标", { exact: true }).selectOption({ label: "漢" });
  await page.getByLabel("公式字体 eastAsia", { exact: true }).fill("monospace");
  await page.getByRole("button", { name: "应用公式字符格式", exact: true }).click();
  await expect(frame.locator('math [data-word-script="eastAsia"]')).toHaveCSS(
    "font-family",
    /monospace/,
  );
  await page.evaluate(() => window.paginationHost.session!.undo());
  await expect(frame.locator('math [data-word-script="eastAsia"]')).toHaveCSS(
    "font-family",
    /serif/,
  );
  expect(await page.evaluate(() => window.paginationHost.session!.snapshot().html)).not.toMatch(
    /<math|<mtable/,
  );
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
  await page.getByLabel("线性公式", { exact: true }).click();
  await page.getByLabel("线性公式", { exact: true }).fill("\\matrix{a&b\\\\c&d}");
  await page.getByRole("button", { name: "按线性输入重建公式", exact: true }).click();
  await page.getByLabel("公式矩阵或数组行数", { exact: true }).fill("3");
  await page.getByLabel("公式矩阵列数", { exact: true }).fill("3");
  await page.getByLabel("公式矩阵各列对齐", { exact: true }).fill("left,center,right");
  await page.getByRole("button", { name: "应用公式组件结构", exact: true }).click();
  await expect(frame.locator("math mtable mtable > mtr")).toHaveCount(3);
  await expect(frame.locator("math mtable mtable > mtr > mtd")).toHaveCount(9);
  const files = await page.evaluate(() =>
      [...window.paginationHost.session!.serialize()].map(
        ([path, bytes]) => [path, [...bytes]] as [string, number[]],
      ),
    ),
    model = openWordDocument(new Map(files.map(([path, bytes]) => [path, new Uint8Array(bytes)])), {
      kind: "document",
    }),
    reopened = openWordDocument(
      importDocx(exportDocx(model).bytes, {
        documentId: "component-reopen",
        unsupportedContent: "reject",
      }),
      { kind: "document" },
    );
  expect(
    reopened
      .snapshot()
      .metadata.word!.equations![0]!.nodes.filter((node) => node.kind === "matrixRow"),
  ).toHaveLength(3);
  expect(
    reopened.snapshot().metadata.word!.equations![0]!.nodes.find((node) => node.kind === "text")!
      .style!.fontSize,
  ).toBe(18);
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
});
function positionedDrawingFiles(columns = false): DocumentPackage {
  const session = createDocumentSession(
      openWordDocument(
        createWordDocument({
          documentId: "drawing-positions",
          pageSetup: {
            unit: "pt",
            width: 300,
            height: 240,
            margins: { top: 30, right: 36, bottom: 42, left: 48 },
          },
        }),
        { kind: "document" },
      ),
    ),
    apply = (operations: import("@doffi/core").DocumentOperation[]) =>
      session.commit(
        session.prepare({
          protocolVersion: 1,
          documentId: "drawing-positions",
          baseRevision: session.snapshot().manifest.revision,
          operations,
        }),
      );
  const text = columns
    ? Array.from({ length: 32 }, (_, index) => "LINE" + index).join("<br>")
    : "ALPHA<br>TARGET<br>LAST";
  apply([
    {
      op: "replaceFragment",
      nodeId: "section-root-1",
      html:
        '<p class="position-first" style="margin:0;line-height:18pt;text-indent:18pt">' +
        text +
        "</p>" +
        (columns
          ? ""
          : '<p class="position-second" style="margin:0;line-height:18pt;break-before:page;text-indent:18pt">' +
            text +
            "</p>"),
    },
    { op: "setWordOptions", nodeId: "document-root", mirrorMargins: true, noColumnBalance: true },
    ...(columns
      ? [
          {
            op: "setWordSection" as const,
            nodeId: "section-root-1",
            changes: { columns: { count: 2, gap: 18 } },
          },
        ]
      : []),
  ]);
  const ids = session.getNode("section-root-1")!.childNodeIds;
  for (const [index, line] of (columns ? [11, 24] : ids.map(() => 1)).entries()) {
    const nodeId = columns ? ids[0]! : ids[index]!,
      text = session.getNode(nodeId)!.text,
      token = columns ? "LINE" + line : "TARGET";
    apply([
      {
        op: "insertWordDrawing",
        nodeId,
        offset: text.indexOf(token),
        geometry: { x: 6, y: 4, width: 12, height: 9, rotation: 0, flipH: false, flipV: false },
        layout: {
          mode: "floating",
          wrap: "none",
          horizontal: columns ? "column" : "page",
          vertical: columns ? "line" : "page",
        },
        object: { kind: "shape", preset: "rect" },
      },
    ]);
  }
  return session.serialize();
}
test("floating coordinates use page regions and character anchors independently of paragraph indentation on mirrored pages", async ({
  page,
}) => {
  await mount(page, positionedDrawingFiles(), "edit");
  const frame = page.frameLocator("iframe");
  for (const horizontal of [
    "page",
    "margin",
    "column",
    "leftMargin",
    "rightMargin",
    "insideMargin",
    "outsideMargin",
    "character",
  ] as const) {
    await page.evaluate(async (horizontal) => {
      const session = window.paginationHost.session!,
        snapshot = session.snapshot();
      session.commit(
        session.prepare({
          protocolVersion: 1,
          documentId: snapshot.manifest.documentId,
          baseRevision: snapshot.manifest.revision,
          operations: snapshot.metadata.word!.drawings!.map((drawing) => ({
            op: "setWordDrawingLayout",
            nodeId: drawing.nodeId,
            layout: { mode: "floating", wrap: "none", horizontal, vertical: "page" },
          })),
        }),
      );
      await window.paginationHost.view!.settled();
    }, horizontal);
    const boxes = await frame.locator("body").evaluate((body) =>
      [...body.querySelectorAll<HTMLElement>(".sdk-word-floating-page [data-word-drawing]")].map(
        (node) => {
          const container = node.closest<HTMLElement>(".sdk-word-floating-page")!,
            paper = body.querySelector<HTMLElement>(
              '.sdk-paper[data-page-id="' + container.dataset.pageId + '"]',
            )!,
            r = node.getBoundingClientRect(),
            p = paper.getBoundingClientRect(),
            paragraph = body.querySelector<HTMLElement>(
              '[data-node-id="' + node.dataset.sdkWordSourceParagraph + '"]',
            )!,
            walker = document.createTreeWalker(paragraph, NodeFilter.SHOW_TEXT);
          let character = 0;
          for (let text = walker.nextNode(); text; text = walker.nextNode()) {
            if (text.parentElement?.closest("[data-word-drawing]")) continue;
            const from = text.textContent?.indexOf("TARGET") ?? -1;
            if (from < 0) continue;
            const range = document.createRange();
            range.setStart(text, from);
            range.setEnd(text, from + 1);
            character = (range.getBoundingClientRect().left - p.left) * 0.75;
            break;
          }
          return {
            x: (r.left - p.left) * 0.75,
            y: (r.top - p.top) * 0.75,
            character,
            hit: !!document
              .elementFromPoint(r.left + r.width / 2, r.top + r.height / 2)
              ?.closest(".sdk-word-floating-page"),
          };
        },
      ),
    );
    expect(boxes).toHaveLength(2);
    for (const [index, box] of boxes.entries()) {
      const left = index === 0 ? 48 : 36,
        right = index === 0 ? 36 : 48,
        base =
          horizontal === "page" || horizontal === "leftMargin"
            ? 0
            : horizontal === "margin" || horizontal === "column"
              ? left
              : horizontal === "rightMargin"
                ? 300 - right
                : horizontal === "insideMargin"
                  ? index === 0
                    ? 0
                    : 300 - right
                  : horizontal === "outsideMargin"
                    ? index === 0
                      ? 300 - right
                      : 0
                    : box.character;
      expect(box.x).toBeCloseTo(base + 6, 0);
      expect(box.y).toBeCloseTo(4, 0);
      expect(box.hit).toBe(true);
    }
    expect(
      await page.evaluate(() => {
        const report = window.paginationHost.layout();
        return window.paginationHost
          .session!.snapshot()
          .metadata.word!.drawings!.map((drawing) => report.nodePages[drawing.nodeId]);
      }),
    ).toEqual([["page-1"], ["page-2"]]);
  }
  expect(await page.evaluate(() => window.paginationHost.session!.snapshot().html)).not.toMatch(
    /sdk-word-floating|sdk-word-anchor|visibility:hidden/,
  );
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
});
test("floating vertical references can paint in both margins and locate the actual anchor line", async ({
  page,
}) => {
  await mount(page, positionedDrawingFiles(), "edit");
  for (const vertical of [
    "page",
    "margin",
    "topMargin",
    "bottomMargin",
    "paragraph",
    "line",
  ] as const) {
    await page.evaluate(async (vertical) => {
      const session = window.paginationHost.session!,
        snapshot = session.snapshot();
      session.commit(
        session.prepare({
          protocolVersion: 1,
          documentId: snapshot.manifest.documentId,
          baseRevision: snapshot.manifest.revision,
          operations: snapshot.metadata.word!.drawings!.map((drawing) => ({
            op: "setWordDrawingLayout",
            nodeId: drawing.nodeId,
            layout: { mode: "floating", wrap: "none", horizontal: "page", vertical },
          })),
        }),
      );
      await window.paginationHost.view!.settled();
    }, vertical);
    const boxes = await page
      .frameLocator("iframe")
      .locator("body")
      .evaluate((body) =>
        [...body.querySelectorAll<HTMLElement>(".sdk-word-floating-page [data-word-drawing]")].map(
          (node) => {
            const paper = body.querySelector<HTMLElement>(
                '.sdk-paper[data-page-id="' +
                  node.closest<HTMLElement>(".sdk-word-floating-page")!.dataset.pageId +
                  '"]',
              )!,
              r = node.getBoundingClientRect(),
              p = paper.getBoundingClientRect(),
              paragraph = body.querySelector<HTMLElement>(
                '[data-node-id="' + node.dataset.sdkWordSourceParagraph + '"]',
              )!;
            let line = 0;
            const walker = document.createTreeWalker(paragraph, NodeFilter.SHOW_TEXT);
            for (let text = walker.nextNode(); text; text = walker.nextNode()) {
              const from = text.textContent?.indexOf("TARGET") ?? -1;
              if (from < 0) continue;
              const range = document.createRange();
              range.setStart(text, from);
              range.setEnd(text, from + 1);
              line = (range.getBoundingClientRect().top - p.top) * 0.75;
              break;
            }
            return {
              y: (r.top - p.top) * 0.75,
              paragraph: (paragraph.getBoundingClientRect().top - p.top) * 0.75,
              line,
              hit: !!document
                .elementFromPoint(r.left + 2, r.top + 2)
                ?.closest(".sdk-word-floating-page"),
            };
          },
        ),
      );
    expect(boxes).toHaveLength(2);
    for (const box of boxes) {
      const base =
        vertical === "page" || vertical === "topMargin"
          ? 0
          : vertical === "margin"
            ? 30
            : vertical === "bottomMargin"
              ? 198
              : vertical === "paragraph"
                ? box.paragraph
                : box.line;
      expect(Math.abs(box.y - base - 4)).toBeLessThan(vertical === "line" ? 3 : 0.6);
      expect(box.hit).toBe(true);
    }
  }
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
});
test("one paragraph can anchor floating drawings in different physical columns and pages without duplicating source atoms", async ({
  page,
}) => {
  await mount(page, positionedDrawingFiles(true), "edit");
  const frame = page.frameLocator("iframe");
  const boxes = await frame.locator("body").evaluate((body) =>
    [...body.querySelectorAll<HTMLElement>(".sdk-word-floating-page [data-word-drawing]")].map(
      (node) => {
        const parent = node.closest<HTMLElement>(".sdk-word-floating-page")!,
          paper = body.querySelector<HTMLElement>(
            '.sdk-paper[data-page-id="' + parent.dataset.pageId + '"]',
          )!,
          p = paper.getBoundingClientRect(),
          r = node.getBoundingClientRect();
        return {
          page: parent.dataset.pageId,
          x: (r.left - p.left) * 0.75,
          y: (r.top - p.top) * 0.75,
        };
      },
    ),
  );
  expect(boxes).toHaveLength(2);
  expect(boxes[0]!.page).toBe("page-1");
  expect(boxes[0]!.x).toBeCloseTo(171, 0);
  expect(boxes[1]!.page).toBe("page-2");
  expect(boxes[1]!.x).toBeCloseTo(42, 0);
  const source = await page.evaluate(() => window.paginationHost.session!.snapshot());
  expect(source.metadata.word!.drawings).toHaveLength(2);
  expect(source.html).not.toMatch(/sdk-word-anchor|sdk-word-floating/);
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
});
test("text anchors over a floating source atom highlight its painted margin instance", async ({
  page,
}) => {
  const session = createDocumentSession(
      openWordDocument(positionedDrawingFiles(), { kind: "document" }),
    ),
    snapshot = session.snapshot(),
    paragraph = session.getNode(snapshot.metadata.sections![0]!.rootNodeId)!.childNodeIds[0]!;
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: snapshot.manifest.documentId,
      baseRevision: snapshot.manifest.revision,
      operations: [
        {
          op: "addAnnotation",
          annotation: {
            annotationId: "floating-anchor",
            documentId: snapshot.manifest.documentId,
            baseRevision: snapshot.manifest.revision,
            status: "open",
            anchor: { kind: "text", nodeId: paragraph, from: 6, to: 7, quote: "\ufffc" },
            instruction: "检查图形",
          },
        },
      ],
    }),
  );
  await mount(page, session.serialize(), "interaction");
  await page.evaluate(() => window.paginationHost.interaction!.focusAnnotation("floating-anchor"));
  const frame = page.frameLocator("iframe"),
    object = await frame
      .locator(".sdk-word-floating-page [data-word-drawing]")
      .first()
      .boundingBox(),
    highlight = await frame.locator("[data-sdk-highlight]").boundingBox();
  expect(object).not.toBeNull();
  expect(highlight).not.toBeNull();
  expect(Math.abs(highlight!.x - object!.x)).toBeLessThan(2);
  expect(Math.abs(highlight!.y - object!.y)).toBeLessThan(2);
  expect(Math.abs(highlight!.width - object!.width)).toBeLessThan(2);
  expect(Math.abs(highlight!.height - object!.height)).toBeLessThan(2);
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
});
function drawingFiles(): DocumentPackage {
  const session = createDocumentSession(
    openWordDocument(createWordDocument({ documentId: "word-drawings" }), { kind: "document" }),
  );
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "word-drawings",
      baseRevision: 0,
      operations: [
        {
          op: "replaceText",
          nodeId: "paragraph-1",
          from: 0,
          to: 0,
          expectedText: "",
          text: "前后正文",
        },
        {
          op: "insertWordDrawing",
          nodeId: "paragraph-1",
          offset: 1,
          geometry: {
            x: 30,
            y: 20,
            width: 120,
            height: 75,
            rotation: 0,
            flipH: false,
            flipV: false,
          },
          layout: { mode: "floating", wrap: "none" },
          object: { kind: "shape", preset: "roundRect", text: "形状文字" },
        },
        {
          op: "insertWordDrawing",
          nodeId: "paragraph-1",
          offset: 0,
          geometry: {
            x: 230,
            y: 20,
            width: 120,
            height: 75,
            rotation: 0,
            flipH: false,
            flipV: false,
          },
          layout: { mode: "floating", wrap: "none" },
          object: { kind: "textBox", text: "文本框文字" },
        },
      ],
    }),
  );
  return session.serialize();
}
for (const layout of ["pages", "continuous"] as const)
  test(
    "Word drawing composition preserves source atoms while neighbouring native input edits " +
      layout,
    async ({ page }) => {
      await mount(page, drawingFiles(), "edit", { layout });
      const frame = page.frameLocator("iframe"),
        paragraph = frame.locator('[data-node-id="paragraph-1"]');
      await expect(frame.locator(".sdk-word-drawing-text:visible")).toHaveCount(2);
      await expect(frame.locator(".sdk-word-drawing-text:visible").first()).toContainText(
        "文本框文字",
      );
      expect(
        await frame
          .locator('[data-word-drawing="shape"]:visible')
          .evaluate((node) => Math.round(node.getBoundingClientRect().width)),
      ).toBe(160);
      await paragraph.evaluate((node) => {
        const walker = document.createTreeWalker(node, 4);
        let text = walker.nextNode();
        while (text && text.parentElement?.closest("[data-word-drawing]")) text = walker.nextNode();
        if (!text) throw new Error("Missing surrounding source text");
        const range = document.createRange();
        range.setStart(text, 1);
        range.collapse(true);
        (node.closest("[data-office-editor]") as HTMLElement).focus();
        const selection = document.getSelection()!;
        selection.removeAllRanges();
        selection.addRange(range);
      });
      await page.keyboard.insertText("EDIT");
      const result = await page.evaluate(() => ({
        snapshot: window.paginationHost.session!.snapshot(),
        text: window.paginationHost.session!.getNode("paragraph-1")!.text,
        errors: window.paginationHost.errors,
      }));
      expect(result.text).toContain("EDIT");
      expect(result.snapshot.metadata.word!.drawings).toHaveLength(2);
      expect(result.snapshot.html).not.toMatch(
        /sdk-word|sdk-preset|contenteditable|<svg|position:absolute/,
      );
      expect(result.errors).toEqual([]);
    },
  );
test("Word objects drag and resize through one-step history and edit their text in the shared story session", async ({
  page,
}) => {
  await mount(page, drawingFiles(), "edit");
  const frame = page.frameLocator("iframe"),
    shape = frame.locator('[data-word-drawing="shape"]:visible'),
    before = await page.evaluate(() => window.paginationHost.session!.snapshot()),
    object = before.metadata.word!.drawings!.find((item) => item.kind === "shape")!,
    box = await shape.boundingBox();
  expect(box).not.toBeNull();
  await page.mouse.move(box!.x + 20, box!.y + 20);
  await page.mouse.down();
  await page.mouse.move(box!.x + 60, box!.y + 40, { steps: 4 });
  await page.mouse.up();
  const moved = await page.evaluate(
    () =>
      window.paginationHost
        .session!.snapshot()
        .metadata.word!.drawings!.find((item) => item.kind === "shape")!.geometry,
  );
  expect(moved.x).toBeCloseTo(object.geometry.x + 30);
  expect(moved.y).toBeCloseTo(object.geometry.y + 15);
  await page.evaluate(() => window.paginationHost.session!.undo());
  expect((await page.evaluate(() => window.paginationHost.session!.snapshot())).html).toBe(
    before.html,
  );
  await shape.click({ position: { x: 10, y: 10 } });
  await page.getByRole("button", { name: "编辑对象文字", exact: true }).click();
  const text = frame.locator("[data-office-editor] p").first();
  await text.click();
  await page.keyboard.press("End");
  await page.keyboard.insertText("编辑");
  await page.evaluate(() =>
    (window.paginationHost.view as import("@doffi/editor").DocumentTextEditor).execute({
      type: "word",
      action: "mainStory",
    }),
  );
  await expect(shape.locator(".sdk-word-drawing-text")).toContainText("形状文字编辑");
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
});
test("moving an aligned floating object starts at its displayed position and releases alignment in the same undo step", async ({
  page,
}) => {
  await mount(page, drawingFiles(), "edit");
  const frame = page.frameLocator("iframe");
  await page.evaluate(async () => {
    const session = window.paginationHost.session!,
      snapshot = session.snapshot(),
      shape = snapshot.metadata.word!.drawings!.find((drawing) => drawing.kind === "shape")!;
    session.commit(
      session.prepare({
        protocolVersion: 1,
        documentId: snapshot.manifest.documentId,
        baseRevision: snapshot.manifest.revision,
        operations: [
          {
            op: "setWordDrawingLayout",
            nodeId: shape.nodeId,
            layout: {
              mode: "floating",
              wrap: "none",
              horizontal: "page",
              vertical: "margin",
              alignH: "center",
              alignV: "center",
            },
          },
        ],
      }),
    );
    await window.paginationHost.view!.settled();
  });
  const shape = frame.locator('[data-word-drawing="shape"]:visible'),
    before = await shape.boundingBox();
  expect(before).not.toBeNull();
  await shape.click({ position: { x: 15, y: 15 } });
  await expect(page.getByLabel("对象水平对齐", { exact: true })).toHaveValue("center");
  await expect(page.getByLabel("对象垂直对齐", { exact: true })).toHaveValue("center");
  await shape.scrollIntoViewIfNeeded();
  const position = await shape.evaluate((node) => {
      const r = node.getBoundingClientRect(),
        p = node.closest(".sdk-word-floating-page")!.getBoundingClientRect();
      return { x: r.left - p.left, y: r.top - p.top };
    }),
    start = await shape.boundingBox();
  await page.mouse.move(start!.x + 20, start!.y + 20);
  await page.mouse.down();
  await page.mouse.move(start!.x + 60, start!.y + 40, { steps: 4 });
  await page.mouse.up();
  await page.evaluate(() => window.paginationHost.view!.settled());
  const moved = await shape.evaluate((node) => {
    const r = node.getBoundingClientRect(),
      p = node.closest(".sdk-word-floating-page")!.getBoundingClientRect();
    return { x: r.left - p.left, y: r.top - p.top };
  });
  expect(moved.x - position.x).toBeCloseTo(40, 0);
  expect(moved.y - position.y).toBeCloseTo(20, 0);
  const state = await page.evaluate(
    () =>
      window.paginationHost
        .session!.snapshot()
        .metadata.word!.drawings!.find((drawing) => drawing.kind === "shape")!,
  );
  expect(state.layout.alignH).toBeUndefined();
  expect(state.layout.alignV).toBeUndefined();
  await page.evaluate(async () => {
    window.paginationHost.session!.undo();
    await window.paginationHost.view!.settled();
  });
  await expect
    .poll(() =>
      shape.evaluate((node, position) => {
        if (!node.isConnected) return false;
        const r = node.getBoundingClientRect(),
          p = node.closest(".sdk-word-floating-page")!.getBoundingClientRect();
        return (
          r.width > 0 &&
          Math.abs(r.left - p.left - position.x) < 0.5 &&
          Math.abs(r.top - p.top - position.y) < 0.5
        );
      }, position),
    )
    .toBe(true);
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
});
test("physical columns keep one source paragraph, precise selection and reversible rich input", async ({
  page,
}) => {
  await mount(page, columnFiles(), "edit");
  const before = await page.evaluate(() => window.paginationHost.session!.snapshot()),
    main = before.metadata.word!.stories.find((story) => story.kind === "main")!;
  const frame = page.frameLocator("iframe"),
    paragraph = frame.locator("[data-office-editor] p").first();
  const positions = await paragraph.evaluate((element) => {
    const base = element.closest(".sdk-page-flow")!.getBoundingClientRect(),
      walker = document.createTreeWalker(element, 4),
      result: { x: number; y: number }[] = [];
    for (let node = walker.nextNode(); node; node = walker.nextNode()) {
      const range = document.createRange();
      range.selectNodeContents(node);
      for (const rect of range.getClientRects())
        if (rect.width > 0.1) result.push({ x: rect.x - base.x, y: rect.y - base.y });
    }
    return result;
  });
  expect(positions.some((rect) => rect.x > 150 && rect.y < 256)).toBe(true);
  expect(positions.some((rect) => rect.y > 344)).toBe(true);
  await paragraph.evaluate((element) => {
    const walker = document.createTreeWalker(element, 4),
      base = element.closest(".sdk-page-flow")!.getBoundingClientRect();
    for (let node = walker.nextNode(); node; node = walker.nextNode()) {
      const range = document.createRange();
      range.selectNodeContents(node);
      if (range.getBoundingClientRect().left - base.left > 150) {
        range.setStart(node, 0);
        range.collapse(true);
        const selection = document.getSelection()!;
        selection.removeAllRanges();
        selection.addRange(range);
        (element.closest("[data-office-editor]") as HTMLElement).focus();
        return;
      }
    }
    throw new Error("Missing second-column text.");
  });
  await page.keyboard.press("Enter");
  await page.keyboard.insertText("EDIT");
  const result = await page.evaluate(() => ({
    snapshot: window.paginationHost.session!.snapshot(),
    text: window.paginationHost.session!.getNode(
      window.paginationHost
        .session!.snapshot()
        .metadata.word!.stories.find((story) => story.kind === "main")!.rootNodeId,
    )!.text,
    errors: window.paginationHost.errors,
  }));
  expect(result.text.replace("EDIT", "")).toBe(
    await page.evaluate((id) => window.paginationHost.model!.getNode(id)!.text, main.rootNodeId),
  );
  expect(result.snapshot.html).not.toMatch(/position:relative|sdk-page|sdk-repeat-header/);
  expect(result.errors).toEqual([]);
});
test("table header projections repeat at every physical column without entering source", async ({
  page,
}) => {
  await mount(page, columnFiles(true));
  const frame = page.frameLocator("iframe"),
    headers = frame.locator(".sdk-repeat-header");
  expect(await headers.count()).toBeGreaterThan(1);
  const boxes = await headers.evaluateAll((nodes) =>
    nodes.map((node) => ({
      x: node.getBoundingClientRect().left,
      y: node.getBoundingClientRect().top,
    })),
  );
  expect(
    boxes.some((box, index) =>
      boxes.some(
        (other, ordinal) =>
          index !== ordinal && Math.abs(box.y - other.y) < 1 && Math.abs(box.x - other.x) > 100,
      ),
    ),
  ).toBe(true);
  expect((await page.evaluate(() => window.paginationHost.model!.snapshot())).html).not.toContain(
    "sdk-repeat-header",
  );
});
test("native list keyboard input adjusts levels, exits an empty list and shares undo history", async ({
  page,
}) => {
  const session = createDocumentSession(
    openWordDocument(createWordDocument({ documentId: "list-input" }), { kind: "document" }),
  );
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "list-input",
      baseRevision: 0,
      operations: [
        {
          op: "replaceText",
          nodeId: "paragraph-1",
          from: 0,
          to: 0,
          expectedText: "",
          text: "First",
        },
      ],
    }),
  );
  await mount(page, session.serialize(), "edit");
  const frame = page.frameLocator("iframe"),
    first = frame.locator('[data-node-id="paragraph-1"]');
  await first.click();
  await page.evaluate(() =>
    (window.paginationHost.view as import("@doffi/editor").DocumentTextEditor).execute({
      type: "word",
      action: "list",
      kind: "numbering",
    }),
  );
  await page.keyboard.press("End");
  await page.keyboard.press("Enter");
  await page.keyboard.press("Tab");
  await expect(frame.locator("[data-office-editor] p").nth(1)).toHaveAttribute(
    "data-word-level",
    "1",
  );
  await page.keyboard.press("Shift+Tab");
  await expect(frame.locator("[data-office-editor] p").nth(1)).toHaveAttribute(
    "data-word-level",
    "0",
  );
  await page.keyboard.press("Enter");
  await expect(frame.locator("[data-office-editor] p").nth(1)).not.toHaveAttribute(
    "data-word-numbering",
  );
  await page.keyboard.press("Control+z");
  await expect(frame.locator("[data-office-editor] p").nth(1)).toHaveAttribute(
    "data-word-numbering",
  );
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
});
test("dragging a source selection at the outer viewport edge scrolls and extends its range", async ({
  page,
}) => {
  await mount(page, columnFiles(), "edit");
  const frame = page.frameLocator("iframe"),
    first = frame.locator("[data-office-editor] p").first(),
    box = await first.boundingBox();
  expect(box).not.toBeNull();
  const stage = page.locator(".stage"),
    bounds = await stage.boundingBox();
  expect(bounds).not.toBeNull();
  await page.mouse.move(box!.x + 8, box!.y + 12);
  await page.mouse.down();
  await page.mouse.move(box!.x + 40, bounds!.y + bounds!.height - 3, { steps: 5 });
  await expect.poll(() => stage.evaluate((node) => node.scrollTop)).toBeGreaterThan(30);
  await page.mouse.up();
  expect(await first.evaluate(() => document.getSelection()!.toString().length)).toBeGreaterThan(
    20,
  );
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
});
test("clear formatting uses source defaults and formatting marks survive input without entering source", async ({
  page,
}) => {
  const session = createDocumentSession(
    openWordDocument(createWordDocument({ documentId: "reset-format" }), { kind: "document" }),
  );
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "reset-format",
      baseRevision: 0,
      operations: [
        {
          op: "replaceText",
          nodeId: "paragraph-1",
          from: 0,
          to: 0,
          expectedText: "",
          text: "a b 😀",
        },
      ],
    }),
  );
  await mount(page, session.serialize(), "edit");
  const frame = page.frameLocator("iframe"),
    paragraph = frame.locator('[data-node-id="paragraph-1"]');
  await paragraph.evaluate((node) => {
    const range = document.createRange();
    range.selectNodeContents(node);
    const selection = document.getSelection()!;
    selection.removeAllRanges();
    selection.addRange(range);
    (node.closest("[data-office-editor]") as HTMLElement).focus();
  });
  await page.evaluate(() => {
    const editor = window.paginationHost.view as import("@doffi/editor").DocumentTextEditor;
    editor.execute({ type: "format", property: "font-size", value: "32pt" });
    editor.execute({ type: "format", property: "font-family", value: '"Liberation Mono"' });
    editor.execute({ type: "word", action: "clearFormat" });
    editor.execute({ type: "word", action: "showMarks", visible: true });
  });
  expect(
    await paragraph.evaluate((node) =>
      [...node.querySelectorAll<HTMLElement>("span")]
        .filter((span) => span.textContent?.trim())
        .some((span) => Math.abs(parseFloat(getComputedStyle(span).fontSize) - 16) < 0.05),
    ),
  ).toBe(true);
  await expect(paragraph.locator(".sdk-space-mark")).toHaveCount(2);
  await paragraph.click();
  await page.keyboard.press("End");
  await page.keyboard.insertText(" next");
  await expect(paragraph.locator(".sdk-space-mark")).toHaveCount(3);
  const snapshot = await page.evaluate(() => window.paginationHost.session!.snapshot());
  expect(snapshot.html).not.toMatch(/data-doffi-show-marks|sdk-space-mark|data-sdk-space-mark/);
  expect(
    await page.evaluate(() => window.paginationHost.session!.getNode("paragraph-1")!.text),
  ).toBe("a b 😀 next");
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
});
test("clearing a whole superscript run removes its old line box and later style changes keep the cleared range editable", async ({
  page,
}) => {
  const session = createDocumentSession(
    openWordDocument(createWordDocument({ documentId: "reset-boxes" }), { kind: "document" }),
  );
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "reset-boxes",
      baseRevision: 0,
      operations: [
        {
          op: "replaceFragment",
          nodeId: "section-root-1",
          html: '<p class="cleared word-normal" data-word-style="Normal"><span style="font-size:32pt;background-color:#ffff00;vertical-align:super;text-decoration:underline">CLEAR</span></p><p class="reference word-normal" data-word-style="Normal">BASE</p>',
        },
      ],
    }),
  );
  const paragraph = session.getNode("section-root-1")!.childNodeIds[0]!;
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "reset-boxes",
      baseRevision: session.snapshot().manifest.revision,
      operations: [
        { op: "setWordReview", nodeId: "document-root", trackChanges: true, author: "A" },
        { op: "resetWordText", nodeId: paragraph, from: 0, to: 5, expectedText: "CLEAR" },
      ],
    }),
  );
  await mount(page, session.serialize(), "edit");
  const frame = page.frameLocator("iframe");
  const boxes = await frame.locator("body").evaluate((body) =>
    [".cleared", ".reference"].map((selector) => {
      const paragraph = body.querySelector<HTMLElement>(selector)!,
        walker = document.createTreeWalker(paragraph, NodeFilter.SHOW_TEXT);
      let text = walker.nextNode();
      while (text && !text.textContent?.trim()) text = walker.nextNode();
      const range = document.createRange();
      range.selectNodeContents(text!);
      const p = paragraph.getBoundingClientRect(),
        r = range.getBoundingClientRect();
      return { height: p.height, top: r.top - p.top };
    }),
  );
  expect(boxes[0]!.height).toBeCloseTo(boxes[1]!.height, 0);
  expect(boxes[0]!.top).toBeCloseTo(boxes[1]!.top, 0);
  await page.evaluate(async () => {
    const session = window.paginationHost.session!,
      snapshot = session.snapshot();
    session.commit(
      session.prepare({
        protocolVersion: 1,
        documentId: snapshot.manifest.documentId,
        baseRevision: snapshot.manifest.revision,
        operations: [
          {
            op: "setWordStyle",
            nodeId: "document-root",
            style: { styleId: "Normal", kind: "paragraph", cssClass: "word-normal", default: true },
            runProperties: { "font-size": "18pt", color: "#112299" },
          },
        ],
      }),
    );
    await window.paginationHost.view!.settled();
  });
  expect(
    await frame
      .locator("[data-word-reset-text]")
      .evaluate((node) => getComputedStyle(node.querySelector(".sdk-word-text-run")!).fontSize),
  ).toBe("24px");
  await frame.locator(".cleared").click();
  await page.keyboard.press("End");
  await page.keyboard.insertText("X");
  await expect
    .poll(() => page.evaluate((id) => window.paginationHost.session!.getNode(id)!.text, paragraph))
    .toBe("CLEARX");
  const snapshot = await page.evaluate(() => window.paginationHost.session!.snapshot());
  expect(snapshot.html).not.toContain("sdk-word-text-run");
  expect(snapshot.metadata.word!.revisions.some((revision) => revision.property === "run")).toBe(
    true,
  );
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
});
test("a partial text reset paints one native run box per surviving format and retains neighbouring superscript geometry", async ({
  page,
}) => {
  const session = createDocumentSession(
      openWordDocument(createWordDocument({ documentId: "partial-reset-boxes" }), {
        kind: "document",
      }),
    ),
    old = "font-size:32pt;background-color:#ffff00;vertical-align:super;text-decoration:underline";
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "partial-reset-boxes",
      baseRevision: 0,
      operations: [
        {
          op: "replaceFragment",
          nodeId: "section-root-1",
          html:
            '<p class="cleared word-normal" data-word-style="Normal"><span style="' +
            old +
            '">ABC</span></p><p class="reference word-normal" data-word-style="Normal"><span style="' +
            old +
            '">A</span><span>B</span><span style="' +
            old +
            '">C</span></p>',
        },
      ],
    }),
  );
  const paragraph = session.getNode("section-root-1")!.childNodeIds[0]!;
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "partial-reset-boxes",
      baseRevision: session.snapshot().manifest.revision,
      operations: [{ op: "resetWordText", nodeId: paragraph, from: 1, to: 2, expectedText: "B" }],
    }),
  );
  await mount(page, session.serialize(), "edit");
  const boxes = await page
    .frameLocator("iframe")
    .locator("body")
    .evaluate((body) =>
      [".cleared", ".reference"].map((selector) => {
        const paragraph = body.querySelector<HTMLElement>(selector)!,
          p = paragraph.getBoundingClientRect(),
          walker = document.createTreeWalker(paragraph, NodeFilter.SHOW_TEXT),
          letters: Record<string, { x: number; y: number; width: number; height: number }> = {};
        for (let node = walker.nextNode(); node; node = walker.nextNode())
          for (const letter of ["A", "B", "C"]) {
            const from = node.textContent?.indexOf(letter) ?? -1;
            if (from < 0) continue;
            const range = document.createRange();
            range.setStart(node, from);
            range.setEnd(node, from + 1);
            const r = range.getBoundingClientRect();
            letters[letter] = { x: r.x - p.x, y: r.y - p.y, width: r.width, height: r.height };
          }
        return { height: p.height, letters };
      }),
    );
  expect(boxes[0]!.height).toBeCloseTo(boxes[1]!.height, 0);
  for (const letter of ["A", "B", "C"])
    for (const key of ["x", "y", "width", "height"] as const)
      expect(boxes[0]!.letters[letter]![key]).toBeCloseTo(boxes[1]!.letters[letter]![key], 0);
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
});
test("partial direct formatting retains native run inheritance and matches independent run boxes", async ({
  page,
}) => {
  const session = createDocumentSession(
      openWordDocument(createWordDocument({ documentId: "native-run-boxes" }), {
        kind: "document",
      }),
    ),
    old = "font-size:32pt;background-color:#ffff00;vertical-align:super;text-decoration:underline";
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "native-run-boxes",
      baseRevision: 0,
      operations: [
        {
          op: "replaceFragment",
          nodeId: "section-root-1",
          html:
            '<p class="edited word-normal" data-word-style="Normal"><span style="' +
            old +
            '">ABC</span></p><p class="reference word-normal" data-word-style="Normal"><span style="' +
            old +
            '">A</span><span style="' +
            old +
            ';font-size:12pt;color:#0000aa">B</span><span style="' +
            old +
            '">C</span></p>',
        },
      ],
    }),
  );
  const id = session.getNode("section-root-1")!.childNodeIds[0]!;
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "native-run-boxes",
      baseRevision: session.snapshot().manifest.revision,
      operations: [
        {
          op: "formatWordText",
          nodeId: id,
          from: 1,
          to: 2,
          expectedText: "B",
          properties: { "font-size": "12pt", color: "#0000aa" },
        },
      ],
    }),
  );
  await mount(page, session.serialize(), "edit");
  const geometry = await page
    .frameLocator("iframe")
    .locator("body")
    .evaluate((body) =>
      [".edited", ".reference"].map((selector) => {
        const root = body.querySelector<HTMLElement>(selector)!,
          p = root.getBoundingClientRect(),
          walker = document.createTreeWalker(root, 4),
          letters: Record<string, number[]> = {};
        for (let node = walker.nextNode(); node; node = walker.nextNode())
          for (const letter of ["A", "B", "C"]) {
            const at = node.textContent!.indexOf(letter);
            if (at < 0) continue;
            const r = document.createRange();
            r.setStart(node, at);
            r.setEnd(node, at + 1);
            const box = r.getBoundingClientRect();
            letters[letter] = [box.x - p.x, box.y - p.y, box.width, box.height];
          }
        return { height: p.height, letters };
      }),
    );
  expect(geometry[0]!.height).toBeCloseTo(geometry[1]!.height, 0);
  for (const letter of ["A", "B", "C"])
    geometry[0]!.letters[letter]!.forEach((value, index) =>
      expect(value).toBeCloseTo(geometry[1]!.letters[letter]![index]!, 0),
    );
  const paint = await page
    .frameLocator("iframe")
    .locator(".edited .sdk-word-text-run")
    .filter({ hasText: "B" })
    .evaluate((node) => {
      const css = getComputedStyle(node);
      return {
        vertical: css.verticalAlign,
        background: css.backgroundColor,
        decoration: css.textDecorationLine,
        color: css.color,
      };
    });
  expect(paint).toEqual({
    vertical: "super",
    background: "rgb(255, 255, 0)",
    decoration: "underline",
    color: "rgb(0, 0, 170)",
  });
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
});
test("shrinking a whole run removes the old font strut and keeps typing and undo in the same source", async ({
  page,
}) => {
  const session = createDocumentSession(
    openWordDocument(createWordDocument({ documentId: "small-run-boxes" }), { kind: "document" }),
  );
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "small-run-boxes",
      baseRevision: 0,
      operations: [
        {
          op: "replaceFragment",
          nodeId: "section-root-1",
          html: '<p class="edited word-normal" data-word-style="Normal"><span style="font-size:40pt">SMALL</span></p><p class="reference word-normal" data-word-style="Normal"><span style="font-size:12pt">SMALL</span></p>',
        },
      ],
    }),
  );
  const id = session.getNode("section-root-1")!.childNodeIds[0]!;
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "small-run-boxes",
      baseRevision: session.snapshot().manifest.revision,
      operations: [
        {
          op: "formatWordText",
          nodeId: id,
          from: 0,
          to: 5,
          expectedText: "SMALL",
          properties: { "font-size": "12pt" },
        },
      ],
    }),
  );
  await mount(page, session.serialize(), "edit");
  const frame = page.frameLocator("iframe");
  const heights = await frame
    .locator("body")
    .evaluate((body) =>
      [".edited", ".reference"].map(
        (selector) => body.querySelector(selector)!.getBoundingClientRect().height,
      ),
    );
  expect(heights[0]).toBeCloseTo(heights[1]!, 0);
  await frame.locator(".edited").click();
  await page.keyboard.press("End");
  await page.keyboard.insertText("X");
  await expect
    .poll(() => page.evaluate((id) => window.paginationHost.session!.getNode(id)!.text, id))
    .toBe("SMALLX");
  await page.evaluate(() => window.paginationHost.session!.undo());
  expect(await page.evaluate((id) => window.paginationHost.session!.getNode(id)!.text, id)).toBe(
    "SMALL",
  );
  expect(await page.evaluate(() => window.paginationHost.session!.snapshot().html)).not.toContain(
    "sdk-word-text-run",
  );
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
});
test("a recalculated field keeps its cleared result font and updating the source cache retains the reset", async ({
  page,
}) => {
  const session = createDocumentSession(
      openWordDocument(createWordDocument({ documentId: "reset-field" }), { kind: "document" }),
    ),
    apply = (operations: import("@doffi/core").DocumentOperation[]) =>
      session.commit(
        session.prepare({
          protocolVersion: 1,
          documentId: "reset-field",
          baseRevision: session.snapshot().manifest.revision,
          operations,
        }),
      );
  apply([
    { op: "replaceText", nodeId: "paragraph-1", from: 0, to: 0, expectedText: "", text: "A" },
    {
      op: "insertWordField",
      nodeId: "paragraph-1",
      offset: 1,
      instruction: "NUMPAGES",
      cachedText: "0",
    },
  ]);
  apply([
    {
      op: "formatWordText",
      nodeId: "paragraph-1",
      from: 0,
      to: 2,
      expectedText: "A0",
      properties: { "font-size": "32pt" },
    },
    { op: "resetWordText", nodeId: "paragraph-1", from: 1, to: 2, expectedText: "0" },
  ]);
  await mount(page, session.serialize(), "edit");
  const display = page.frameLocator("iframe").locator(".sdk-field-display");
  await expect(display).toHaveText("1");
  expect(await display.evaluate((node) => getComputedStyle(node).fontSize)).toBe("16px");
  expect(
    await page.evaluate(() => window.paginationHost.session!.getNode("paragraph-1")!.text),
  ).toBe("A0");
  await page.evaluate(() => window.paginationHost.updateFields());
  await expect(display).toHaveText("1");
  expect(await display.evaluate((node) => getComputedStyle(node).fontSize)).toBe("16px");
  expect(
    await page.evaluate(() => window.paginationHost.session!.getNode("paragraph-1")!.text),
  ).toBe("A1");
  expect(await page.evaluate(() => window.paginationHost.session!.snapshot().html)).not.toMatch(
    /sdk-field-display|sdk-word-text-run/,
  );
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
});
test("source projection strips late owned wrappers while preserving bound roots and unowned lookalikes", async ({
  page,
}) => {
  await mount(page, pagedFiles('<p data-node-id="text">文档</p>'));
  const result = await page.evaluate(() => window.paginationHost.lateProjection());
  expect(result.html).toBe(result.original);
  expect(result.instance).toBe(result.original);
  expect(result.shellMissing).toBe(true);
  expect(result.copy).toBe('<span data-node-id="run" style="color:#123456">原文</span>');
  expect(result.unowned).toContain(
    '<span class="sdk-inline-flow" style="display: block;">源内容</span>',
  );
});
test("read-only pagination requires no editing session and preserves the complete package source", async ({
  page,
}) => {
  const files = pagedFiles(
    Array.from(
      { length: 40 },
      (_, index) => '<p data-node-id="p-' + index + '">LINE-' + index + "</p>",
    ).join(""),
  );
  await mount(page, files);
  const state = await page.evaluate(() => ({
    report: window.paginationHost.layout(),
    session: !!window.paginationHost.session,
    html: window.paginationHost.model!.snapshot().html,
  }));
  expect(state.session).toBe(false);
  expect(state.report.pages).toHaveLength(4);
  const frame = page.frameLocator("iframe");
  await expect(frame.locator('[contenteditable="true"]')).toHaveCount(0);
  await expect(frame.locator("script,input,textarea")).toHaveCount(0);
  expect(state.html).not.toMatch(/sdk-page|column-height|contenteditable/);
  expect(state.html).toContain("LINE-39");
  const geometry = await frame
    .locator('[data-node-id="p-12"]')
    .evaluate((node) => ({ y: node.getBoundingClientRect().top, text: node.textContent }));
  expect(geometry.y).toBeGreaterThan(300);
  expect(state.report.nodePages["p-12"]).toEqual(["page-2"]);
});
test("one paragraph spans pages and maps a cross-page selection to one canonical UTF-16 range", async ({
  page,
}) => {
  await mount(
    page,
    pagedFiles(
      '<p data-node-id="long">' +
        Array.from({ length: 40 }, (_, index) => "LINE-" + index + " 😀").join("<br>") +
        "</p>",
    ),
    "interaction",
  );
  const frame = page.frameLocator("iframe");
  await frame.locator('[data-node-id="long"]').evaluate((node) => {
    const range = document.createRange();
    range.selectNodeContents(node);
    const selected = document.getSelection()!;
    selected.removeAllRanges();
    selected.addRange(range);
    document.dispatchEvent(new Event("selectionchange"));
  });
  await expect
    .poll(() => page.evaluate(() => window.paginationHost.interaction!.selection()?.anchor))
    .toEqual(
      await page.evaluate(() => ({
        kind: "text",
        nodeId: "long",
        from: 0,
        to: window.paginationHost.session!.getNode("long")!.text.length,
        quote: window.paginationHost.session!.getNode("long")!.text,
      })),
    );
  expect(
    (await page.evaluate(() => window.paginationHost.layout().nodePages.long))!.length,
  ).toBeGreaterThan(1);
});
test("hiding a preview preserves its accepted pages and revealing it remeasures the visible canvas", async ({
  page,
}) => {
  const errors: string[] = [];
  page.on("pageerror", (error) => errors.push(error.message));
  await mount(
    page,
    pagedFiles(
      Array.from(
        { length: 40 },
        (_, index) => `<p data-node-id="p-${index}">LINE-${index}</p>`,
      ).join(""),
    ),
  );
  const original = await page.evaluate(() => window.paginationHost.layout());
  expect(original.pages).toHaveLength(4);
  const hidden = await page.evaluate(async () => {
    (document.body.firstElementChild as HTMLElement).hidden = true;
    return window.paginationHost.view!.settled();
  });
  expect(hidden.pages).toEqual(original.pages);
  expect(hidden.nodePages).toEqual(original.nodePages);
  await page.evaluate(() => {
    (document.body.firstElementChild as HTMLElement).hidden = false;
  });
  await expect
    .poll(() => page.evaluate(async () => (await window.paginationHost.view!.settled()).pages))
    .toEqual(original.pages);
  await expect
    .poll(() =>
      page
        .frameLocator("iframe")
        .locator('[data-node-id="p-12"]')
        .evaluate((node) => node.getBoundingClientRect().top),
    )
    .toBeGreaterThan(300);
  expect(errors).toEqual([]);
});
test("editing reflows pages immediately and undo restores layout without persisting page wrappers", async ({
  page,
}) => {
  await mount(page, pagedFiles('<p data-node-id="a">Start</p>'), "edit");
  const before = await page.evaluate(() => window.paginationHost.layout().pages.length);
  await page
    .frameLocator("iframe")
    .locator('[data-node-id="a"]')
    .evaluate((node) => {
      const range = document.createRange();
      range.selectNodeContents(node);
      range.collapse(false);
      const selection = document.getSelection()!;
      selection.removeAllRanges();
      selection.addRange(range);
      (document.querySelector("[data-office-editor]") as HTMLElement).focus();
    });
  await page.keyboard.insertText(" long text".repeat(300));
  expect(await page.evaluate(() => window.paginationHost.layout().pages.length)).toBeGreaterThan(
    before,
  );
  expect(await page.evaluate(() => window.paginationHost.session!.snapshot().html)).not.toMatch(
    /sdk-page|column-height|contenteditable/,
  );
  await page.keyboard.press("Control+z");
  expect(await page.evaluate(() => window.paginationHost.layout().pages.length)).toBe(before);
  expect(await page.evaluate(() => window.paginationHost.session!.getNode("a")!.text)).toBe(
    "Start",
  );
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
});
test("repeated table headers remain a display projection while body cells can be edited", async ({
  page,
}) => {
  await mount(page, longTable(), "edit");
  const pages = await page.evaluate(() => window.paginationHost.layout().pages.length);
  expect(pages).toBeGreaterThan(1);
  await expect(page.frameLocator("iframe").locator(".sdk-repeat-header")).toHaveCount(pages - 1);
  await page
    .frameLocator("iframe")
    .locator('[data-node-id="cell-20"]')
    .first()
    .evaluate((node) => {
      const range = document.createRange();
      range.selectNodeContents(node);
      range.collapse(false);
      const selected = document.getSelection()!;
      selected.removeAllRanges();
      selected.addRange(range);
      (document.querySelector("[data-office-editor]") as HTMLElement).focus();
    });
  await page.keyboard.insertText(" EDITED");
  expect(await page.evaluate(() => window.paginationHost.session!.getNode("cell-20")!.text)).toBe(
    "UNIQUE-20 EDITED",
  );
  const html = await page.evaluate(() => window.paginationHost.session!.snapshot().html);
  expect(html.match(/REPEATED HEADER/g)).toHaveLength(1);
  expect(html).not.toMatch(/sdk-repeat|box-decoration-break|contenteditable/);
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
});
test("page fields and TOC are derived; explicit cache update is one core undo step", async ({
  page,
}) => {
  await mount(page, fieldsFiles(), "interaction");
  const state = await page.evaluate(() => ({
    report: window.paginationHost.layout(),
    html: window.paginationHost.session!.snapshot().html,
    revision: window.paginationHost.session!.snapshot().manifest.revision,
  }));
  expect(state.report.pages).toHaveLength(2);
  expect(state.report.fields.find((field) => field.kind === "NUMPAGES")?.value).toBe("2");
  expect(state.report.fields.find((field) => field.kind === "TOC")?.value).toBe(
    "Alpha\t1\nBeta\t2",
  );
  expect(state.html).not.toContain("sdk-field");
  await page.evaluate(() => window.paginationHost.updateFields());
  expect(
    await page.evaluate(() => window.paginationHost.session!.snapshot().manifest.revision),
  ).toBe(state.revision + 1);
  expect(await page.evaluate(() => window.paginationHost.session!.history().undoSteps)).toBe(1);
  await page.evaluate(() => window.paginationHost.session!.undo());
  expect(await page.evaluate(() => window.paginationHost.session!.snapshot().html)).toBe(
    state.html,
  );
});
test("continuous preview remains available and page budgets fail explicitly", async ({ page }) => {
  const files = pagedFiles(
    Array.from(
      { length: 40 },
      (_, index) => '<p data-node-id="p-' + index + '">LINE-' + index + "</p>",
    ).join(""),
  );
  await mount(page, files, "preview", { layout: "continuous" });
  expect(await page.evaluate(() => window.paginationHost.layout().layout)).toBe("continuous");
  await expect(page.frameLocator("iframe").locator(".sdk-paper")).toHaveCount(0);
  await expect(mount(page, files, "preview", { limits: { maxPages: 1 } })).rejects.toThrow(
    "page budget",
  );
});

const prefix = (count: number): string =>
  Array.from({ length: count }, (_, index) => `<p data-node-id="prefix-${index}">LINE</p>`).join(
    "",
  );
async function atStart(page: Page, id: string): Promise<void> {
  await page
    .frameLocator("iframe")
    .locator(`[data-node-id="${id}"]`)
    .first()
    .evaluate((node) => {
      const walker = document.createTreeWalker(node, NodeFilter.SHOW_TEXT),
        text = walker.nextNode()!;
      const range = document.createRange();
      range.setStart(text, 0);
      range.collapse(true);
      const selected = document.getSelection()!;
      selected.removeAllRanges();
      selected.addRange(range);
      (document.querySelector("[data-office-editor]") as HTMLElement).focus();
    });
}
test("headings stay with following text and short paragraphs respect orphan and widow limits", async ({
  page,
}) => {
  await mount(
    page,
    pagedFiles(
      prefix(11) +
        '<h1 data-node-id="heading" style="font-size:12pt;line-height:16pt;break-after:avoid">HEADING</h1><p data-node-id="after">FOLLOWING</p>',
    ),
  );
  expect(await page.evaluate(() => window.paginationHost.layout().nodePages.heading)).toEqual([
    "page-2",
  ]);
  expect(await page.evaluate(() => window.paginationHost.layout().nodePages.after)).toEqual([
    "page-2",
  ]);
  await mount(
    page,
    pagedFiles(
      prefix(11) +
        '<p data-node-id="short" style="orphans:2;widows:2">FIRST<br>SECOND<br>THIRD</p>',
    ),
  );
  expect(await page.evaluate(() => window.paginationHost.layout().nodePages.short)).toEqual([
    "page-2",
  ]);
  await mount(
    page,
    pagedFiles(
      prefix(9) +
        '<p data-node-id="four" style="orphans:2;widows:2">FIRST<br>SECOND<br>THIRD<br>FOURTH</p>',
    ),
  );
  const y = await page
    .frameLocator("iframe")
    .locator('[data-node-id="four"]')
    .evaluate((node) =>
      [...node.childNodes]
        .filter((item) => item.nodeType === 3 && item.textContent)
        .map((item) => {
          const range = document.createRange();
          range.selectNodeContents(item);
          return range.getBoundingClientRect().top;
        }),
    );
  expect(y[1]! - y[0]!).toBeLessThan(30);
  expect(y[2]! - y[1]!).toBeGreaterThan(60);
  expect(y[3]! - y[2]!).toBeLessThan(30);
});
test("unidentified hard page breaks stay canonical after edit, save and reopen", async ({
  page,
}) => {
  await mount(
    page,
    pagedFiles('<p data-node-id="a">AAA<br data-word-break="page">BBB</p>'),
    "edit",
  );
  expect(await page.evaluate(() => window.paginationHost.layout().pages.length)).toBe(2);
  await atStart(page, "a");
  await page.keyboard.insertText("X");
  expect(await page.evaluate(() => window.paginationHost.session!.getNode("a")!.text)).toBe(
    "XAAA\nBBB",
  );
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
  const entries = await page.evaluate(() =>
    [...window.paginationHost.session!.serialize()].map(
      ([path, bytes]) => [path, [...bytes]] as [string, number[]],
    ),
  );
  await page.evaluate((files) => window.paginationHost.mount(files, "preview"), entries);
  expect(await page.evaluate(() => window.paginationHost.layout().pages.length)).toBe(2);
  expect(await page.evaluate(() => window.paginationHost.model!.snapshot().html)).not.toMatch(
    /display:\s*none|sdk-page-gap/,
  );
});
test("direct container text is paginated without losing its source identity", async ({ page }) => {
  await mount(
    page,
    pagedFiles(
      '<div data-node-id="long">' +
        Array.from({ length: 40 }, (_, index) => "LINE-" + index).join("<br>") +
        "</div>",
    ),
  );
  expect(await page.evaluate(() => window.paginationHost.layout().pages.length)).toBeGreaterThan(1);
  expect(
    await page.evaluate(() => window.paginationHost.layout().nodePages.long!.length),
  ).toBeGreaterThan(1);
  expect(await page.frameLocator("iframe").locator('[data-node-id="long"]').count()).toBe(1);
});
test("section metadata order never reorders source or saved edits", async ({ page }) => {
  const files = pagedFiles(
    '<section data-node-id="first"><p data-node-id="a">A</p></section><section data-node-id="second"><p data-node-id="b">B</p></section>',
  );
  const metadata = JSON.parse(new TextDecoder().decode(files.get("metadata.json")));
  metadata.sections = [
    { sectionId: "second", rootNodeId: "second", breakType: "nextPage" },
    { sectionId: "first", rootNodeId: "first" },
  ];
  files.set("metadata.json", new TextEncoder().encode(JSON.stringify(metadata)));
  await mount(page, files, "edit");
  await atStart(page, "a");
  await page.keyboard.insertText("X");
  const state = await page.evaluate(() => ({
    text: window.paginationHost.session!.getNode("root")!.text,
    html: window.paginationHost.session!.snapshot().html,
    errors: window.paginationHost.errors,
  }));
  expect(state.text).toBe("XAB");
  expect(state.html.indexOf('data-node-id="first"')).toBeLessThan(
    state.html.indexOf('data-node-id="second"'),
  );
  expect(state.errors).toEqual([]);
});
test("deletion without target-range support crosses view gaps and preserves graphemes", async ({
  page,
}) => {
  await mount(
    page,
    pagedFiles('<p data-node-id="a">AB漢󠄀😀</p><p data-node-id="b">NEXT</p>'),
    "edit",
  );
  await page
    .frameLocator("iframe")
    .locator('[data-node-id="a"]')
    .evaluate((node) => {
      Object.defineProperty(InputEvent.prototype, "getTargetRanges", {
        configurable: true,
        value: undefined,
      });
      const range = document.createRange();
      range.selectNodeContents(node);
      range.collapse(false);
      const selected = document.getSelection()!;
      selected.removeAllRanges();
      selected.addRange(range);
      (document.querySelector("[data-office-editor]") as HTMLElement).focus();
    });
  await page.keyboard.press("Backspace");
  await page.keyboard.press("Backspace");
  expect(await page.evaluate(() => window.paginationHost.session!.getNode("a")!.text)).toBe("AB");
  await atStart(page, "b");
  await page.keyboard.press("Backspace");
  expect(await page.evaluate(() => window.paginationHost.session!.getNode("a")!.text)).toBe(
    "ABNEXT",
  );
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
});
test("long table rows fragment while continued pages reserve repeated-header space", async ({
  page,
}) => {
  const lines = Array.from({ length: 30 }, (_, index) => "CELL-" + index).join("<br>");
  await mount(
    page,
    pagedFiles(
      '<table data-node-id="table"><tr data-node-id="header" data-word-repeat-header="true"><th data-node-id="h1">HEAD</th><th data-node-id="h2">HEAD2</th></tr><tr data-node-id="long-row"><td data-node-id="left">' +
        lines +
        '</td><td data-node-id="right">' +
        lines +
        '</td></tr><tr data-node-id="tail"><td data-node-id="tail-left">END</td><td data-node-id="tail-right">END2</td></tr></table>',
    ),
  );
  const report = await page.evaluate(() => window.paginationHost.layout());
  expect(report.pages.length).toBeGreaterThan(2);
  await expect(page.frameLocator("iframe").locator(".sdk-repeat-header")).toHaveCount(
    report.pages.length - 1,
  );
  const overlaps = await page
    .frameLocator("iframe")
    .locator('[data-node-id="left"]')
    .evaluate((node) => {
      const headerRects = [...document.querySelectorAll(".sdk-repeat-header")].map((header) =>
          header.getBoundingClientRect(),
        ),
        walker = document.createTreeWalker(node, NodeFilter.SHOW_TEXT);
      let item;
      const bad: string[] = [];
      while ((item = walker.nextNode())) {
        if (!item.textContent) continue;
        const range = document.createRange();
        range.selectNodeContents(item);
        for (const rect of range.getClientRects())
          if (
            headerRects.some(
              (header) => rect.top < header.bottom - 0.5 && rect.bottom > header.top + 0.5,
            )
          )
            bad.push(item.textContent);
      }
      return bad;
    });
  expect(overlaps).toEqual([]);
});
test("mixed container inline text remains visible and authored copy excludes page furniture", async ({
  page,
}) => {
  await mount(
    page,
    pagedFiles(
      '<div data-node-id="mixed">' +
        Array.from({ length: 30 }, (_, index) => "DIRECT-" + index).join("<br>") +
        '<p data-node-id="end">END</p></div>',
    ),
  );
  expect(await page.evaluate(() => window.paginationHost.layout().pages.length)).toBeGreaterThan(2);
  const copy = await page
    .frameLocator("iframe")
    .locator('[data-node-id="mixed"]')
    .evaluate((node) => {
      const range = document.createRange();
      range.selectNodeContents(node);
      const selected = document.getSelection()!;
      selected.removeAllRanges();
      selected.addRange(range);
      const data = new Map<string, string>(),
        event = new Event("copy", { bubbles: true, cancelable: true });
      Object.defineProperty(event, "clipboardData", {
        value: { setData: (type: string, value: string) => data.set(type, value) },
      });
      document.dispatchEvent(event);
      return {
        html: data.get("text/html") ?? "",
        text: data.get("text/plain") ?? "",
        prevented: event.defaultPrevented,
      };
    });
  expect(copy.prevented).toBe(true);
  expect(copy.html).not.toMatch(/sdk-page|sdk-inline-flow|contenteditable|clip-path/);
  expect(copy.text).toContain("DIRECT-29");
  expect(copy.text).toContain("END");
});
test("host interaction events work while CSP blocks executable document content", async ({
  page,
}) => {
  await mount(page, pagedFiles('<p data-node-id="a">SAFE</p>'));
  await page.evaluate(() => {
    (window as unknown as { sourceExecuted: number; sourceViolations: number }).sourceExecuted = 0;
    (window as unknown as { sourceViolations: number }).sourceViolations = 0;
    const frame = document.querySelector("iframe")!.contentDocument!;
    frame.addEventListener("securitypolicyviolation", (event) => {
      if (event.violatedDirective.startsWith("script-src"))
        (window as unknown as { sourceViolations: number }).sourceViolations++;
    });
    const script = frame.createElement("script");
    script.textContent = "parent.sourceExecuted++";
    frame.body.append(script);
    const image = frame.createElement("img");
    image.setAttribute("onerror", "parent.sourceExecuted++");
    image.src = "invalid-source:";
    frame.body.append(image);
  });
  await expect
    .poll(() =>
      page.evaluate(() => (window as unknown as { sourceViolations: number }).sourceViolations),
    )
    .toBeGreaterThanOrEqual(2);
  expect(
    await page.evaluate(() => (window as unknown as { sourceExecuted: number }).sourceExecuted),
  ).toBe(0);
  expect(
    await page
      .frameLocator("iframe")
      .locator('meta[http-equiv="Content-Security-Policy"]')
      .getAttribute("content"),
  ).toContain("default-src 'none'");
});
test("overheight headers stay outside body text and report their clipping", async ({ page }) => {
  const files = pagedFiles(
    '<section data-node-id="main" data-office-story="main"><section data-node-id="sec"><p data-node-id="body-text">BODY FIRST LINE</p></section></section><section data-node-id="head" data-office-story="header"><p data-node-id="header-a">HEADER 1</p><p data-node-id="header-b">HEADER 2</p><p data-node-id="header-c">HEADER 3</p></section>',
  );
  const metadata = JSON.parse(new TextDecoder().decode(files.get("metadata.json")));
  metadata.pageSetup = {
    ...metadata.pageSetup,
    height: 300,
    margins: { top: 72, bottom: 72, left: 24, right: 24 },
  };
  metadata.sections = [
    { sectionId: "s", rootNodeId: "sec", headers: { default: "hdr" }, headerDistance: 36 },
  ];
  metadata.word = {
    stories: [
      { storyId: "main", kind: "main", rootNodeId: "main" },
      { storyId: "hdr", kind: "header", rootNodeId: "head" },
    ],
    styles: [],
    numbering: [],
    hyperlinks: [],
    bookmarks: [],
    fields: [],
    tabs: [],
    tables: [],
    revisions: [],
  };
  files.set("metadata.json", new TextEncoder().encode(JSON.stringify(metadata)));
  await mount(page, files);
  expect(
    (await page.evaluate(() => window.paginationHost.layout().diagnostics)).some(
      (item) => item.code === "ContentOverflow" && item.nodeId === "head",
    ),
  ).toBe(true);
  const bounds = await page
    .frameLocator("iframe")
    .locator("body")
    .evaluate(() => ({
      header: document.querySelector(".sdk-paper .sdk-page-story")!.getBoundingClientRect().bottom,
      body: document.querySelector('[data-node-id="body-text"]')!.getBoundingClientRect().top,
    }));
  expect(bounds.header).toBeLessThanOrEqual(bounds.body + 0.5);
});
test("copy keeps nested block boundaries and source attributes cannot impersonate the editor", async ({
  page,
}) => {
  const files = pagedFiles(
    '<section data-node-id="nested"><p data-node-id="first">FIRST</p><p data-node-id="second">SECOND</p></section><table data-node-id="grid"><tr data-node-id="r"><td data-node-id="c">C</td><td data-node-id="d">D</td></tr></table>',
  );
  files.set(
    "content.html",
    new TextEncoder().encode(
      new TextDecoder()
        .decode(files.get("content.html"))
        .replace('data-node-id="root"', 'data-node-id="root" data-office-editor="source-data"'),
    ),
  );
  await mount(page, files);
  const copied = await page
    .frameLocator("iframe")
    .locator('[data-node-id="root"]')
    .evaluate((node) => {
      const range = document.createRange();
      range.selectNodeContents(node);
      document.getSelection()!.removeAllRanges();
      document.getSelection()!.addRange(range);
      const values: Record<string, string> = {},
        event = new Event("copy", { bubbles: true, cancelable: true });
      Object.defineProperty(event, "clipboardData", {
        value: {
          setData: (type: string, value: string) => {
            values[type] = value;
          },
        },
      });
      document.dispatchEvent(event);
      return {
        text: values["text/plain"],
        html: values["text/html"],
        prevented: event.defaultPrevented,
        editable: (node as HTMLElement).isContentEditable,
      };
    });
  expect(copied.editable).toBe(false);
  expect(copied.prevented).toBe(true);
  expect(copied.text).toBe("FIRST\nSECOND\nC\tD");
  expect(copied.html).not.toContain("sdk-page");
});
for (const mode of ["preview", "edit"] as const)
  test(
    "complete field display text edges copy current cache with its editable instruction: " + mode,
    async ({ page }) => {
      const source = createDocumentSession(
        openWordDocument(createWordDocument({ documentId: "field-display-copy" }), {
          kind: "document",
        }),
      );
      source.commit(
        source.prepare({
          protocolVersion: 1,
          documentId: "field-display-copy",
          baseRevision: 0,
          operations: [
            {
              op: "insertWordField",
              nodeId: "paragraph-1",
              offset: 0,
              instruction: " PAGE ",
              cachedText: "PRIVATE-OLD-CACHE",
            },
          ],
        }),
      );
      await mount(page, source.serialize(), mode);
      const copied = await page
        .frameLocator("iframe")
        .locator(".sdk-field-display")
        .first()
        .evaluate((node) => {
          const walker = document.createTreeWalker(node, NodeFilter.SHOW_TEXT),
            text = walker.nextNode()!;
          const range = document.createRange();
          range.setStart(text, 0);
          range.setEnd(text, text.textContent!.length);
          const selection = document.getSelection()!;
          selection.removeAllRanges();
          selection.addRange(range);
          const data = new DataTransfer(),
            event = new ClipboardEvent("copy", { bubbles: true, cancelable: true });
          Object.defineProperty(event, "clipboardData", { value: data });
          node.dispatchEvent(event);
          return { html: data.getData("text/html"), text: data.getData("text/plain") };
        });
      expect(copied.text).toBe("1");
      expect(copied.html).not.toContain("PRIVATE-OLD-CACHE");
      const match = /data-doffi-word-clipboard="([^"]+)"/.exec(copied.html)!;
      expect(match).toBeTruthy();
      const format = JSON.parse(Buffer.from(match[1]!, "base64").toString("utf8"));
      expect(format.version).toBe(2);
      expect(format.semantics.fields).toHaveLength(1);
      expect(format.semantics.fields[0].instruction.trim()).toBe("PAGE");
      expect(source.getNode("paragraph-1")!.text).toBe("PRIVATE-OLD-CACHE");
    },
  );
for (const mode of ["preview", "edit"] as const)
  test("partial field display copy stays within the selected range: " + mode, async ({ page }) => {
    await mount(page, fieldsFiles(), mode);
    const copied = await page
      .frameLocator("iframe")
      .locator(".sdk-field-display")
      .filter({ hasText: "Alpha" })
      .first()
      .evaluate((node) => {
        const walker = document.createTreeWalker(node, NodeFilter.SHOW_TEXT),
          text = walker.nextNode()!;
        const range = document.createRange();
        range.setStart(text, 0);
        range.setEnd(text, 3);
        document.getSelection()!.removeAllRanges();
        document.getSelection()!.addRange(range);
        const values: Record<string, string> = {},
          event = new Event("copy", { bubbles: true, cancelable: true });
        Object.defineProperty(event, "clipboardData", {
          value: {
            setData: (type: string, value: string) => {
              values[type] = value;
            },
          },
        });
        node.dispatchEvent(event);
        return values;
      });
    expect(copied["text/plain"]).toBe("Alp");
    expect(copied["text/html"]).not.toContain("Beta");
    expect(copied["text/html"]).not.toMatch(/sdk-field|data-word-field|data-node-id/);
  });
test("disposing a preview during settled invalidates its pending report", async ({ page }) => {
  await mount(page, pagedFiles('<p data-node-id="a">TEXT</p>'));
  const result = await page.evaluate(async () => {
    const view = window.paginationHost.view as import("@doffi/editor").DocumentPreview;
    const pending = view.settled();
    await Promise.resolve();
    view.dispose();
    return pending.then(
      () => "success",
      (error: { code: string }) => error.code,
    );
  });
  expect(result).toBe("InvalidOperation");
});
test("odd-page section breaks create a blank sheet using the new geometry and restart printed numbering", async ({
  page,
}) => {
  const files = pagedFiles(
    '<section data-node-id="first"><p data-node-id="a">FIRST</p></section><section data-node-id="second"><p data-node-id="b">SECOND</p></section>',
  );
  const metadata = JSON.parse(new TextDecoder().decode(files.get("metadata.json")));
  metadata.sections = [
    { sectionId: "first", rootNodeId: "first" },
    {
      sectionId: "second",
      rootNodeId: "second",
      breakType: "oddPage",
      pageNumberStart: 8,
      pageNumberFormat: "lowerRoman",
      pageSetup: { ...metadata.pageSetup, width: 240, height: 300 },
    },
  ];
  files.set("metadata.json", new TextEncoder().encode(JSON.stringify(metadata)));
  await mount(page, files);
  const report = await page.evaluate(() => window.paginationHost.layout());
  expect(
    report.pages.map((item) => [item.blank, item.setup.width, item.setup.height, item.label]),
  ).toEqual([
    [false, 300, 240, "1"],
    [true, 240, 300, "2"],
    [false, 240, 300, "viii"],
  ]);
  expect(report.nodePages.a).toEqual(["page-1"]);
  expect(report.nodePages.b).toEqual(["page-3"]);
});
test("unitless inherited Word line spacing works without Typed OM and preserves absolute overrides", async ({
  page,
}) => {
  await page.addInitScript(() => {
    Object.defineProperty(HTMLElement.prototype, "computedStyleMap", {
      configurable: true,
      value: undefined,
    });
  });
  const files = pagedFiles(
    '<p data-node-id="auto" data-word-line-rule="auto">AUTO</p><p data-node-id="absolute" data-word-line-rule="auto" style="line-height:30px">ABSOLUTE</p>',
  );
  const metadata = JSON.parse(new TextDecoder().decode(files.get("metadata.json")));
  metadata.word = {
    stories: [{ storyId: "main", kind: "main", rootNodeId: "root" }],
    styles: [],
    numbering: [],
    hyperlinks: [],
    bookmarks: [],
    fields: [],
    tabs: [],
    tables: [],
    revisions: [],
  };
  files.set("metadata.json", new TextEncoder().encode(JSON.stringify(metadata)));
  files.set(
    "content.html",
    new TextEncoder().encode(
      new TextDecoder()
        .decode(files.get("content.html"))
        .replace(
          'data-office-kind="document"',
          'data-office-kind="document" data-office-story="main"',
        ),
    ),
  );
  files.set(
    "styles.css",
    new TextEncoder().encode(
      'article{font-family:"Liberation Serif",serif;font-size:12pt}p{line-height:1.5;margin:0}',
    ),
  );
  await mount(page, files);
  const measured = await page
    .frameLocator("iframe")
    .locator('[data-node-id="auto"]')
    .evaluate((node) => {
      const probe = document.createElement("span");
      probe.style.cssText =
        "position:absolute;display:block;line-height:normal;white-space:pre;font:inherit";
      probe.style.lineHeight = "normal";
      probe.textContent = "汉M";
      node.append(probe);
      const normal = probe.getBoundingClientRect().height;
      probe.remove();
      return { line: parseFloat(getComputedStyle(node).lineHeight), normal };
    });
  expect(measured.line).toBeCloseTo(measured.normal * 1.5, 1);
  await expect(page.frameLocator("iframe").locator('[data-node-id="absolute"]')).toHaveCSS(
    "line-height",
    "30px",
  );
});
for (const layout of ["pages", "continuous"] as const)
  test(
    "editing preserves source content before, between and after sections: " + layout,
    async ({ page }) => {
      const files = pagedFiles(
        '<p data-node-id="lead">L</p><section data-node-id="first"><p data-node-id="a">A</p></section><p data-node-id="middle">M</p><section data-node-id="second"><p data-node-id="b">B</p></section><p data-node-id="tail">T</p>',
      );
      const metadata = JSON.parse(new TextDecoder().decode(files.get("metadata.json")));
      metadata.sections = [
        { sectionId: "first", rootNodeId: "first" },
        { sectionId: "second", rootNodeId: "second", breakType: "nextPage" },
      ];
      files.set("metadata.json", new TextEncoder().encode(JSON.stringify(metadata)));
      await mount(page, files, "edit", { layout });
      await atStart(page, "a");
      await page.keyboard.insertText("X");
      const state = await page.evaluate(() => ({
        text: window.paginationHost.session!.getNode("root")!.text,
        html: window.paginationHost.session!.snapshot().html,
        errors: window.paginationHost.errors,
      }));
      expect(state.text).toBe("LXAMBT");
      expect(state.errors).toEqual([]);
      expect(state.html).not.toMatch(/sdk-page|edit-\d|<div/);
      const structure = await page.evaluate(
        () => window.paginationHost.session!.getNode("root")!.childNodeIds,
      );
      expect(structure).toEqual(["lead", "first", "middle", "second", "tail"]);
      await page.keyboard.press("Control+z");
      expect(await page.evaluate(() => window.paginationHost.session!.getNode("root")!.text)).toBe(
        "LAMBT",
      );
    },
  );
test("formatted table whitespace does not become extra plain-text rows or cells", async ({
  page,
}) => {
  await mount(
    page,
    pagedFiles(
      '<table data-node-id="table">\n <tbody>\n <tr data-node-id="row">\n  <td data-node-id="c">C</td>\n  <td data-node-id="d">D</td>\n </tr>\n </tbody>\n</table>',
    ),
  );
  const copied = await page
    .frameLocator("iframe")
    .locator('[data-node-id="table"]')
    .evaluate((node) => {
      const range = document.createRange();
      range.selectNode(node);
      document.getSelection()!.removeAllRanges();
      document.getSelection()!.addRange(range);
      const values: Record<string, string> = {},
        event = new Event("copy", { bubbles: true, cancelable: true });
      Object.defineProperty(event, "clipboardData", {
        value: {
          setData: (type: string, value: string) => {
            values[type] = value;
          },
        },
      });
      node.dispatchEvent(event);
      return values;
    });
  expect(copied["text/plain"]).toBe("C\tD");
});

test("content controls support UI choices, checkbox clicks, undo and surrounding native text input", async ({
  page,
}) => {
  const session = createDocumentSession(
    openWordDocument(createWordDocument({ documentId: "controls-input" }), { kind: "document" }),
  );
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "controls-input",
      baseRevision: 0,
      operations: [
        {
          op: "replaceText",
          nodeId: "paragraph-1",
          from: 0,
          to: 0,
          expectedText: "",
          text: "前后",
        },
      ],
    }),
  );
  await mount(page, session.serialize(), "edit");
  const frame = page.frameLocator("iframe"),
    paragraph = frame.locator('[data-node-id="paragraph-1"]');
  await paragraph.click();
  await page.keyboard.press("End");
  await page.evaluate(() =>
    (window.paginationHost.view as import("@doffi/editor").DocumentTextEditor).execute({
      type: "word",
      action: "insertControl",
    }),
  );
  await page.getByLabel("控件类型", { exact: true }).selectOption("checkBox");
  await page.getByLabel("控件标题", { exact: true }).fill("同意");
  await page.getByRole("button", { name: "插入内容控件", exact: true }).click();
  await expect(frame.getByRole("checkbox", { name: "同意", exact: true })).toHaveAttribute(
    "aria-checked",
    "false",
  );
  await page.getByRole("button", { name: "关闭内容控件设置", exact: true }).click();
  await frame.getByRole("checkbox", { name: "同意", exact: true }).click();
  await expect(frame.getByRole("checkbox", { name: "同意", exact: true })).toHaveAttribute(
    "aria-checked",
    "true",
  );
  await page.evaluate(() => window.paginationHost.session!.undo());
  await expect(frame.getByRole("checkbox", { name: "同意", exact: true })).toHaveAttribute(
    "aria-checked",
    "false",
  );
  await paragraph.click({ position: { x: 4, y: 5 } });
  await page.keyboard.press("Home");
  await page.keyboard.insertText("EDIT");
  expect(
    await page.evaluate(() => window.paginationHost.session!.getNode("paragraph-1")!.text),
  ).toBe("EDIT前后☐");
  const result = await page.evaluate(() => ({
    html: window.paginationHost.session!.snapshot().html,
    errors: window.paginationHost.errors,
  }));
  expect(result.html).not.toMatch(/aria-checked|contenteditable|data-word-control-lock/);
  expect(result.errors).toEqual([]);
});

test("Word chart UI edits combo data, secondary axes and titles through reversible shared rendering", async ({
  page,
}) => {
  const session = createDocumentSession(
    openWordDocument(createWordDocument({ documentId: "chart-input" }), { kind: "document" }),
  );
  await mount(page, session.serialize(), "edit");
  const frame = page.frameLocator("iframe");
  await frame.locator('[data-node-id="paragraph-1"]').click();
  await page.evaluate(() =>
    (window.paginationHost.view as import("@doffi/editor").DocumentTextEditor).execute({
      type: "word",
      action: "insertChart",
    }),
  );
  await page.getByLabel("图表类型", { exact: true }).selectOption("combo");
  await page.getByLabel("系列 2 类型", { exact: true }).selectOption("line");
  await page.getByLabel("系列 2 坐标轴", { exact: true }).selectOption("secondary");
  await page.getByLabel("系列 1 值 1", { exact: true }).fill("9");
  await page.getByLabel("图表标题", { exact: true }).fill("营收趋势");
  await page.getByRole("button", { name: "插入图表", exact: true }).click();
  await expect(frame.locator("[data-sdk-word-chart=combo]")).toHaveCount(1);
  await expect(frame.locator("[data-sdk-word-chart=combo]")).toContainText("营收趋势");
  expect(await frame.locator("[data-sdk-word-chart=combo] rect[data-series]").count()).toBe(3);
  expect(await frame.locator("[data-sdk-word-chart=combo] path[data-series]").count()).toBe(1);
  await page.getByLabel("系列 1 值 1", { exact: true }).fill("12");
  await page.getByRole("button", { name: "应用图表修改", exact: true }).click();
  expect(
    await page.evaluate(() => {
      const drawing = window.paginationHost
        .session!.snapshot()
        .metadata.word!.drawings!.find((drawing) => drawing.kind === "chart")!;
      return drawing.chart.series[0]!.values[0];
    }),
  ).toBe(12);
  await page.evaluate(() => window.paginationHost.session!.undo());
  expect(
    await page.evaluate(() => {
      const drawing = window.paginationHost
        .session!.snapshot()
        .metadata.word!.drawings!.find((drawing) => drawing.kind === "chart")!;
      return drawing.chart.series[0]!.values[0];
    }),
  ).toBe(9);
  const result = await page.evaluate(() => ({
    html: window.paginationHost.session!.snapshot().html,
    errors: window.paginationHost.errors,
  }));
  expect(result.html).not.toMatch(/<svg|data-sdk-word-chart|营收趋势/);
  expect(result.errors).toEqual([]);
});

test("chart font roles remain source data while missing preview faces are reported as substitutions", async ({
  page,
}) => {
  const session = createDocumentSession(
    openWordDocument(createWordDocument({ documentId: "chart-fonts" }), { kind: "document" }),
  );
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "chart-fonts",
      baseRevision: 0,
      operations: [
        {
          op: "insertWordDrawing",
          nodeId: "paragraph-1",
          offset: 0,
          geometry: {
            x: 0,
            y: 0,
            width: 360,
            height: 220,
            rotation: 0,
            flipH: false,
            flipV: false,
          },
          layout: { mode: "inline" },
          object: {
            kind: "chart",
            chart: {
              kind: "column",
              title: "中文 x 漢\u{e0100}",
              legend: "none",
              showLabels: true,
              categories: ["中文"],
              series: [{ seriesId: "s", name: "系列", values: [5] }],
              textStyle: {
                fonts: {
                  ascii: "Doffi Missing Latin 7321",
                  hAnsi: "Doffi Missing Latin 7321",
                  eastAsia: "Doffi Missing East Asia 7321",
                },
                fontSize: 11,
              },
            },
          },
        },
      ],
    }),
  );
  await mount(page, session.serialize(), "edit");
  const report = await page.evaluate(() => window.paginationHost.layout());
  expect(
    report.fonts.some(
      (font) => font.requestedFamily === "Doffi Missing East Asia 7321" && font.substituted,
    ),
  ).toBe(true);
  expect(
    report.fonts.some(
      (font) => font.requestedFamily === "Doffi Missing Latin 7321" && font.substituted,
    ),
  ).toBe(true);
  const snapshot = await page.evaluate(() => window.paginationHost.session!.snapshot()),
    drawing = snapshot.metadata.word!.drawings!.find((drawing) => drawing.kind === "chart")!;
  expect(drawing.chart.textStyle!.fonts!.eastAsia).toBe("Doffi Missing East Asia 7321");
  expect(snapshot.html).not.toContain("data-sdk-word-chart");
  await expect(page.frameLocator("iframe").locator("[data-sdk-word-chart]")).toContainText(
    "漢\u{e0100}",
  );
});

test("Word review UI tracks replacements, keeps the caret after retained deletions and rejects with shared undo", async ({
  page,
}) => {
  const session = createDocumentSession(
    openWordDocument(createWordDocument({ documentId: "review-input" }), { kind: "document" }),
  );
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "review-input",
      baseRevision: 0,
      operations: [
        {
          op: "replaceText",
          nodeId: "paragraph-1",
          from: 0,
          to: 0,
          expectedText: "",
          text: "前😀后",
        },
      ],
    }),
  );
  await mount(page, session.serialize(), "edit");
  await page.evaluate(() =>
    (window.paginationHost.view as import("@doffi/editor").DocumentTextEditor).execute({
      type: "word",
      action: "review",
    }),
  );
  await page.getByLabel("记录修订", { exact: true }).check();
  await page.getByLabel("修订作者", { exact: true }).fill("甲");
  await page.getByRole("button", { name: "应用修订设置", exact: true }).click();
  await page.getByRole("button", { name: "关闭修订审阅", exact: true }).click();
  const frame = page.frameLocator("iframe"),
    paragraph = frame.locator('[data-node-id="paragraph-1"]');
  await paragraph.click();
  await paragraph.evaluate((node) => {
    const walker = document.createTreeWalker(node, NodeFilter.SHOW_TEXT),
      range = document.createRange();
    let offset = 0,
      start = false;
    for (let text = walker.nextNode(); text; text = walker.nextNode()) {
      const end = offset + (text.textContent?.length ?? 0);
      if (!start && 1 >= offset && 1 < end) {
        range.setStart(text, 1 - offset);
        start = true;
      }
      if (start && 3 <= end) {
        range.setEnd(text, 3 - offset);
        break;
      }
      offset = end;
    }
    window.getSelection()!.removeAllRanges();
    window.getSelection()!.addRange(range);
    node.dispatchEvent(new Event("selectionchange", { bubbles: true }));
  });
  await page.keyboard.insertText("新");
  await expect(paragraph.locator("del")).toHaveText("😀");
  await expect(paragraph.locator("ins")).toHaveText("新");
  await page.keyboard.insertText("添");
  await expect(paragraph.locator("ins")).toHaveText("新添");
  expect(
    await page.evaluate(() => window.paginationHost.session!.getNode("paragraph-1")!.text),
  ).toBe("前😀新添后");
  await page.evaluate(() =>
    (window.paginationHost.view as import("@doffi/editor").DocumentTextEditor).execute({
      type: "word",
      action: "review",
    }),
  );
  await page.getByLabel("修订显示", { exact: true }).selectOption("final");
  await expect(paragraph.locator("del")).toBeHidden();
  await page.getByRole("button", { name: "拒绝全部显示的修订", exact: true }).click();
  expect(
    await page.evaluate(() => window.paginationHost.session!.getNode("paragraph-1")!.text),
  ).toBe("前😀后");
  await page.evaluate(() => window.paginationHost.session!.undo());
  expect(
    await page.evaluate(() => window.paginationHost.session!.getNode("paragraph-1")!.text),
  ).toBe("前😀新添后");
  expect(await page.evaluate(() => window.paginationHost.session!.snapshot().html)).not.toContain(
    "data-sdk-word-revision",
  );
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
});

test("tracked native Enter and Backspace preserve paragraph marks and fold a new break through the same history", async ({
  page,
}) => {
  const session = createDocumentSession(
    openWordDocument(createWordDocument({ documentId: "tracked-paragraph-input" }), {
      kind: "document",
    }),
  );
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "tracked-paragraph-input",
      baseRevision: 0,
      operations: [
        { op: "replaceText", nodeId: "paragraph-1", from: 0, to: 0, expectedText: "", text: "AB" },
        { op: "setWordReview", nodeId: "document-root", trackChanges: true, author: "甲" },
      ],
    }),
  );
  await mount(page, session.serialize(), "edit");
  const frame = page.frameLocator("iframe");
  await frame.locator('[data-node-id="paragraph-1"]').click();
  await page.keyboard.press("Home");
  await page.keyboard.press("ArrowRight");
  await page.keyboard.press("Enter");
  await expect(frame.locator("p")).toHaveCount(2);
  expect(
    await page.evaluate(() =>
      window.paginationHost
        .session!.snapshot()
        .metadata.word!.revisions.map((revision) => revision.kind),
    ),
  ).toEqual(["paragraphInsert"]);
  await page.keyboard.press("Backspace");
  await expect(frame.locator("p")).toHaveCount(1);
  await expect(frame.locator("p")).toHaveText("AB");
  expect(
    await page.evaluate(() => window.paginationHost.session!.snapshot().metadata.word!.revisions),
  ).toEqual([]);
  await page.evaluate(() => window.paginationHost.session!.undo());
  await expect(frame.locator("p")).toHaveCount(2);
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
});

test("native compatibility cache paints in the editor and surrounding typing leaves its opaque source intact", async ({
  page,
}) => {
  const { readFileSync } = await import("node:fs"),
    { importDocx, readDocumentZip } = await import("@doffi/office"),
    { writeOpcZip } = await import("../../packages/office/src/document-zip.js");
  const files = new Map(
      readDocumentZip(
        new Uint8Array(
          readFileSync(new URL("../../examples/docx/03-images.docx", import.meta.url)),
        ),
      ),
    ),
    decoder = new TextDecoder(),
    source = decoder.decode(files.get("word/document.xml")!),
    drawing = source.match(/<w:drawing[\s>][\s\S]*?<\/w:drawing>/)![0];
  files.set(
    "word/document.xml",
    new TextEncoder().encode(
      source.replace(
        drawing,
        '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:future"><mc:Choice Requires="x"><x:widget/></mc:Choice><mc:Fallback>' +
          drawing +
          "</mc:Fallback></mc:AlternateContent>",
      ),
    ),
  );
  await mount(page, importDocx(writeOpcZip(files), { documentId: "native-cache" }), "edit");
  const frame = page.frameLocator("iframe"),
    cache = frame.locator("[data-sdk-native-preview] img");
  await expect(cache).toHaveCount(1);
  expect(
    await cache.evaluate((image: HTMLImageElement) => image.complete && image.naturalWidth > 0),
  ).toBe(true);
  const paragraph = frame.locator("p").first();
  await paragraph.click();
  await page.keyboard.press("Home");
  await page.keyboard.type("缓存周围编辑");
  await expect
    .poll(() => page.evaluate(() => window.paginationHost.session!.snapshot().html))
    .toContain("缓存周围编辑");
  const html = await page.evaluate(() => window.paginationHost.session!.snapshot().html);
  expect(html).not.toContain("data-sdk-native-preview");
  expect(html).not.toContain("blob:");
  expect(await frame.locator("[data-sdk-native-preview] img").count()).toBe(1);
  await page.evaluate(() => {
    const session = window.paginationHost.session!,
      snapshot = session.snapshot(),
      section = snapshot.metadata.sections![0]!.rootNodeId;
    session.commit(
      session.prepare({
        protocolVersion: 1,
        documentId: snapshot.manifest.documentId,
        baseRevision: snapshot.manifest.revision,
        operations: [
          {
            op: "insertWordStory",
            nodeId: section,
            kind: "header",
            variant: "default",
            html: "<p>独立页眉</p>",
          },
        ],
      }),
    );
    const story = session
      .snapshot()
      .metadata.word!.stories.find((story) => story.kind === "header")!;
    (window.paginationHost.view as import("@doffi/editor").DocumentTextEditor).editStory(
      story.storyId,
    );
  });
  await expect(frame.locator("[data-sdk-native-preview] img")).toHaveCount(0);
  const header = frame.locator("p").filter({ hasText: "独立页眉" });
  await header.click();
  await page.keyboard.press("End");
  await page.keyboard.type("可编辑");
  await expect
    .poll(() => page.evaluate(() => window.paginationHost.session!.snapshot().html))
    .toContain("独立页眉可编辑");
  await page.evaluate(() =>
    (window.paginationHost.view as import("@doffi/editor").DocumentTextEditor).editStory(),
  );
  await expect(frame.locator("[data-sdk-native-preview] img")).toHaveCount(1);
});

test("style management edits named inheritance and all outline levels while Enter uses the following style", async ({
  page,
}) => {
  await mount(page, createWordDocument({ documentId: "style-ui" }), "edit");
  const frame = page.frameLocator("iframe"),
    paragraph = frame.locator('[data-node-id="paragraph-1"]');
  await paragraph.click();
  await page.keyboard.type("样式标题");
  await page.evaluate(() =>
    (window.paginationHost.view as import("@doffi/editor").DocumentTextEditor).execute({
      type: "word",
      action: "styleSettings",
    }),
  );
  const panel = page.getByRole("region", { name: "样式管理" });
  await panel.getByLabel("已有样式", { exact: true }).selectOption("Heading1");
  await panel.getByLabel("样式显示名称", { exact: true }).fill("自定义九级");
  await panel.getByLabel("大纲级别", { exact: true }).selectOption({ value: "8" });
  await panel.getByLabel("样式字号（磅）", { exact: true }).fill("18");
  await panel.getByRole("button", { name: "保存样式", exact: true }).click();
  await panel.getByRole("button", { name: "应用到选区", exact: true }).click();
  await expect(frame.locator('[data-node-id="paragraph-1"]')).toHaveAttribute(
    "data-word-style",
    "Heading1",
  );
  await expect
    .poll(() =>
      page.evaluate(
        () =>
          window.paginationHost
            .session!.snapshot()
            .metadata.word!.styles.find((style) => style.styleId === "Heading1")!.outlineLevel,
      ),
    )
    .toBe(8);
  await frame.locator('[data-node-id="paragraph-1"]').click();
  await page.keyboard.press("End");
  await page.keyboard.press("Enter");
  await page.keyboard.type("后续正文");
  await expect
    .poll(() => page.evaluate(() => window.paginationHost.session!.snapshot().html))
    .toContain("后续正文");
  await expect(frame.locator("p").filter({ hasText: "后续正文" })).toHaveAttribute(
    "data-word-style",
    "Normal",
  );
});

test("tracked table rows display both review states and reject with the shared grid history", async ({
  page,
}) => {
  const session = createDocumentSession(
    openWordDocument(createWordDocument({ documentId: "row-ui" }), { kind: "document" }),
  );
  const apply = (operations: import("@doffi/core").DocumentOperation[]) =>
    session.commit(
      session.prepare({
        protocolVersion: 1,
        documentId: "row-ui",
        baseRevision: session.snapshot().manifest.revision,
        operations,
      }),
    );
  apply([{ op: "insertWordTable", nodeId: "paragraph-1", position: "after", rows: 3, columns: 2 }]);
  const table = session.snapshot().metadata.word!.tables[0]!.nodeId;
  apply([
    {
      op: "editWordTable",
      nodeId: table,
      change: {
        kind: "setCells",
        row: 0,
        column: 0,
        values: [
          ["A", "a"],
          ["B", "b"],
          ["C", "c"],
        ],
      },
    },
    { op: "setWordReview", nodeId: "document-root", trackChanges: true, author: "甲" },
    { op: "editWordTable", nodeId: table, change: { kind: "deleteRows", at: 1, count: 1 } },
    { op: "editWordTable", nodeId: table, change: { kind: "insertRows", at: 1, count: 1 } },
  ]);
  await mount(page, session.serialize(), "edit");
  await page.evaluate(() =>
    (window.paginationHost.view as import("@doffi/editor").DocumentTextEditor).execute({
      type: "word",
      action: "review",
    }),
  );
  const frame = page.frameLocator("iframe");
  await expect(frame.locator("tr[data-node-id]:visible")).toHaveCount(4);
  await page.getByLabel("修订显示", { exact: true }).selectOption({ value: "final" });
  await expect(frame.locator("tr[data-node-id]:visible")).toHaveCount(3);
  await page.getByLabel("修订显示", { exact: true }).selectOption({ value: "original" });
  await expect(frame.locator("tr[data-node-id]:visible")).toHaveCount(3);
  await page.getByLabel("修订显示", { exact: true }).selectOption({ value: "all" });
  await page.getByRole("button", { name: "拒绝全部显示的修订", exact: true }).click();
  await expect(frame.locator("tr[data-node-id]:visible")).toHaveCount(3);
  await expect
    .poll(() =>
      page.evaluate(
        () => window.paginationHost.session!.snapshot().metadata.word!.revisions.length,
      ),
    )
    .toBe(0);
  await page.evaluate(() => window.paginationHost.session!.undo());
  await expect(frame.locator("tr[data-node-id]:visible")).toHaveCount(4);
  const html = await page.evaluate(() => window.paginationHost.session!.snapshot().html);
  expect(html).not.toContain("data-sdk-word-deleted");
  expect(html).not.toContain("data-sdk-word-inserted");
});

test("page margin rulers drag and use keyboard increments through section transactions and undo", async ({
  page,
}) => {
  await mount(page, createWordDocument({ documentId: "margin-rulers" }), "edit");
  await page.frameLocator("iframe").locator('[data-node-id="paragraph-1"]').click();
  await page.evaluate(() =>
    (window.paginationHost.view as import("@doffi/editor").DocumentTextEditor).setRulers(true),
  );
  const left = page.getByRole("slider", { name: "页面左边距", exact: true });
  await expect(left).toBeVisible();
  await left.focus();
  await left.press("ArrowRight");
  await expect
    .poll(() =>
      page.evaluate(
        () =>
          window.paginationHost.session!.snapshot().metadata.sections![0]!.pageSetup!.margins.left,
      ),
    )
    .toBe(73);
  const box = (await left.boundingBox())!;
  await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
  await page.mouse.down();
  await page.mouse.move(box.x + box.width / 2 + 8, box.y + box.height / 2);
  await page.mouse.up();
  await expect
    .poll(() =>
      page.evaluate(
        () =>
          window.paginationHost.session!.snapshot().metadata.sections![0]!.pageSetup!.margins.left,
      ),
    )
    .toBe(79);
  await page.evaluate(() => window.paginationHost.session!.undo());
  await expect
    .poll(() =>
      page.evaluate(
        () =>
          window.paginationHost.session!.snapshot().metadata.sections![0]!.pageSetup!.margins.left,
      ),
    )
    .toBe(73);
  const top = page.getByRole("slider", { name: "页面上边距", exact: true });
  await top.focus();
  await top.press("ArrowDown");
  await expect
    .poll(() =>
      page.evaluate(
        () =>
          window.paginationHost.session!.snapshot().metadata.sections![0]!.pageSetup!.margins.top,
      ),
    )
    .toBe(73);
  expect(await page.evaluate(() => window.paginationHost.session!.snapshot().html)).not.toContain(
    "data-ruler-margin",
  );
});

test("mirrored physical pages swap text, header and gutter margins independently of printed page labels", async ({
  page,
}) => {
  const session = createDocumentSession(
    openWordDocument(
      createWordDocument({
        documentId: "mirrors",
        pageSetup: {
          unit: "pt",
          width: 400,
          height: 250,
          margins: { left: 40, right: 80, top: 25, bottom: 25 },
        },
      }),
      { kind: "document" },
    ),
  );
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "mirrors",
      baseRevision: 0,
      operations: [
        {
          op: "replaceText",
          nodeId: "paragraph-1",
          from: 0,
          to: 0,
          expectedText: "",
          text: "第一页",
        },
        {
          op: "insertFragment",
          nodeId: "paragraph-1",
          position: "after",
          html: '<p style="break-before:page">第二页</p>',
        },
        {
          op: "setWordSection",
          nodeId: "section-root-1",
          changes: { gutter: 10, gutterPosition: "left", pageNumberStart: 4, headerDistance: 5 },
        },
        {
          op: "insertWordStory",
          nodeId: "section-root-1",
          kind: "header",
          variant: "default",
          html: "<p>页眉</p>",
        },
        { op: "setWordOptions", nodeId: "document-root", mirrorMargins: true },
      ],
    }),
  );
  await mount(page, session.serialize(), "edit");
  const frame = page.frameLocator("iframe");
  await expect.poll(() => page.evaluate(() => window.paginationHost.layout().pages.length)).toBe(2);
  const pages = await page.evaluate(() => window.paginationHost.layout().pages);
  expect(pages.map((page) => page.number)).toEqual([4, 5]);
  expect(pages.map((page) => [page.setup.margins.left, page.setup.margins.right])).toEqual([
    [50, 80],
    [80, 50],
  ]);
  const offsets = await frame.locator("body").evaluate((body) => {
    const nodes = [...body.querySelectorAll<HTMLElement>('[data-office-story="main"] p')].filter(
        (node) => node.textContent === "第一页" || node.textContent === "第二页",
      ),
      papers = [...body.querySelectorAll<HTMLElement>(".sdk-paper")];
    return nodes.map(
      (node, index) =>
        (node.getBoundingClientRect().left - papers[index]!.getBoundingClientRect().left) * 0.75,
    );
  });
  expect(offsets[0]).toBeCloseTo(50, 0);
  expect(offsets[1]).toBeCloseTo(80, 0);
  const headers = await frame
    .locator(".sdk-paper .sdk-page-story")
    .evaluateAll((nodes) => nodes.map((node) => (node as HTMLElement).style.left));
  expect(headers).toEqual(["50pt", "80pt"]);
  const second = frame.locator("p").filter({ hasText: "第二页" }).first();
  await second.click();
  await page.keyboard.press("End");
  await page.keyboard.type("输入");
  await expect
    .poll(() => page.evaluate(() => window.paginationHost.session!.snapshot().html))
    .toContain("第二页输入");
  expect(await page.evaluate(() => window.paginationHost.session!.snapshot().html)).not.toContain(
    "sdk-column",
  );
  await page.evaluate(() =>
    (window.paginationHost.view as import("@doffi/editor").DocumentTextEditor).setRulers(true),
  );
  const left = page.getByRole("slider", { name: "页面左边距", exact: true });
  await expect(left).toHaveAttribute("aria-valuenow", "80");
  await left.focus();
  await left.press("ArrowRight");
  await expect
    .poll(() =>
      page.evaluate(
        () =>
          window.paginationHost.session!.snapshot().metadata.sections![0]!.pageSetup!.margins.right,
      ),
    )
    .toBe(81);
  expect(
    await page.evaluate(
      () =>
        window.paginationHost.session!.snapshot().metadata.sections![0]!.pageSetup!.margins.left,
    ),
  ).toBe(40);
});

test("font controls apply script roles, complex size and language in one tracked undoable source change", async ({
  page,
}) => {
  const session = createDocumentSession(
    openWordDocument(createWordDocument({ documentId: "font-ui" }), { kind: "document" }),
  );
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "font-ui",
      baseRevision: 0,
      operations: [
        {
          op: "replaceText",
          nodeId: "paragraph-1",
          from: 0,
          to: 0,
          expectedText: "",
          text: "Latin 漢 العربية",
        },
        { op: "setWordReview", nodeId: "document-root", trackChanges: true, author: "甲" },
      ],
    }),
  );
  await mount(page, session.serialize(), "edit");
  const frame = page.frameLocator("iframe"),
    paragraph = frame.locator('[data-node-id="paragraph-1"]');
  await paragraph.click();
  await paragraph.evaluate((node) => {
    const range = node.ownerDocument.createRange();
    range.selectNodeContents(node);
    const selection = node.ownerDocument.getSelection()!;
    selection.removeAllRanges();
    selection.addRange(range);
  });
  await page.evaluate(() =>
    (window.paginationHost.view as import("@doffi/editor").DocumentTextEditor).execute({
      type: "word",
      action: "fontSettings",
    }),
  );
  const panel = page.getByRole("region", { name: "字符格式", exact: true });
  await panel.getByLabel("ASCII 字体", { exact: true }).fill("Liberation Serif");
  await panel.getByLabel("东亚字体", { exact: true }).fill("Noto Serif CJK SC");
  await panel.getByLabel("复杂文字字体", { exact: true }).fill("Liberation Sans");
  await panel.getByLabel("复杂文字字号（磅）", { exact: true }).fill("16");
  await panel.getByLabel("复杂文字字重", { exact: true }).selectOption({ value: "1" });
  await panel.getByLabel("西文语言", { exact: true }).fill("en-GB");
  await panel.getByLabel("东亚语言", { exact: true }).fill("ja-JP");
  await panel.getByLabel("复杂文字语言", { exact: true }).fill("ar-SA");
  await panel.getByRole("button", { name: "应用字符格式", exact: true }).click();
  await expect
    .poll(() =>
      page.evaluate(
        () => window.paginationHost.session!.snapshot().metadata.word!.revisions.length,
      ),
    )
    .toBe(1);
  await expect(
    frame
      .locator('[data-sdk-word-script-view][data-word-script="eastAsia"]')
      .filter({ hasText: "漢" }),
  ).toHaveAttribute("lang", "ja-JP");
  const complex = frame
    .locator('[data-sdk-word-script-view][data-word-script="cs"]')
    .filter({ hasText: "العربية" });
  await expect(complex).toHaveAttribute("lang", "ar-SA");
  expect(
    await complex.evaluate(
      (node) => parseFloat(node.ownerDocument.defaultView!.getComputedStyle(node).fontSize) * 0.75,
    ),
  ).toBeCloseTo(16, 1);
  const html = await page.evaluate(() => window.paginationHost.session!.snapshot().html);
  expect(html).not.toContain("data-sdk-word-script-view");
  expect(html).not.toContain('lang="ja-JP"');
  expect(html).toContain("--word-language-eastAsia");
  await page.evaluate(() => window.paginationHost.session!.undo());
  await expect
    .poll(() =>
      page.evaluate(
        () => window.paginationHost.session!.snapshot().metadata.word!.revisions.length,
      ),
    )
    .toBe(0);
  expect(await page.evaluate(() => window.paginationHost.session!.snapshot().html)).not.toContain(
    "--word-language-eastAsia",
  );
});

test("native cell revision views retain siblings, review columns and keep undo in the source session", async ({
  page,
}) => {
  const session = createDocumentSession(
      openWordDocument(createWordDocument({ documentId: "cell-review-ui" }), { kind: "document" }),
    ),
    apply = (operations: import("@doffi/core").DocumentOperation[]) =>
      session.commit(
        session.prepare({
          protocolVersion: 1,
          documentId: "cell-review-ui",
          baseRevision: session.snapshot().manifest.revision,
          operations,
        }),
      );
  apply([
    {
      op: "insertWordTable",
      nodeId: "paragraph-1",
      position: "after",
      rows: 2,
      columns: 3,
      width: 216,
    },
  ]);
  const table = session.snapshot().metadata.word!.tables[0]!.nodeId;
  apply([
    {
      op: "editWordTable",
      nodeId: table,
      change: {
        kind: "setCells",
        row: 0,
        column: 0,
        values: [
          ["A", "B", "C"],
          ["D", "E", "F"],
        ],
      },
    },
    { op: "setWordReview", nodeId: "document-root", trackChanges: true, author: "A" },
    { op: "editWordTable", nodeId: table, change: { kind: "deleteColumns", at: 1, count: 1 } },
  ]);
  await mount(page, session.serialize(), "edit");
  const frame = page.frameLocator("iframe"),
    tableView = frame.locator('[data-node-id="' + table + '"]');
  await expect(tableView.locator("[data-sdk-word-deleted]")).toHaveCount(2);
  await page.evaluate(() =>
    (window.paginationHost.view as import("@doffi/editor").DocumentTextEditor).execute({
      type: "word",
      action: "review",
    }),
  );
  await page.getByLabel("修订显示", { exact: true }).selectOption("final");
  await expect(tableView.locator("[data-sdk-word-deleted]").first()).toBeHidden();
  expect(
    await tableView.evaluate((node) => (node as HTMLElement).innerText.replace(/\s/g, "")),
  ).toBe("ACDF");
  await page.getByLabel("修订显示", { exact: true }).selectOption("original");
  await expect(tableView.locator("[data-sdk-word-deleted]").first()).toBeVisible();
  expect(
    await tableView.evaluate((node) => (node as HTMLElement).innerText.replace(/\s/g, "")),
  ).toBe("ABCDEF");
  await page.getByLabel("修订显示", { exact: true }).selectOption("final");
  await page.getByRole("button", { name: "接受全部显示的修订", exact: true }).click();
  await expect(tableView.locator("tr").first().locator("td")).toHaveCount(2);
  expect(
    await page.evaluate(
      () => window.paginationHost.session!.snapshot().metadata.word!.tables[0]!.grid,
    ),
  ).toEqual([108, 108]);
  await page.evaluate(() => window.paginationHost.session!.undo());
  await expect(tableView.locator("tr").first().locator("td")).toHaveCount(3);
  expect(await page.evaluate(() => window.paginationHost.session!.snapshot().html)).not.toContain(
    "data-sdk-word",
  );
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
});

test("original column insertion view restores unequal widths and merged spans while input preserves current source geometry", async ({
  page,
}) => {
  const session = createDocumentSession(
      openWordDocument(createWordDocument({ documentId: "column-original-ui" }), {
        kind: "document",
      }),
    ),
    apply = (operations: import("@doffi/core").DocumentOperation[]) =>
      session.commit(
        session.prepare({
          protocolVersion: 1,
          documentId: "column-original-ui",
          baseRevision: session.snapshot().manifest.revision,
          operations,
        }),
      );
  apply([
    {
      op: "insertWordTable",
      nodeId: "paragraph-1",
      position: "after",
      rows: 2,
      columns: 3,
      width: 216,
    },
  ]);
  const table = session.snapshot().metadata.word!.tables[0]!.nodeId;
  apply([
    {
      op: "editWordTable",
      nodeId: table,
      change: {
        kind: "setCells",
        row: 0,
        column: 0,
        values: [
          ["A", "B", "C"],
          ["D", "E", "F"],
        ],
      },
    },
    { op: "editWordTable", nodeId: table, change: { kind: "resize", grid: [40, 60, 116] } },
    {
      op: "editWordTable",
      nodeId: table,
      change: { kind: "mergeCells", row: 0, column: 0, rows: 1, columns: 2 },
    },
  ]);
  await mount(page, session.serialize(), "edit");
  const geometry = async () =>
    page
      .frameLocator("iframe")
      .locator('[data-node-id="' + table + '"] tr')
      .evaluateAll((rows) =>
        rows.flatMap((row) =>
          [...(row as HTMLTableRowElement).cells]
            .filter((cell) => getComputedStyle(cell).display !== "none")
            .map((cell) => ({ width: cell.getBoundingClientRect().width, span: cell.colSpan })),
        ),
      );
  const original = await geometry();
  apply([
    { op: "setWordReview", nodeId: "document-root", trackChanges: true, author: "A" },
    { op: "editWordTable", nodeId: table, change: { kind: "insertColumns", at: 1, count: 1 } },
  ]);
  const currentGrid = session.snapshot().metadata.word!.tables[0]!.grid;
  await mount(page, session.serialize(), "edit");
  await page.evaluate(() =>
    (window.paginationHost.view as import("@doffi/editor").DocumentTextEditor).execute({
      type: "word",
      action: "review",
    }),
  );
  await page.getByLabel("修订显示", { exact: true }).selectOption("original");
  const projected = await geometry();
  expect(projected.map((cell) => cell.span)).toEqual(original.map((cell) => cell.span));
  projected.forEach((cell, index) => expect(cell.width).toBeCloseTo(original[index]!.width, 0));
  const frame = page.frameLocator("iframe"),
    last = frame
      .locator('[data-node-id="' + table + '"] tr')
      .last()
      .locator("td")
      .last();
  await last.locator("p").click();
  await page.keyboard.press("End");
  await page.keyboard.insertText("Z");
  await expect
    .poll(() => page.evaluate((id) => window.paginationHost.session!.getNode(id)!.text, table))
    .toContain("FZ");
  expect(
    await page.evaluate(
      () => window.paginationHost.session!.snapshot().metadata.word!.tables[0]!.grid,
    ),
  ).toEqual(currentGrid);
  expect(
    await page.evaluate((id) => window.paginationHost.session!.getNode(id)!.html, table),
  ).toContain('colspan="3"');
  await page.evaluate(() => window.paginationHost.session!.undo());
  await page.getByRole("button", { name: "拒绝全部显示的修订", exact: true }).click();
  expect(
    await page.evaluate(
      () => window.paginationHost.session!.snapshot().metadata.word!.tables[0]!.grid,
    ),
  ).toEqual([40, 60, 116]);
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
});

test("table panel creation and deletion enter native table revision review and retain undo", async ({
  page,
}) => {
  const session = createDocumentSession(
    openWordDocument(createWordDocument({ documentId: "whole-table-ui" }), { kind: "document" }),
  );
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "whole-table-ui",
      baseRevision: 0,
      operations: [
        { op: "setWordReview", nodeId: "document-root", trackChanges: true, author: "A" },
      ],
    }),
  );
  await mount(page, session.serialize(), "edit");
  const frame = page.frameLocator("iframe");
  await frame.locator('[data-node-id="paragraph-1"]').click();
  await page.evaluate(() =>
    (window.paginationHost.view as import("@doffi/editor").DocumentTextEditor).execute({
      type: "word",
      action: "insertTable",
    }),
  );
  await page.getByLabel("表格行数", { exact: true }).fill("2");
  await page.getByLabel("表格列数", { exact: true }).fill("2");
  await page.getByRole("button", { name: "插入表格", exact: true }).click();
  expect(
    await page.evaluate(() =>
      window.paginationHost.session!.snapshot().metadata.word!.revisions.map((item) => item.kind),
    ),
  ).toEqual(["rowInsert", "rowInsert"]);
  await page.evaluate(() =>
    (window.paginationHost.view as import("@doffi/editor").DocumentTextEditor).execute({
      type: "word",
      action: "review",
    }),
  );
  await page.getByRole("button", { name: "接受全部显示的修订", exact: true }).click();
  const table = await page.evaluate(
    () => window.paginationHost.session!.snapshot().metadata.word!.tables[0]!.nodeId,
  );
  await frame
    .locator('[data-node-id="' + table + '"] td p')
    .first()
    .click();
  await page.evaluate(() =>
    (window.paginationHost.view as import("@doffi/editor").DocumentTextEditor).execute({
      type: "word",
      action: "tableSettings",
    }),
  );
  await page.getByRole("button", { name: "删除表格", exact: true }).click();
  expect(
    await page.evaluate(() =>
      window.paginationHost.session!.snapshot().metadata.word!.revisions.map((item) => item.kind),
    ),
  ).toEqual(["rowDelete", "rowDelete"]);
  expect(
    await page.evaluate(() => window.paginationHost.session!.snapshot().metadata.word!.tables),
  ).toHaveLength(1);
  await page.getByRole("button", { name: "拒绝全部显示的修订", exact: true }).click();
  expect(
    await page.evaluate(() => window.paginationHost.session!.snapshot().metadata.word!.tables),
  ).toHaveLength(1);
  await page.evaluate(() => window.paginationHost.session!.undo());
  expect(
    await page.evaluate(() => window.paginationHost.session!.snapshot().metadata.word!.revisions),
  ).toHaveLength(2);
  await page.getByRole("button", { name: "接受全部显示的修订", exact: true }).click();
  expect(
    await page.evaluate(() => window.paginationHost.session!.snapshot().metadata.word!.tables),
  ).toEqual([]);
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
});

test("tracked native run views preserve superscript geometry and restore authored color highlight and underline in clean views", async ({
  page,
}) => {
  const run = (value: string, size: number) =>
      '<w:r><w:rPr><w:color w:val="AA0000"/><w:sz w:val="' +
      size +
      '"/><w:highlight w:val="yellow"/><w:u w:val="single"/><w:vertAlign w:val="superscript"/></w:rPr><w:t>' +
      value +
      "</w:t></w:r>",
    files = importDocx(
      mutateDocx("01-text-numbering.docx", {
        "word/styles.xml": () =>
          nativeStyles('<w:style w:type="paragraph" w:styleId="Normal" w:default="1"/>'),
        "word/document.xml": (source) =>
          nativeBody(
            source,
            "<w:p>" +
              run("ABC", 64) +
              "</w:p><w:p>" +
              run("A", 64) +
              run("X", 24) +
              run("C", 64) +
              "</w:p>",
          ),
      }),
      { documentId: "tracked-runs-ui", unsupportedContent: "reject" },
    ),
    session = createDocumentSession(openWordDocument(files, { kind: "document" })),
    story = session.snapshot().metadata.word!.stories[0]!,
    nodes = [story.rootNodeId];
  for (let index = 0; index < nodes.length; index++)
    nodes.push(...session.getNode(nodes[index]!)!.childNodeIds);
  const paragraph = nodes.find(
      (id) => session.getNode(id)!.tagName === "p" && session.getNode(id)!.text === "ABC",
    )!,
    reference = nodes.find(
      (id) => session.getNode(id)!.tagName === "p" && session.getNode(id)!.text === "AXC",
    )!,
    root = session.getNode(story.rootNodeId)!.parentNodeId!;
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "tracked-runs-ui",
      baseRevision: 0,
      operations: [
        { op: "setWordReview", nodeId: root, trackChanges: true, author: "A" },
        {
          op: "trackWordText",
          nodeId: paragraph,
          from: 1,
          to: 2,
          expectedText: "B",
          text: "X",
          properties: { "font-size": "12pt" },
        },
      ],
    }),
  );
  await mount(page, session.serialize(), "edit");
  const frame = page.frameLocator("iframe"),
    current = frame.locator('[data-node-id="' + paragraph + '"]');
  await expect(current.locator('[data-sdk-word-text-review="insert"]')).toHaveCSS(
    "color",
    "rgb(0, 100, 0)",
  );
  await expect(current.locator('[data-sdk-word-text-review="delete"]')).toHaveCSS(
    "text-decoration-line",
    /underline.*line-through|line-through.*underline/,
  );
  await page.evaluate(() =>
    (window.paginationHost.view as import("@doffi/editor").DocumentTextEditor).execute({
      type: "word",
      action: "review",
    }),
  );
  await page.getByLabel("修订显示", { exact: true }).selectOption("final");
  await expect(current.locator("del")).toBeHidden();
  const inserted = current.locator('[data-sdk-word-text-review="insert"]');
  await expect(inserted).toHaveCSS("color", "rgb(170, 0, 0)");
  await expect(inserted).toHaveCSS("background-color", "rgb(255, 255, 0)");
  await expect(inserted).toHaveCSS("text-decoration-line", "underline");
  const bounds = await frame.locator("[data-office-editor]").evaluate(
    (_, ids) =>
      ids.map((id) => {
        const paragraph = document.querySelector('[data-node-id="' + id + '"]')!,
          walker = document.createTreeWalker(paragraph, NodeFilter.SHOW_TEXT),
          glyphs: DOMRect[] = [];
        for (let text = walker.nextNode(); text; text = walker.nextNode()) {
          if (!(text.parentElement as HTMLElement).checkVisibility()) continue;
          const range = document.createRange();
          range.selectNodeContents(text);
          glyphs.push(range.getBoundingClientRect());
        }
        const origin = paragraph.getBoundingClientRect();
        return {
          height: origin.height,
          glyphs: glyphs.map((rect) => ({ y: rect.y - origin.y, height: rect.height })),
        };
      }),
    [paragraph, reference],
  );
  expect(bounds[0]!.height).toBeCloseTo(bounds[1]!.height, 0);
  expect(bounds[0]!.glyphs).toHaveLength(3);
  bounds[0]!.glyphs.forEach((glyph, index) => {
    expect(glyph.y).toBeCloseTo(bounds[1]!.glyphs[index]!.y, 0);
    expect(glyph.height).toBeCloseTo(bounds[1]!.glyphs[index]!.height, 0);
  });
  await page.getByLabel("修订显示", { exact: true }).selectOption("original");
  const deleted = current.locator('[data-sdk-word-text-review="delete"]');
  await expect(deleted).toHaveCSS("color", "rgb(170, 0, 0)");
  await expect(deleted).toHaveCSS("background-color", "rgb(255, 255, 0)");
  await expect(deleted).toHaveCSS("vertical-align", "super");
  await page.getByRole("button", { name: "拒绝全部显示的修订", exact: true }).click();
  expect(
    await page.evaluate((id) => window.paginationHost.session!.getNode(id)!.text, paragraph),
  ).toBe("ABC");
  expect(await page.evaluate(() => window.paginationHost.session!.snapshot().html)).not.toMatch(
    /sdk-word-authored|sdk-word-text-review|sdk-word-text-run/,
  );
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
});

test("paragraph mark font controls preserve body text, expose original history and seed empty paragraph typing from source fonts", async ({
  page,
}) => {
  const files = importDocx(
      mutateDocx("01-text-numbering.docx", {
        "word/styles.xml": () =>
          nativeStyles('<w:style w:type="paragraph" w:styleId="Normal" w:default="1"/>'),
        "word/document.xml": (source) =>
          nativeBody(
            source,
            '<w:p><w:pPr><w:rPr><w:sz w:val="28"/><w:color w:val="880000"/><w:rPrChange w:id="7" w:author="A"><w:rPr><w:sz w:val="20"/><w:color w:val="000088"/></w:rPr></w:rPrChange></w:rPr></w:pPr><w:r><w:t>BODY</w:t></w:r></w:p><w:p><w:pPr><w:rPr><w:sz w:val="24"/><w:color w:val="000088"/></w:rPr></w:pPr></w:p>',
          ),
      }),
      { documentId: "mark-font-ui", unsupportedContent: "reject" },
    ),
    session = createDocumentSession(openWordDocument(files, { kind: "document" })),
    nodes = [session.snapshot().metadata.word!.stories[0]!.rootNodeId];
  for (let index = 0; index < nodes.length; index++)
    nodes.push(...session.getNode(nodes[index]!)!.childNodeIds);
  const paragraphs = nodes.filter((id) => session.getNode(id)!.tagName === "p"),
    root = session.getNode(nodes[0]!)!.parentNodeId!;
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "mark-font-ui",
      baseRevision: 0,
      operations: [{ op: "setWordReview", nodeId: root, trackChanges: true, author: "B" }],
    }),
  );
  await mount(page, session.serialize(), "edit");
  const frame = page.frameLocator("iframe"),
    first = frame.locator('[data-node-id="' + paragraphs[0] + '"]'),
    empty = frame.locator('[data-node-id="' + paragraphs[1] + '"]');
  await page.evaluate(() =>
    (window.paginationHost.view as import("@doffi/editor").DocumentTextEditor).execute({
      type: "word",
      action: "showMarks",
      visible: true,
    }),
  );
  await first.click();
  await page.evaluate(() =>
    (window.paginationHost.view as import("@doffi/editor").DocumentTextEditor).execute({
      type: "word",
      action: "fontSettings",
    }),
  );
  await page.getByLabel("仅设置段落标记", { exact: true }).check();
  await page.getByLabel("字符字号（磅）", { exact: true }).fill("18");
  await page.getByRole("button", { name: "应用字符格式", exact: true }).click();
  expect(
    await first.evaluate((node) => parseFloat(getComputedStyle(node, "::after").fontSize) * 0.75),
  ).toBeCloseTo(18, 1);
  expect(
    await first.evaluate((node) => parseFloat(getComputedStyle(node).fontSize) * 0.75),
  ).toBeCloseTo(12, 1);
  await page.getByRole("button", { name: "关闭字符格式", exact: true }).click();
  await page.evaluate(() =>
    (window.paginationHost.view as import("@doffi/editor").DocumentTextEditor).execute({
      type: "word",
      action: "review",
    }),
  );
  await page.getByLabel("修订显示", { exact: true }).selectOption("original");
  expect(
    await first.evaluate((node) => parseFloat(getComputedStyle(node, "::after").fontSize) * 0.75),
  ).toBeCloseTo(10, 1);
  await page.getByLabel("修订显示", { exact: true }).selectOption("all");
  await page.getByRole("button", { name: "关闭修订审阅", exact: true }).click();
  await empty.click();
  await page.evaluate(() =>
    (window.paginationHost.view as import("@doffi/editor").DocumentTextEditor).execute({
      type: "word",
      action: "fontSettings",
    }),
  );
  await page.getByLabel("仅设置段落标记", { exact: true }).check();
  await page.getByLabel("字符字号（磅）", { exact: true }).fill("24");
  await page.getByLabel("字符颜色", { exact: true }).fill("#880000");
  await page.getByRole("button", { name: "应用字符格式", exact: true }).click();
  await page.getByRole("button", { name: "关闭字符格式", exact: true }).click();
  expect(
    await empty.evaluate((node) => parseFloat(getComputedStyle(node).fontSize) * 0.75),
  ).toBeCloseTo(24, 1);
  await empty.click();
  await page.keyboard.insertText("新");
  await expect
    .poll(() =>
      page.evaluate((id) => window.paginationHost.session!.getNode(id!)!.text, paragraphs[1]),
    )
    .toBe("新");
  const inserted = empty.locator("ins .sdk-word-text-run").last();
  await expect(inserted).toHaveCSS("font-size", "32px");
  await page.evaluate(() => window.paginationHost.session!.undo());
  expect(
    await empty.evaluate((node) => parseFloat(getComputedStyle(node).fontSize) * 0.75),
  ).toBeCloseTo(24, 1);
  await page.evaluate(() => window.paginationHost.session!.undo());
  expect(
    await empty.evaluate((node) => parseFloat(getComputedStyle(node).fontSize) * 0.75),
  ).toBeCloseTo(12, 1);
  expect(await page.evaluate(() => window.paginationHost.session!.snapshot().html)).not.toContain(
    "--sdk-word-mark",
  );
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
});

test("table merge controls explain the Word tracking behavior and keep the operation undoable without extra revisions", async ({
  page,
}) => {
  const session = createDocumentSession(
    openWordDocument(createWordDocument({ documentId: "merge-policy-ui" }), { kind: "document" }),
  );
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "merge-policy-ui",
      baseRevision: 0,
      operations: [
        { op: "insertWordTable", nodeId: "paragraph-1", position: "after", rows: 2, columns: 2 },
        { op: "setWordReview", nodeId: "document-root", trackChanges: true, author: "A" },
      ],
    }),
  );
  const table = session.snapshot().metadata.word!.tables[0]!.nodeId;
  await mount(page, session.serialize(), "edit");
  await page
    .frameLocator("iframe")
    .locator('[data-node-id="' + table + '"] td p')
    .first()
    .click();
  await page.evaluate(() =>
    (window.paginationHost.view as import("@doffi/editor").DocumentTextEditor).execute({
      type: "word",
      action: "tableSettings",
    }),
  );
  await expect(page.getByRole("region", { name: "表格设置" })).toContainText("不会记录为修订");
  await page.getByLabel("选区行数", { exact: true }).fill("1");
  await page.getByLabel("选区列数", { exact: true }).fill("2");
  await page.getByRole("button", { name: "合并单元格", exact: true }).click();
  expect(
    await page.evaluate((id) => window.paginationHost.session!.getNode(id)!.html, table),
  ).toContain('colspan="2"');
  expect(
    await page.evaluate(() => window.paginationHost.session!.snapshot().metadata.word!.revisions),
  ).toEqual([]);
  await page.evaluate(() => window.paginationHost.session!.undo());
  expect(
    await page.evaluate((id) => window.paginationHost.session!.getNode(id)!.html, table),
  ).not.toContain("colspan");
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
});

test("East Asian spacing switches share exact source text across runs, line ends, input and undo", async ({
  page,
}) => {
  const session = createDocumentSession(
      openWordDocument(createWordDocument({ documentId: "typography-ui" }), { kind: "document" }),
    ),
    text = "汉A汉1汉Ａ汉１汉 A汉。A汉𠀀\u{E0100}B";
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "typography-ui",
      baseRevision: 0,
      operations: [
        {
          op: "replaceFragment",
          nodeId: "section-root-1",
          html:
            [
              ["both", "1", "1"],
              ["letters", "1", "0"],
              ["numbers", "0", "1"],
              ["off", "0", "0"],
            ]
              .map(
                ([name, de, dn]) =>
                  '<p class="space-' +
                  name +
                  '" style="font-size:16px;line-height:24px;margin:0;--word-auto-space-de:' +
                  de +
                  ";--word-auto-space-dn:" +
                  dn +
                  '"><span>汉</span><strong>A</strong>' +
                  text.slice(2) +
                  "</p>",
              )
              .join("") +
            '<p class="space-wrap" style="width:16px;font-size:16px;line-height:24px;margin:0">汉A汉A</p>',
        },
      ],
    }),
  );
  await mount(page, session.serialize(), "edit");
  const frame = page.frameLocator("iframe");
  for (const [name, count] of [
    ["both", 7],
    ["letters", 5],
    ["numbers", 2],
    ["off", 0],
  ] as const) {
    const paragraph = frame.locator(".space-" + name);
    await expect(paragraph.locator("[data-sdk-word-autospace]")).toHaveCount(count);
    const widths = await paragraph
      .locator("[data-sdk-word-autospace]")
      .evaluateAll((nodes) => nodes.map((node) => node.getBoundingClientRect().width));
    for (const width of widths) expect(width).toBeCloseTo(2, 1);
  }
  // A generated gap collapses at a wrapped line edge rather than indenting A.
  const wrap = frame.locator(".space-wrap");
  const bounds = await wrap.evaluate((root) => {
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT),
      left = root.getBoundingClientRect().left,
      letters: number[] = [];
    for (let text = walker.nextNode(); text; text = walker.nextNode()) {
      if (text.parentElement?.closest("[data-sdk-word-autospace]")) continue;
      for (let at = 0; at < (text.textContent?.length ?? 0); at++)
        if (text.textContent![at] === "A") {
          const range = document.createRange();
          range.setStart(text, at);
          range.setEnd(text, at + 1);
          letters.push(range.getBoundingClientRect().left - left);
        }
    }
    return letters;
  });
  for (const left of bounds) expect(left).toBeLessThan(0.5);
  const both = frame.locator(".space-both"),
    id = await both.getAttribute("data-node-id");
  await both.evaluate((root) => {
    (root.closest("[data-office-editor]") as HTMLElement).focus();
    const walker = document.createTreeWalker(root.querySelector("strong")!, NodeFilter.SHOW_TEXT);
    let node: Node | null;
    do {
      node = walker.nextNode();
    } while (node && node.textContent !== "A");
    if (!node) throw new Error("Missing A");
    const range = document.createRange();
    range.selectNodeContents(node);
    range.collapse(false);
    const selection = document.getSelection()!;
    selection.removeAllRanges();
    selection.addRange(range);
  });
  await page.keyboard.insertText("中");
  await expect
    .poll(() => page.evaluate((id) => window.paginationHost.session!.getNode(id!)!.text, id))
    .toBe("汉A中" + text.slice(2));
  expect(await page.evaluate(() => window.paginationHost.session!.snapshot().html)).not.toContain(
    "sdk-word-autospace",
  );
  expect(await page.evaluate(() => window.paginationHost.session!.snapshot().html)).not.toContain(
    "word-spacing",
  );
  await page.evaluate(() => window.paginationHost.session!.undo());
  expect(await page.evaluate((id) => window.paginationHost.session!.getNode(id!)!.text, id)).toBe(
    text,
  );
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
});

test("paragraph word wrapping respects the native Boolean and retains Chinese line-end restrictions", async ({
  page,
}) => {
  const session = createDocumentSession(
    openWordDocument(createWordDocument({ documentId: "word-wrap-ui" }), { kind: "document" }),
  );
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "word-wrap-ui",
      baseRevision: 0,
      operations: [
        {
          op: "replaceFragment",
          nodeId: "section-root-1",
          html: '<p class="whole-word" style="font-size:16px;font-family:monospace;width:80px;line-height:24px;margin:0;--word-word-wrap:1">ABC ABCDE</p><p class="whole-loose" style="font-size:16px;font-family:monospace;width:80px;line-height:24px;margin:0;--word-word-wrap:1;--word-kinsoku:0">ABC ABCDE</p><p class="break-word" style="font-size:16px;font-family:monospace;width:80px;line-height:24px;margin:0;--word-word-wrap:0">ABC ABCDE</p><p class="kinsoku-on" style="font-size:16px;font-family:sans-serif;width:32px;line-height:24px;margin:0;--word-kinsoku:1">汉字，中文</p><p class="kinsoku-off" style="font-size:16px;font-family:sans-serif;width:32px;line-height:24px;margin:0;--word-kinsoku:0">汉字，中文</p>',
        },
      ],
    }),
  );
  await mount(page, session.serialize());
  const frame = page.frameLocator("iframe"),
    lines = async (selector: string) =>
      frame.locator(selector).evaluate((root) => {
        const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT),
          result: number[] = [];
        for (let text = walker.nextNode(); text; text = walker.nextNode()) {
          if (text.parentElement?.closest("[data-sdk-word-autospace]")) continue;
          for (let at = 0; at < (text.textContent?.length ?? 0); at++) {
            const range = document.createRange();
            range.setStart(text, at);
            range.setEnd(text, at + 1);
            result.push(range.getBoundingClientRect().top);
          }
        }
        return result;
      });
  const whole = await lines(".whole-word"),
    broken = await lines(".break-word");
  expect(whole[4]).toBeGreaterThan(whole[0]! + 10);
  const loose = await lines(".whole-loose");
  expect(loose[4]).toBeGreaterThan(loose[0]! + 10);
  expect(broken[4]).toBeCloseTo(broken[0]!, 1);
  const on = await lines(".kinsoku-on"),
    off = await lines(".kinsoku-off");
  expect(on[2]).toBeCloseTo(on[1]!, 1);
  expect(off[2]).toBeGreaterThan(off[1]! + 10);
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
});

test("automatic spacing follows visible revision text and clean views retain exact input offsets", async ({
  page,
}) => {
  const session = createDocumentSession(
    openWordDocument(createWordDocument({ documentId: "spacing-review" }), { kind: "document" }),
  );
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "spacing-review",
      baseRevision: 0,
      operations: [
        {
          op: "replaceText",
          nodeId: "paragraph-1",
          from: 0,
          to: 0,
          expectedText: "",
          text: "A汉B",
        },
        { op: "setWordReview", nodeId: "document-root", trackChanges: true, author: "A" },
        {
          op: "trackWordText",
          nodeId: "paragraph-1",
          from: 1,
          to: 2,
          expectedText: "汉",
          text: "",
        },
      ],
    }),
  );
  await mount(page, session.serialize(), "edit");
  const paragraph = page.frameLocator("iframe").locator('[data-node-id="paragraph-1"]');
  await expect(paragraph.locator("[data-sdk-word-autospace]")).toHaveCount(2);
  await page.evaluate(() =>
    (window.paginationHost.view as import("@doffi/editor").DocumentTextEditor).execute({
      type: "word",
      action: "review",
    }),
  );
  await page.getByLabel("修订显示", { exact: true }).selectOption("final");
  await page.evaluate(() => window.paginationHost.view!.settled());
  await expect(paragraph.locator("[data-sdk-word-autospace]")).toHaveCount(0);
  await page.getByLabel("修订显示", { exact: true }).selectOption("original");
  await page.evaluate(() => window.paginationHost.view!.settled());
  await expect(paragraph.locator("[data-sdk-word-autospace]")).toHaveCount(2);
  expect(
    await page.evaluate(() => window.paginationHost.session!.getNode("paragraph-1")!.text),
  ).toBe("A汉B");
  expect(await page.evaluate(() => window.paginationHost.session!.snapshot().html)).not.toMatch(
    /sdk-word-autospace|text-autospace|word-spacing/,
  );
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
});

test("Enter inherits the current paragraph mark while retaining its original format history on the tail", async ({
  page,
}) => {
  const files = importDocx(
      mutateDocx("01-text-numbering.docx", {
        "word/styles.xml": () =>
          nativeStyles('<w:style w:type="paragraph" w:styleId="Normal" w:default="1"/>'),
        "word/document.xml": (source) =>
          nativeBody(
            source,
            '<w:p><w:pPr><w:rPr><w:color w:val="880000"/><w:sz w:val="48"/><w:rPrChange w:id="7" w:author="A"><w:rPr><w:sz w:val="20"/></w:rPr></w:rPrChange></w:rPr></w:pPr><w:r><w:t>BODY</w:t></w:r></w:p>',
          ),
      }),
      { documentId: "split-mark-ui", unsupportedContent: "reject" },
    ),
    session = createDocumentSession(openWordDocument(files, { kind: "document" })),
    story = session.snapshot().metadata.word!.stories[0]!,
    nodes = [story.rootNodeId];
  for (let i = 0; i < nodes.length; i++) nodes.push(...session.getNode(nodes[i]!)!.childNodeIds);
  const paragraph = nodes.find((id) => session.getNode(id)!.tagName === "p")!,
    section = session.getNode(paragraph)!.parentNodeId!,
    root = session.getNode(story.rootNodeId)!.parentNodeId!;
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "split-mark-ui",
      baseRevision: 0,
      operations: [{ op: "setWordReview", nodeId: root, trackChanges: true, author: "B" }],
    }),
  );
  await mount(page, session.serialize(), "edit");
  const first = page.frameLocator("iframe").locator('[data-node-id="' + paragraph + '"]');
  await first.evaluate((root) => {
    (root.closest("[data-office-editor]") as HTMLElement).focus();
    const range = document.createRange();
    range.setStart(root, 0);
    range.collapse(true);
    const selection = document.getSelection()!;
    selection.removeAllRanges();
    selection.addRange(range);
  });
  await page.keyboard.press("Enter");
  await expect
    .poll(() =>
      page.evaluate(
        (id) => window.paginationHost.session!.getNode(id)!.childNodeIds.length,
        section,
      ),
    )
    .toBe(2);
  await expect(first).toHaveCSS("font-size", "32px");
  await first.click();
  await page.keyboard.insertText("新");
  await expect
    .poll(() => page.evaluate((id) => window.paginationHost.session!.getNode(id)!.text, paragraph))
    .toBe("新");
  await expect(first.locator("ins .sdk-word-text-run")).toHaveCSS("font-size", "32px");
  expect(
    await page.evaluate(
      () =>
        window.paginationHost
          .session!.snapshot()
          .metadata.word!.revisions.filter((item) => item.property === "run").length,
    ),
  ).toBe(1);
  await page.evaluate(() => window.paginationHost.session!.undo());
  await page.evaluate(() => window.paginationHost.session!.undo());
  expect(
    await page.evaluate(
      (id) => window.paginationHost.session!.getNode(id)!.childNodeIds.length,
      section,
    ),
  ).toBe(1);
  expect(
    await page.evaluate((id) => window.paginationHost.session!.getNode(id)!.text, paragraph),
  ).toBe("BODY");
  expect(
    await page.evaluate(
      () =>
        window.paginationHost
          .session!.snapshot()
          .metadata.word!.revisions.filter((item) => item.property === "run").length,
    ),
  ).toBe(1);
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
});

test("tracked typing uses the actual caret run at a format boundary and keeps continuing insertion together", async ({
  page,
}) => {
  const files = importDocx(
      mutateDocx("01-text-numbering.docx", {
        "word/styles.xml": () =>
          nativeStyles('<w:style w:type="paragraph" w:styleId="Normal" w:default="1"/>'),
        "word/document.xml": (source) =>
          nativeBody(
            source,
            '<w:p><w:r><w:rPr><w:b/><w:color w:val="880000"/></w:rPr><w:t>A</w:t></w:r><w:r><w:rPr><w:i/><w:color w:val="000088"/></w:rPr><w:t>B</w:t></w:r></w:p>',
          ),
      }),
      { documentId: "context-ui", unsupportedContent: "reject" },
    ),
    session = createDocumentSession(openWordDocument(files, { kind: "document" })),
    story = session.snapshot().metadata.word!.stories[0]!,
    nodes = [story.rootNodeId];
  for (let i = 0; i < nodes.length; i++) nodes.push(...session.getNode(nodes[i]!)!.childNodeIds);
  const paragraph = nodes.find((id) => session.getNode(id)!.tagName === "p")!,
    root = session.getNode(story.rootNodeId)!.parentNodeId!;
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "context-ui",
      baseRevision: 0,
      operations: [{ op: "setWordReview", nodeId: root, trackChanges: true, author: "A" }],
    }),
  );
  await mount(page, session.serialize(), "edit");
  const p = page.frameLocator("iframe").locator('[data-node-id="' + paragraph + '"]');
  const caret = async (token: string, end: boolean) =>
    p.evaluate(
      (root, { token, end }) => {
        (root.closest("[data-office-editor]") as HTMLElement).focus();
        const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
        for (let node = walker.nextNode(); node; node = walker.nextNode()) {
          if (node.textContent !== token) continue;
          const range = document.createRange();
          range.setStart(node, end ? token.length : 0);
          range.collapse(true);
          const selection = document.getSelection()!;
          selection.removeAllRanges();
          selection.addRange(range);
          return;
        }
        throw new Error("Missing caret token " + token);
      },
      { token, end },
    );
  await caret("A", true);
  await page.keyboard.insertText("X");
  await expect
    .poll(() => page.evaluate((id) => window.paginationHost.session!.getNode(id)!.text, paragraph))
    .toBe("AXB");
  await expect(p.locator("ins .sdk-word-text-run")).toHaveCSS("font-weight", "700");
  await caret("X", true);
  await page.keyboard.insertText("Z");
  await expect
    .poll(() => page.evaluate((id) => window.paginationHost.session!.getNode(id)!.text, paragraph))
    .toBe("AXZB");
  expect(
    await page.evaluate(
      () =>
        window.paginationHost
          .session!.snapshot()
          .metadata.word!.revisions.filter((item) => item.kind === "insert").length,
    ),
  ).toBe(1);
  await caret("B", false);
  await page.keyboard.insertText("Y");
  await expect
    .poll(() => page.evaluate((id) => window.paginationHost.session!.getNode(id)!.text, paragraph))
    .toBe("AXZYB");
  const last = p.locator("ins .sdk-word-text-run").last();
  await expect(last).toHaveCSS("font-style", "italic");
  await expect(last).toHaveCSS("font-weight", "400");
  expect(
    await page.evaluate(
      () =>
        window.paginationHost
          .session!.snapshot()
          .metadata.word!.revisions.filter((item) => item.kind === "insert").length,
    ),
  ).toBe(2);
  await page.evaluate(() => window.paginationHost.session!.undo());
  await page.evaluate(() => window.paginationHost.session!.undo());
  await page.evaluate(() => window.paginationHost.session!.undo());
  expect(
    await page.evaluate((id) => window.paginationHost.session!.getNode(id)!.text, paragraph),
  ).toBe("AB");
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
});

test("clipboard HTML copies authorized image bytes and pastes mixed text and images as one undo step across documents", async ({
  page,
}) => {
  const encoded = Buffer.from(
      readFileSync(new URL("../../examples/fidelity/assets/sample.png", import.meta.url)),
    ).toString("base64"),
    files = createWordDocument({ documentId: "clipboard-image" });
  await mount(page, files, "edit");
  const p = page.frameLocator("iframe").locator('[data-node-id="paragraph-1"]');
  await p.click();
  const html =
    '<strong>前</strong><img src="data:image/png;base64,' +
    encoded +
    '" style="width:12pt;height:8pt" alt="粘贴图片"><em>后</em><img src="data:image/png;base64,' +
    encoded +
    '" style="width:8pt;height:8pt" alt="第二张图片">';
  const paste = async (html: string) =>
    page
      .frameLocator("iframe")
      .locator("[data-office-editor]")
      .evaluate((root, html) => {
        const transfer = new DataTransfer();
        transfer.setData("text/html", html);
        const encoded = /data:image\/png;base64,([^\"]+)/.exec(html)?.[1];
        if (encoded)
          transfer.items.add(
            new File(
              [Uint8Array.from(atob(encoded), (character) => character.charCodeAt(0))],
              "clipboard.png",
              { type: "image/png" },
            ),
          );
        const event = new ClipboardEvent("paste", {
          bubbles: true,
          cancelable: true,
          clipboardData: transfer,
        });
        if (event.clipboardData !== transfer)
          Object.defineProperty(event, "clipboardData", { value: transfer });
        root.dispatchEvent(event);
      }, html);
  await paste(html);
  await expect
    .poll(() => page.evaluate(() => window.paginationHost.session!.getNode("paragraph-1")!.text))
    .toBe("前\ufffc后\ufffc");
  expect(
    await page.evaluate(() => window.paginationHost.session!.snapshot().manifest.assets.length),
  ).toBe(1);
  expect(await page.evaluate(() => window.paginationHost.session!.snapshot().html)).not.toContain(
    "data:image",
  );
  const copied = await p.evaluate((root) => {
    (root.closest("[data-office-editor]") as HTMLElement).focus();
    const range = document.createRange();
    range.selectNodeContents(root);
    const selection = document.getSelection()!;
    selection.removeAllRanges();
    selection.addRange(range);
    const transfer = new DataTransfer();
    const event = new ClipboardEvent("copy", {
      bubbles: true,
      cancelable: true,
      clipboardData: transfer,
    });
    if (event.clipboardData !== transfer)
      Object.defineProperty(event, "clipboardData", { value: transfer });
    root.dispatchEvent(event);
    return { html: transfer.getData("text/html"), text: transfer.getData("text/plain") };
  });
  expect(copied.html).toContain("data:image/png;base64," + encoded);
  expect(copied.html).not.toContain("data-node-id");
  expect(copied.text).toBe("前\ufffc后\ufffc");
  const entries = await page.evaluate(() =>
    [...window.paginationHost.session!.serialize()].map(
      ([path, bytes]) => [path, [...bytes]] as [string, number[]],
    ),
  );
  let native = exportDocx(
    openWordDocument(new Map(entries.map(([path, bytes]) => [path, new Uint8Array(bytes)])), {
      kind: "document",
    }),
  ).bytes;
  for (let i = 0; i < 2; i++) {
    const reopened = openWordDocument(
      importDocx(native, { documentId: "clipboard-native", unsupportedContent: "reject" }),
      { kind: "document" },
    );
    expect(reopened.snapshot().html).toContain("粘贴图片");
    expect(
      reopened.snapshot().metadata.word!.drawings!.filter((item) => item.kind === "image"),
    ).toHaveLength(2);
    native = exportDocx(reopened).bytes;
  }
  await page.evaluate(() => window.paginationHost.session!.undo());
  expect(
    await page.evaluate(() => window.paginationHost.session!.getNode("paragraph-1")!.text),
  ).toBe("");
  expect(
    await page.evaluate(() => window.paginationHost.session!.snapshot().manifest.assets.length),
  ).toBe(0);
  await mount(page, createWordDocument({ documentId: "clipboard-other" }), "edit");
  await p.click();
  await paste(copied.html);
  await expect
    .poll(() => page.evaluate(() => window.paginationHost.session!.getNode("paragraph-1")!.text))
    .toBe("前\ufffc后\ufffc");
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
});

test("malformed clipboard image bytes fail before changing text, resources or revision", async ({
  page,
}) => {
  await mount(page, createWordDocument({ documentId: "clipboard-invalid" }), "edit");
  await page.frameLocator("iframe").locator('[data-node-id="paragraph-1"]').click();
  const before = await page.evaluate(() => window.paginationHost.session!.snapshot());
  await page
    .frameLocator("iframe")
    .locator("[data-office-editor]")
    .evaluate((root) => {
      const transfer = new DataTransfer();
      transfer.setData("text/html", '前<img src="data:image/png;base64,YWJj">后');
      const event = new ClipboardEvent("paste", {
        bubbles: true,
        cancelable: true,
        clipboardData: transfer,
      });
      if (event.clipboardData !== transfer)
        Object.defineProperty(event, "clipboardData", { value: transfer });
      root.dispatchEvent(event);
    });
  expect(await page.evaluate(() => window.paginationHost.session!.snapshot())).toEqual(before);
  expect((await page.evaluate(() => window.paginationHost.errors)).length).toBe(1);
});

test("clipboard source formatting isolates conflicting named styles and keeps later style edits and undo linked", async ({
  page,
  browserName,
}) => {
  const files = importDocx(
    mutateDocx("01-text-numbering.docx", {
      "word/styles.xml": () =>
        nativeStyles(
          '<w:style w:type="paragraph" w:styleId="Normal" w:default="1"/><w:style w:type="paragraph" w:styleId="SourceParagraph"><w:name w:val="来源段落"/><w:basedOn w:val="Normal"/><w:rPr><w:i/><w:sz w:val="28"/></w:rPr></w:style><w:style w:type="character" w:styleId="Clash"><w:name w:val="同名强调"/><w:rPr><w:b/><w:color w:val="880000"/><w:rFonts w:ascii="SourceTypeface" w:hAnsi="SourceTypeface"/></w:rPr></w:style>',
        ),
      "word/document.xml": (source) =>
        nativeBody(
          source,
          '<w:p><w:pPr><w:pStyle w:val="SourceParagraph"/><w:rPr><w:sz w:val="48"/></w:rPr></w:pPr><w:r><w:rPr><w:rStyle w:val="Clash"/></w:rPr><w:t>SOURCE</w:t></w:r></w:p><w:p><w:r><w:t>UNSELECTED-PRIVATE-TEXT</w:t></w:r></w:p>',
        ),
    }),
    { documentId: "style-copy", unsupportedContent: "reject" },
  );
  await mount(page, files, "edit");
  let copied = await page
    .frameLocator("iframe")
    .locator("[data-office-editor] p")
    .first()
    .evaluate((root) => {
      (root.closest("[data-office-editor]") as HTMLElement).focus();
      const range = document.createRange();
      range.selectNodeContents(root);
      const selection = document.getSelection()!;
      selection.removeAllRanges();
      selection.addRange(range);
      const transfer = new DataTransfer(),
        event = new ClipboardEvent("copy", {
          bubbles: true,
          cancelable: true,
          clipboardData: transfer,
        });
      if (event.clipboardData !== transfer)
        Object.defineProperty(event, "clipboardData", { value: transfer });
      root.dispatchEvent(event);
      return transfer.getData("text/html");
    });
  expect(copied).toContain("data-doffi-word-clipboard");
  expect(copied).toContain("SourceTypeface");
  expect(copied).not.toContain("data-word-paragraph-mark");
  expect(copied).not.toContain("UNSELECTED-PRIVATE-TEXT");
  if (browserName === "chromium") {
    await page.context().grantPermissions(["clipboard-read", "clipboard-write"]);
    await page.keyboard.press("ControlOrMeta+C");
    copied = await page.evaluate(async () => {
      const item = (await navigator.clipboard.read()).find((item) =>
        item.types.includes("text/html"),
      )!;
      return (await item.getType("text/html")).text();
    });
    expect(copied).toContain("data-doffi-word-clipboard");
  }
  const target = createDocumentSession(
    openWordDocument(createWordDocument({ documentId: "style-paste" }), { kind: "document" }),
  );
  target.commit(
    target.prepare({
      protocolVersion: 1,
      documentId: "style-paste",
      baseRevision: 0,
      operations: [
        {
          op: "replaceFragment",
          nodeId: "section-root-1",
          html: '<p class="existing-style">DEST</p><p class="paste-style"></p>',
        },
        {
          op: "setWordStyle",
          nodeId: "document-root",
          style: {
            styleId: "Clash",
            name: "同名强调",
            kind: "character",
            cssClass: "target-clash",
          },
          runProperties: { color: "#000088" },
        },
      ],
    }),
  );
  const first = target.getNode("section-root-1")!.childNodeIds[0]!;
  target.commit(
    target.prepare({
      protocolVersion: 1,
      documentId: "style-paste",
      baseRevision: 1,
      operations: [
        {
          op: "applyWordCharacterStyle",
          nodeId: first,
          from: 0,
          to: 4,
          expectedText: "DEST",
          styleId: "Clash",
        },
      ],
    }),
  );
  const before = target.snapshot(),
    destination = target.getNode("section-root-1")!.childNodeIds[1]!;
  await mount(page, target.serialize(), "edit");
  const frame = page.frameLocator("iframe"),
    p = frame.locator(".paste-style");
  await p.click();
  await p.evaluate((root, html) => {
    const transfer = new DataTransfer();
    transfer.setData("text/html", html);
    const event = new ClipboardEvent("paste", {
      bubbles: true,
      cancelable: true,
      clipboardData: transfer,
    });
    if (event.clipboardData !== transfer)
      Object.defineProperty(event, "clipboardData", { value: transfer });
    root.dispatchEvent(event);
  }, copied);
  await expect
    .poll(() =>
      page.evaluate((id) => window.paginationHost.session!.getNode(id)!.text, destination),
    )
    .toBe("SOURCE");
  await expect(frame.locator(".existing-style .sdk-word-text-run")).toHaveCSS(
    "color",
    "rgb(0, 0, 136)",
  );
  await expect(p.getByText("SOURCE", { exact: true })).toHaveCSS("color", "rgb(136, 0, 0)");
  await expect(p.getByText("SOURCE", { exact: true })).toHaveCSS("font-weight", "700");
  await expect(p.getByText("SOURCE", { exact: true })).toHaveCSS("font-style", "italic");
  expect(await page.evaluate(() => window.paginationHost.session!.snapshot().html)).not.toMatch(
    /data-doffi-word-clipboard|data-doffi-clipboard-node/,
  );
  await p.evaluate((root) => {
    (root.closest("[data-office-editor]") as HTMLElement).focus();
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
    for (let node = walker.nextNode(); node; node = walker.nextNode()) {
      if (node.textContent !== "SOURCE") continue;
      const range = document.createRange();
      range.setStart(node, 6);
      range.collapse(true);
      const selection = document.getSelection()!;
      selection.removeAllRanges();
      selection.addRange(range);
      return;
    }
    throw new Error("Missing pasted text");
  });
  await page.keyboard.insertText("!");
  await expect
    .poll(() =>
      page.evaluate((id) => window.paginationHost.session!.getNode(id)!.text, destination),
    )
    .toBe("SOURCE!");
  await page.evaluate(() => window.paginationHost.session!.undo());
  const copiedStyle = await page.evaluate(
    () =>
      window.paginationHost
        .session!.snapshot()
        .metadata.word!.styles.find(
          (style) => style.name === "同名强调" && style.styleId !== "Clash",
        )!.styleId,
  );
  await page.evaluate((id) => {
    const session = window.paginationHost.session!,
      snapshot = session.snapshot(),
      style = snapshot.metadata.word!.styles.find((style) => style.styleId === id)!;
    session.commit(
      session.prepare({
        protocolVersion: 1,
        documentId: snapshot.manifest.documentId,
        baseRevision: snapshot.manifest.revision,
        operations: [
          {
            op: "setWordStyle",
            nodeId: "document-root",
            style: JSON.parse(JSON.stringify(style)),
            runProperties: { color: "#008800" },
          },
        ],
      }),
    );
  }, copiedStyle);
  await expect(p.getByText("SOURCE", { exact: true })).toHaveCSS("color", "rgb(0, 136, 0)");
  await expect(frame.locator(".existing-style .sdk-word-text-run")).toHaveCSS(
    "color",
    "rgb(0, 0, 136)",
  );
  const entries = await page.evaluate(() =>
    [...window.paginationHost.session!.serialize()].map(
      ([path, bytes]) => [path, [...bytes]] as [string, number[]],
    ),
  );
  let native = exportDocx(
    openWordDocument(new Map(entries.map(([path, bytes]) => [path, new Uint8Array(bytes)])), {
      kind: "document",
    }),
  ).bytes;
  for (let i = 0; i < 2; i++) {
    const reopened = openWordDocument(
      importDocx(native, { documentId: "styled-clipboard-native", unsupportedContent: "reject" }),
      { kind: "document" },
    );
    expect(
      reopened.snapshot().metadata.word!.styles.filter((style) => style.name === "同名强调"),
    ).toHaveLength(2);
    native = exportDocx(reopened).bytes;
  }
  await page.evaluate(() => window.paginationHost.session!.undo());
  await page.evaluate(() => window.paginationHost.session!.undo());
  expect(
    await page.evaluate(() => window.paginationHost.session!.snapshot().metadata.word!.styles),
  ).toEqual(before.metadata.word!.styles);
  expect(
    await page.evaluate((id) => window.paginationHost.session!.getNode(id)!.text, destination),
  ).toBe("");
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
});

test("read-only copy preserves named formatting and same-document paste reuses identical style definitions", async ({
  page,
}) => {
  const session = createDocumentSession(
    openWordDocument(createWordDocument({ documentId: "clipboard-reuse" }), { kind: "document" }),
  );
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "clipboard-reuse",
      baseRevision: 0,
      operations: [
        {
          op: "replaceText",
          nodeId: "paragraph-1",
          from: 0,
          to: 0,
          expectedText: "",
          text: "REUSE",
        },
        {
          op: "setWordStyle",
          nodeId: "document-root",
          style: {
            styleId: "emphasis",
            name: "复用样式",
            kind: "character",
            cssClass: "reuse-emphasis",
          },
          runProperties: { color: "#880000", "font-weight": "bold" },
        },
        {
          op: "applyWordCharacterStyle",
          nodeId: "paragraph-1",
          from: 0,
          to: 5,
          expectedText: "REUSE",
          styleId: "emphasis",
        },
      ],
    }),
  );
  const before = session.snapshot();
  await mount(page, session.serialize(), "preview");
  const copied = await page
    .frameLocator("iframe")
    .locator('[data-node-id="paragraph-1"]')
    .evaluate((root) => {
      const range = document.createRange();
      range.selectNodeContents(root);
      const selection = document.getSelection()!;
      selection.removeAllRanges();
      selection.addRange(range);
      const transfer = new DataTransfer(),
        event = new ClipboardEvent("copy", {
          bubbles: true,
          cancelable: true,
          clipboardData: transfer,
        });
      if (event.clipboardData !== transfer)
        Object.defineProperty(event, "clipboardData", { value: transfer });
      root.dispatchEvent(event);
      return transfer.getData("text/html");
    });
  expect(copied).toContain("data-doffi-word-clipboard");
  await mount(page, session.serialize(), "edit");
  const p = page.frameLocator("iframe").locator('[data-node-id="paragraph-1"]');
  await p.evaluate((root, html) => {
    (root.closest("[data-office-editor]") as HTMLElement).focus();
    const range = document.createRange();
    range.selectNodeContents(root);
    range.collapse(false);
    const selection = document.getSelection()!;
    selection.removeAllRanges();
    selection.addRange(range);
    const transfer = new DataTransfer(),
      event = new ClipboardEvent("paste", {
        bubbles: true,
        cancelable: true,
        clipboardData: transfer,
      });
    transfer.setData("text/html", html);
    if (event.clipboardData !== transfer)
      Object.defineProperty(event, "clipboardData", { value: transfer });
    root.dispatchEvent(event);
  }, copied);
  await expect
    .poll(() => page.evaluate(() => window.paginationHost.session!.getNode("paragraph-1")!.text))
    .toBe("REUSEREUSE");
  expect(
    await page.evaluate(() => window.paginationHost.session!.snapshot().metadata.word!.styles),
  ).toEqual(before.metadata.word!.styles);
  await page.evaluate(() => {
    const session = window.paginationHost.session!,
      snapshot = session.snapshot(),
      style = snapshot.metadata.word!.styles.find((style) => style.styleId === "emphasis")!;
    session.commit(
      session.prepare({
        protocolVersion: 1,
        documentId: snapshot.manifest.documentId,
        baseRevision: snapshot.manifest.revision,
        operations: [
          {
            op: "setWordStyle",
            nodeId: "document-root",
            style: JSON.parse(JSON.stringify(style)),
            runProperties: { color: "#008800" },
          },
        ],
      }),
    );
  });
  const texts = p.getByText("REUSE", { exact: true });
  await expect(texts).toHaveCount(2);
  for (const text of await texts.all()) await expect(text).toHaveCSS("color", "rgb(0, 136, 0)");
  await page.evaluate(() => window.paginationHost.session!.undo());
  await page.evaluate(() => window.paginationHost.session!.undo());
  expect(
    await page.evaluate(() => window.paginationHost.session!.getNode("paragraph-1")!.text),
  ).toBe("REUSE");
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
});

test("invalid or cyclic clipboard format packages leave source styles, text and revision unchanged", async ({
  page,
}) => {
  await mount(page, createWordDocument({ documentId: "clipboard-format-invalid" }), "edit");
  const p = page.frameLocator("iframe").locator('[data-node-id="paragraph-1"]');
  await p.click();
  const before = await page.evaluate(() => window.paginationHost.session!.snapshot());
  const base = { format: "doffi-word-clipboard", version: 1, styles: [], nodes: {} },
    definition = (id: string, parent: string) => ({
      style: { styleId: id, cssClass: "class-" + id, kind: "character", basedOn: parent },
      p: {},
      r: {},
    });
  for (const packet of [
    { ...base, version: 99 },
    { ...base, styles: [definition("a", "b"), definition("b", "a")] },
    {
      ...base,
      styles: [
        {
          ...definition("a", ""),
          style: { styleId: "a", cssClass: "class-a", kind: "character" },
          r: { color: { unexpected: true } },
        },
      ],
    },
  ]) {
    await p.evaluate((root, packet) => {
      const bytes = new TextEncoder().encode(JSON.stringify(packet)),
        encoded = btoa(String.fromCharCode(...bytes)),
        transfer = new DataTransfer(),
        event = new ClipboardEvent("paste", {
          bubbles: true,
          cancelable: true,
          clipboardData: transfer,
        });
      transfer.setData(
        "text/html",
        '<span data-doffi-word-clipboard="' + encoded + '">DROP</span>',
      );
      if (event.clipboardData !== transfer)
        Object.defineProperty(event, "clipboardData", { value: transfer });
      root.dispatchEvent(event);
    }, packet);
    expect(await page.evaluate(() => window.paginationHost.session!.snapshot())).toEqual(before);
  }
  expect((await page.evaluate(() => window.paginationHost.errors)).length).toBe(3);
});

test("tracked formatted clipboard HTML keeps source runs and images through review and undo", async ({
  page,
}) => {
  const session = createDocumentSession(
    openWordDocument(createWordDocument({ documentId: "tracked-clipboard" }), { kind: "document" }),
  );
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "tracked-clipboard",
      baseRevision: 0,
      operations: [
        { op: "replaceText", nodeId: "paragraph-1", from: 0, to: 0, expectedText: "", text: "AB" },
        { op: "setWordReview", nodeId: "document-root", trackChanges: true, author: "A" },
      ],
    }),
  );
  await mount(page, session.serialize(), "edit");
  const p = page.frameLocator("iframe").locator('[data-node-id="paragraph-1"]'),
    encoded = Buffer.from(
      readFileSync(new URL("../../examples/fidelity/assets/sample.png", import.meta.url)),
    ).toString("base64");
  await p.evaluate((root) => {
    (root.closest("[data-office-editor]") as HTMLElement).focus();
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
    for (let node = walker.nextNode(); node; node = walker.nextNode()) {
      if (node.textContent !== "AB") continue;
      const range = document.createRange();
      range.setStart(node, 1);
      range.setEnd(node, 2);
      const selection = document.getSelection()!;
      selection.removeAllRanges();
      selection.addRange(range);
      return;
    }
    throw new Error("Missing AB");
  });
  await p.evaluate((root, encoded) => {
    const transfer = new DataTransfer(),
      event = new ClipboardEvent("paste", {
        bubbles: true,
        cancelable: true,
        clipboardData: transfer,
      });
    transfer.setData(
      "text/html",
      '<strong style="font-size:24pt;color:#880000">X</strong><em>Y</em><img src="data:image/png;base64,' +
        encoded +
        '" style="width:12pt;height:8pt">',
    );
    if (event.clipboardData !== transfer)
      Object.defineProperty(event, "clipboardData", { value: transfer });
    root.dispatchEvent(event);
  }, encoded);
  await expect
    .poll(() => page.evaluate(() => window.paginationHost.session!.getNode("paragraph-1")!.text))
    .toBe("ABXY\ufffc");
  await expect(p.getByText("X", { exact: true })).toHaveCSS("font-size", "32px");
  await expect(p.getByText("Y", { exact: true })).toHaveCSS("font-style", "italic");
  expect(
    await page.evaluate(() =>
      window.paginationHost.session!.snapshot().metadata.word!.revisions.map((item) => item.kind),
    ),
  ).toEqual(["delete", "insert"]);
  expect(
    await page.evaluate(() => window.paginationHost.session!.snapshot().manifest.assets.length),
  ).toBe(1);
  await page.evaluate(() =>
    (window.paginationHost.view as import("@doffi/editor").DocumentTextEditor).execute({
      type: "word",
      action: "review",
    }),
  );
  await page.getByRole("button", { name: "拒绝全部显示的修订", exact: true }).click();
  await expect
    .poll(() => page.evaluate(() => window.paginationHost.session!.getNode("paragraph-1")!.text))
    .toBe("AB");
  expect(
    await page.evaluate(() => window.paginationHost.session!.snapshot().manifest.assets.length),
  ).toBe(0);
  await page.evaluate(() => window.paginationHost.session!.undo());
  expect(
    await page.evaluate(() => window.paginationHost.session!.getNode("paragraph-1")!.text),
  ).toBe("ABXY\ufffc");
  await page.evaluate(() => window.paginationHost.session!.undo());
  expect(
    await page.evaluate(() => window.paginationHost.session!.getNode("paragraph-1")!.text),
  ).toBe("AB");
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
});

test("tracked cross-run replacement keeps the first format through continuing input, native reopen and undo", async ({
  page,
}) => {
  const files = importDocx(
      mutateDocx("01-text-numbering.docx", {
        "word/styles.xml": () =>
          nativeStyles('<w:style w:type="paragraph" w:styleId="Normal" w:default="1"/>'),
        "word/document.xml": (source) =>
          nativeBody(
            source,
            '<w:p><w:r><w:rPr><w:b/><w:color w:val="880000"/></w:rPr><w:t>A</w:t></w:r><w:r><w:rPr><w:i/><w:color w:val="000088"/></w:rPr><w:t>B</w:t></w:r></w:p>',
          ),
      }),
      { documentId: "replacement-ui", unsupportedContent: "reject" },
    ),
    session = createDocumentSession(openWordDocument(files, { kind: "document" })),
    story = session.snapshot().metadata.word!.stories[0]!,
    nodes = [story.rootNodeId];
  for (let i = 0; i < nodes.length; i++) nodes.push(...session.getNode(nodes[i]!)!.childNodeIds);
  const paragraph = nodes.find((id) => session.getNode(id)!.tagName === "p")!,
    root = session.getNode(story.rootNodeId)!.parentNodeId!;
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "replacement-ui",
      baseRevision: 0,
      operations: [{ op: "setWordReview", nodeId: root, trackChanges: true, author: "A" }],
    }),
  );
  await mount(page, session.serialize(), "edit");
  const p = page.frameLocator("iframe").locator('[data-node-id="' + paragraph + '"]');
  await p.evaluate((root) => {
    (root.closest("[data-office-editor]") as HTMLElement).focus();
    const range = document.createRange();
    range.selectNodeContents(root);
    const selection = document.getSelection()!;
    selection.removeAllRanges();
    selection.addRange(range);
  });
  await page.keyboard.insertText("X");
  await expect
    .poll(() => page.evaluate((id) => window.paginationHost.session!.getNode(id)!.text, paragraph))
    .toBe("ABX");
  const inserted = p.locator("ins .sdk-word-text-run").last();
  await expect(inserted).toHaveCSS("font-weight", "700");
  await expect(inserted).toHaveCSS("font-style", "normal");
  await page.keyboard.insertText("Z");
  await expect
    .poll(() => page.evaluate((id) => window.paginationHost.session!.getNode(id)!.text, paragraph))
    .toBe("ABXZ");
  expect(
    await page.evaluate(
      () =>
        window.paginationHost
          .session!.snapshot()
          .metadata.word!.revisions.filter((item) => item.kind === "insert").length,
    ),
  ).toBe(1);
  const saved = await page.evaluate(() =>
    [...window.paginationHost.session!.serialize()].map(
      ([path, bytes]) => [path, [...bytes]] as [string, number[]],
    ),
  );
  const native = openWordDocument(
    importDocx(
      exportDocx(
        openWordDocument(new Map(saved.map(([path, bytes]) => [path, new Uint8Array(bytes)])), {
          kind: "document",
        }),
      ).bytes,
      { documentId: "replacement-browser-native", unsupportedContent: "reject" },
    ),
    { kind: "document" },
  );
  expect(native.getNode(native.snapshot().metadata.word!.stories[0]!.rootNodeId)!.text).toBe(
    "ABXZ",
  );
  await page.evaluate(() => window.paginationHost.session!.undo());
  await page.evaluate(() => window.paginationHost.session!.undo());
  expect(
    await page.evaluate((id) => window.paginationHost.session!.getNode(id)!.text, paragraph),
  ).toBe("AB");
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
});

test("independent merged-cell segments paint their physical borders and shading while split cells remain editable and undoable", async ({
  page,
}) => {
  const files = importDocx(continuationCellFixture(), {
    documentId: "physical-cell-ui",
    unsupportedContent: "reject",
  });
  await mount(page, files, "edit");
  const frame = page.frameLocator("iframe"),
    cell = frame.locator('td[rowspan="2"]'),
    segments = cell.locator(".sdk-word-continuation-segment");
  await expect(segments).toHaveCount(2);
  await expect(segments.nth(0)).toHaveCSS("background-color", "rgb(221, 238, 255)");
  await expect(segments.nth(1)).toHaveCSS("background-color", "rgb(255, 255, 136)");
  await expect(segments.nth(1)).toHaveCSS("border-bottom-style", "double");
  await expect(segments.nth(1)).toHaveCSS("border-left-color", "rgb(136, 0, 0)");
  const geometry = await segments.evaluateAll((nodes) =>
    nodes.map((node) => {
      const box = node.getBoundingClientRect();
      return { top: box.top, bottom: box.bottom, height: box.height };
    }),
  );
  expect(geometry[0]!.height * 0.75).toBeCloseTo(40, 0);
  expect(geometry[1]!.top).toBeCloseTo(geometry[0]!.bottom, 0);
  await cell.locator("p").click();
  await page.keyboard.press("End");
  await page.keyboard.insertText("EDIT");
  expect(await page.evaluate(() => window.paginationHost.session!.snapshot().html)).not.toContain(
    "sdk-word-continuation",
  );
  await page.evaluate(() =>
    (window.paginationHost.view as import("@doffi/editor").DocumentTextEditor).execute({
      type: "word",
      action: "tableSettings",
    }),
  );
  await page.getByRole("button", { name: "拆分单元格", exact: true }).click();
  const separate = frame.locator("tr").nth(1).locator("td").first();
  await expect(separate).toHaveAttribute("data-word-text-direction", "tbRl");
  await expect(separate).toHaveCSS("border-bottom-style", "double");
  await expect(separate).toHaveCSS("background-color", "rgb(255, 255, 136)");
  await expect(separate).toHaveCSS("padding-left", "12px");
  await expect(separate).toHaveCSS("vertical-align", "bottom");
  await page.getByRole("button", { name: "关闭表格设置", exact: true }).click();
  const paragraph = separate.locator("p"),
    id = (await paragraph.getAttribute("data-node-id"))!;
  await paragraph.click();
  await page.keyboard.insertText("NEW");
  await expect
    .poll(() => page.evaluate((id) => window.paginationHost.session!.getNode(id)!.text, id))
    .toBe("NEW");
  await page.evaluate(() => window.paginationHost.session!.undo());
  await page.evaluate(() => window.paginationHost.session!.undo());
  await expect(frame.locator('td[rowspan="2"] .sdk-word-continuation-segment')).toHaveCount(2);
  await page.evaluate(() => window.paginationHost.session!.undo());
  expect(await page.evaluate(() => window.paginationHost.session!.snapshot().html)).not.toContain(
    "EDIT",
  );
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
});

test("right-to-left tables mirror logical borders, retain nested direction and resize in visual coordinates with undo", async ({
  page,
}) => {
  const files = importDocx(tableBidiFixture(), {
    documentId: "bidi-ui",
    unsupportedContent: "reject",
  });
  await mount(page, files, "edit");
  const frame = page.frameLocator("iframe"),
    table = frame.locator("[data-office-editor] table").first(),
    row = table.locator(":scope > tbody > tr,:scope > tr").first(),
    first = row.locator(":scope > td").first(),
    second = row.locator(":scope > td").nth(1);
  expect((await first.boundingBox())!.x).toBeGreaterThan((await second.boundingBox())!.x);
  await expect(first).toHaveCSS("border-right-color", "rgb(0, 0, 136)");
  await expect(first).toHaveCSS("border-right-style", "solid");
  await expect(first).toHaveCSS("border-left-color", "rgb(136, 0, 0)");
  await expect(first).toHaveCSS("padding-right", "12px");
  await expect(first).toHaveCSS("padding-left", "4px");
  await expect(first).toHaveCSS("direction", "ltr");
  const inner = table.locator("table"),
    left = inner.locator("td").first(),
    right = inner.locator("td").last();
  expect((await left.boundingBox())!.x).toBeLessThan((await right.boundingBox())!.x);
  await first.hover();
  const guide = frame.getByRole("button", { name: "调整第 1 列宽度", exact: true }),
    guideBox = await guide.boundingBox(),
    cellBox = await first.boundingBox();
  expect(guideBox!.x + 4).toBeCloseTo(cellBox!.x, 0);
  await guide.focus();
  await page.keyboard.press("ArrowLeft");
  expect(
    (
      await page.evaluate(
        () =>
          window.paginationHost
            .session!.snapshot()
            .metadata.word!.tables.find(
              (info) => window.paginationHost.session!.getNode(info.nodeId)!.text === "ABLRCD",
            )!.grid,
      )
    )[0],
  ).toBe(81);
  await page.evaluate(() => window.paginationHost.session!.undo());
  await first.locator("p").click();
  await page.keyboard.press("End");
  await page.keyboard.insertText("EDIT");
  expect(await page.evaluate(() => window.paginationHost.session!.snapshot().html)).not.toMatch(
    /direction:rtl|sdk-word-table-bidi|padding-right:9pt/,
  );
  await page.evaluate(() =>
    (window.paginationHost.view as import("@doffi/editor").DocumentTextEditor).execute({
      type: "word",
      action: "tableSettings",
    }),
  );
  await page.getByLabel("表格排列方向", { exact: true }).selectOption("0");
  await page.getByRole("button", { name: "应用表格方向", exact: true }).click();
  expect((await first.boundingBox())!.x).toBeLessThan((await second.boundingBox())!.x);
  await expect(first).toHaveCSS("border-left-color", "rgb(0, 0, 136)");
  await page.evaluate(() => window.paginationHost.session!.undo());
  expect((await first.boundingBox())!.x).toBeGreaterThan((await second.boundingBox())!.x);
  await page.evaluate(() => window.paginationHost.session!.undo());
  expect(await page.evaluate(() => window.paginationHost.session!.snapshot().html)).not.toContain(
    "EDIT",
  );
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
});

test("nonzero table cell spacing separates borders, isolates nested tables and updates through the panel with input and undo", async ({
  page,
}) => {
  await mount(
    page,
    importDocx(tableSpacingFixture(), { documentId: "spacing-ui", unsupportedContent: "reject" }),
    "edit",
  );
  const frame = page.frameLocator("iframe"),
    table = frame.locator("[data-office-editor] table").first(),
    row = table.locator(":scope > tbody > tr,:scope > tr").first(),
    first = row.locator(":scope > td").first(),
    second = row.locator(":scope > td").nth(1),
    inner = table.locator("table");
  await expect(table).toHaveCSS("border-collapse", "separate");
  await expect(table).toHaveCSS("border-spacing", "8px");
  await expect(inner).toHaveCSS("border-spacing", "0px");
  const gap = async () => {
    const a = (await first.boundingBox())!,
      b = (await second.boundingBox())!;
    return (b.x - a.x - a.width) * 0.75;
  };
  expect(await gap()).toBeCloseTo(6, 1);
  await first.locator("p").click();
  await page.keyboard.press("End");
  await page.keyboard.insertText("EDIT");
  expect(await page.evaluate(() => window.paginationHost.session!.snapshot().html)).not.toContain(
    "sdk-word-table-bidi",
  );
  await page.evaluate(() =>
    (window.paginationHost.view as import("@doffi/editor").DocumentTextEditor).execute({
      type: "word",
      action: "tableSettings",
    }),
  );
  await page.getByLabel("表格单元格间距（磅）", { exact: true }).fill("3");
  await page.getByRole("button", { name: "应用表格间距", exact: true }).click();
  await expect(table).toHaveCSS("border-spacing", "4px");
  expect(await gap()).toBeCloseTo(3, 1);
  await page.evaluate(() => window.paginationHost.session!.undo());
  expect(await gap()).toBeCloseTo(6, 1);
  const id = (await table.getAttribute("data-node-id"))!;
  await page.evaluate((id) => {
    const session = window.paginationHost.session!,
      snapshot = session.snapshot();
    session.commit(
      session.prepare({
        protocolVersion: 1,
        documentId: snapshot.manifest.documentId,
        baseRevision: snapshot.manifest.revision,
        operations: [
          { op: "formatWordTable", nodeId: id, properties: { "--word-table-cell-spacing": "2pt" } },
        ],
      }),
    );
  }, id);
  expect(await table.evaluate((node) => (node as HTMLElement).style.borderSpacing)).toBe("2pt");
  expect(Math.abs((await gap()) - 2)).toBeLessThan(0.76);
  const entries = await page.evaluate(() =>
    [...window.paginationHost.session!.serialize()].map(
      ([path, bytes]) => [path, [...bytes]] as [string, number[]],
    ),
  );
  const native = openWordDocument(
    importDocx(
      exportDocx(
        openWordDocument(new Map(entries.map(([path, bytes]) => [path, new Uint8Array(bytes)])), {
          kind: "document",
        }),
      ).bytes,
      { documentId: "spacing-browser-native", unsupportedContent: "reject" },
    ),
    { kind: "document" },
  );
  expect(native.snapshot().html).toContain("AEDIT");
  expect(native.snapshot().html).toContain("--word-table-cell-spacing:2pt");
  await page.evaluate(() => window.paginationHost.session!.undo());
  await page.evaluate(() => window.paginationHost.session!.undo());
  expect(await page.evaluate(() => window.paginationHost.session!.snapshot().html)).not.toContain(
    "EDIT",
  );
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
});

test("row spacing zero and nonzero overrides stay inside table bounds and share tracking, input and undo", async ({
  page,
}) => {
  const session = createDocumentSession(
      openWordDocument(
        importDocx(rowSpacingFixture(), {
          documentId: "row-spacing-ui",
          unsupportedContent: "reject",
        }),
        { kind: "document" },
      ),
    ),
    root = session.getNode(session.snapshot().metadata.word!.stories[0]!.rootNodeId)!.parentNodeId!;
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "row-spacing-ui",
      baseRevision: 0,
      operations: [{ op: "setWordReview", nodeId: root, trackChanges: true, author: "A" }],
    }),
  );
  await mount(page, session.serialize(), "edit");
  const frame = page.frameLocator("iframe"),
    table = frame.locator("table").first(),
    rows = table.locator("tr"),
    gap = async (index: number) => {
      const paint = rows.nth(index).locator(".sdk-word-cell-spacing-paint");
      const a = (await paint.nth(0).boundingBox())!,
        b = (await paint.nth(1).boundingBox())!;
      return (b.x - a.x - a.width) * 0.75;
    };
  await expect(table.locator(".sdk-word-cell-spacing-paint")).toHaveCount(6);
  expect(await gap(0)).toBeCloseTo(0, 1);
  expect(await gap(1)).toBeCloseTo(12, 1);
  expect(await gap(2)).toBeCloseTo(6, 1);
  const width = (await table.boundingBox())!.width,
    paragraph = rows.first().locator("p").first();
  await paragraph.click();
  await page.evaluate(() =>
    (window.paginationHost.view as import("@doffi/editor").DocumentTextEditor).execute({
      type: "word",
      action: "tableSettings",
    }),
  );
  await page.getByLabel("当前行间距来源", { exact: true }).selectOption("custom");
  await page.getByLabel("当前行单元格间距（磅）", { exact: true }).fill("3");
  await page.getByRole("button", { name: "应用当前行间距", exact: true }).click();
  expect(await gap(0)).toBeCloseTo(3, 1);
  expect((await table.boundingBox())!.width).toBeCloseTo(width, 1);
  await page.getByRole("button", { name: "关闭表格设置", exact: true }).click();
  await paragraph.click();
  await page.keyboard.press("End");
  await page.keyboard.insertText("EDIT");
  const id = (await paragraph.getAttribute("data-node-id"))!;
  await expect
    .poll(() => page.evaluate((id) => window.paginationHost.session!.getNode(id)!.text, id))
    .toBe("A0EDIT");
  expect(await page.evaluate(() => window.paginationHost.session!.snapshot().html)).not.toMatch(
    /sdk-word-row-spacing|calc\(100%|sdk-word-cell-spacing/,
  );
  await page.evaluate(() =>
    (window.paginationHost.view as import("@doffi/editor").DocumentTextEditor).execute({
      type: "word",
      action: "review",
    }),
  );
  await page.getByLabel("修订显示", { exact: true }).selectOption("original");
  expect(await gap(0)).toBeCloseTo(0, 1);
  await page.getByLabel("修订显示", { exact: true }).selectOption("all");
  expect(await gap(0)).toBeCloseTo(3, 1);
  const entries = await page.evaluate(() =>
    [...window.paginationHost.session!.serialize()].map(
      ([path, bytes]) => [path, [...bytes]] as [string, number[]],
    ),
  );
  const native = openWordDocument(
    importDocx(
      exportDocx(
        openWordDocument(new Map(entries.map(([path, bytes]) => [path, new Uint8Array(bytes)])), {
          kind: "document",
        }),
      ).bytes,
      { documentId: "row-space-browser-native", unsupportedContent: "reject" },
    ),
    { kind: "document" },
  );
  expect(native.snapshot().html).toContain("--word-row-cell-spacing:3pt");
  await page.evaluate(() => window.paginationHost.session!.undo());
  await page.evaluate(() => window.paginationHost.session!.undo());
  expect(await gap(0)).toBeCloseTo(0, 1);
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
});

test("row spacing remains a reversible view when a long native table paginates and receives input", async ({
  page,
}) => {
  const session = createDocumentSession(
      openWordDocument(
        importDocx(rowSpacingFixture(24), {
          documentId: "row-spacing-pages",
          unsupportedContent: "reject",
        }),
        { kind: "document" },
      ),
    ),
    section = session.snapshot().metadata.sections![0]!.rootNodeId;
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "row-spacing-pages",
      baseRevision: 0,
      operations: [
        {
          op: "setWordSection",
          nodeId: section,
          changes: {
            pageSetup: {
              unit: "pt",
              width: 300,
              height: 220,
              margins: { top: 15, right: 15, bottom: 15, left: 15 },
            },
          },
        },
      ],
    }),
  );
  await mount(page, session.serialize(), "edit");
  expect((await page.evaluate(() => window.paginationHost.layout())).pages.length).toBeGreaterThan(
    2,
  );
  const paragraph = page.frameLocator("iframe").locator("p").filter({ hasText: "A18" }).first(),
    id = (await paragraph.getAttribute("data-node-id"))!;
  await paragraph.click();
  await page.keyboard.press("End");
  await page.keyboard.insertText("X");
  await expect
    .poll(() => page.evaluate((id) => window.paginationHost.session!.getNode(id)!.text, id))
    .toBe("A18X");
  expect(await page.evaluate(() => window.paginationHost.session!.snapshot().html)).not.toMatch(
    /sdk-word-row-spacing|sdk-word-cell-spacing|position:absolute/,
  );
  await page.evaluate(() => window.paginationHost.session!.undo());
  expect(await page.evaluate((id) => window.paginationHost.session!.getNode(id)!.text, id)).toBe(
    "A18",
  );
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
});

test("row-specific spacing composes with RTL merged continuations without duplicating source text", async ({
  page,
}) => {
  const session = createDocumentSession(
      openWordDocument(
        importDocx(continuationCellFixture(), {
          documentId: "row-space-combined",
          unsupportedContent: "reject",
        }),
        { kind: "document" },
      ),
    ),
    tableId = session.snapshot().metadata.word!.tables[0]!.nodeId;
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "row-space-combined",
      baseRevision: 0,
      operations: [
        {
          op: "formatWordTable",
          nodeId: tableId,
          properties: { "--word-table-bidi": "1" },
          rowIndex: 0,
          rowCellSpacing: 3,
        },
        { op: "formatWordTable", nodeId: tableId, rowIndex: 1, rowCellSpacing: 0 },
      ],
    }),
  );
  await mount(page, session.serialize(), "edit");
  const frame = page.frameLocator("iframe"),
    table = frame.locator("table").first(),
    merged = table.locator('td[rowspan="2"]'),
    segments = merged.locator(".sdk-word-continuation-segment");
  await expect(segments).toHaveCount(2);
  await expect(segments.last()).toHaveCSS("background-color", "rgb(255, 255, 136)");
  await expect(segments.last()).toHaveCSS("border-right-color", "rgb(136, 0, 0)");
  const outer = await merged.evaluate((node) => node.getBoundingClientRect().toJSON()),
    ink = await merged.locator("p").evaluate((node) => {
      const range = document.createRange();
      range.selectNodeContents(node);
      return range.getBoundingClientRect().toJSON();
    });
  expect(ink.left).toBeGreaterThan(outer!.x);
  expect(ink.right).toBeLessThan(outer!.x + outer!.width);
  await merged.locator("p").click();
  await page.keyboard.press("End");
  await page.keyboard.insertText("X");
  expect(await page.evaluate(() => window.paginationHost.session!.snapshot().html)).toContain(
    "MERGEDX",
  );
  expect(await page.evaluate(() => window.paginationHost.session!.snapshot().html)).not.toMatch(
    /sdk-word-row-spacing|sdk-word-continuation-segment|margin-left:-/,
  );
  await page.evaluate(() => window.paginationHost.session!.undo());
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
});

test("row table exceptions sit below direct row spacing and reappear after clearing or reviewing an override", async ({
  page,
}) => {
  const session = createDocumentSession(
      openWordDocument(
        importDocx(rowExceptionFixture(), {
          documentId: "row-ex-ui",
          unsupportedContent: "reject",
        }),
        { kind: "document" },
      ),
    ),
    root = session.getNode(session.snapshot().metadata.word!.stories[0]!.rootNodeId)!.parentNodeId!;
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "row-ex-ui",
      baseRevision: 0,
      operations: [{ op: "setWordReview", nodeId: root, trackChanges: true, author: "A" }],
    }),
  );
  await mount(page, session.serialize(), "edit");
  const frame = page.frameLocator("iframe"),
    table = frame.locator("table").first(),
    rows = table.locator("tr"),
    gap = async (index: number) => {
      const paints = rows.nth(index).locator(".sdk-word-cell-spacing-paint"),
        a = (await paints.first().boundingBox())!,
        b = (await paints.last().boundingBox())!;
      return (b.x - a.x - a.width) * 0.75;
    };
  expect(await gap(0)).toBeCloseTo(0, 1);
  expect(await gap(1)).toBeCloseTo(12, 1);
  expect(await gap(2)).toBeCloseTo(4, 1);
  await rows.last().locator("p").first().click();
  await page.evaluate(() =>
    (window.paginationHost.view as import("@doffi/editor").DocumentTextEditor).execute({
      type: "word",
      action: "tableSettings",
    }),
  );
  await expect(page.getByLabel("当前行间距来源", { exact: true })).toHaveValue("inherit");
  await expect(page.getByLabel("当前行单元格间距（磅）", { exact: true })).toHaveValue("4");
  await page.getByRole("button", { name: "关闭表格设置", exact: true }).click();
  const width = (await table.boundingBox())!.width;
  await rows.first().locator("p").first().click();
  await page.evaluate(() =>
    (window.paginationHost.view as import("@doffi/editor").DocumentTextEditor).execute({
      type: "word",
      action: "tableSettings",
    }),
  );
  await page.getByLabel("当前行间距来源", { exact: true }).selectOption("inherit");
  await page.getByRole("button", { name: "应用当前行间距", exact: true }).click();
  expect(await gap(0)).toBeCloseTo(4, 1);
  expect((await table.boundingBox())!.width).toBeCloseTo(width, 1);
  await page.getByRole("button", { name: "关闭表格设置", exact: true }).click();
  await page.evaluate(() =>
    (window.paginationHost.view as import("@doffi/editor").DocumentTextEditor).execute({
      type: "word",
      action: "review",
    }),
  );
  await page.getByLabel("修订显示", { exact: true }).selectOption("original");
  expect(await gap(0)).toBeCloseTo(0, 1);
  await page.getByLabel("修订显示", { exact: true }).selectOption("all");
  expect(await gap(0)).toBeCloseTo(4, 1);
  await page.evaluate(() => window.paginationHost.session!.undo());
  expect(await gap(0)).toBeCloseTo(0, 1);
  expect(await page.evaluate(() => window.paginationHost.session!.snapshot().html)).toContain(
    "word-row-exception",
  );
  expect(await page.evaluate(() => window.paginationHost.session!.snapshot().html)).not.toMatch(
    /sdk-word-row-spacing|sdk-word-cell-spacing/,
  );
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
});

for (const inline of [false, true])
  test(
    "native object comments retain their anchors while surrounding text stays editable " + inline,
    async ({ page }) => {
      const files = importDocx(nativeCommentFixture("inside", inline), {
        documentId: "native-comment-ui",
      });
      await mount(page, files, "edit");
      const data = await page.evaluate(() => {
        const snapshot = window.paginationHost.session!.snapshot();
        return {
          atom: snapshot.metadata.word!.nativeContent!.fragments.find(
            (item) => item.kind !== "property",
          )!.nodeId!,
          annotation: snapshot.annotations[0]!.annotationId,
        };
      });
      const frame = page.frameLocator("iframe"),
        atom = frame.locator('[data-node-id="' + data.atom + '"]');
      await expect(atom).toBeVisible();
      const paragraph = frame.locator("p").filter({ hasText: "AFTER" }).last();
      await paragraph.click();
      await page.keyboard.press("End");
      await page.keyboard.insertText("EDIT");
      await expect
        .poll(() => page.evaluate(() => window.paginationHost.session!.snapshot().html))
        .toContain("AFTEREDIT");
      await expect(atom).toHaveCount(1);
      expect(
        await page.evaluate(
          (id) =>
            window.paginationHost
              .session!.getAnnotations()
              .find((item) => item.annotationId === id)!.status,
          data.annotation,
        ),
      ).toBe("open");
      await page.evaluate(() => window.paginationHost.session!.undo());
      expect(
        await page.evaluate(() => window.paginationHost.session!.snapshot().html),
      ).not.toContain("AFTEREDIT");
      expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
    },
  );

test("floating tables wrap following text, retain cell input and undo native positioning controls", async ({
  page,
}) => {
  await mount(
    page,
    importDocx(floatingTableFixture(), {
      documentId: "floating-table-ui",
      unsupportedContent: "reject",
    }),
    "edit",
  );
  const frame = page.frameLocator("iframe"),
    table = frame.locator("table[data-node-id]").first(),
    wrapper = frame.locator(".sdk-word-floating-table");
  await expect(wrapper).toHaveCount(1);
  const sourceStyle = await page.evaluate(() => {
    const session = window.paginationHost.session!;
    return session.getNode(session.snapshot().metadata.word!.tables[0]!.nodeId)!.attributes.style;
  });
  const geometry = await table.evaluate((table) => {
    const p = table.parentElement!.nextElementSibling!,
      range = document.createRange();
    range.selectNodeContents(p);
    const text = range.getClientRects()[0]!,
      box = table.getBoundingClientRect();
    return {
      left: box.left,
      right: box.right,
      top: box.top,
      textLeft: text.left,
      textTop: text.top,
    };
  });
  expect(geometry.textLeft).toBeGreaterThan(geometry.right);
  expect(Math.abs(geometry.textTop - geometry.top)).toBeLessThan(25);
  await table.locator("p").click();
  await page.keyboard.press("End");
  await page.keyboard.insertText("EDIT");
  await expect(table).toContainText("CELLEDIT");
  expect(
    await page.evaluate(() => {
      const session = window.paginationHost.session!;
      return session.getNode(session.snapshot().metadata.word!.tables[0]!.nodeId)!.attributes.style;
    }),
  ).toBe(sourceStyle);
  expect(await page.evaluate(() => window.paginationHost.session!.snapshot().html)).not.toContain(
    "sdk-word-floating-table",
  );
  await page.evaluate(() =>
    (window.paginationHost.view as import("@doffi/editor").DocumentTextEditor).execute({
      type: "word",
      action: "tableSettings",
    }),
  );
  await page.getByLabel("表格文字环绕", { exact: true }).selectOption("0");
  await page.getByRole("button", { name: "应用表格位置", exact: true }).click();
  await expect(wrapper).toHaveCount(0);
  await page.evaluate(() => window.paginationHost.session!.undo());
  await expect(wrapper).toHaveCount(1);
  await page.evaluate(() => window.paginationHost.session!.undo());
  await expect(table).toContainText("CELL");
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
});

test("floating table page anchors ignore RTL mirroring and retain physical page positions", async ({
  page,
}) => {
  await mount(
    page,
    importDocx(
      floatingTableFixture(
        '<w:tblpPr w:horzAnchor="page" w:vertAnchor="page" w:tblpX="1441" w:tblpY="2881"/><w:bidiVisual/>',
      ),
      { documentId: "floating-page-ui", unsupportedContent: "reject" },
    ),
    "edit",
  );
  const table = page.frameLocator("iframe").locator("table[data-node-id]").first(),
    geometry = await table.evaluate((table) => {
      const paper = document.querySelector(".sdk-paper")!.getBoundingClientRect(),
        box = table.getBoundingClientRect();
      return { x: (box.left - paper.left) * 0.75, y: (box.top - paper.top) * 0.75 };
    });
  expect(geometry.x).toBeCloseTo(72, 0);
  expect(geometry.y).toBeCloseTo(144, 0);
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
});

test("floating tables continue rows across pages while following body text remains editable", async ({
  page,
}) => {
  const bytes = floatingTableFixture(),
    files = new Map(
      importDocx(bytes, { documentId: "floating-long-ui", unsupportedContent: "reject" }),
    ),
    session = createDocumentSession(openWordDocument(files, { kind: "document" })),
    id = session.snapshot().metadata.word!.tables[0]!.nodeId;
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "floating-long-ui",
      baseRevision: 0,
      operations: [
        { op: "editWordTable", nodeId: id, change: { kind: "insertRows", at: 1, count: 45 } },
      ],
    }),
  );
  await mount(page, session.serialize(), "edit");
  const layout = await page.evaluate(() => window.paginationHost.layout());
  expect(layout.pages.length).toBeGreaterThan(1);
  await expect(
    page
      .frameLocator("iframe")
      .locator(".sdk-word-floating-table > table > tbody > tr:not(.sdk-page-gap)"),
  ).toHaveCount(46);
  const tail = page.frameLocator("iframe").locator("p").filter({ hasText: "AFTER text" }).last();
  await tail.evaluate((p) => {
    (p.closest("[data-office-editor]") as HTMLElement).focus();
    const range = document.createRange();
    range.selectNodeContents(p);
    range.collapse(false);
    const selection = document.getSelection()!;
    selection.removeAllRanges();
    selection.addRange(range);
  });
  await page.keyboard.insertText("TAIL-EDIT");
  await expect
    .poll(() => page.evaluate(() => window.paginationHost.session!.snapshot().html))
    .toContain("TAIL-EDIT");
  await page.evaluate(() => window.paginationHost.session!.undo());
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
});

for (const overlap of ["overlap", "never"])
  test("floating tables honor explicit native overlap " + overlap, async ({ page }) => {
    await mount(
      page,
      importDocx(
        floatingTableFixture(
          '<w:tblpPr w:horzAnchor="page" w:vertAnchor="page" w:tblpX="1441" w:tblpY="1441"/><w:tblOverlap w:val="' +
            overlap +
            '"/>',
          4,
          2,
        ),
        { documentId: "floating-overlap-ui", unsupportedContent: "reject" },
      ),
      "edit",
    );
    const boxes = await page
      .frameLocator("iframe")
      .locator(".sdk-word-floating-table > table")
      .evaluateAll((tables) =>
        tables.map((table) => {
          const box = table.getBoundingClientRect();
          return { left: box.left, top: box.top, bottom: box.bottom };
        }),
      );
    expect(boxes).toHaveLength(2);
    expect(boxes[0]!.left).toBeCloseTo(boxes[1]!.left, 0);
    if (overlap === "overlap") expect(boxes[0]!.top).toBeCloseTo(boxes[1]!.top, 0);
    else expect(boxes[1]!.top).toBeGreaterThanOrEqual(boxes[0]!.bottom - 0.1);
    expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
  });

test("floating header tables retain page coordinates in every repeated page instance", async ({
  page,
}) => {
  await mount(
    page,
    importDocx(floatingHeaderFixture(), {
      documentId: "floating-header-ui",
      unsupportedContent: "reject",
    }),
  );
  const pages = (await page.evaluate(() => window.paginationHost.layout())).pages;
  expect(pages.length).toBeGreaterThan(1);
  const instances = await page
    .frameLocator("iframe")
    .locator("body")
    .evaluate((body) =>
      [
        ...body.querySelectorAll<HTMLTableElement>(".sdk-paper .sdk-word-floating-table > table"),
      ].map((table) => {
        const paper = table.closest(".sdk-paper")!.getBoundingClientRect(),
          box = table.getBoundingClientRect();
        return {
          x: (box.left - paper.left) * 0.75,
          y: (box.top - paper.top) * 0.75,
          width: box.width * 0.75,
          connected: table.isConnected,
          display: getComputedStyle(table).display,
          carrierDisplay: getComputedStyle(table.parentElement!).display,
          style: table.parentElement!.getAttribute("style"),
          anchors: table.closest(".sdk-paper")!.querySelectorAll("[data-sdk-word-table-anchor]")
            .length,
        };
      }),
    );
  expect(instances).toHaveLength(pages.length);
  expect(
    (await page.evaluate(() => window.paginationHost.layout())).diagnostics.filter(
      (item) => item.code === "ContentOverflow",
    ),
  ).toEqual([]);
  for (const item of instances) {
    expect(item.x, JSON.stringify(item)).toBeCloseTo(72, 0);
    expect(item.y).toBeCloseTo(24, 0);
    expect(item.width).toBeGreaterThan(100);
  }
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
});

test("drawing settings retain independent text distances and expose native anchor switches in one undo step", async ({
  page,
}) => {
  const session = createDocumentSession(openWordDocument(drawingFiles(), { kind: "document" })),
    snapshot = session.snapshot(),
    shape = snapshot.metadata.word!.drawings!.find((item) => item.kind === "shape")!;
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: snapshot.manifest.documentId,
      baseRevision: snapshot.manifest.revision,
      operations: [
        {
          op: "setWordDrawingLayout",
          nodeId: shape.nodeId,
          layout: {
            ...shape.layout,
            wrapPolygon: shape.layout.wrapPolygon?.map((point) => ({ ...point })),
            distances: { left: 1, right: 2, top: 3, bottom: 4 },
            allowOverlap: false,
            layoutInCell: false,
            locked: true,
          },
        },
      ],
    }),
  );
  await mount(page, session.serialize(), "edit");
  await page
    .frameLocator("iframe")
    .locator('[data-word-drawing="shape"]:visible')
    .click({ position: { x: 10, y: 10 } });
  for (const [label, value] of [
    ["左环绕距离（磅）", "1"],
    ["右环绕距离（磅）", "2"],
    ["上环绕距离（磅）", "3"],
    ["下环绕距离（磅）", "4"],
    ["锁定对象锚点", "1"],
    ["对象允许重叠", "0"],
    ["对象在单元格内布局", "0"],
  ] as const)
    await expect(page.getByLabel(label, { exact: true })).toHaveValue(value);
  await page.getByLabel("对象水平位置（磅）", { exact: true }).fill("45");
  await page.getByRole("button", { name: "应用对象位置和布局", exact: true }).click();
  const changed = await page.evaluate(
    () =>
      window.paginationHost
        .session!.snapshot()
        .metadata.word!.drawings!.find((item) => item.kind === "shape")!,
  );
  expect(changed.layout.distances).toEqual({ left: 1, right: 2, top: 3, bottom: 4 });
  expect(changed.layout).toMatchObject({ allowOverlap: false, layoutInCell: false, locked: true });
  expect(changed.geometry.x).toBe(45);
  await page.evaluate(() => window.paginationHost.session!.undo());
  expect(
    (
      await page.evaluate(
        () =>
          window.paginationHost
            .session!.snapshot()
            .metadata.word!.drawings!.find((item) => item.kind === "shape")!,
      )
    ).geometry.x,
  ).toBe(shape.geometry.x);
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
});

test("contextual paragraph spacing follows identical styles and cell boundaries without altering saved margins", async ({
  page,
}) => {
  const { contextualSpacingFixture } = await import(
    "../../packages/office/tests/contextual-spacing-fixture.js"
  );
  await mount(
    page,
    importDocx(contextualSpacingFixture(), {
      documentId: "contextual-ui",
      unsupportedContent: "reject",
    }),
    "edit",
  );
  const frame = page.frameLocator("iframe"),
    paragraph = (text: string) =>
      frame.locator("p").filter({ hasText: new RegExp("^" + text + "$") });
  await expect(paragraph("A")).toHaveCSS("margin-top", "16px");
  await expect(paragraph("A")).toHaveCSS("margin-bottom", "0px");
  await expect(paragraph("B")).toHaveCSS("margin-top", "0px");
  await expect(paragraph("B")).toHaveCSS("margin-bottom", "24px");
  await expect(paragraph("C")).toHaveCSS("margin-top", "16px");
  await expect(paragraph("D")).toHaveCSS("margin-top", "16px");
  await expect(paragraph("Y")).toHaveCSS("margin-bottom", "0px");
  await expect(paragraph("Z")).toHaveCSS("margin-top", "0px");
  const id = (await paragraph("B").getAttribute("data-node-id"))!,
    before = await page.evaluate(
      (id) => window.paginationHost.session!.getNode(id)!.attributes.style,
      id,
    );
  await paragraph("B").click();
  await page.evaluate(() =>
    (window.paginationHost.view as import("@doffi/editor").DocumentTextEditor).execute({
      type: "word",
      action: "paragraph",
      properties: { "--word-contextual-spacing": "0" },
    }),
  );
  await expect(paragraph("B")).toHaveCSS("margin-top", "16px");
  await page.evaluate(() => window.paginationHost.session!.undo());
  await expect(paragraph("B")).toHaveCSS("margin-top", "0px");
  await paragraph("B").click();
  await page.keyboard.press("End");
  await page.keyboard.insertText("EDIT");
  expect(
    await page.evaluate((id) => window.paginationHost.session!.getNode(id)!.attributes.style, id),
  ).toBe(before);
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
});

for (const fixture of ["style", "continuation", "nested"] as const)
  test(
    "whole-table semantic clipboard preserves isolated styles, merge formats and editing " +
      fixture,
    async ({ page }, testInfo) => {
      const { tableStylesFixture } = await import(
          "../../packages/office/tests/table-style-fixture.js"
        ),
        { continuationCellFixture } = await import(
          "../../packages/office/tests/continuation-cell-fixture.js"
        ),
        { nestedTableContentFixture } = await import(
          "../../packages/office/tests/word-table-content-fixture.js"
        );
      await mount(
        page,
        importDocx(
          fixture === "style"
            ? tableStylesFixture()
            : fixture === "nested"
              ? nestedTableContentFixture()
              : continuationCellFixture(),
          { documentId: "table-copy-ui-source", unsupportedContent: "reject" },
        ),
        fixture === "style" ? "preview" : "edit",
      );
      const sourceTable = page.frameLocator("iframe").locator("table").first(),
        appearance = await sourceTable.locator("td").evaluateAll((cells) =>
          cells.map((cell) => {
            const p = cell.querySelector("p")!,
              style = getComputedStyle(p);
            return {
              background: getComputedStyle(cell).backgroundColor,
              color: style.color,
              font: style.fontFamily,
              weight: style.fontWeight,
            };
          }),
        );
      const segments = await sourceTable
        .locator(".sdk-word-continuation-segment")
        .evaluateAll((nodes) =>
          nodes.map((node) => {
            const css = getComputedStyle(node);
            return {
              background: css.backgroundColor,
              border: css.borderBottomStyle,
              color: css.borderLeftColor,
            };
          }),
        );
      const copied = await sourceTable.evaluate((table) => {
          const range = document.createRange();
          range.selectNode(table);
          const selection = document.getSelection()!;
          selection.removeAllRanges();
          selection.addRange(range);
          const data = new DataTransfer(),
            event = new ClipboardEvent("copy", { bubbles: true, cancelable: true });
          Object.defineProperty(event, "clipboardData", { value: data });
          table.dispatchEvent(event);
          return { html: data.getData("text/html"), text: data.getData("text/plain") };
        }),
        payload = JSON.parse(
          Buffer.from(
            /data-doffi-word-clipboard="([^"]+)"/.exec(copied.html)![1]!,
            "base64",
          ).toString("utf8"),
        );
      expect(payload.semantics.tables).toHaveLength(fixture === "nested" ? 2 : 1);
      if (fixture === "continuation") expect(payload.semantics.extraClasses).toBeTruthy();
      const target = createDocumentSession(
        openWordDocument(createWordDocument({ documentId: "table-copy-ui-target" }), {
          kind: "document",
        }),
      );
      target.commit(
        target.prepare({
          protocolVersion: 1,
          documentId: "table-copy-ui-target",
          baseRevision: 0,
          operations: [
            {
              op: "replaceText",
              nodeId: "paragraph-1",
              from: 0,
              to: 0,
              expectedText: "",
              text: "AB",
            },
            {
              op: "setWordStyle",
              nodeId: "document-root",
              style: { styleId: "Red", kind: "table", cssClass: "target-red" },
              cellProperties: { "background-color": "#00ffff" },
            },
            {
              op: "setWordReview",
              nodeId: "document-root",
              trackChanges: fixture !== "style",
              author: "Receiver",
            },
          ],
        }),
      );
      await mount(page, target.serialize(), "edit");
      await page
        .frameLocator("iframe")
        .locator('[data-node-id="paragraph-1"]')
        .evaluate((paragraph, copied) => {
          (paragraph.closest("[data-office-editor]") as HTMLElement).focus();
          const text = document.createTreeWalker(paragraph, NodeFilter.SHOW_TEXT).nextNode()!,
            range = document.createRange();
          range.setStart(text, 1);
          range.collapse(true);
          const selection = document.getSelection()!;
          selection.removeAllRanges();
          selection.addRange(range);
          const data = new DataTransfer();
          data.setData("text/html", copied.html);
          data.setData("text/plain", copied.text);
          const event = new ClipboardEvent("paste", { bubbles: true, cancelable: true });
          Object.defineProperty(event, "clipboardData", { value: data });
          paragraph.dispatchEvent(event);
        }, copied);
      await expect
        .poll(() =>
          page.evaluate(
            () => window.paginationHost.session!.snapshot().metadata.word!.tables.length,
          ),
        )
        .toBe(fixture === "nested" ? 2 : 1);
      await page.evaluate(() => window.paginationHost.view!.settled());
      const table = page.frameLocator("iframe").locator("table").first();
      expect(
        await table.locator("td").evaluateAll((cells) =>
          cells.map((cell) => {
            const p = cell.querySelector("p")!,
              style = getComputedStyle(p);
            return {
              background: getComputedStyle(cell).backgroundColor,
              color: style.color,
              font: style.fontFamily,
              weight: style.fontWeight,
            };
          }),
        ),
      ).toEqual(appearance);
      if (fixture === "style") expect(await table.getAttribute("data-word-style")).not.toBe("Red");
      else if (fixture === "continuation")
        expect(
          (
            await page.evaluate(
              () =>
                window.paginationHost.session!.snapshot().metadata.word!.tables[0]!
                  .continuationFormats,
            )
          )?.[0]?.textDirection,
        ).toBe("tbRl");
      expect(
        await table.locator(".sdk-word-continuation-segment").evaluateAll((nodes) =>
          nodes.map((node) => {
            const css = getComputedStyle(node);
            return {
              background: css.backgroundColor,
              border: css.borderBottomStyle,
              color: css.borderLeftColor,
            };
          }),
        ),
      ).toEqual(segments);
      const last =
        fixture === "nested"
          ? table.locator("table p").filter({ hasText: "INNER-B" })
          : table.locator("p").last();
      await last.click();
      await page.keyboard.press("End");
      await page.keyboard.insertText("EDIT");
      await expect(last).toContainText("EDIT");
      expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
      const entries = await page.evaluate(() =>
        [...window.paginationHost.session!.serialize()].map(
          ([path, bytes]) => [path, [...bytes]] as [string, number[]],
        ),
      );
      let native = exportDocx(
        openWordDocument(new Map(entries.map(([path, bytes]) => [path, new Uint8Array(bytes)])), {
          kind: "document",
        }),
      ).bytes;
      const nativePath = testInfo.outputPath("pasted-table.docx");
      writeFileSync(nativePath, native);
      await testInfo.attach("pasted-table.docx", {
        path: nativePath,
        contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      });
      for (let pass = 0; pass < 2; pass++) {
        const reopened = openWordDocument(
          importDocx(native, { documentId: "table-copy-ui-native", unsupportedContent: "reject" }),
          { kind: "document" },
        );
        expect(reopened.snapshot().metadata.word!.tables).toHaveLength(
          fixture === "nested" ? 2 : 1,
        );
        expect(
          reopened.getNode(reopened.snapshot().metadata.word!.stories[0]!.rootNodeId)!.text,
        ).toContain("EDIT");
        native = exportDocx(reopened).bytes;
      }
      await page.evaluate(() => window.paginationHost.session!.undo());
      if (fixture !== "style") {
        await page.evaluate(() =>
          (window.paginationHost.view as import("@doffi/editor").DocumentTextEditor).execute({
            type: "word",
            action: "review",
          }),
        );
        await page.getByRole("button", { name: "拒绝全部显示的修订", exact: true }).click();
        await expect(page.frameLocator("iframe").locator("table")).toHaveCount(0);
        expect(
          await page.evaluate(() => window.paginationHost.session!.getNode("main-story")!.text),
        ).toBe("AB");
        await page.evaluate(() => window.paginationHost.session!.undo());
        await expect(page.frameLocator("iframe").locator("table")).toHaveCount(
          fixture === "nested" ? 2 : 1,
        );
      }
      await page.evaluate(() => window.paginationHost.session!.undo());
      expect(
        await page.evaluate(() => window.paginationHost.session!.getNode("main-story")!.text),
      ).toBe("AB");
      expect(
        await page.evaluate(() => window.paginationHost.session!.snapshot().metadata.word!.tables),
      ).toHaveLength(0);
    },
  );
for (const semantic of [false, true])
  for (const tracked of [false, true])
    test(
      "semantic block clipboard preserves paragraph structure, fields, continued input and one paste undo " +
        semantic +
        " " +
        tracked,
      async ({ page }) => {
        const source = createDocumentSession(
          openWordDocument(createWordDocument({ documentId: "block-copy-ui-source" }), {
            kind: "document",
          }),
        );
        source.commit(
          source.prepare({
            protocolVersion: 1,
            documentId: "block-copy-ui-source",
            baseRevision: 0,
            operations: [
              {
                op: "replaceText",
                nodeId: "paragraph-1",
                from: 0,
                to: 0,
                expectedText: "",
                text: "FIRST",
              },
              ...(semantic
                ? [
                    {
                      op: "insertWordField" as const,
                      nodeId: "paragraph-1",
                      offset: 5,
                      instruction: " PAGE ",
                      cachedText: "9",
                    },
                  ]
                : []),
              {
                op: "insertFragment",
                nodeId: "section-root-1",
                position: "end",
                html: "<p>SECOND</p>",
              },
            ],
          }),
        );
        await mount(page, source.serialize(), "edit");
        const copied = await page
          .frameLocator("iframe")
          .locator('[data-node-id="section-root-1"]')
          .evaluate((section) => {
            (section.closest("[data-office-editor]") as HTMLElement).focus();
            const range = document.createRange();
            range.selectNodeContents(section);
            const selection = document.getSelection()!;
            selection.removeAllRanges();
            selection.addRange(range);
            const data = new DataTransfer(),
              event = new ClipboardEvent("copy", { bubbles: true, cancelable: true });
            Object.defineProperty(event, "clipboardData", { value: data });
            section.dispatchEvent(event);
            return { html: data.getData("text/html"), text: data.getData("text/plain") };
          });
        const target = createDocumentSession(
          openWordDocument(createWordDocument({ documentId: "block-copy-ui-target" }), {
            kind: "document",
          }),
        );
        target.commit(
          target.prepare({
            protocolVersion: 1,
            documentId: "block-copy-ui-target",
            baseRevision: 0,
            operations: [
              {
                op: "replaceText",
                nodeId: "paragraph-1",
                from: 0,
                to: 0,
                expectedText: "",
                text: "ABCD",
              },
              {
                op: "setWordReview",
                nodeId: "document-root",
                trackChanges: tracked,
                author: "Receiver",
              },
            ],
          }),
        );
        await mount(page, target.serialize(), "edit");
        await page
          .frameLocator("iframe")
          .locator('[data-node-id="paragraph-1"]')
          .evaluate((paragraph, copied) => {
            (paragraph.closest("[data-office-editor]") as HTMLElement).focus();
            const text = document.createTreeWalker(paragraph, NodeFilter.SHOW_TEXT).nextNode()!,
              range = document.createRange();
            range.setStart(text, 1);
            range.setEnd(text, 3);
            const selection = document.getSelection()!;
            selection.removeAllRanges();
            selection.addRange(range);
            const data = new DataTransfer();
            data.setData("text/html", copied.html);
            data.setData("text/plain", copied.text);
            const event = new ClipboardEvent("paste", { bubbles: true, cancelable: true });
            Object.defineProperty(event, "clipboardData", { value: data });
            paragraph.dispatchEvent(event);
          }, copied);
        await expect
          .poll(() =>
            page.evaluate(
              () => window.paginationHost.session!.snapshot().metadata.word!.fields.length,
            ),
          )
          .toBe(semantic ? 1 : 0);
        expect(
          (await page.evaluate(() => window.paginationHost.session!.snapshot().html)).match(
            /<p\b/g,
          ),
        ).toHaveLength(4);
        await page.evaluate(() => window.paginationHost.view!.settled());
        await page.keyboard.insertText("EDIT");
        await expect
          .poll(() =>
            page.evaluate(() => window.paginationHost.session!.getNode("main-story")!.text),
          )
          .toContain("SECONDEDITD");
        expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
        await page.evaluate(() => window.paginationHost.session!.undo());
        await page.evaluate(() => window.paginationHost.session!.undo());
        expect(
          await page.evaluate(() => window.paginationHost.session!.getNode("main-story")!.text),
        ).toBe("ABCD");
        expect(
          await page.evaluate(
            () => window.paginationHost.session!.snapshot().metadata.word!.fields,
          ),
        ).toHaveLength(0);
      },
    );
for (const complete of [true, false])
  test(
    "semantic clipboard isolates bookmark targets and reports unresolved partial references " +
      complete,
    async ({ page }) => {
      const { wordBookmarkContentFixture } = await import(
        "../../packages/office/tests/word-bookmark-content-fixture.js"
      );
      await mount(page, wordBookmarkContentFixture().files, complete ? "preview" : "edit");
      const copied = await page
          .frameLocator("iframe")
          .locator('[data-node-id="paragraph-1"]')
          .evaluate((paragraph, complete) => {
            const range = document.createRange();
            range.selectNodeContents(paragraph);
            if (!complete)
              range.setStartAfter(paragraph.querySelector('[data-word-bookmark="end"]')!);
            const selection = document.getSelection()!;
            selection.removeAllRanges();
            selection.addRange(range);
            const data = new DataTransfer(),
              event = new ClipboardEvent("copy", { bubbles: true, cancelable: true });
            Object.defineProperty(event, "clipboardData", { value: data });
            paragraph.dispatchEvent(event);
            return { html: data.getData("text/html"), text: data.getData("text/plain") };
          }, complete),
        payload = JSON.parse(
          Buffer.from(
            /data-doffi-word-clipboard="([^"]+)"/.exec(copied.html)![1]!,
            "base64",
          ).toString("utf8"),
        );
      expect(payload.semantics.bookmarks).toHaveLength(complete ? 1 : 0);
      expect(payload.semantics.hyperlinks).toHaveLength(1);
      expect(payload.semantics.fields).toHaveLength(1);
      const target = createDocumentSession(
        openWordDocument(createWordDocument({ documentId: "bookmark-copy-ui" }), {
          kind: "document",
        }),
      );
      target.commit(
        target.prepare({
          protocolVersion: 1,
          documentId: "bookmark-copy-ui",
          baseRevision: 0,
          operations: [
            {
              op: "replaceText",
              nodeId: "paragraph-1",
              from: 0,
              to: 0,
              expectedText: "",
              text: "ORIGINAL",
            },
            {
              op: "insertWordBookmark",
              nodeId: "paragraph-1",
              from: 0,
              to: 8,
              expectedText: "ORIGINAL",
              name: "Anchor",
            },
          ],
        }),
      );
      await mount(page, target.serialize(), "edit");
      await page
        .frameLocator("iframe")
        .locator('[data-node-id="paragraph-1"]')
        .evaluate((paragraph, copied) => {
          (paragraph.closest("[data-office-editor]") as HTMLElement).focus();
          const range = document.createRange();
          range.selectNodeContents(paragraph);
          range.collapse(false);
          const selection = document.getSelection()!;
          selection.removeAllRanges();
          selection.addRange(range);
          const data = new DataTransfer();
          data.setData("text/html", copied.html);
          data.setData("text/plain", copied.text);
          const event = new ClipboardEvent("paste", { bubbles: true, cancelable: true });
          Object.defineProperty(event, "clipboardData", { value: data });
          paragraph.dispatchEvent(event);
        }, copied);
      await expect
        .poll(() =>
          page.evaluate(
            () => window.paginationHost.session!.snapshot().metadata.word!.fields.length,
          ),
        )
        .toBe(1);
      await page.evaluate(() => window.paginationHost.view!.settled());
      const word = await page.evaluate(
          () => window.paginationHost.session!.snapshot().metadata.word!,
        ),
        name = word.hyperlinks[0]!.bookmark!;
      expect(name).not.toBe("Anchor");
      expect(word.fields[0]!.instruction).toContain(name);
      expect(word.bookmarks.some((item) => item.name === name)).toBe(complete);
      expect(
        (await page.evaluate(() => window.paginationHost.layout())).diagnostics.some(
          (item) => item.code === "UnresolvedReference",
        ),
      ).toBe(!complete);
      if (complete)
        await expect(page.frameLocator("iframe").locator(".sdk-field-display").first()).toHaveText(
          "SOURCE",
        );
      expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
      await page.evaluate(() => window.paginationHost.session!.undo());
      expect(
        await page.evaluate(() => window.paginationHost.session!.getNode("paragraph-1")!.text),
      ).toBe("ORIGINAL");
      expect(
        await page.evaluate(() =>
          window.paginationHost
            .session!.snapshot()
            .metadata.word!.bookmarks.map((item) => item.name),
        ),
      ).toEqual(["Anchor"]);
    },
  );
for (const trackChanges of [false, true])
  test(
    "semantic clipboard copies fields, equations and controls across documents with fresh IDs and one undo " +
      trackChanges,
    async ({ page }) => {
      const source = createDocumentSession(
        openWordDocument(createWordDocument({ documentId: "semantic-source" }), {
          kind: "document",
        }),
      );
      source.commit(
        source.prepare({
          protocolVersion: 1,
          documentId: "semantic-source",
          baseRevision: 0,
          operations: [
            {
              op: "insertWordEquation",
              nodeId: "paragraph-1",
              offset: 0,
              math: { rootMathId: "m0", nodes: [{ mathId: "m0", kind: "text", text: "x" }] },
            },
            {
              op: "insertWordControl",
              nodeId: "paragraph-1",
              control: { kind: "checkBox", level: "inline", checked: true },
            },
            {
              op: "insertWordField",
              nodeId: "paragraph-1",
              offset: 0,
              instruction: " PAGE ",
              cachedText: "7",
            },
            {
              op: "insertFragment",
              nodeId: "section-root-1",
              position: "end",
              html: "<p>UNSELECTED-PRIVATE-CONTENT</p>",
            },
          ],
        }),
      );
      await mount(page, source.serialize(), "edit");
      const copied = await page
        .frameLocator("iframe")
        .locator('[data-node-id="paragraph-1"]')
        .evaluate((paragraph) => {
          (paragraph.closest("[data-office-editor]") as HTMLElement).focus();
          const range = document.createRange();
          range.selectNodeContents(paragraph);
          const selection = document.getSelection()!;
          selection.removeAllRanges();
          selection.addRange(range);
          const data = new DataTransfer(),
            event = new ClipboardEvent("copy", { bubbles: true, cancelable: true });
          Object.defineProperty(event, "clipboardData", { value: data });
          paragraph.dispatchEvent(event);
          return { html: data.getData("text/html"), text: data.getData("text/plain") };
        });
      const match = /data-doffi-word-clipboard="([^"]+)"/.exec(copied.html)!;
      const payload = JSON.parse(Buffer.from(match[1]!, "base64").toString("utf8"));
      expect(payload.version, copied.html.replace(/data-doffi-word-clipboard="[^"]+"/, "")).toBe(2);
      expect(payload.semantics.fields).toHaveLength(1);
      expect(payload.semantics.equations).toHaveLength(1);
      expect(payload.semantics.controls).toHaveLength(1);
      expect(JSON.stringify(payload)).not.toContain("UNSELECTED-PRIVATE-CONTENT");
      const target = createDocumentSession(
        openWordDocument(createWordDocument({ documentId: "semantic-target" }), {
          kind: "document",
        }),
      );
      target.commit(
        target.prepare({
          protocolVersion: 1,
          documentId: "semantic-target",
          baseRevision: 0,
          operations: [
            {
              op: "replaceText",
              nodeId: "paragraph-1",
              from: 0,
              to: 0,
              expectedText: "",
              text: "AB",
            },
            { op: "setWordReview", nodeId: "document-root", trackChanges, author: "Receiver" },
          ],
        }),
      );
      await mount(page, target.serialize(), "edit");
      const p = page.frameLocator("iframe").locator('[data-node-id="paragraph-1"]');
      await p.evaluate((paragraph, copied) => {
        (paragraph.closest("[data-office-editor]") as HTMLElement).focus();
        const walker = document.createTreeWalker(paragraph, NodeFilter.SHOW_TEXT),
          node = walker.nextNode()!;
        const range = document.createRange();
        range.setStart(node, 1);
        range.collapse(true);
        const selection = document.getSelection()!;
        selection.removeAllRanges();
        selection.addRange(range);
        const data = new DataTransfer();
        data.setData("text/html", copied.html);
        data.setData("text/plain", copied.text);
        const event = new ClipboardEvent("paste", { bubbles: true, cancelable: true });
        Object.defineProperty(event, "clipboardData", { value: data });
        paragraph.dispatchEvent(event);
      }, copied);
      await expect
        .poll(() =>
          page.evaluate(
            () => window.paginationHost.session!.snapshot().metadata.word!.fields.length,
          ),
        )
        .toBe(1);
      const snapshot = await page.evaluate(() => window.paginationHost.session!.snapshot());
      expect(snapshot.metadata.word!.equations).toHaveLength(1);
      expect(snapshot.metadata.word!.controls).toHaveLength(1);
      expect(snapshot.metadata.word!.fields[0]!.nodeId).not.toBe(
        payload.semantics.fields[0].nodeId,
      );
      expect(snapshot.html).not.toContain("data-doffi-");
      expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
      await page.evaluate(() => window.paginationHost.session!.undo());
      expect(
        await page.evaluate(() => window.paginationHost.session!.getNode("paragraph-1")!.text),
      ).toBe("AB");
      expect(
        await page.evaluate(() => window.paginationHost.session!.snapshot().metadata.word!.fields),
      ).toHaveLength(0);
      await page.evaluate(() => window.paginationHost.session!.redo());
      expect(
        await page.evaluate(() => window.paginationHost.session!.snapshot().metadata.word!.fields),
      ).toHaveLength(1);
      if (trackChanges) {
        await page.evaluate(() => {
          const session = window.paginationHost.session!,
            snapshot = session.snapshot();
          session.commit(
            session.prepare({
              protocolVersion: 1,
              documentId: snapshot.manifest.documentId,
              baseRevision: snapshot.manifest.revision,
              operations: [
                { op: "reviewWordRevisions", nodeId: "document-root", decision: "reject" },
              ],
            }),
          );
        });
        expect(
          await page.evaluate(() => window.paginationHost.session!.getNode("paragraph-1")!.text),
        ).toBe("AB");
        expect(
          await page.evaluate(
            () => window.paginationHost.session!.snapshot().metadata.word!.fields,
          ),
        ).toHaveLength(0);
      }
    },
  );

registerWordListClipboardTests(mount);
registerWordCrossClipboardTests(mount);
registerWordMoveClipboardTests(mount);
registerWordObjectClipboardTests(mount);
registerWordNoteClipboardTests(mount);
