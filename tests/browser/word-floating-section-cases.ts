import { writeFileSync } from "node:fs";
import { expect, test } from "@playwright/test";
import type { Page, TestInfo } from "@playwright/test";
import { openDocument } from "@doffi/core";
import type { DocumentPackage } from "@doffi/core";
import type { DocumentPreview } from "@doffi/render";
import { exportDocx, importDocx } from "@doffi/office";
import { floatingSectionFixture } from "../../packages/office/tests/floating-section-fixture.js";

type Mount = (page: Page, files: DocumentPackage, mode: "edit") => Promise<void>;
type FixtureKind = Parameters<typeof floatingSectionFixture>[0];

function openWord(files: DocumentPackage) {
  return openDocument(files, { kind: "document" });
}

function read(bytes: Uint8Array) {
  return importDocx(bytes, { documentId: "floating-sections", unsupportedContent: "reject" });
}

async function capture(page: Page, info: TestInfo) {
  const { entries, fonts, count } = await page.evaluate(() => {
    const host = window.paginationHost;
    return {
      entries: [...host.session!.serialize()].map(
        ([path, bytes]) => [path, [...bytes]] as [string, number[]],
      ),
      fonts: host.layout().fonts,
      count: host.layout().pages.length,
    };
  });
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
  await page.screenshot({ path: info.outputPath("floating-sections.png") });
  await page.evaluate((entries) => window.paginationHost.mount(entries, "preview"), entries);
  const print = await page.evaluate(() => {
    const document = (window.paginationHost.view as DocumentPreview).printDocument();
    return {
      html: document.documentElement.outerHTML,
      pages: document.querySelectorAll(".sdk-print-page").length,
    };
  });
  expect(print.pages).toBe(count);
  const model = openWord(new Map(entries.map(([path, bytes]) => [path, Uint8Array.from(bytes)])));
  const expected = model.snapshot().metadata,
    native = exportDocx(model).bytes;
  expect(model.snapshot().html).not.toMatch(
    /sdk-word-floating-table|sdk-word-center-exclusion|sdk-section-band/,
  );
  const drawingValues = (metadata: typeof expected) =>
    metadata.word!.drawings!.map(({ geometry, layout }) => ({ geometry, layout }));
  for (let pass = 0, bytes = native; pass < 2; pass++) {
    const reopened = openWord(read(bytes)),
      metadata = reopened.snapshot().metadata;
    expect(metadata.sections?.map(({ breakType, columns }) => ({ breakType, columns }))).toEqual(
      expected.sections?.map(({ breakType, columns }) => ({ breakType, columns })),
    );
    expect(drawingValues(metadata)).toEqual(drawingValues(expected));
    expect(metadata.word?.tables.length).toBe(expected.word?.tables.length);
    bytes = exportDocx(reopened).bytes;
  }
  writeFileSync(info.outputPath("floating-sections.docx"), native);
  writeFileSync(info.outputPath("font-report.json"), JSON.stringify(fonts, null, 2));
  writeFileSync(info.outputPath("floating-sections.print.html"), print.html);
}

async function paint(page: Page) {
  return page
    .frameLocator("iframe")
    .locator("body")
    .evaluate((body) => {
      const paper = [...body.querySelectorAll(".sdk-paper")].map((node) =>
        node.getBoundingClientRect(),
      );
      const box = (node: Element) => {
        const r = node.getBoundingClientRect(),
          index = paper.findIndex((p) => r.top >= p.top - 0.2 && r.top < p.bottom - 0.2),
          p = paper[index]!;
        return {
          index,
          x: (r.left - p.left) * 0.75,
          y: (r.top - p.top) * 0.75,
          bottom: (r.bottom - p.top) * 0.75,
        };
      };
      return {
        objects: [...body.querySelectorAll(".sdk-word-floating-page > [data-sdk-word-object]")].map(
          box,
        ),
        tables: [...body.querySelectorAll(".sdk-word-floating-table > table")]
          .filter((node) => node.getClientRects().length > 0)
          .map(box),
      };
    });
}

export function registerWordFloatingSectionTests(mount: Mount) {
  const open = (page: Page, kind: FixtureKind) =>
    mount(page, read(floatingSectionFixture(kind)), "edit");
  test("R18 actual unequal columns survive next-column and continuous section changes with mixed floats", async ({
    page,
  }, info) => {
    await open(page, "bands");
    const { objects, tables } = await paint(page);
    expect(objects).toHaveLength(3);
    expect(tables).toHaveLength(1);
    expect(objects.map((item) => item.index)).toEqual([0, 0, 0]);
    expect(objects[0]!.x).toBeCloseTo(66, 0);
    expect(objects[1]!.x).toBeCloseTo(252, 0);
    expect(objects[2]!.x).toBeCloseTo(180, 0);
    expect(tables[0]!.x).toBeCloseTo(242, 0);
    expect(tables[0]!.y).toBeGreaterThanOrEqual(objects[1]!.bottom - 0.1);
    expect(await page.evaluate(() => window.paginationHost.layout().pages.length)).toBe(1);
    const [first, second] = await page
      .frameLocator("iframe")
      .locator(".sdk-section-band")
      .evaluateAll((bands) =>
        bands.slice(0, 2).map((band, index) => {
          const paragraph = [...band.querySelectorAll("p")].find((p) =>
              p.textContent?.includes(index ? "SECOND" : "FIRST"),
            )!,
            walker = document.createTreeWalker(paragraph, NodeFilter.SHOW_TEXT);
          for (let text = walker.nextNode(); text; text = walker.nextNode()) {
            if (!text.textContent?.includes(index ? "SECOND" : "FIRST")) continue;
            const range = document.createRange();
            range.selectNodeContents(text);
            return range.getBoundingClientRect().top;
          }
          throw new Error("Missing section text");
        }),
      );
    expect(first! - second!).toBeGreaterThan(18);
    await capture(page, info);
  });

  test("R18 positioned floating rows continue across unequal columns and pages with live cell input", async ({
    page,
  }, info) => {
    await open(page, "rows");
    const frame = page.frameLocator("iframe"),
      rows = frame.locator(".sdk-word-floating-table > table > tbody > tr[data-node-id]");
    await expect(rows).toHaveCount(31);
    const layout = await page.evaluate(() => window.paginationHost.layout());
    expect(layout.pages.length).toBeGreaterThan(1);
    const positions = await rows.evaluateAll((rows) => {
      const pages = [...document.querySelectorAll(".sdk-paper")].map((p) =>
        p.getBoundingClientRect(),
      );
      return rows.map((row) => {
        const box = row.getBoundingClientRect(),
          index = pages.findIndex((p) => box.top >= p.top && box.top < p.bottom);
        return {
          index,
          x: (box.left - pages[index]!.left) * 0.75,
          bottom: (box.bottom - pages[index]!.top) * 0.75,
        };
      });
    });
    expect(positions[0]!.index).toBe(0);
    expect(positions[0]!.x).toBeCloseTo(242, 0);
    expect(positions[16]!.x).toBeCloseTo(56, 0);
    expect(positions.at(-1)!.index).toBeGreaterThan(0);
    const tableId = await frame
      .locator(".sdk-word-floating-table > table")
      .getAttribute("data-node-id");
    expect(layout.nodePages[tableId!]).toEqual([
      ...new Set(positions.map(({ index }) => layout.pages[index]!.pageId)),
    ]);
    expect(await frame.locator(".sdk-repeat-header").count()).toBeGreaterThan(0);
    const row = rows.nth(16);
    const { objects } = await paint(page);
    expect(objects[0]!.index).toBe(positions[16]!.index);
    expect(objects[0]!.x).toBeCloseTo(positions[16]!.x + 10, 0);
    expect(objects[1]!.index).toBe(positions.at(-1)!.index);
    expect(objects[1]!.y).toBeGreaterThanOrEqual(positions.at(-1)!.bottom - 0.2);
    const source = await page.evaluate(() => window.paginationHost.session!.snapshot().html);
    await row.locator("p").click({ position: { x: 3, y: 9 } });
    await page.keyboard.press(process.platform === "darwin" ? "Meta+ArrowRight" : "End");
    await page.keyboard.insertText("编辑");
    await expect(row).toContainText(/16\s*编辑/);
    const paragraphId = await row.locator("p").getAttribute("data-node-id");
    const entries = await page.evaluate((id) => {
      const session = window.paginationHost.session!;
      if (!session.getNode(id!)?.text.endsWith("ROW 16编辑"))
        throw new Error("Cell edit did not reach the exact source text.");
      return [...session.serialize()].map(([path, bytes]) => [path, [...bytes]] as const);
    }, paragraphId);
    const bytes = exportDocx(
      openWord(new Map(entries.map(([path, bytes]) => [path, Uint8Array.from(bytes)]))),
    ).bytes;
    writeFileSync(info.outputPath("floating-sections-edited.docx"), bytes);
    let reopened = openWord(read(bytes));
    expect(reopened.snapshot().html).toContain("编辑");
    reopened = openWord(read(exportDocx(reopened).bytes));
    expect(reopened.snapshot().html).toContain("编辑");
    await page.evaluate(() => window.paginationHost.session!.undo());
    await expect
      .poll(() => page.evaluate(() => window.paginationHost.session!.snapshot().html))
      .toBe(source);
    await capture(page, info);
  });

  test("R18 repeated header and footer position and avoid mixed floats per physical mirrored page", async ({
    page,
  }, info) => {
    await open(page, "stories");
    await page.evaluate(() => window.paginationHost.view!.settled());
    const measured = await page
      .frameLocator("iframe")
      .locator(".sdk-paper")
      .evaluateAll((pages) =>
        pages.map((page) => {
          const paper = page.getBoundingClientRect();
          const rect = (node: Element) => {
            const r = node.getBoundingClientRect();
            return {
              x: (r.left - paper.left) * 0.75,
              top: (r.top - paper.top) * 0.75,
              bottom: (r.bottom - paper.top) * 0.75,
            };
          };
          const shapes = [
              ...page.querySelectorAll(":scope > .sdk-page-story-paint > [data-sdk-word-object]"),
            ],
            tables = [
              ...page.querySelectorAll(
                ":scope > .sdk-page-story-paint > .sdk-word-floating-table > table",
              ),
            ],
            nested = shapes.find(
              (shape) =>
                (shape as HTMLElement).dataset.sdkWordStoryAnchor ===
                tables[0]!.querySelector<HTMLElement>("p")?.dataset.nodeId,
            )!,
            obstacle = shapes[0]!.getBoundingClientRect(),
            header = page.querySelector('.sdk-page-story[data-office-story="header"]')!,
            walker = document.createTreeWalker(header, NodeFilter.SHOW_TEXT);
          let glyphs = 0,
            overlapping = 0;
          for (let node = walker.nextNode(); node; node = walker.nextNode()) {
            for (let i = 0; i < node.textContent!.length; i++) {
              if (/\s/.test(node.textContent![i]!)) continue;
              const range = document.createRange();
              range.setStart(node, i);
              range.setEnd(node, i + 1);
              const glyph = range.getBoundingClientRect();
              if (glyph.bottom <= obstacle.top + 0.8 || glyph.top >= obstacle.bottom - 0.8)
                continue;
              glyphs++;
              if (glyph.right > obstacle.left + 0.8 && glyph.left < obstacle.right - 0.8)
                overlapping++;
            }
          }
          return {
            glyphs,
            overlapping,
            shapes: shapes.filter((shape) => shape !== nested).map(rect),
            tables: tables.map(rect),
            nested: rect(nested),
          };
        }),
      );
    expect(measured).toHaveLength(3);
    for (const [index, { glyphs, overlapping, shapes, tables, nested }] of measured.entries()) {
      expect(glyphs).toBeGreaterThan(0);
      expect(overlapping).toBe(0);
      expect(shapes).toHaveLength(2);
      expect(tables).toHaveLength(2);
      expect(nested.x).toBeCloseTo(tables[0]!.x + 10, 0);
      expect(nested.top).toBeGreaterThanOrEqual(tables[0]!.top - 0.1);
      expect(nested.top).toBeLessThan(tables[0]!.bottom);
      for (const [i, shape] of shapes.entries()) {
        expect(shape.x).toBeCloseTo(index === 1 ? 192 : 168, 0);
        expect(shape.top).toBeCloseTo(i === 0 ? 12 : 204, 0);
        expect(tables[i]!.x).toBeCloseTo(index === 1 ? 182 : 158, 0);
        expect(tables[i]!.top).toBeGreaterThanOrEqual(shape.bottom - 0.1);
      }
    }
    await capture(page, info);
  });

  test("R18 differing page and next-column geometries retain native sections and diagnose the preview page boundary", async ({
    page,
  }, info) => {
    await open(page, "geometry");
    const { pages, diagnostics } = await page.evaluate(() => window.paginationHost.layout());
    expect(pages.map((page) => page.setup.height)).toEqual([260, 290, 290]);
    expect(
      diagnostics
        .filter(({ code }) => code === "SectionGeometryChange")
        .map(({ message }) => message),
    ).toEqual([
      "A continuous section with different page geometry starts on a new page.",
      "A next-column section with different column geometry starts on a new page.",
    ]);
    expect((await paint(page)).objects.map(({ index }) => index)).toEqual([0, 1, 2]);
    await capture(page, info);
  });
}
