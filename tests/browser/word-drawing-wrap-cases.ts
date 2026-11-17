import { expect, test } from "@playwright/test";
import type { Page, TestInfo } from "@playwright/test";
import type { DocumentPackage } from "@doffi/core";
import { createDocumentSession, openDocument } from "@doffi/core";
import type { DocumentPreviewOptions } from "@doffi/editor";
import { exportDocx, importDocx } from "@doffi/office";
import { writeFileSync } from "node:fs";
import { drawingWrapFixture } from "../../packages/office/tests/drawing-wrap-fixture.js";
type Mount = (
  page: Page,
  files: DocumentPackage,
  mode?: "edit" | "preview" | "interaction",
  options?: DocumentPreviewOptions,
) => Promise<void>;
async function inspect(page: Page) {
  return page
    .frameLocator("iframe")
    .locator(".sdk-word-floating-layer")
    .evaluate((layer) => {
      const doc = layer.ownerDocument,
        frames = [...layer.querySelectorAll<HTMLElement>("[data-sdk-word-object]")],
        objects = frames.map((frame) => {
          const r = frame.getBoundingClientRect();
          return {
            left: r.left,
            right: r.right,
            top: r.top,
            bottom: r.bottom,
            id: frame.dataset.sdkWordObject!,
          };
        }),
        rects: { left: number; right: number; top: number; bottom: number; cell: string | null }[] =
          [];
      for (const p of doc.querySelectorAll<HTMLElement>('[data-office-story="main"] p')) {
        if (p.closest("[data-sdk-word-object],.sdk-paper")) continue;
        const walker = doc.createTreeWalker(p, NodeFilter.SHOW_TEXT);
        for (let node = walker.nextNode(); node; node = walker.nextNode()) {
          if (!node.textContent || node.parentElement?.closest("[data-sdk-word-object]")) continue;
          if (doc.defaultView!.getComputedStyle(node.parentElement!).visibility !== "visible")
            continue;
          const range = doc.createRange();
          range.selectNodeContents(node);
          for (const r of range.getClientRects())
            if (r.width > 0.5)
              rects.push({
                left: r.left,
                right: r.right,
                top: r.top,
                bottom: r.bottom,
                cell: p.closest("td,th")?.getAttribute("data-node-id") ?? null,
              });
        }
      }
      const cells = [...doc.querySelectorAll<HTMLElement>('[data-office-story="main"] td')].map(
        (cell) => {
          const r = cell.getBoundingClientRect();
          return {
            id: cell.dataset.nodeId,
            left: r.left,
            right: r.right,
            top: r.top,
            bottom: r.bottom,
          };
        },
      );
      return {
        objects,
        rects,
        cells,
        spaces: doc.querySelectorAll(".sdk-word-center-exclusion").length,
      };
    });
}
async function artifacts(page: Page, info: TestInfo) {
  const entries = await page.evaluate(() =>
    [...window.paginationHost.session!.serialize()].map(
      ([path, bytes]) => [path, [...bytes]] as const,
    ),
  );
  const files = new Map(entries.map(([path, bytes]) => [path, new Uint8Array(bytes)])),
    model = openDocument(files, { kind: "document" }),
    bytes = exportDocx(model).bytes;
  writeFileSync(info.outputPath("drawing-wrap.docx"), bytes);
  for (let pass = 0, native = bytes; pass < 2; pass++) {
    const reopened = openDocument(
      importDocx(native, { documentId: "wrap-browser-reopen", unsupportedContent: "reject" }),
      { kind: "document" },
    );
    expect(reopened.snapshot().metadata.word!.drawings).toHaveLength(
      model.snapshot().metadata.word!.drawings!.length,
    );
    native = exportDocx(reopened).bytes;
  }
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
  writeFileSync(
    info.outputPath("font-report.json"),
    JSON.stringify(await page.evaluate(() => window.paginationHost.layout().fonts), null, 2),
  );
  await page.screenshot({ path: info.outputPath("drawing-wrap.png") });
}
export function registerWordDrawingWrapTests(mount: Mount): void {
  for (const side of ["bothSides", "left", "right", "largest"] as const)
    test("R16 native square wrap uses the declared text side " + side, async ({ page }, info) => {
      const files = importDocx(drawingWrapFixture({ side }), {
        documentId: "wrap-side",
        unsupportedContent: "reject",
      });
      await mount(page, files, "edit");
      const seen = await inspect(page),
        object = seen.objects[0]!;
      writeFileSync(info.outputPath("wrap-layout.json"), JSON.stringify(seen, null, 2));
      const rects = seen.rects.filter(
          (r) => r.bottom > object.top + 4 && r.top < object.bottom - 4,
        ),
        left = rects.filter((r) => r.right <= object.left + 1),
        right = rects.filter((r) => r.left >= object.right - 1);
      expect(seen.spaces).toBeGreaterThan(0);
      expect(rects.filter((r) => r.left < object.right - 1 && r.right > object.left + 1)).toEqual(
        [],
      );
      for (const rect of left) expect(rect.right).toBeLessThanOrEqual(object.left - 3 / 0.75 + 0.6);
      for (const rect of right)
        expect(rect.left).toBeGreaterThanOrEqual(object.right + 4 / 0.75 - 0.6);
      if (side === "bothSides") {
        expect(left.length).toBeGreaterThan(0);
        expect(right.length).toBeGreaterThan(0);
      } else if (side === "right") {
        expect(left).toEqual([]);
        expect(right.length).toBeGreaterThan(0);
      } else {
        expect(right).toEqual([]);
        expect(left.length).toBeGreaterThan(0);
      }
      const pages = (await page.evaluate(() => window.paginationHost.layout())).pages.length;
      for (let i = 0; i < 2; i++) await page.evaluate(() => window.paginationHost.view!.settled());
      expect((await page.evaluate(() => window.paginationHost.layout())).pages).toHaveLength(pages);
      await artifacts(page, info);
    });
  for (const wrap of ["tight", "through"] as const)
    test(
      "R16 native polygon wrapping distinguishes internal openings " + wrap,
      async ({ page }, info) => {
        await mount(
          page,
          importDocx(drawingWrapFixture({ wrap }), {
            documentId: "wrap-contour",
            unsupportedContent: "reject",
          }),
          "edit",
        );
        const seen = await inspect(page),
          object = seen.objects[0]!,
          inner = seen.rects.filter(
            (r) =>
              r.top > object.top + (object.bottom - object.top) * 0.35 &&
              r.bottom < object.bottom - 3 &&
              r.left > object.left + (object.right - object.left) * 0.28 &&
              r.right < object.left + (object.right - object.left) * 0.72,
          );
        expect(inner.length)[wrap === "through" ? "toBeGreaterThan" : "toBe"](0);
        for (const r of seen.rects.filter(
          (r) => r.bottom > object.top + 3 && r.top < object.bottom - 1,
        )) {
          const entirelyBelowTop =
            r.top >= object.top + ((object.bottom - object.top) * 5000) / 21600 + 3 / 0.75;
          const inOpening =
            entirelyBelowTop &&
            wrap === "through" &&
            r.left >=
              object.left + ((object.right - object.left) * 5600) / 21600 + 4 / 0.75 - 0.6 &&
            r.right <=
              object.left + ((object.right - object.left) * 16000) / 21600 - 3 / 0.75 + 0.6;
          expect(
            inOpening ||
              r.right <= object.left - 3 / 0.75 + 0.6 ||
              r.left >= object.right + 4 / 0.75 - 0.6,
          ).toBe(true);
        }
        await artifacts(page, info);
      },
    );
  test("R16 diamond polygon and rotation change per-line boundaries", async ({ page }, info) => {
    await mount(
      page,
      importDocx(drawingWrapFixture({ wrap: "tight", polygon: "diamond", rotation: 15 }), {
        documentId: "wrap-diamond",
        unsupportedContent: "reject",
      }),
      "edit",
    );
    const seen = await inspect(page),
      object = seen.objects[0]!;
    expect(
      seen.rects.some(
        (r) =>
          r.top > object.top &&
          r.bottom < object.top + (object.bottom - object.top) * 0.35 &&
          r.right > object.left + 5 &&
          r.right < object.left + (object.right - object.left) * 0.4,
      ),
    ).toBe(true);
    await artifacts(page, info);
  });
  test("R16 top-and-bottom excludes the full paragraph width", async ({ page }, info) => {
    await mount(
      page,
      importDocx(drawingWrapFixture({ wrap: "topAndBottom" }), {
        documentId: "wrap-top-bottom",
        unsupportedContent: "reject",
      }),
      "edit",
    );
    const seen = await inspect(page),
      object = seen.objects[0]!;
    expect(
      seen.rects.filter((r) => r.bottom > object.top + 1 && r.top < object.bottom - 1),
    ).toEqual([]);
    expect(seen.rects.some((r) => r.top >= object.bottom)).toBe(true);
    await artifacts(page, info);
  });
  for (const allowOverlap of [true, false])
    for (const layout of ["pages", "continuous"] as const)
      test(
        "R16 floating collisions obey both native overlap flags " + allowOverlap + " " + layout,
        async ({ page }, info) => {
          await mount(
            page,
            importDocx(drawingWrapFixture({ wrap: "none", count: 3, allowOverlap }), {
              documentId: "wrap-collision",
              unsupportedContent: "reject",
            }),
            "edit",
            { layout },
          );
          const seen = await inspect(page),
            sorted = [...seen.objects].sort((a, b) => a.top - b.top);
          expect(sorted).toHaveLength(3);
          if (allowOverlap) expect(sorted[2]!.top - sorted[0]!.top).toBeLessThan(1);
          else
            for (let i = 1; i < sorted.length; i++)
              expect(sorted[i]!.top).toBeGreaterThanOrEqual(sorted[i - 1]!.bottom - 0.5);
          await artifacts(page, info);
        },
      );
  for (const layoutInCell of [true, false])
    test(
      "R16 floating cell layout controls containing region and row space " + layoutInCell,
      async ({ page }, info) => {
        await mount(
          page,
          importDocx(drawingWrapFixture({ inCell: true, layoutInCell }), {
            documentId: "wrap-cell",
            unsupportedContent: "reject",
          }),
          "edit",
        );
        const seen = await inspect(page),
          object = seen.objects[0]!,
          cell = seen.cells[0]!;
        if (layoutInCell) {
          expect(object.left).toBeGreaterThanOrEqual(cell.left - 1);
          expect(object.right).toBeLessThanOrEqual(cell.right + 1);
          expect(cell.bottom).toBeGreaterThanOrEqual(object.bottom - 1);
          expect(
            seen.rects.filter(
              (r) =>
                r.cell === cell.id &&
                r.right > object.left + 1 &&
                r.left < object.right - 1 &&
                r.bottom > object.top + 1 &&
                r.top < object.bottom - 1,
            ),
          ).toEqual([]);
          const neighbor = seen.rects.filter((r) => r.cell === seen.cells[1]!.id);
          expect(neighbor.some((r) => r.bottom > object.top && r.top < object.bottom)).toBe(true);
        } else expect(object.left).toBeGreaterThan(cell.left + 60);
        const frame = page.frameLocator("iframe"),
          p = frame.locator('[data-office-story="main"] td p').first(),
          id = await p.getAttribute("data-node-id"),
          before = await page.evaluate(
            (id) => window.paginationHost.session!.getNode(id!)!.text,
            id,
          );
        const point = await p.evaluate((p) => {
          const doc = p.ownerDocument,
            walker = doc.createTreeWalker(p, NodeFilter.SHOW_TEXT),
            box = p.getBoundingClientRect();
          let result: { x: number; y: number } | undefined;
          for (let node = walker.nextNode(); node; node = walker.nextNode()) {
            if (!node.textContent || node.parentElement?.closest("[data-sdk-word-object]"))
              continue;
            const range = doc.createRange();
            range.selectNodeContents(node);
            for (const rect of range.getClientRects())
              if (rect.width > 1)
                result = {
                  x: rect.left - box.left + Math.min(3, rect.width / 2),
                  y: (rect.top + rect.bottom) / 2 - box.top,
                };
          }
          return result!;
        });
        await p.click({ position: point });
        await page.keyboard.press("End");
        await page.keyboard.insertText("单元格输入");
        await expect
          .poll(() => page.evaluate((id) => window.paginationHost.session!.getNode(id!)!.text, id))
          .toContain("单元格输入");
        expect(
          (
            await page.evaluate((id) => window.paginationHost.session!.getNode(id!)!.text, id)
          ).replace("单元格输入", ""),
        ).toBe(before);
        await page.evaluate(() => window.paginationHost.session!.undo());
        await artifacts(page, info);
      },
    );
  test("R16 contour wrapping also measures the continuous view", async ({ page }, info) => {
    await mount(
      page,
      importDocx(drawingWrapFixture({ wrap: "through" }), {
        documentId: "wrap-continuous",
        unsupportedContent: "reject",
      }),
      "edit",
      { layout: "continuous" },
    );
    const seen = await inspect(page),
      object = seen.objects[0]!;
    expect(seen.spaces).toBeGreaterThan(0);
    expect(
      seen.rects.some(
        (r) =>
          r.left > object.left + (object.right - object.left) * 0.3 &&
          r.right < object.left + (object.right - object.left) * 0.7 &&
          r.top > object.top + (object.bottom - object.top) * 0.4 &&
          r.bottom < object.bottom,
      ),
    ).toBe(true);
    await artifacts(page, info);
  });
  test("R16 nested groups keep their geometry and text story while the root contour wraps", async ({
    page,
  }, info) => {
    const session = createDocumentSession(
        openDocument(
          importDocx(drawingWrapFixture({ wrap: "through", count: 2 }), {
            documentId: "wrap-group",
            unsupportedContent: "reject",
          }),
          { kind: "document" },
        ),
      ),
      apply = (operations: import("@doffi/core").DocumentOperation[]) =>
        session.commit(
          session.prepare({
            protocolVersion: 1,
            documentId: "wrap-group",
            baseRevision: session.snapshot().manifest.revision,
            operations,
          }),
        ),
      initial = session.snapshot().metadata.word!.drawings!,
      root = session.getNode(initial[0]!.nodeId)!.parentNodeId!,
      paragraph =
        session.getNode(root)!.tagName === "p" ? root : session.getNode(root)!.parentNodeId!;
    apply([{ op: "groupWordDrawings", nodeId: paragraph, nodeIds: initial.map((d) => d.nodeId) }]);
    const inner = session.snapshot().metadata.word!.drawings!.find((d) => d.kind === "group")!;
    apply([
      {
        op: "insertWordDrawing",
        nodeId: paragraph,
        offset: 0,
        geometry: { ...inner.geometry, x: 40, y: 15, width: 70, height: 30 },
        layout: {
          ...inner.layout,
          wrapPolygon: inner.layout.wrapPolygon?.map((point) => ({ ...point })),
        },
        object: { kind: "textBox", text: "组内文字" },
      },
    ]);
    const textBox = session.snapshot().metadata.word!.drawings!.find((d) => d.kind === "textBox")!;
    apply([
      { op: "groupWordDrawings", nodeId: paragraph, nodeIds: [inner.nodeId, textBox.nodeId] },
    ]);
    const source = session.snapshot();
    await mount(page, session.serialize(), "edit");
    const frame = page
      .frameLocator("iframe")
      .locator(".sdk-word-floating-layer [data-sdk-word-object]")
      .first();
    await expect(frame.locator(".sdk-word-drawing-text")).toContainText("组内文字");
    expect((await inspect(page)).spaces).toBeGreaterThan(0);
    expect(await page.evaluate(() => window.paginationHost.session!.snapshot().html)).toBe(
      source.html,
    );
    expect(
      await page.evaluate(() => window.paginationHost.session!.snapshot().metadata.word!.drawings),
    ).toEqual(source.metadata.word!.drawings);
    await artifacts(page, info);
  });
  for (const variant of ["right", "grid"] as const)
    test(
      "R16 through contours compose with alignment and character grids " + variant,
      async ({ page }, info) => {
        const session = createDocumentSession(
            openDocument(
              importDocx(drawingWrapFixture({ wrap: "through" }), {
                documentId: "wrap-combination",
                unsupportedContent: "reject",
              }),
              { kind: "document" },
            ),
          ),
          section = session.snapshot().metadata.sections![0]!,
          paragraph = session
            .getNode(section.rootNodeId)!
            .childNodeIds.find((id) => session.getNode(id)!.tagName === "p")!;
        session.commit(
          session.prepare({
            protocolVersion: 1,
            documentId: "wrap-combination",
            baseRevision: 0,
            operations:
              variant === "right"
                ? [
                    {
                      op: "setWordParagraph",
                      nodeId: paragraph,
                      properties: { "text-align": "right" },
                    },
                  ]
                : [
                    {
                      op: "setWordSection",
                      nodeId: section.rootNodeId,
                      changes: {
                        grid: { type: "linesAndChars", linePitch: 16, characterSpacing: 0 },
                      },
                    },
                  ],
          }),
        );
        const original = session.snapshot().html;
        await mount(page, session.serialize(), "edit");
        const seen = await inspect(page),
          object = seen.objects[0]!;
        expect(seen.spaces).toBeGreaterThan(0);
        for (const r of seen.rects.filter(
          (r) => r.top > object.top + 40 && r.bottom < object.bottom - 5,
        ))
          expect(
            r.right <= object.left + 1 ||
              r.left >= object.right - 1 ||
              (r.left >= object.left + ((object.right - object.left) * 5600) / 21600 - 1 &&
                r.right <= object.left + ((object.right - object.left) * 16000) / 21600 + 1),
          ).toBe(true);
        expect(await page.evaluate(() => window.paginationHost.session!.snapshot().html)).toBe(
          original,
        );
        await artifacts(page, info);
      },
    );
  test("R16 collision overflow is explicit while native positions stay unchanged", async ({
    page,
  }, info) => {
    const files = importDocx(drawingWrapFixture({ wrap: "none", count: 6, allowOverlap: false }), {
        documentId: "wrap-overflow",
        unsupportedContent: "reject",
      }),
      before = openDocument(files, { kind: "document" }).snapshot();
    await mount(page, files, "edit");
    expect(
      (await page.evaluate(() => window.paginationHost.layout())).diagnostics.some(
        (d) => d.code === "ContentOverflow",
      ),
    ).toBe(true);
    expect(
      await page.evaluate(() => window.paginationHost.session!.snapshot().metadata.word!.drawings),
    ).toEqual(before.metadata.word!.drawings);
    await artifacts(page, info);
  });
  test("R16 contour settings preserve exact input, undo and shared print", async ({
    page,
  }, info) => {
    const files = importDocx(drawingWrapFixture({ wrap: "tight" }), {
      documentId: "wrap-controls",
      unsupportedContent: "reject",
    });
    await mount(page, files, "edit");
    const original = await page.evaluate(() => window.paginationHost.session!.snapshot().html),
      object = page
        .frameLocator("iframe")
        .locator(".sdk-word-floating-layer [data-sdk-word-object]")
        .first();
    await object.click({ position: { x: 10, y: 10 } });
    await page.getByLabel("文字环绕", { exact: true }).selectOption("through");
    await page
      .getByLabel("环绕边界点（x,y；21600 为完整宽高）", { exact: true })
      .fill("10800,0;21600,10800;10800,21600;0,10800");
    await page.getByRole("button", { name: "应用对象位置和布局", exact: true }).click();
    expect(
      await page.evaluate(
        () => window.paginationHost.session!.snapshot().metadata.word!.drawings![0]!.layout,
      ),
    ).toMatchObject({
      wrap: "through",
      wrapPolygonEdited: true,
      wrapPolygon: [
        { x: 10800, y: 0 },
        { x: 21600, y: 10800 },
        { x: 10800, y: 21600 },
        { x: 0, y: 10800 },
      ],
    });
    await page.getByRole("button", { name: "关闭绘图设置", exact: true }).click();
    const p = page.frameLocator("iframe").locator('[data-office-story="main"] p').first(),
      id = await p.getAttribute("data-node-id"),
      before = await page.evaluate((id) => window.paginationHost.session!.getNode(id!)!.text, id);
    await p.click({ position: { x: 10, y: 12 } });
    await page.keyboard.press("End");
    await page.keyboard.insertText("实际输入");
    await expect
      .poll(() => page.evaluate((id) => window.paginationHost.session!.getNode(id!)!.text, id))
      .toContain("实际输入");
    expect(
      (await page.evaluate((id) => window.paginationHost.session!.getNode(id!)!.text, id)).replace(
        "实际输入",
        "",
      ),
    ).toBe(before);
    await page.evaluate(() => window.paginationHost.session!.undo());
    expect(await page.evaluate(() => window.paginationHost.session!.snapshot().html)).toBe(
      original,
    );
    await artifacts(page, info);
    const entries = await page.evaluate(() =>
      [...window.paginationHost.session!.serialize()].map(
        ([path, bytes]) => [path, [...bytes]] as [string, number[]],
      ),
    );
    await page.evaluate((entries) => window.paginationHost.mount(entries, "preview"), entries);
    const print = await page.evaluate(() => {
      const output = (
        window.paginationHost.view as import("@doffi/editor").DocumentPreview
      ).printDocument();
      return {
        html: output.documentElement.outerHTML,
        spaces: output.querySelectorAll(".sdk-word-center-exclusion").length,
      };
    });
    expect(print.spaces).toBeGreaterThan(0);
    writeFileSync(info.outputPath("drawing-wrap-print.html"), print.html);
  });
}
