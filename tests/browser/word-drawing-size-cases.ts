import { expect, test } from "@playwright/test";
import type { Page, TestInfo, Locator } from "@playwright/test";
import { createWordDocument, createDocumentSession, openDocument } from "@doffi/core";
import type { DocumentOperation, DocumentPackage } from "@doffi/core";
import type { DocumentPreviewOptions } from "@doffi/editor";
import { importDocx, exportDocx } from "@doffi/office";
import { writeFileSync } from "node:fs";
import {
  drawingSizeFixture,
  mirroredDrawingSizeFixture,
} from "../../packages/office/tests/drawing-size-fixture.js";
import { centerWrapFixture } from "../../packages/office/tests/center-wrap-fixture.js";
type Mount = (
  page: Page,
  files: DocumentPackage,
  mode?: "edit" | "preview" | "interaction",
  options?: DocumentPreviewOptions,
) => Promise<void>;
const object = (page: Page) =>
  page
    .frameLocator("iframe")
    .locator(".sdk-word-floating-layer > .sdk-word-floating-page > [data-sdk-word-object]")
    .first();
async function dimensions(frame: Locator) {
  return frame.evaluate((frame) => {
    const r = frame.getBoundingClientRect();
    return {
      width: Number(frame.dataset.sdkWordWidth),
      height: Number(frame.dataset.sdkWordHeight),
      left: r.left,
      top: r.top,
      right: r.right,
      bottom: r.bottom,
      viewBox: frame.querySelector(":scope > svg")?.getAttribute("viewBox"),
    };
  });
}
async function apply(page: Page, operations: DocumentOperation[]) {
  await page.evaluate((operations) => {
    const s = window.paginationHost.session!,
      snap = s.snapshot();
    s.commit(
      s.prepare({
        protocolVersion: 1,
        documentId: snap.manifest.documentId,
        baseRevision: snap.manifest.revision,
        operations,
      }),
    );
  }, operations);
  await page.evaluate(() => window.paginationHost.view!.settled());
}
async function artifacts(page: Page, info: TestInfo) {
  const entries = await page.evaluate(() =>
      [...window.paginationHost.session!.serialize()].map(
        ([path, bytes]) => [path, [...bytes]] as const,
      ),
    ),
    model = openDocument(new Map(entries.map(([path, bytes]) => [path, new Uint8Array(bytes)])), {
      kind: "document",
    }),
    expected = model.snapshot().metadata.word!.drawings!,
    bytes = exportDocx(model).bytes;
  writeFileSync(info.outputPath("drawing-size.docx"), bytes);
  for (let pass = 0, native = bytes; pass < 2; pass++) {
    const reopened = openDocument(
      importDocx(native, { documentId: "size-browser-reopened", unsupportedContent: "reject" }),
      { kind: "document" },
    );
    const sizes = (drawings: typeof expected) =>
      drawings
        .map((d) =>
          JSON.stringify([d.kind, d.layout.relativeWidth ?? null, d.layout.relativeHeight ?? null]),
        )
        .sort();
    expect(sizes(reopened.snapshot().metadata.word!.drawings!)).toEqual(sizes(expected));
    native = exportDocx(reopened).bytes;
  }
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
  writeFileSync(
    info.outputPath("font-report.json"),
    JSON.stringify(await page.evaluate(() => window.paginationHost.layout().fonts), null, 2),
  );
  await page.screenshot({ path: info.outputPath("drawing-size.png") });
}
function kindFiles(kind: "textBox" | "group" | "chart" | "image"): DocumentPackage {
  const session = createDocumentSession(
      openDocument(
        kind === "image"
          ? importDocx(centerWrapFixture("image", 4), {
              documentId: "size-kind",
              unsupportedContent: "reject",
            })
          : createWordDocument({
              documentId: "size-kind",
              pageSetup: {
                unit: "pt",
                width: 420,
                height: 500,
                margins: { top: 36, left: 36, right: 36, bottom: 36 },
              },
            }),
        { kind: "document" },
      ),
    ),
    apply = (operations: DocumentOperation[]) =>
      session.commit(
        session.prepare({
          protocolVersion: 1,
          documentId: "size-kind",
          baseRevision: session.snapshot().manifest.revision,
          operations,
        }),
      ),
    geometry = { x: 0, y: 12, width: 120, height: 75, rotation: 0, flipH: false, flipV: false },
    layout = { mode: "floating" as const, wrap: "square" as const };
  if (kind !== "image") {
    apply([
      {
        op: "replaceText",
        nodeId: "paragraph-1",
        from: 0,
        to: 0,
        expectedText: "",
        text: "正文保持准确 UTF-16 😀 𠀀\u{e0100} 文本。".repeat(12),
      },
      {
        op: "insertWordDrawing",
        nodeId: "paragraph-1",
        offset: 0,
        geometry,
        layout,
        object:
          kind === "chart"
            ? {
                kind: "chart",
                chart: {
                  kind: "column",
                  title: "相对图表",
                  legend: "none",
                  minimum: 0,
                  maximum: 4,
                  majorUnit: 2,
                  showLabels: false,
                  categories: ["A", "B"],
                  series: [{ seriesId: "s1", name: "数据", values: [2, 4] }],
                },
              }
            : { kind: "textBox", text: "独立对象文字 ABC 中文。".repeat(6) },
      },
    ]);
    if (kind === "group") {
      apply([
        {
          op: "insertWordDrawing",
          nodeId: "paragraph-1",
          offset: 0,
          geometry: { ...geometry, x: 130 },
          layout,
          object: { kind: "shape", preset: "ellipse", text: "组合内文字" },
        },
      ]);
      apply([
        {
          op: "groupWordDrawings",
          nodeId: "paragraph-1",
          nodeIds: session.snapshot().metadata.word!.drawings!.map((d) => d.nodeId),
        },
      ]);
    }
  }
  const root = session
    .snapshot()
    .metadata.word!.drawings!.find((d) => (kind === "group" ? d.kind === "group" : true))!;
  apply([
    {
      op: "setWordDrawingLayout",
      nodeId: root.nodeId,
      layout: {
        ...root.layout,
        wrapPolygon: root.layout.wrapPolygon?.map((p) => ({ ...p })),
        relativeWidth: { relativeFrom: "page", percent: 50 },
        relativeHeight: { relativeFrom: "margin", percent: 25 },
      },
    },
  ]);
  return session.serialize();
}
export function registerWordDrawingSizeTests(mount: Mount): void {
  for (const layout of ["pages", "continuous"] as const)
    test(
      "R17 relative native sizes reflow on page changes without rewriting source extents " + layout,
      async ({ page }, info) => {
        const files = importDocx(drawingSizeFixture(), {
            documentId: "size-page",
            unsupportedContent: "reject",
          }),
          before = openDocument(files, { kind: "document" }).snapshot();
        await mount(page, files, "edit", { layout });
        expect(await dimensions(object(page))).toMatchObject({
          width: 210,
          height: 107,
          viewBox: "0 0 210 107",
        });
        expect(await page.evaluate(() => window.paginationHost.session!.snapshot().html)).toBe(
          before.html,
        );
        await apply(page, [
          {
            op: "setWordSection",
            nodeId: before.metadata.sections![0]!.rootNodeId,
            changes: { pageSetup: { ...before.metadata.pageSetup!, width: 520, height: 600 } },
          },
        ]);
        expect(await dimensions(object(page))).toMatchObject({
          width: 260,
          height: 132,
          viewBox: "0 0 260 132",
        });
        expect(
          await page.evaluate(
            () => window.paginationHost.session!.snapshot().metadata.word!.drawings![0]!.geometry,
          ),
        ).toEqual(before.metadata.word!.drawings![0]!.geometry);
        await page.evaluate(() => window.paginationHost.session!.undo());
        await page.evaluate(() => window.paginationHost.view!.settled());
        expect(await dimensions(object(page))).toMatchObject({ width: 210, height: 107 });
        await artifacts(page, info);
      },
    );
  for (const from of ["leftMargin", "rightMargin", "insideMargin", "outsideMargin"] as const)
    test(
      "R17 native margin percentages resolve per physical mirrored page " + from,
      async ({ page }, info) => {
        await mount(
          page,
          importDocx(mirroredDrawingSizeFixture(from), {
            documentId: "size-mirrored",
            unsupportedContent: "reject",
          }),
          "edit",
        );
        const measured = await page
          .frameLocator("iframe")
          .locator(".sdk-word-floating-layer > .sdk-word-floating-page > [data-sdk-word-object]")
          .evaluateAll((nodes) =>
            nodes.map((node) => Number((node as HTMLElement).dataset.sdkWordWidth)),
          );
        writeFileSync(
          info.outputPath("size-pages.json"),
          JSON.stringify(
            {
              measured,
              layout: await page.evaluate(() => window.paginationHost.layout()),
              mirror: await page.evaluate(
                () => window.paginationHost.session!.snapshot().metadata.word!.mirrorMargins,
              ),
            },
            null,
            2,
          ),
        );
        expect(measured).toEqual(
          from === "leftMargin"
            ? [36, 72]
            : from === "rightMargin"
              ? [72, 36]
              : from === "insideMargin"
                ? [36, 36]
                : [72, 72],
        );
        await artifacts(page, info);
      },
    );
  for (const wrap of ["square", "through"] as const)
    test(
      "R17 measured dimensions govern polygon and rectangular text exclusion " + wrap,
      async ({ page }, info) => {
        await mount(
          page,
          importDocx(drawingSizeFixture({ wrap }), {
            documentId: "size-wrap",
            unsupportedContent: "reject",
          }),
          "edit",
        );
        const measured = await object(page).evaluate((frame) => {
          const doc = frame.ownerDocument,
            b = frame.getBoundingClientRect(),
            rects: { left: number; right: number; top: number; bottom: number }[] = [];
          for (const p of doc.querySelectorAll('[data-office-story="main"] p')) {
            const w = doc.createTreeWalker(p, NodeFilter.SHOW_TEXT);
            for (let n = w.nextNode(); n; n = w.nextNode()) {
              if (n.parentElement?.closest("[data-sdk-word-object]")) continue;
              const r = doc.createRange();
              r.selectNodeContents(n);
              for (const box of r.getClientRects())
                if (box.width > 0.5 && box.bottom > b.top + 3 && box.top < b.bottom - 3)
                  rects.push({
                    left: box.left,
                    right: box.right,
                    top: box.top,
                    bottom: box.bottom,
                  });
            }
          }
          return { left: b.left, right: b.right, top: b.top, bottom: b.bottom, rects };
        });
        if (wrap === "square")
          expect(
            measured.rects.filter(
              (r) => r.right > measured.left + 1 && r.left < measured.right - 1,
            ),
          ).toEqual([]);
        else
          expect(
            measured.rects.some(
              (r) =>
                r.left > measured.left + (measured.right - measured.left) * 0.28 &&
                r.right < measured.left + (measured.right - measured.left) * 0.72 &&
                r.top > measured.top + (measured.bottom - measured.top) * 0.4,
            ),
          ).toBe(true);
        await artifacts(page, info);
      },
    );
  for (const kind of ["image", "textBox", "group", "chart"] as const)
    test(
      "R17 relative paint recomposes physical details and independent stories " + kind,
      async ({ page }, info) => {
        const files = kindFiles(kind),
          before = openDocument(files, { kind: "document" }).snapshot();
        await mount(page, files, "edit");
        const frame = object(page);
        expect(await dimensions(frame)).toMatchObject({ width: 210, height: 107 });
        if (kind === "group") {
          await expect(frame).toContainText("独立对象文字");
          expect(
            await frame
              .locator(":scope > span")
              .first()
              .evaluate((node) => getComputedStyle(node).transform),
          ).not.toBe("none");
        } else expect((await dimensions(frame)).viewBox).toBe("0 0 210 107");
        if (kind === "textBox") {
          const text = frame.locator(".sdk-word-drawing-text");
          await expect(text).toContainText("独立对象文字");
          const insets = await text.evaluate((n) => ({
            width: getComputedStyle(n).width,
            padding: getComputedStyle(n).paddingLeft,
            font: getComputedStyle(n).fontSize,
          }));
          expect(parseFloat(insets.width) * 0.75).toBeCloseTo(210);
          expect(parseFloat(insets.padding) * 0.75).toBeCloseTo(7.2);
          expect(parseFloat(insets.font)).toBe(16);
        }
        if (kind === "chart") await expect(frame.locator("[data-sdk-word-chart]")).toBeVisible();
        if (kind === "image")
          expect(
            await frame.locator("img").evaluate((n) => n.getBoundingClientRect().width * 0.75),
          ).toBeCloseTo(210, 1);
        expect(await page.evaluate(() => window.paginationHost.session!.snapshot().html)).toBe(
          before.html,
        );
        await artifacts(page, info);
      },
    );
  test("R17 size controls switch axes independently and direct resize bakes the displayed size with one undo", async ({
    page,
  }, info) => {
    await mount(
      page,
      importDocx(drawingSizeFixture(), {
        documentId: "size-controls",
        unsupportedContent: "reject",
      }),
      "edit",
    );
    let frame = object(page);
    await frame.click({ position: { x: 10, y: 10 } });
    await expect(page.getByLabel("对象宽度（磅）", { exact: true })).toHaveValue("210");
    await page.getByLabel("相对宽度参照", { exact: true }).selectOption("margin");
    await page.getByLabel("相对宽度（%，0 使用固定尺寸）", { exact: true }).fill("25");
    await page.getByLabel("对象高度方式", { exact: true }).selectOption("absolute");
    await page.getByLabel("对象高度（磅）", { exact: true }).fill("80");
    await page.getByRole("button", { name: "应用对象位置和布局", exact: true }).click();
    expect(await dimensions(object(page))).toMatchObject({ width: 87, height: 80 });
    const before = await page.evaluate(() => window.paginationHost.session!.snapshot());
    expect(before.metadata.word!.drawings![0]!.layout.relativeHeight).toBeUndefined();
    await page.getByRole("button", { name: "关闭绘图设置", exact: true }).click();
    frame = object(page);
    const r = (await frame.boundingBox())!;
    await page.mouse.move(r.x + r.width - 3, r.y + r.height - 3);
    await page.mouse.down();
    await page.mouse.move(r.x + r.width + 13, r.y + r.height + 13, { steps: 3 });
    await page.mouse.up();
    const resized = await page.evaluate(
      () => window.paginationHost.session!.snapshot().metadata.word!.drawings![0]!,
    );
    expect(resized.geometry.width).toBeCloseTo(99);
    expect(resized.geometry.height).toBeCloseTo(92);
    expect(resized.layout.relativeWidth).toBeUndefined();
    await page.getByRole("button", { name: "关闭绘图设置", exact: true }).click();
    await page.evaluate(() => window.paginationHost.session!.undo());
    expect(await page.evaluate(() => window.paginationHost.session!.snapshot().metadata)).toEqual(
      before.metadata,
    );
    await page.evaluate(() => window.paginationHost.view!.settled());
    expect(await dimensions(object(page))).toMatchObject({ width: 87, height: 80 });
    await artifacts(page, info);
  });
  test("R17 relative objects retain precise surrounding text input, undo and shared print", async ({
    page,
  }, info) => {
    await mount(
      page,
      importDocx(drawingSizeFixture({ wrap: "square" }), {
        documentId: "size-input",
        unsupportedContent: "reject",
      }),
      "edit",
    );
    const p = page.frameLocator("iframe").locator('[data-office-story="main"] p').first(),
      id = await p.getAttribute("data-node-id"),
      before = await page.evaluate((id) => window.paginationHost.session!.getNode(id!)!.text, id);
    await p.click({ position: { x: 8, y: 12 } });
    await page.keyboard.press("End");
    await page.keyboard.insertText("相对输入");
    await expect
      .poll(() => page.evaluate((id) => window.paginationHost.session!.getNode(id!)!.text, id))
      .toContain("相对输入");
    expect(
      (await page.evaluate((id) => window.paginationHost.session!.getNode(id!)!.text, id)).replace(
        "相对输入",
        "",
      ),
    ).toBe(before);
    await page.evaluate(() => window.paginationHost.session!.undo());
    await artifacts(page, info);
    const entries = await page.evaluate(() =>
      [...window.paginationHost.session!.serialize()].map(
        ([path, bytes]) => [path, [...bytes]] as [string, number[]],
      ),
    );
    await page.evaluate((entries) => window.paginationHost.mount(entries, "preview"), entries);
    const print = await page.evaluate(() => {
      const d = (
        window.paginationHost.view as import("@doffi/editor").DocumentPreview
      ).printDocument();
      return {
        html: d.documentElement.outerHTML,
        spaces: d.querySelectorAll(".sdk-word-center-exclusion").length,
      };
    });
    expect(print.spaces).toBeGreaterThan(0);
    expect(print.html).toContain("0 0 210 107");
    writeFileSync(info.outputPath("drawing-size-print.html"), print.html);
  });
  test("R17 relative collision bounds use the measured height", async ({ page }, info) => {
    const session = createDocumentSession(
        openDocument(
          importDocx(drawingSizeFixture({ count: 2 }), {
            documentId: "size-collision",
            unsupportedContent: "reject",
          }),
          { kind: "document" },
        ),
      ),
      snap = session.snapshot();
    session.commit(
      session.prepare({
        protocolVersion: 1,
        documentId: "size-collision",
        baseRevision: 0,
        operations: snap.metadata.word!.drawings!.map((d) => ({
          op: "setWordDrawingLayout" as const,
          nodeId: d.nodeId,
          layout: {
            ...d.layout,
            wrapPolygon: d.layout.wrapPolygon?.map((p) => ({ ...p })),
            allowOverlap: false,
          },
        })),
      }),
    );
    await mount(page, session.serialize(), "edit");
    const rects = await page
      .frameLocator("iframe")
      .locator(".sdk-word-floating-layer > .sdk-word-floating-page > [data-sdk-word-object]")
      .evaluateAll((nodes) =>
        nodes
          .map((n) => {
            const r = n.getBoundingClientRect();
            return { top: r.top, bottom: r.bottom };
          })
          .sort((a, b) => a.top - b.top),
      );
    expect(rects).toHaveLength(2);
    expect(rects[1]!.top).toBeGreaterThanOrEqual(rects[0]!.bottom - 0.5);
    await artifacts(page, info);
  });
  test("R17 bound connectors follow relative target endpoints in the derived canvas", async ({
    page,
  }, info) => {
    const session = createDocumentSession(
        openDocument(
          createWordDocument({
            documentId: "size-connectors",
            pageSetup: {
              unit: "pt",
              width: 420,
              height: 500,
              margins: { top: 36, bottom: 36, left: 36, right: 36 },
            },
          }),
          { kind: "document" },
        ),
      ),
      apply = (operations: DocumentOperation[]) =>
        session.commit(
          session.prepare({
            protocolVersion: 1,
            documentId: "size-connectors",
            baseRevision: session.snapshot().manifest.revision,
            operations,
          }),
        ),
      g = { x: 10, y: 12, width: 160, height: 100, rotation: 0, flipH: false, flipV: false };
    apply([
      {
        op: "insertWordDrawing",
        nodeId: "paragraph-1",
        offset: 0,
        geometry: g,
        layout: {
          mode: "floating",
          wrap: "none",
          relativeWidth: { relativeFrom: "page", percent: 50 },
        },
        object: { kind: "shape", preset: "rect" },
      },
      {
        op: "insertWordDrawing",
        nodeId: "paragraph-1",
        offset: 0,
        geometry: { ...g, x: 300, width: 60 },
        layout: { mode: "floating", wrap: "none" },
        object: { kind: "shape", preset: "rect" },
      },
    ]);
    const [a, b] = session.snapshot().metadata.word!.drawings!;
    apply([
      {
        op: "insertWordDrawing",
        nodeId: "paragraph-1",
        offset: 0,
        geometry: g,
        layout: { mode: "floating", wrap: "none" },
        object: {
          kind: "connector",
          connector: {
            routing: "straight",
            startNodeId: a!.nodeId,
            endNodeId: b!.nodeId,
            startSite: 3,
            endSite: 1,
          },
        },
      },
    ]);
    const connector = session
      .snapshot()
      .metadata.word!.drawings!.find((d) => d.kind === "connector")!;
    await mount(page, session.serialize(), "edit");
    const frame = page
      .frameLocator("iframe")
      .locator('.sdk-word-floating-layer [data-sdk-word-object="' + connector.nodeId + '"]');
    expect((await dimensions(frame)).width).toBeCloseTo(80);
    expect(connector.geometry.width).toBeCloseTo(130);
    await artifacts(page, info);
  });
  test("R17 ungrouping a relative canvas materializes the displayed extent with one undo", async ({
    page,
  }, info) => {
    await mount(page, kindFiles("group"), "edit");
    const before = await page.evaluate(() => window.paginationHost.session!.snapshot());
    await object(page).click({ position: { x: 10, y: 10 } });
    await page.getByRole("button", { name: "取消组合", exact: true }).click();
    const ungrouped = await page.evaluate(
      () => window.paginationHost.session!.snapshot().metadata.word!.drawings!,
    );
    expect(ungrouped).toHaveLength(2);
    expect(ungrouped.every((d) => !d.layout.relativeWidth && !d.layout.relativeHeight)).toBe(true);
    for (const drawing of ungrouped) expect(drawing.geometry.height).toBeCloseTo(107);
    await page.evaluate(() => window.paginationHost.session!.undo());
    expect(await page.evaluate(() => window.paginationHost.session!.snapshot().metadata)).toEqual(
      before.metadata,
    );
    await artifacts(page, info);
  });
  test("R17 moving a relative object keeps both percentages", async ({ page }, info) => {
    await mount(
      page,
      importDocx(drawingSizeFixture(), { documentId: "size-move", unsupportedContent: "reject" }),
      "edit",
    );
    const before = await page.evaluate(() => window.paginationHost.session!.snapshot());
    const rect = (await object(page).boundingBox())!;
    await page.mouse.move(rect.x + 15, rect.y + 15);
    await page.mouse.down();
    await page.mouse.move(rect.x + 31, rect.y + 23, { steps: 3 });
    await page.mouse.up();
    const after = await page.evaluate(
      () => window.paginationHost.session!.snapshot().metadata.word!.drawings![0]!.layout,
    );
    expect(after.relativeWidth).toEqual(before.metadata.word!.drawings![0]!.layout.relativeWidth);
    expect(after.relativeHeight).toEqual(before.metadata.word!.drawings![0]!.layout.relativeHeight);
    expect(after.alignH).toBeUndefined();
    await page.getByRole("button", { name: "关闭绘图设置", exact: true }).click();
    await page.evaluate(() => window.paginationHost.session!.undo());
    expect(await page.evaluate(() => window.paginationHost.session!.snapshot().metadata)).toEqual(
      before.metadata,
    );
    await artifacts(page, info);
  });
  test("R17 relative header objects resize in each physical page instance", async ({
    page,
  }, info) => {
    const setup = {
        unit: "pt" as const,
        width: 420,
        height: 300,
        margins: { top: 54, bottom: 36, left: 36, right: 72 },
      },
      session = createDocumentSession(
        openDocument(createWordDocument({ documentId: "size-header", pageSetup: setup }), {
          kind: "document",
        }),
      ),
      apply = (operations: DocumentOperation[]) =>
        session.commit(
          session.prepare({
            protocolVersion: 1,
            documentId: "size-header",
            baseRevision: session.snapshot().manifest.revision,
            operations,
          }),
        );
    apply([
      {
        op: "replaceText",
        nodeId: "paragraph-1",
        from: 0,
        to: 0,
        expectedText: "",
        text: "多页正文与相对页眉。".repeat(160),
      },
      { op: "setWordOptions", nodeId: "document-root", mirrorMargins: true },
      {
        op: "insertWordStory",
        nodeId: "section-root-1",
        kind: "header",
        variant: "default",
        html: "<p>页眉</p>",
      },
    ]);
    const header = session
        .snapshot()
        .metadata.word!.stories.find((story) => story.kind === "header")!,
      p = session.getNode(header.rootNodeId)!.childNodeIds[0]!;
    apply([
      {
        op: "insertWordDrawing",
        nodeId: p,
        offset: 0,
        geometry: { x: 0, y: 0, width: 120, height: 75, rotation: 0, flipH: false, flipV: false },
        layout: {
          mode: "floating",
          wrap: "square",
          relativeWidth: { relativeFrom: "leftMargin", percent: 100 },
          relativeHeight: { relativeFrom: "topMargin", percent: 50 },
        },
        object: { kind: "shape", preset: "rect" },
      },
    ]);
    const before = session.snapshot();
    await mount(page, session.serialize(), "edit");
    const widths = await page
      .frameLocator("iframe")
      .locator(".sdk-paper [data-sdk-word-object]")
      .evaluateAll((nodes) =>
        nodes.map((n) => ({
          page: n.closest<HTMLElement>(".sdk-paper")!.dataset.pageId,
          width: Number((n as HTMLElement).dataset.sdkWordWidth),
          height: Number((n as HTMLElement).dataset.sdkWordHeight),
        })),
      );
    expect(widths.length).toBeGreaterThan(1);
    for (const [index, instance] of widths.entries()) {
      expect(instance.width).toBe(index % 2 ? 72 : 36);
      expect(instance.height).toBe(27);
    }
    expect(await page.evaluate(() => window.paginationHost.session!.snapshot().html)).toBe(
      before.html,
    );
    await artifacts(page, info);
  });
  test("R17 relative objects in cells use page size references while cell layout constrains placement", async ({
    page,
  }, info) => {
    await mount(
      page,
      importDocx(
        drawingSizeFixture({
          inCell: true,
          width: "20000",
          heightFrom: "page",
          height: "20000",
          wrap: "square",
        }),
        { documentId: "size-cell", unsupportedContent: "reject" },
      ),
      "edit",
    );
    expect(await dimensions(object(page))).toMatchObject({ width: 84, height: 100 });
    const state = await object(page).evaluate((frame) => {
      const id = frame.dataset.sdkWordWrapCell!,
        cell = frame.ownerDocument.querySelector<HTMLElement>('td[data-node-id="' + id + '"]')!,
        a = frame.getBoundingClientRect(),
        b = cell.getBoundingClientRect();
      return {
        left: a.left,
        right: a.right,
        bottom: a.bottom,
        cellLeft: b.left,
        cellRight: b.right,
        cellBottom: b.bottom,
      };
    });
    expect(state.left).toBeGreaterThanOrEqual(state.cellLeft);
    expect(state.right).toBeLessThanOrEqual(state.cellRight);
    expect(state.bottom).toBeLessThanOrEqual(state.cellBottom);
    await artifacts(page, info);
  });
}
