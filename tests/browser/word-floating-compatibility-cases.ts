import { expect, test } from "@playwright/test";
import type { Page } from "@playwright/test";
import { openDocument } from "@doffi/core";
import type { DocumentPackage } from "@doffi/core";
import { exportDocx, importDocx } from "@doffi/office";
import { floatingCompatibilityFixture, drawingBodyDefaultsFixture } from "../../packages/office/tests/floating-compatibility-fixture.js";
import type {} from "./fixtures/pagination-host.js";
import { wordProducerDrawingFixture } from "../../packages/office/tests/word-producer-drawing-fixture.js";
import { writeFileSync } from "node:fs";

export function registerWordFloatingCompatibilityTests(mount: (page: Page, files: DocumentPackage, mode?: "preview" | "edit") => Promise<void>): void {
  test("floating text edits in place without an automatic settings form, with clean source, undo and editable native reopen", async ({ page }, testInfo) => {
    const files = importDocx(wordProducerDrawingFixture({ scale: 90 }), { documentId: "in-place-text", unsupportedContent: "reject" });
    await mount(page, files, "edit");
    const frame = page.frameLocator("iframe"), object = frame.locator('.sdk-word-floating-layer [data-word-drawing="textBox"]'),
      before = await page.evaluate(() => window.paginationHost.session!.snapshot()),
      paper = await frame.locator(".sdk-paper").first().boundingBox(), box = await object.boundingBox();
    await page.evaluate(() => {
      const host = window.paginationHost as typeof window.paginationHost & { stale?: ReturnType<import("@doffi/core").DocumentEditingSession["prepare"]> },
        session = host.session!, snapshot = session.snapshot(),
        paragraph = [...snapshot.html.matchAll(/data-node-id="([^"]+)"/g)].map(match => session.getNode(match[1]!)).find(node => node?.tagName === "p" && !!node.text && !node.text.includes("\ufffc"))!;
      host.stale = session.prepare({ protocolVersion: 1, documentId: snapshot.manifest.documentId, baseRevision: snapshot.manifest.revision,
        operations: [{ op: "replaceText", nodeId: paragraph.nodeId, from: 0, to: 0, expectedText: "", text: "STALE" }] });
    });
    await object.click({ position: { x: 15, y: 15 } });
    await expect(object).toHaveAttribute("data-sdk-word-selected", "");
    await expect(page.getByRole("region", { name: "绘图设置", exact: true })).toHaveCount(0);
    await page.evaluate(() => (window.paginationHost.view as import("@doffi/editor").DocumentTextEditor).execute({ type: "word", action: "drawingSettings" }));
    await expect(page.getByRole("dialog", { name: "形状格式", exact: true })).toBeVisible();
    expect((await frame.locator(".sdk-paper").first().boundingBox())!.y).toBe(paper!.y);
    await page.getByRole("button", { name: "关闭绘图设置", exact: true }).click();
    await object.evaluate(node => Object.defineProperty(node.ownerDocument, "caretRangeFromPoint", { value: undefined, configurable: true }));
    await object.dblclick({ position: { x: 15, y: 15 } });
    const editor = object.locator("[data-office-editor]");
    await expect(editor).toBeVisible();
    expect((await object.boundingBox())!.y).toBeCloseTo(box!.y, 1);
    await expect(frame.locator(".sdk-paper")).toHaveCount(1);
    await expect(page.getByRole("region", { name: "绘图设置", exact: true })).toHaveCount(0);
    await editor.press("ArrowLeft");
    expect((await page.evaluate(() => window.paginationHost.session!.snapshot())).metadata.word!.drawings).toEqual(before.metadata.word!.drawings);
    const points = await editor.evaluate(root => {
      const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT), points: { x: number; y: number; prefix: string }[] = [];
      let prefix = "";
      for (let text = walker.nextNode(); text && points.length < 3; text = walker.nextNode()) {
        if (!text.textContent || text.parentElement?.classList.contains("sdk-word-kinsoku-boundary")) continue;
        for (let offset = 0; offset < text.textContent.length && points.length < 3; offset++) {
          const character = text.textContent[offset]!; if (!/[\p{Script=Han}]/u.test(character)) continue;
          prefix += character;
          const range = document.createRange(); range.setStart(text, offset); range.setEnd(text, offset + 1);
          const rect = range.getBoundingClientRect(); points.push({ x: rect.right - 1, y: (rect.top + rect.bottom) / 2, prefix });
        }
      }
      return points;
    });
    expect(points.map(point => point.prefix)).toEqual(["合", "合成", "合成文"]);
    const frameBox = (await page.locator("iframe").boundingBox())!;
    for (const point of points) {
      await page.mouse.click(frameBox.x + point.x, frameBox.y + point.y);
      expect(await editor.evaluate(root => {
        const selection = document.getSelection()!, prefix = document.createRange();
        prefix.selectNodeContents(root); prefix.setEnd(selection.focusNode!, selection.focusOffset);
        return { collapsed: selection.isCollapsed, prefix: prefix.toString().replace(/[\u2060\u200b]/g, "") };
      })).toEqual({ collapsed: true, prefix: point.prefix });
    }
    await page.mouse.click(frameBox.x + points[0]!.x, frameBox.y + points[0]!.y);
    await page.keyboard.insertText("点");
    await expect.poll(() => page.evaluate(() => window.paginationHost.session!.snapshot().html)).toContain("合点");
    await page.evaluate(() => window.paginationHost.session!.undo());
    await expect.poll(() => page.evaluate(() => window.paginationHost.session!.snapshot().html)).toBe(before.html);
    await editor.evaluate(root => {
      root.focus(); const range = document.createRange(); range.selectNodeContents(root.querySelector("p")!); range.collapse(false);
      const selection = document.getSelection()!; selection.removeAllRanges(); selection.addRange(range);
    });
    await page.keyboard.insertText("原位EDIT😀");
    await expect.poll(() => page.evaluate(() => window.paginationHost.session!.snapshot().html)).toContain("原位EDIT😀");
    const edited = await page.evaluate(() => window.paginationHost.session!.snapshot());
    expect(await page.evaluate(() => {
      const host = window.paginationHost as typeof window.paginationHost & { stale?: ReturnType<import("@doffi/core").DocumentEditingSession["prepare"]> };
      try { host.session!.commit(host.stale!); return "committed"; }
      catch (error) { return (error as { code: string }).code; }
      finally { delete host.stale; }
    })).toBe("RevisionConflict");
    expect(await page.evaluate(() => window.paginationHost.session!.snapshot().html)).toBe(edited.html);
    expect(edited.html).not.toMatch(/data-sdk-|data-office-editor|sdk-word-run-width/);
    expect(edited.metadata.word!.drawings).toEqual(before.metadata.word!.drawings);
    await page.evaluate(() => window.paginationHost.session!.undo());
    await expect.poll(() => page.evaluate(() => window.paginationHost.session!.snapshot().html)).toBe(before.html);
    await page.evaluate(() => window.paginationHost.session!.redo());
    await expect.poll(() => page.evaluate(() => window.paginationHost.session!.snapshot().html)).toBe(edited.html);
    const beforeEscape = await object.boundingBox();
    await editor.press("Escape");
    expect((await object.boundingBox())!.y).toBeCloseTo(beforeEscape!.y, 1);
    await expect(frame.locator('[data-office-editor][data-sdk-word-text-story]')).toHaveCount(0);
    await expect(object).toContainText(/原位\s*EDIT😀/);
    await expect(object.locator(".sdk-word-run-width .sdk-word-run-width")).toHaveCount(0);
    const entries = await page.evaluate(() => [...window.paginationHost.session!.serialize()].map(([path, bytes]) => [path, Array.from(bytes)] as const)),
      current = new Map(entries.map(([path, bytes]) => [path, new Uint8Array(bytes)])), native = exportDocx(openDocument(current, { kind: "document" })).bytes;
    const output = testInfo.outputPath("inplace-edit.docx"); writeFileSync(output, native);
    await testInfo.attach("inplace-edit.docx", { path: output, contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
    for (let pass = 0, bytes = native; pass < 2; pass++) {
      const reopened = openDocument(importDocx(bytes, { documentId: "in-place-reopen", unsupportedContent: "reject" }), { kind: "document" });
      const story = reopened.snapshot().metadata.word!.stories.find(item => item.kind === "textBox")!;
      expect(reopened.getNode(story.rootNodeId)!.text).toContain("原位EDIT😀");
      expect(reopened.snapshot().metadata.word!.drawings![0]!.kind).toBe("textBox");
      bytes = exportDocx(reopened).bytes;
    }
    expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
  });
  test("producer floating text stays editable with script fonts, scale, locks, hidden objects and table story isolation", async ({ page }) => {
    for (const options of [{ scale: 75 }, { hidden: true }, { connector: true }, { oversized: true, scale: 90 }, { scale: 90, offset: -5 }]) {
      const files = importDocx(wordProducerDrawingFixture(options), { documentId: "producer-frame", unsupportedContent: "reject" }),
        reopened = importDocx(exportDocx(openDocument(files, { kind: "document" })).bytes, { documentId: "producer-frame-reopen", unsupportedContent: "reject" });
      for (const source of [files, reopened]) for (const mode of ["preview", "edit"] as const) {
        await mount(page, source, mode);
        const frame = page.frameLocator("iframe"), object = frame.locator(".sdk-word-floating-layer [data-sdk-word-object]");
        if ("hidden" in options) { await expect(object).toHaveCount(0); continue; }
        await expect(object).toHaveCount(1);
        if ("scale" in options) {
          const scales = object.locator(".sdk-word-run-width-text");
          expect(await scales.count()).toBeGreaterThan(0);
          expect(await scales.first().evaluate(node => new DOMMatrix(getComputedStyle(node).transform).a)).toBeCloseTo(options.scale! / 100, 3);
          expect(await object.textContent()).toContain("合成文本框");
        }
        if ("offset" in options) {
          const x = await object.evaluate(node => (node.getBoundingClientRect().left - node.ownerDocument.querySelector(".sdk-paper")!.getBoundingClientRect().left) * .75);
          expect(x).toBeCloseTo(36, 1);
        }
        expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
      }
    }
  });
  test("native shape text margins retain their defaults and explicit zero through both shared surfaces", async ({ page }) => {
    for (const body of ["", 'lIns="0" tIns="25400"']) {
      const files = importDocx(drawingBodyDefaultsFixture(body), { documentId: "shape-insets", unsupportedContent: "reject" }),
        reopened = importDocx(exportDocx(openDocument(files, { kind: "document" })).bytes, { documentId: "shape-insets-reopen", unsupportedContent: "reject" });
      for (const source of [files, reopened]) for (const mode of ["preview", "edit"] as const) {
        await mount(page, source, mode);
        const box = page.frameLocator("iframe").locator(".sdk-word-floating-layer .sdk-word-drawing-text");
        expect(await box.evaluate(node => parseFloat(getComputedStyle(node).paddingLeft) * .75)).toBeCloseTo(body ? 0 : 7.2, 2);
        expect(await box.evaluate(node => parseFloat(getComputedStyle(node).paddingTop) * .75)).toBeCloseTo(body ? 2 : 3.6, 2);
        expect(await box.evaluate(node => parseFloat(getComputedStyle(node).paddingRight) * .75)).toBeCloseTo(7.2, 2);
      }
    }
  });
  for (const compatibility of [14, 15] as const)
    test(`native floating cell scope and physical table origin in compatibility ${compatibility}`, async ({ page }) => {
      for (const kind of ["drawing", "cell", "cell-outside", "cell-none"] as const) for (const preceding of kind === "cell-outside" ? [undefined, "paragraph", "row"] as const : [undefined]) {
        const files = importDocx(floatingCompatibilityFixture(compatibility, kind, preceding), { documentId: "float-mode", unsupportedContent: "reject" }),
          native = exportDocx(openDocument(files, { kind: "document" })).bytes,
          reopened = importDocx(native, { documentId: "float-mode-reopen", unsupportedContent: "reject" });
        expect(openDocument(reopened, { kind: "document" }).snapshot().metadata.word!.drawings![0]!.layout.layoutInCell).toBe(kind === "cell");
        for (const source of [files, reopened]) for (const mode of ["preview", "edit"] as const) {
          await mount(page, source, mode);
          const rect = await page.frameLocator("iframe").locator(".sdk-word-floating-layer [data-sdk-word-object]").first().evaluate(node => {
            const r = node.getBoundingClientRect(), paper = node.ownerDocument.querySelector(".sdk-paper")!.getBoundingClientRect();
            return { x: (r.x - paper.x) * .75, y: (r.y - paper.y) * .75, width: r.width * .75, height: r.height * .75 };
          });
          const inCell = kind !== "drawing" && (kind === "cell" || compatibility === 15);
          expect(rect.x, kind + "/" + mode).toBeCloseTo(inCell ? 43 : 130, 1);
          const offset = preceding === "row" ? 44 : preceding === "paragraph" && inCell ? 16 : 0;
          expect(rect.y, kind + "/" + mode + "/" + preceding).toBeCloseTo((inCell ? 52 : 48) + offset, 1);
          expect(rect.width).toBeCloseTo(160, 1);
          expect(rect.height).toBeCloseTo(100, 1);
          if (inCell && kind !== "cell-none" && !preceding) {
            const textTop = await page.frameLocator("iframe").locator("td").first().evaluate(cell => {
              const walker = document.createTreeWalker(cell, NodeFilter.SHOW_TEXT);
              for (let text = walker.nextNode(); text; text = walker.nextNode()) {
                if (!text.textContent?.trim() || text.parentElement?.closest("[data-sdk-word-object]")) continue;
                const range = document.createRange(); range.setStart(text, 0); range.setEnd(text, 1);
                const r = range.getBoundingClientRect(), paper = document.querySelector(".sdk-paper")!.getBoundingClientRect();
                if (r.width > 0) return (r.top - paper.top) * .75;
              }
              return NaN;
            });
            expect(textTop).toBeGreaterThan(rect.y + rect.height);
            expect(textTop - rect.y - rect.height).toBeLessThan(10);
          }
          if (kind === "cell-outside" && source === files && mode === "edit") {
            const before = await page.evaluate(() => window.paginationHost.session!.snapshot().html),
              paragraph = page.frameLocator("iframe").locator("td p").last();
            await paragraph.evaluate(p => {
              (p.closest("[contenteditable]") as HTMLElement).focus();
              const range = document.createRange(); range.selectNodeContents(p); range.collapse(false);
              const selection = document.getSelection()!; selection.removeAllRanges(); selection.addRange(range);
            });
            await page.keyboard.insertText("X");
            await expect.poll(() => page.evaluate(() => window.paginationHost.session!.snapshot().html)).not.toBe(before);
            await page.evaluate(() => window.paginationHost.session!.undo());
            expect(await page.evaluate(() => window.paginationHost.session!.snapshot().html)).toBe(before);
          }
          expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
        }
      }
    });
}
