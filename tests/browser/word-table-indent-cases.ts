import { expect, test } from "@playwright/test";
import type { Page } from "@playwright/test";
import { createDocumentSession, openDocument } from "@doffi/core";
import type { DocumentPackage } from "@doffi/core";
import type { DocumentPreviewOptions } from "@doffi/editor";
import { exportDocx, importDocx } from "@doffi/office";
import { tableIndentFixture } from "../../packages/office/tests/table-indent-fixture.js";
import type {} from "./fixtures/pagination-host.js";

export function registerWordTableIndentTests(mount: (page: Page, files: DocumentPackage,
  mode?: "preview" | "edit" | "interaction", options?: DocumentPreviewOptions) => Promise<void>): void {
  test("native table indentation resolves compatibility, inherited styles and leading-cell overrides through preview, editing and native reopen", async ({ page }) => {
    for (const options of [{}, { mode: 12 }, { mode: 14, styled: true }, { mode: 15, styled: true },
      { mode: 14, styled: true, cellOverride: true }, { mode: 14, rtl: true, cellOverride: true }, { mode: 14, nested: true },
      { mode: 14, styled: true, ignoredIndent: true }]) {
      const files = importDocx(tableIndentFixture(options), { documentId: "table-indent", unsupportedContent: "reject" });
      const native = exportDocx(openDocument(files, { kind: "document" })).bytes;
      const reopened = importDocx(native, { documentId: "table-indent-reopen", unsupportedContent: "reject" });
      const expected = options.mode === 15 ? 12 : 12 - (options.cellOverride ? 5 : 7);
      for (const source of [files, reopened]) for (const mode of ["preview", "edit"] as const) {
        await mount(page, source, mode, { layout: "continuous" });
        const table = page.frameLocator("iframe").locator("table").first();
        const actual = await table.evaluate((node, rtl) => parseFloat(getComputedStyle(node).getPropertyValue(rtl ? "margin-right" : "margin-left")), options.rtl ?? false);
        expect(actual).toBeCloseTo(expected * 4 / 3, 2);
        if (options.nested) await expect(table.locator("table")).toHaveCSS("margin-left", "0px");
        if (source === files && mode === "edit" && options.mode === 12) {
          const before = await page.evaluate(() => window.paginationHost.session!.snapshot().html);
          const id = await table.getAttribute("data-node-id");
          const originalStyle = await page.evaluate(id => window.paginationHost.session!.getNode(id!)!.attributes.style, id);
          const cell = page.frameLocator("iframe").getByText("LEADING", { exact: true });
          await cell.click(); await cell.press("End"); await cell.press("X");
          expect(await page.evaluate(id => window.paginationHost.session!.getNode(id!)!.attributes.style, id)).toBe(originalStyle);
          await page.evaluate(() => window.paginationHost.session!.undo());
          expect(await page.evaluate(() => window.paginationHost.session!.snapshot().html)).toBe(before);
        }
        expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
      }
    }
  });

  test("physical CSS table margins survive legacy DOCX export and input undo without acquiring a second inset", async ({ page }) => {
    const files = new Map(importDocx(tableIndentFixture({ mode: 12 }), { documentId: "physical-indent" }));
    files.set("styles.css", new TextEncoder().encode(new TextDecoder().decode(files.get("styles.css"))
      .replaceAll("--word-table-indent-origin:native", "--word-table-indent-origin:initial")));
    files.set("content.html", new TextEncoder().encode(new TextDecoder().decode(files.get("content.html"))
      .replaceAll("--word-table-indent-origin:native", "--word-table-indent-origin:initial")));
    const session = createDocumentSession(openDocument(files, { kind: "document" }));
    const tableId = session.snapshot().metadata.word!.tables[0]!.nodeId;
    session.commit(session.prepare({ protocolVersion: 1, documentId: "physical-indent", baseRevision: 0,
      operations: [{ op: "formatWordTable", nodeId: tableId, properties: { "margin-left": "18pt" } }] }));
    const before = session.snapshot().html;
    const reopened = importDocx(exportDocx(session).bytes, { documentId: "physical-reopen", unsupportedContent: "reject" });
    for (const source of [session.serialize(), reopened]) for (const mode of ["preview", "edit"] as const) {
      await mount(page, source, mode, { layout: "continuous" });
      await expect(page.frameLocator("iframe").locator("table").first()).toHaveCSS("margin-left", "24px");
    }
    await mount(page, session.serialize(), "edit", { layout: "continuous" });
    const cell = page.frameLocator("iframe").getByText("LEADING", { exact: true });
    await cell.click(); await cell.press("End"); await cell.press("X");
    await page.evaluate(() => window.paginationHost.session!.undo());
    expect(await page.evaluate(() => window.paginationHost.session!.snapshot().html)).toBe(before);
    await expect(page.frameLocator("iframe").locator("table").first()).toHaveCSS("margin-left", "24px");
    expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
    session.commit(session.prepare({ protocolVersion: 1, documentId: "physical-indent", baseRevision: session.snapshot().manifest.revision,
      operations: [{ op: "formatWordTable", nodeId: tableId, properties: { "margin-left": "initial" } }] }));
    const reset = importDocx(exportDocx(session).bytes, { documentId: "physical-reset", unsupportedContent: "reject" });
    for (const mode of ["preview", "edit"] as const) {
      await mount(page, reset, mode, { layout: "continuous" });
      await expect(page.frameLocator("iframe").locator("table").first()).toHaveCSS("margin-left", "0px");
    }
  });
}
