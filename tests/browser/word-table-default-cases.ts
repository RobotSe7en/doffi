import { expect, test } from "@playwright/test";
import type { Page } from "@playwright/test";
import { openDocument } from "@doffi/core";
import type { DocumentPackage } from "@doffi/core";
import type { DocumentPreviewOptions } from "@doffi/editor";
import { exportDocx, importDocx } from "@doffi/office";
import { tableDefaultCases, tableDefaultFixture } from "../../packages/office/tests/table-default-fixture.js";

export function registerWordTableDefaultTests(mount: (page: Page, files: DocumentPackage,
  mode?: "preview" | "edit" | "interaction", options?: DocumentPreviewOptions) => Promise<void>): void {
  test("native stock table defaults differ from custom and absent styles through preview, editing and editable DOCX reopen", async ({ page }) => {
    test.setTimeout(90_000);
    for (const { kind, padding } of tableDefaultCases) for (const modeNumber of [14, 15]) {
      const files = importDocx(tableDefaultFixture(kind, { mode: modeNumber }), {
        documentId: "native-table-defaults", unsupportedContent: "reject",
      });
      const reopened = importDocx(exportDocx(openDocument(files, { kind: "document" })).bytes, {
        documentId: "native-defaults-reopen", unsupportedContent: "reject",
      });
      for (const source of [files, reopened]) for (const mode of ["preview", "edit"] as const) {
        await mount(page, source, mode, { layout: "continuous" });
        const frame = page.frameLocator("iframe"), table = frame.locator("table");
        await expect(frame.locator("td").first()).toHaveCSS("padding-left", padding * 4 / 3 + "px");
        const margin = await table.evaluate(node => parseFloat(getComputedStyle(node).marginLeft));
        expect(margin).toBeCloseTo((modeNumber === 14 ? -padding : 0) * 4 / 3, 2);
        expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
      }
    }
  });
}
