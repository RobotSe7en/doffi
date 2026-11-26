import { expect, test } from "@playwright/test";
import type { Page } from "@playwright/test";
import { openDocument } from "@doffi/core";
import type { DocumentPackage } from "@doffi/core";
import { exportDocx, importDocx } from "@doffi/office";
import { rowExceptionScopeFixture } from "../../packages/office/tests/row-exception-scope-fixture.js";
import type {} from "./fixtures/pagination-host.js";

export function registerWordRowExceptionScopeTests(mount: (page: Page, files: DocumentPackage, mode?: "preview" | "edit") => Promise<void>): void {
  for (const compatibility of [14, 15] as const)
    test(`Word row exception scopes retain one grid in compatibility ${compatibility}`, async ({ page }) => {
      for (const kind of ["indent", "width", "align", "first-indent", "first-width", "first-align"] as const) for (const merged of [false, true]) for (const directAlignment of kind.endsWith("align") ? [false, true] : [false]) {
        const files = importDocx(rowExceptionScopeFixture(compatibility, kind, merged, directAlignment), { documentId: "row-scope", unsupportedContent: "reject" });
        const reopened = importDocx(exportDocx(openDocument(files, { kind: "document" })).bytes, { documentId: "scope-reopen", unsupportedContent: "reject" });
        for (const source of [files, reopened]) for (const mode of ["preview", "edit"] as const) {
          await mount(page, source, mode);
          const actual = await page.frameLocator("iframe").locator("body").evaluate(body => {
            const x = (p: Element) => {
              const r = document.createRange(); r.selectNodeContents(p); return r.getBoundingClientRect().left;
            }, reference = x([...body.querySelectorAll("p")].find(p => p.textContent === "REFERENCE")!);
            return Object.fromEntries([...body.querySelectorAll("td p")].map(p => [p.textContent, (x(p) - reference) * .75]));
          });
          const globalWidth = kind === "first-width" || (kind === "first-align"),
            track = globalWidth ? 60 : 80;
          for (let row = 0; row < 3; row++) {
            let start = (compatibility === 14 ? 12 : 19.5) + (kind === "first-indent" ? 18 : 0);
            if (kind === "indent" && row > 0) start += row === 1 ? 4 : -4;
            if ((kind === "align" && compatibility === 14 && row > 0)
              || (kind === "first-align" && (compatibility === 15 || row === 0)))
              start = kind === "align" && row === 2 ? 225.3 : (451.3 - track * 3) / 2 + 7;
            for (let col = 0; col < 3; col++) {
              if (merged && row === 1 && col === 0) continue;
              // Word's PDF grid is quantized to 0.24pt; browser stroke snapping
              // and glyph advances add a sub-point difference at this scale.
              expect(Math.abs(actual[`R${row}C${col}`]! - start - col * track), `${kind}/${mode}/row${row}col${col}`).toBeLessThan(1.1);
            }
          }
          if (kind === "width" && !merged && source === files && mode === "edit") {
            const before = await page.evaluate(() => window.paginationHost.session!.snapshot().html),
              table = page.frameLocator("iframe").locator("table"),
              width = await table.evaluate(node => node.getBoundingClientRect().width);
            await page.evaluate(() => {
              const s = window.paginationHost.session!, snap = s.snapshot();
              s.commit(s.prepare({ protocolVersion: 1, documentId: snap.manifest.documentId,
                baseRevision: snap.manifest.revision, operations: [{ op: "editWordTable",
                  nodeId: snap.metadata.word!.tables[0]!.nodeId, change: { kind: "deleteRows", at: 0, count: 1 } }] }));
            });
            await expect.poll(() => table.evaluate(node => node.getBoundingClientRect().width)).toBeLessThan(width - 50);
            await page.evaluate(() => window.paginationHost.session!.undo());
            expect(await page.evaluate(() => window.paginationHost.session!.snapshot().html)).toBe(before);
            expect(await table.evaluate(node => node.getBoundingClientRect().width)).toBeCloseTo(width, 1);
          }
          expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
        }
      }
    });
}
