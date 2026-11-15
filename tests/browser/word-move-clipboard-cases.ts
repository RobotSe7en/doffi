import { expect, test } from "@playwright/test";
import type { Page } from "@playwright/test";
import { writeFileSync } from "node:fs";
import { createDocumentSession, createWordDocument, openDocument } from "@doffi/core";
import type { DocumentPackage } from "@doffi/core";
import { exportDocx, importDocx } from "@doffi/office";

export function registerWordMoveClipboardTests(
  mount: (page: Page, files: DocumentPackage, mode?: "edit" | "preview") => Promise<void>,
): void {
  for (const block of [false, true])
    for (const tracked of [false, true])
      test(
        "system cut and paste creates grouped move revisions " + block + " " + tracked,
        async ({ page }, testInfo) => {
          const documentId = "browser-move",
            session = createDocumentSession(
              openDocument(createWordDocument({ documentId }), { kind: "document" }),
            );
          session.commit(
            session.prepare({
              protocolVersion: 1,
              documentId,
              baseRevision: 0,
              operations: [
                {
                  op: "replaceFragment",
                  nodeId: "section-root-1",
                  html: block
                    ? "<p>ALEFT</p><p>MID</p><p>RIGHTZ</p><p>TARGET</p>"
                    : "<p>A😀MOVEZ</p><p>TARGET</p>",
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
          const before = session.snapshot(),
            paragraphs = [...before.html.matchAll(/data-node-id="([^"]+)"/g)]
              .map((m) => session.getNode(m[1]!)!)
              .filter((n) => n.tagName === "p");
          await mount(page, session.serialize(), "edit");
          const copied = await page
            .frameLocator("iframe")
            .locator('[data-node-id="' + paragraphs[0]!.nodeId + '"]')
            .evaluate(
              (first, args) => {
                const root = first.closest("[data-office-editor]") as HTMLElement;
                root.focus();
                const last = root.querySelector('[data-node-id="' + args.last + '"]')!,
                  range = document.createRange();
                range.setStart(first.firstChild!, 1);
                range.setEnd(last.firstChild!, args.block ? 5 : 7);
                const selection = document.getSelection()!;
                selection.removeAllRanges();
                selection.addRange(range);
                const data = new DataTransfer(),
                  event = new ClipboardEvent("cut", { bubbles: true, cancelable: true });
                Object.defineProperty(event, "clipboardData", { value: data });
                first.dispatchEvent(event);
                return { html: data.getData("text/html"), text: data.getData("text/plain") };
              },
              { last: paragraphs[block ? 2 : 0]!.nodeId, block },
            );
          expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
          expect(copied.html).toContain("data-doffi-word-cut");
          await page
            .frameLocator("iframe")
            .locator('[data-node-id="' + paragraphs.at(-1)!.nodeId + '"]')
            .evaluate((target, copied) => {
              const range = document.createRange();
              range.setStart(target.firstChild!, 3);
              range.collapse(true);
              const selection = document.getSelection()!;
              selection.removeAllRanges();
              selection.addRange(range);
              const data = new DataTransfer();
              data.setData("text/html", copied.html);
              data.setData("text/plain", copied.text);
              const event = new ClipboardEvent("paste", { bubbles: true, cancelable: true });
              Object.defineProperty(event, "clipboardData", { value: data });
              target.dispatchEvent(event);
            }, copied);
          expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
          await expect
            .poll(() =>
              page.evaluate(() => window.paginationHost.session!.snapshot().manifest.revision),
            )
            .toBe(before.manifest.revision + 2);
          expect(
            (await page.evaluate(() => window.paginationHost.session!.snapshot())).metadata.word!
              .moves ?? [],
          ).toHaveLength(tracked ? 2 : 0);
          const entries = await page.evaluate(() =>
            [...window.paginationHost.session!.serialize()].map(
              ([p, b]) => [p, [...b]] as [string, number[]],
            ),
          );
          const bytes = exportDocx(
            openDocument(new Map(entries.map(([p, b]) => [p, new Uint8Array(b)])), {
              kind: "document",
            }),
          ).bytes;
          const output = testInfo.outputPath("move-clipboard.docx");
          writeFileSync(output, bytes);
          await testInfo.attach("move-clipboard.docx", {
            path: output,
            contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
          });
          const reopened = openDocument(
            importDocx(bytes, { documentId, unsupportedContent: "reject" }),
            { kind: "document" },
          );
          expect(reopened.snapshot().metadata.word!.moves ?? []).toHaveLength(tracked ? 2 : 0);
          if (tracked) {
            await page.evaluate(() =>
              (window.paginationHost.view as import("@doffi/editor").DocumentTextEditor).execute({
                type: "word",
                action: "review",
              }),
            );
            await page.getByRole("button", { name: "接受全部显示的修订", exact: true }).click();
            expect(
              await page.evaluate(() => window.paginationHost.session!.getNode("main-story")!.text),
            ).toBe(block ? "AZTARLEFTMIDRIGHTGET" : "AZTAR😀MOVEGET");
            await page.evaluate(() => window.paginationHost.session!.undo());
            await page.getByRole("button", { name: "拒绝全部显示的修订", exact: true }).click();
            expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
            expect(
              await page.evaluate(() => window.paginationHost.session!.getNode("main-story")!.text),
            ).toBe(block ? "ALEFTMIDRIGHTZTARGET" : "A😀MOVEZTARGET");
            await page.evaluate(() => window.paginationHost.session!.undo());
          } else
            expect(
              await page.evaluate(() => window.paginationHost.session!.getNode("main-story")!.text),
            ).toBe(block ? "AZTARLEFTMIDRIGHTGET" : "AZTAR😀MOVEGET");
          await page.evaluate(() => {
            window.paginationHost.session!.undo();
            window.paginationHost.session!.undo();
          });
          expect(await page.evaluate(() => window.paginationHost.session!.snapshot().html)).toBe(
            before.html,
          );
        },
      );
}
