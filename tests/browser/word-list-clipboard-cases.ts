import { expect, test } from "@playwright/test";
import type { Page } from "@playwright/test";
import { writeFileSync } from "node:fs";
import { createDocumentSession, createWordDocument, openDocument } from "@doffi/core";
import type { DocumentPackage } from "@doffi/core";
import { exportDocx, importDocx } from "@doffi/office";
import { nativeListContentFixture } from "../../packages/office/tests/word-list-content-fixture.js";

export function registerWordListClipboardTests(
  mount: (page: Page, files: DocumentPackage, mode?: "preview" | "edit") => Promise<void>,
): void {
  for (const variant of ["native", "html", "continuation"] as const)
    for (const tracked of [false, true]) {
      test(
        "list semantic clipboard isolates numbering, preserves markers, continues input and reviews " +
          variant +
          " " +
          tracked,
        async ({ page }, testInfo) => {
          let files: DocumentPackage;
          if (variant === "html") {
            const source = createDocumentSession(
              openDocument(createWordDocument({ documentId: "html-list-source" }), {
                kind: "document",
              }),
            );
            source.commit(
              source.prepare({
                protocolVersion: 1,
                documentId: "html-list-source",
                baseRevision: 0,
                operations: [
                  {
                    op: "replaceFragment",
                    nodeId: "section-root-1",
                    html: '<ol start="4" type="A" style="font-family: Liberation Serif;color:#880000"><li>ONE<ul><li>SUB</li></ul><p>TAIL</p></li><li value="9">TWO</li><li>THREE</li></ol><p>PRIVATE-BODY</p>',
                  },
                ],
              }),
            );
            files = source.serialize();
          } else
            files = importDocx(nativeListContentFixture(), {
              documentId: "native-list-source",
              unsupportedContent: "reject",
            });
          const sourceModel = openDocument(files, { kind: "document" });
          const sourceNodes = [
            ...sourceModel.snapshot().html.matchAll(/data-node-id="([^"]+)"/g),
          ].map((match) => sourceModel.getNode(match[1]!)!);
          const first =
            variant === "continuation"
              ? sourceModel.snapshot().metadata.word!.tables[0]!.nodeId
              : variant === "html"
                ? sourceNodes.find((node) => node.tagName === "ol")!.nodeId
                : sourceNodes.find((node) => node.tagName === "p" && node.text === "FIRST")!.nodeId;
          const last =
            variant === "native"
              ? sourceNodes.find((node) => node.tagName === "p" && node.text === "OFF")!.nodeId
              : first;
          await mount(page, files, tracked ? "edit" : "preview");
          const copied = await page
            .frameLocator("iframe")
            .locator('[data-office-kind="document"]')
            .evaluate(
              (root, ids) => {
                const first = root.querySelector('[data-node-id="' + ids.first + '"]')!;
                const last = root.querySelector('[data-node-id="' + ids.last + '"]')!;
                const range = document.createRange();
                range.setStartBefore(first);
                range.setEndAfter(last);
                const selection = document.getSelection()!;
                selection.removeAllRanges();
                selection.addRange(range);
                const data = new DataTransfer(),
                  event = new ClipboardEvent("copy", { bubbles: true, cancelable: true });
                Object.defineProperty(event, "clipboardData", { value: data });
                first.dispatchEvent(event);
                return { html: data.getData("text/html"), text: data.getData("text/plain") };
              },
              { first, last },
            );
          const payload = JSON.parse(
            Buffer.from(
              /data-doffi-word-clipboard="([^"]+)"/.exec(copied.html)![1]!,
              "base64",
            ).toString("utf8"),
          );
          expect(copied.html + JSON.stringify(payload)).not.toMatch(/PRIVATE-BODY|PRIVATE-MARKER/);
          if (variant !== "html")
            expect(
              payload.semantics.numbering.map((item: { numberingId: string }) => item.numberingId),
            ).toEqual(variant === "native" ? ["1", "2"] : ["2"]);
          const documentId = "list-copy-target";
          const target = createDocumentSession(
            openDocument(createWordDocument({ documentId }), { kind: "document" }),
          );
          target.commit(
            target.prepare({
              protocolVersion: 1,
              documentId,
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
                  op: "setWordNumbering",
                  nodeId: "document-root",
                  numberingId: "1",
                  levels: [
                    {
                      level: 0,
                      start: 91,
                      format: "decimal",
                      text: "%1.",
                      suffix: "tab",
                      cssClass: "receiver-marker",
                    },
                  ],
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
          const before = target.snapshot();
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
                () => window.paginationHost.session!.snapshot().metadata.word!.numbering.length,
              ),
            )
            .toBe(variant === "continuation" ? 2 : variant === "native" ? 3 : 4);
          const paragraphs = await page.evaluate(() =>
            [...window.paginationHost.session!.snapshot().html.matchAll(/data-node-id="([^"]+)"/g)]
              .map((match) => window.paginationHost.session!.getNode(match[1]!)!)
              .filter((node) => node.tagName === "p")
              .map((node) => ({
                id: node.nodeId,
                text: node.text,
                numbering: node.attributes["data-word-numbering"],
              })),
          );
          const find = (text: string) => paragraphs.find((node) => node.text === text)!;
          const expected =
            variant === "native"
              ? { FIRST: "V)", CHILD: "V.b) ", SECOND: "VI)", RESTART: "X)" }
              : variant === "html"
                ? { ONE: "D.", SUB: "•", TWO: "I.", THREE: "J." }
                : {};
          for (const [text, marker] of Object.entries(expected)) {
            const item = find(text);
            expect(item.numbering).not.toBe("1");
            await expect(
              page
                .frameLocator("iframe")
                .locator('[data-node-id="' + item.id + '"] .sdk-number-marker'),
            ).toHaveAttribute("data-sdk-number-text", marker);
          }
          if (variant === "native") {
            await expect(
              page
                .frameLocator("iframe")
                .locator('[data-node-id="' + find("FIRST").id + '"] .sdk-number-marker'),
            ).toHaveCSS("color", "rgb(136, 0, 0)");
            expect(find("OFF").numbering).toBeUndefined();
          }
          if (variant === "continuation") {
            const reference = await page.evaluate(
              () =>
                window.paginationHost.session!.snapshot().metadata.word!.tables[0]!
                  .continuationFormats![0]!.emptyParagraphs![0]!.numbering!,
            );
            expect(reference.numberingId).not.toBe("2");
            expect(reference.level).toBe(1);
          }
          const pasted = await page.evaluate(() => window.paginationHost.session!.snapshot().html);
          const edited = find(
            variant === "native" ? "RESTART" : variant === "html" ? "THREE" : "R2",
          );
          const paragraph = page
            .frameLocator("iframe")
            .locator('[data-node-id="' + edited.id + '"]');
          await paragraph.click();
          await page.keyboard.press("End");
          await page.keyboard.insertText("EDIT");
          await expect
            .poll(() =>
              page.evaluate((id) => window.paginationHost.session!.getNode(id)!.text, edited.id),
            )
            .toBe(edited.text + "EDIT");
          if (variant !== "continuation") {
            await page.keyboard.press("Enter");
            expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
            await page.keyboard.insertText("NEXT");
            await expect
              .poll(() =>
                page.evaluate(
                  () =>
                    [
                      ...window.paginationHost
                        .session!.snapshot()
                        .html.matchAll(/data-node-id="([^"]+)"/g),
                    ]
                      .map((match) => window.paginationHost.session!.getNode(match[1]!)!)
                      .find((node) => node.tagName === "p" && node.text === "NEXT")?.attributes[
                      "data-word-numbering"
                    ],
                ),
              )
              .toBe(edited.numbering);
            await expect(
              page
                .frameLocator("iframe")
                .locator("p")
                .filter({ hasText: /^NEXT$/ })
                .locator(".sdk-number-marker"),
            ).toHaveAttribute("data-sdk-number-text", variant === "native" ? "XI)" : "K.");
          }
          const entries = await page.evaluate(() =>
            [...window.paginationHost.session!.serialize()].map(
              ([path, bytes]) => [path, [...bytes]] as [string, number[]],
            ),
          );
          let bytes = exportDocx(
            openDocument(new Map(entries.map(([path, bytes]) => [path, new Uint8Array(bytes)])), {
              kind: "document",
            }),
          ).bytes;
          const output = testInfo.outputPath("list-clipboard.docx");
          writeFileSync(output, bytes);
          await testInfo.attach("list-clipboard.docx", {
            path: output,
            contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
          });
          for (let pass = 0; pass < 2; pass++) {
            const reopened = openDocument(
              importDocx(bytes, { documentId, unsupportedContent: "reject" }),
              { kind: "document" },
            );
            expect(
              reopened.getNode(reopened.snapshot().metadata.word!.stories[0]!.rootNodeId)!.text,
            ).toContain(edited.text + "EDIT");
            bytes = exportDocx(reopened).bytes;
          }
          if (tracked) {
            const editedHtml = await page.evaluate(
              () => window.paginationHost.session!.snapshot().html,
            );
            await page.evaluate(() =>
              (window.paginationHost.view as import("@doffi/editor").DocumentTextEditor).execute({
                type: "word",
                action: "review",
              }),
            );
            await page.getByRole("button", { name: "拒绝全部显示的修订", exact: true }).click();
            expect(
              await page.evaluate(() => window.paginationHost.session!.getNode("main-story")!.text),
            ).toBe("AB");
            await page.evaluate(() => window.paginationHost.session!.undo());
            expect(await page.evaluate(() => window.paginationHost.session!.snapshot().html)).toBe(
              editedHtml,
            );
          }
          for (let index = 0; index < (variant === "continuation" ? 1 : 3); index++)
            await page.evaluate(() => window.paginationHost.session!.undo());
          expect(await page.evaluate(() => window.paginationHost.session!.snapshot().html)).toBe(
            pasted,
          );
          await page.evaluate(() => window.paginationHost.session!.undo());
          expect(await page.evaluate(() => window.paginationHost.session!.snapshot().html)).toBe(
            before.html,
          );
          expect(
            await page.evaluate(() => window.paginationHost.session!.snapshot().metadata.word),
          ).toEqual(before.metadata.word);
          expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
        },
      );
    }
}
