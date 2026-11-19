import { expect, test } from "@playwright/test";
import type { Page } from "@playwright/test";
import { writeFileSync } from "node:fs";
import { createDocumentSession, createWordDocument, openDocument } from "@doffi/core";
import type { DocumentPackage } from "@doffi/core";
import { exportDocx, importDocx } from "@doffi/office";

export function registerWordCrossClipboardTests(
  mount: (page: Page, files: DocumentPackage, mode?: "preview" | "edit") => Promise<void>,
): void {
  for (const tracked of [false, true])
    for (const block of [false, true])
      test(
        "cross paragraph clipboard replaces exact receiver source " + tracked + " " + block,
        async ({ page }, testInfo) => {
          const documentId = "cross-browser",
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
                  html: "<p>AA😀BB</p><p>CCDD</p><p>EE禰\u{e0100}FF</p>",
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
          const before = session.snapshot();
          const paragraphs = [...before.html.matchAll(/data-node-id="([^"]+)"/g)]
            .map((m) => session.getNode(m[1]!)!)
            .filter((n) => n.tagName === "p");
          await mount(page, session.serialize(), "edit");
          const detail = await page
            .frameLocator("iframe")
            .locator('[data-node-id="' + paragraphs[0]!.nodeId + '"]')
            .evaluate(
              (first, args) => {
                const root = first.closest("[data-office-editor]") as HTMLElement;
                root.focus();
                const last = root.querySelector('[data-node-id="' + args.last + '"]')!;
                const range = document.createRange();
                const point = (element: Element, offset: number): [Node, number] => {
                  const walker = document.createTreeWalker(element, NodeFilter.SHOW_TEXT);
                  let node: Node | null;
                  while ((node = walker.nextNode())) {
                    if (
                      node.parentElement?.closest(
                        "[data-sdk-word-autospace],.sdk-number-marker,[data-sdk-word-editing-mark]",
                      )
                    )
                      continue;
                    if (offset <= node.textContent!.length) return [node, offset];
                    offset -= node.textContent!.length;
                  }
                  throw new Error("Missing source boundary");
                };
                range.setStart(...point(first, 2));
                range.setEnd(...point(last, 5));
                const selection = document.getSelection()!;
                selection.removeAllRanges();
                selection.addRange(range);
                const data = new DataTransfer();
                data.setData(
                  "text/html",
                  args.block ? "<p>NEW</p><p>END</p>" : "<strong>NEW</strong>",
                );
                const event = new ClipboardEvent("paste", { bubbles: true, cancelable: true });
                Object.defineProperty(event, "clipboardData", { value: data });
                const detail = {
                  first: first.outerHTML,
                  last: last.outerHTML,
                  selected: range.toString(),
                  start: range.startOffset,
                  end: range.endOffset,
                };
                first.dispatchEvent(event);
                return detail;
              },
              { last: paragraphs[2]!.nodeId, block },
            );
          expect(
            await page.evaluate(() => window.paginationHost.errors),
            JSON.stringify(detail),
          ).toEqual([]);
          await expect
            .poll(() =>
              page.evaluate(() => window.paginationHost.session!.snapshot().manifest.revision),
            )
            .toBe(before.manifest.revision + 1);
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
          const output = testInfo.outputPath("cross-clipboard.docx");
          writeFileSync(output, bytes);
          await testInfo.attach("cross-clipboard.docx", {
            path: output,
            contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
          });
          const reopened = createDocumentSession(
            openDocument(importDocx(bytes, { documentId, unsupportedContent: "reject" }), {
              kind: "document",
            }),
          );
          if (tracked) {
            let root = reopened.getNode(reopened.snapshot().metadata.word!.stories[0]!.rootNodeId)!;
            while (root.parentNodeId) root = reopened.getNode(root.parentNodeId)!;
            reopened.commit(
              reopened.prepare({
                protocolVersion: 1,
                documentId,
                baseRevision: reopened.snapshot().manifest.revision,
                operations: [
                  { op: "reviewWordRevisions", nodeId: root.nodeId, decision: "accept" },
                ],
              }),
            );
          }
          expect(
            reopened.getNode(reopened.snapshot().metadata.word!.stories[0]!.rootNodeId)!.text,
          ).toBe(block ? "AANEWENDFF" : "AANEWFF");
          if (tracked) {
            await page.evaluate(() => {
              const s = window.paginationHost.session!;
              s.commit(
                s.prepare({
                  protocolVersion: 1,
                  documentId: s.snapshot().manifest.documentId,
                  baseRevision: s.snapshot().manifest.revision,
                  operations: [
                    { op: "reviewWordRevisions", nodeId: "document-root", decision: "reject" },
                  ],
                }),
              );
            });
            expect(
              await page.evaluate(() => window.paginationHost.session!.getNode("main-story")!.text),
            ).toBe("AA😀BBCCDDEE禰\u{e0100}FF");
            await page.evaluate(() => window.paginationHost.session!.undo());
          }
          await page.evaluate(() => window.paginationHost.session!.undo());
          expect(await page.evaluate(() => window.paginationHost.session!.snapshot().html)).toBe(
            before.html,
          );
        },
      );
}
