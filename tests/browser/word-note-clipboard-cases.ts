import { expect, test } from "@playwright/test";
import type { Page } from "@playwright/test";
import { writeFileSync } from "node:fs";
import { createWordDocument, createDocumentSession, openDocument } from "@doffi/core";
import type { DocumentPackage, WordContentSemantics } from "@doffi/core";
import { exportDocx, importDocx } from "@doffi/office";
import { wordNoteContentFixture } from "../../packages/office/tests/word-note-content-fixture.js";

export function registerWordNoteClipboardTests(
  mount: (page: Page, files: DocumentPackage, mode?: "edit" | "preview") => Promise<void>,
): void {
  for (const both of [false, true])
    for (const tracked of [false, true])
      test(
        "note semantic system clipboard copies selected ownership and editable bodies " +
          both +
          " " +
          tracked,
        async ({ page }, testInfo) => {
          const source = wordNoteContentFixture(true);
          await mount(page, source.session.serialize(), tracked ? "edit" : "preview");
          const copied = await page
            .frameLocator("iframe")
            .locator('[data-node-id="' + source.footId + '"]')
            .evaluate(
              (first, args) => {
                const root = first.closest('[data-office-kind="document"]')!,
                  last = root.querySelector('[data-node-id="' + args.last + '"]')!;
                const range = document.createRange();
                range.setStartBefore(first);
                range.setEndAfter(args.both ? last : first);
                const selection = document.getSelection()!;
                selection.removeAllRanges();
                selection.addRange(range);
                const data = new DataTransfer(),
                  event = new ClipboardEvent("copy", { bubbles: true, cancelable: true });
                Object.defineProperty(event, "clipboardData", { value: data });
                first.dispatchEvent(event);
                return { html: data.getData("text/html"), text: data.getData("text/plain") };
              },
              { both, last: source.endId },
            );
          const payload = JSON.parse(
            Buffer.from(
              /data-doffi-word-clipboard="([^"]+)"/.exec(copied.html)![1]!,
              "base64",
            ).toString("utf8"),
          ) as { semantics: WordContentSemantics };
          expect(JSON.stringify(payload) + copied.html).not.toMatch(/PRIVATE-/);
          expect(payload.semantics.notes).toHaveLength(both ? 2 : 1);
          expect(payload.semantics.stories).toHaveLength(both ? 3 : 2);
          expect(payload.semantics.fields).toHaveLength(both ? 2 : 1);
          if (!both) expect(JSON.stringify(payload)).not.toMatch(/END-BODY|NOTE-CELL/);
          const documentId = "note-browser-target";
          let target = createDocumentSession(
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
                  op: "insertWordNote",
                  nodeId: "paragraph-1",
                  offset: 1,
                  kind: "footnote",
                  text: "RECEIVER-NOTE",
                },
              ],
            }),
          );
          target = createDocumentSession(
            openDocument(
              importDocx(exportDocx(target).bytes, { documentId, unsupportedContent: "reject" }),
              { kind: "document" },
            ),
          );
          const targetP = [...target.snapshot().html.matchAll(/data-node-id="([^"]+)"/g)]
            .map((m) => target.getNode(m[1]!)!)
            .find((n) => n.tagName === "p" && n.text === "A\ufffcB")!.nodeId;
          let root = target.getNode(targetP)!;
          while (root.parentNodeId) root = target.getNode(root.parentNodeId)!;
          target.commit(
            target.prepare({
              protocolVersion: 1,
              documentId,
              baseRevision: target.snapshot().manifest.revision,
              operations: [
                {
                  op: "setWordReview",
                  nodeId: root.nodeId,
                  trackChanges: tracked,
                  author: "Receiver",
                },
              ],
            }),
          );
          const before = target.snapshot(),
            rootId = root.nodeId;
          await mount(page, target.serialize(), "edit");
          await page
            .frameLocator("iframe")
            .locator('[data-node-id="' + targetP + '"]')
            .evaluate((paragraph, copied) => {
              (paragraph.closest("[data-office-editor]") as HTMLElement).focus();
              const range = document.createRange();
              range.setStart(paragraph.firstChild!, 1);
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
          expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
          const pasted = await page.evaluate(() => window.paginationHost.session!.snapshot()),
            word = pasted.metadata.word!,
            normal = word.notes!.filter((n) => (n.type ?? "normal") === "normal");
          expect(normal).toHaveLength(both ? 3 : 2);
          const oldIds = new Set(before.metadata.word!.notes!.map((n) => n.noteId));
          expect(
            normal
              .filter((n) => !oldIds.has(n.noteId))
              .every(
                (n) =>
                  !payload.semantics.notes!.some(
                    (s) => s.noteId === n.noteId || s.storyId === n.storyId,
                  ),
              ),
          ).toBe(true);
          expect(word.fields).toHaveLength(both ? 2 : 1);
          expect(word.drawings).toHaveLength(2);
          const foot = normal.find((n) => n.kind === "footnote" && !oldIds.has(n.noteId))!,
            story = word.stories.find((s) => s.storyId === foot.storyId)!;
          await page.evaluate(() =>
            (window.paginationHost.view as import("@doffi/editor").DocumentTextEditor).execute({
              type: "word",
              action: "referenceSettings",
            }),
          );
          await page.getByRole("button", { name: "编辑 " + foot.noteId, exact: true }).click();
          const paragraph = page
            .frameLocator("iframe")
            .locator('[data-node-id="' + story.rootNodeId + '"] p')
            .filter({ hasText: /LIST-BODY/ });
          await paragraph.evaluate((p) => {
            (p.closest("[data-office-editor]") as HTMLElement).focus();
            const range = document.createRange();
            range.selectNodeContents(p);
            range.collapse(false);
            const selection = document.getSelection()!;
            selection.removeAllRanges();
            selection.addRange(range);
          });
          await page.keyboard.insertText("EDIT");
          await expect
            .poll(() =>
              page.evaluate(
                (id) => window.paginationHost.session!.getNode(id)!.text,
                story.rootNodeId,
              ),
            )
            .toContain("LIST-BODYEDIT");
          const entries = await page.evaluate(() =>
            [...window.paginationHost.session!.serialize()].map(
              ([p, b]) => [p, [...b]] as [string, number[]],
            ),
          );
          let bytes = exportDocx(
            openDocument(new Map(entries.map(([p, b]) => [p, new Uint8Array(b)])), {
              kind: "document",
            }),
          ).bytes;
          const output = testInfo.outputPath("note-clipboard.docx");
          writeFileSync(output, bytes);
          await testInfo.attach("note-clipboard.docx", {
            path: output,
            contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
          });
          for (let pass = 0; pass < 2; pass++) {
            const reopened = openDocument(
              importDocx(bytes, { documentId, unsupportedContent: "reject" }),
              { kind: "document" },
            );
            expect(reopened.snapshot().metadata.word!.fields).toHaveLength(both ? 2 : 1);
            expect(
              reopened
                .snapshot()
                .metadata.word!.stories.some(
                  (s) =>
                    s.kind === "footnote" &&
                    reopened.getNode(s.rootNodeId)!.text.includes("LIST-BODYEDIT"),
                ),
            ).toBe(true);
            expect(
              reopened.snapshot().metadata.word!.noteReferences!.some((r) => r.customMark === "*"),
            ).toBe(true);
            bytes = exportDocx(reopened).bytes;
          }
          await page.evaluate(() => window.paginationHost.session!.undo());
          if (tracked) {
            await page.evaluate((rootId) => {
              const s = window.paginationHost.session!;
              s.commit(
                s.prepare({
                  protocolVersion: 1,
                  documentId: s.snapshot().manifest.documentId,
                  baseRevision: s.snapshot().manifest.revision,
                  operations: [{ op: "reviewWordRevisions", nodeId: rootId, decision: "reject" }],
                }),
              );
            }, rootId);
            expect(
              (
                await page.evaluate(() => window.paginationHost.session!.snapshot())
              ).metadata.word!.notes!.filter((n) => (n.type ?? "normal") === "normal"),
            ).toHaveLength(1);
            await page.evaluate(() => window.paginationHost.session!.undo());
          }
          await page.evaluate(() => window.paginationHost.session!.undo());
          expect(await page.evaluate(() => window.paginationHost.session!.snapshot().html)).toBe(
            before.html,
          );
          expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
        },
      );
}
