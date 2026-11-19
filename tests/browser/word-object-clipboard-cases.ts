import { expect, test } from "@playwright/test";
import type { Page } from "@playwright/test";
import { writeFileSync } from "node:fs";
import {
  createWordDocument,
  createDocumentSession,
  openDocument,
  wordDrawingTransforms,
} from "@doffi/core";
import type { DocumentPackage, WordContentSemantics } from "@doffi/core";
import { exportDocx, importDocx } from "@doffi/office";
import { wordObjectContentFixture } from "../../packages/office/tests/word-object-content-fixture.js";

export function registerWordObjectClipboardTests(
  mount: (page: Page, files: DocumentPackage, mode?: "edit" | "preview") => Promise<void>,
): void {
  for (const partial of [false, true])
    for (const tracked of [false, true])
      test(
        "object semantic system clipboard preserves current graph and owned text " +
          partial +
          " " +
          tracked,
        async ({ page }, testInfo) => {
          const source = wordObjectContentFixture(true),
            sourceWord = source.session.snapshot().metadata.word!;
          const selected = partial
            ? sourceWord.drawings!.find((d) => d.kind === "shape")!.nodeId
            : source.paragraphId;
          await mount(page, source.session.serialize(), tracked ? "edit" : "preview");
          const copied = await page
            .frameLocator("iframe")
            .locator('[data-node-id="' + selected + '"]')
            .evaluate((node, partial) => {
              const range = document.createRange();
              if (partial) {
                range.selectNodeContents(node.closest("[data-office-editor]")!);
                range.collapse(true);
              } else range.selectNode(node);
              const selection = document.getSelection()!;
              selection.removeAllRanges();
              selection.addRange(range);
              if (partial) {
                node.dispatchEvent(
                  new PointerEvent("pointerdown", { bubbles: true, altKey: true }),
                );
                node.ownerDocument.dispatchEvent(new PointerEvent("pointerup", { bubbles: true }));
              }
              const data = new DataTransfer(),
                event = new ClipboardEvent("copy", { bubbles: true, cancelable: true });
              Object.defineProperty(event, "clipboardData", { value: data });
              node.dispatchEvent(event);
              return { html: data.getData("text/html"), text: data.getData("text/plain") };
            }, partial && tracked);
          const payload = JSON.parse(
            Buffer.from(
              /data-doffi-word-clipboard="([^"]+)"/.exec(copied.html)![1]!,
              "base64",
            ).toString("utf8"),
          ) as { semantics: WordContentSemantics; resources?: { data: string }[] };
          expect(JSON.stringify(payload) + copied.html).not.toMatch(
            /PRIVATE-BODY|PRIVATE-TEXT-STORY|PRIVATE-WORKBOOK-CELL/,
          );
          expect(payload.semantics.stories).toHaveLength(partial ? 1 : 2);
          if (partial) expect(JSON.stringify(payload)).not.toMatch(/BOX-BODY|CURRENT-CHART/);
          for (const resource of payload.resources ?? [])
            expect(Buffer.from(resource.data, "base64").toString()).not.toContain(
              "PRIVATE-WORKBOOK-CELL",
            );
          const documentId = "object-clipboard-target",
            target = createDocumentSession(
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
          const entries = await page.evaluate(() =>
            [...window.paginationHost.session!.serialize()].map(
              ([p, b]) => [p, [...b]] as [string, number[]],
            ),
          );
          const model = openDocument(new Map(entries.map(([p, b]) => [p, new Uint8Array(b)])), {
              kind: "document",
            }),
            word = model.snapshot().metadata.word!,
            drawings = word.drawings!;
          expect(drawings).toHaveLength(payload.semantics.drawings!.length);
          expect(word.stories.filter((s) => s.kind === "textBox")).toHaveLength(partial ? 1 : 2);
          const sourceTransforms = wordDrawingTransforms(sourceWord, source.session),
            targetTransforms = wordDrawingTransforms(word, model);
          for (let i = 0; i < drawings.length; i++)
            for (let n = 0; n < 6; n++)
              expect(targetTransforms.world(drawings[i]!.nodeId)[n]).toBeCloseTo(
                sourceTransforms.world(payload.semantics.drawings![i]!.nodeId)[n]!,
                6,
              );
          if (!partial) {
            const chart = drawings.find((d) => d.kind === "chart")!;
            if (chart.kind !== "chart") throw new Error("Missing chart");
            expect(chart.chart.series[0]!.values).toEqual([-7, 23, 8]);
            expect(chart.chart.native?.workbookPart).toBeUndefined();
            const connector = drawings.find((d) => d.kind === "connector")!;
            if (connector.kind !== "connector") throw new Error("Missing connector");
            expect(drawings.some((d) => d.nodeId === connector.connector.startNodeId)).toBe(true);
          }
          let bytes = exportDocx(model).bytes;
          const output = testInfo.outputPath("object-clipboard.docx");
          writeFileSync(output, bytes);
          await testInfo.attach("object-clipboard.docx", {
            path: output,
            contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
          });
          for (let pass = 0; pass < 2; pass++) {
            const reopened = openDocument(
              importDocx(bytes, { documentId, unsupportedContent: "reject" }),
              { kind: "document" },
            );
            expect(reopened.snapshot().metadata.word!.drawings).toHaveLength(drawings.length);
            expect(
              reopened
                .snapshot()
                .metadata.word!.fields.some((f) => f.instruction.trim() === "PAGE"),
            ).toBe(true);
            bytes = exportDocx(reopened).bytes;
          }
          // Continue typing in the pasted object's owned story through the UI.
          const shape = drawings.find((d) => d.kind === "shape")!;
          if (!("textStoryId" in shape) || !shape.textStoryId)
            throw new Error("Missing text owner");
          const story = word.stories.find((s) => s.storyId === shape.textStoryId)!;
          await page
            .frameLocator("iframe")
            .locator('[data-node-id="' + shape.nodeId + '"]')
            .evaluate((node) => {
              node.dispatchEvent(new PointerEvent("pointerdown", { bubbles: true, altKey: true }));
              node.ownerDocument.dispatchEvent(new PointerEvent("pointerup", { bubbles: true }));
            });
          await page.evaluate(() =>
            (window.paginationHost.view as import("@doffi/editor").DocumentTextEditor).execute({
              type: "word",
              action: "drawingSettings",
            }),
          );
          await page.getByRole("button", { name: "编辑对象文字", exact: true }).click();
          const paragraph = page
            .frameLocator("iframe")
            .locator('[data-node-id="' + story.rootNodeId + '"] p')
            .filter({ hasText: /OWNED/ });
          await paragraph.evaluate((p) => {
            (p.closest("[data-office-editor]") as HTMLElement).focus();
            const r = document.createRange();
            r.selectNodeContents(p);
            r.collapse(false);
            const s = document.getSelection()!;
            s.removeAllRanges();
            s.addRange(r);
          });
          await page.keyboard.insertText("EDIT");
          await expect
            .poll(() =>
              page.evaluate(
                (id) => window.paginationHost.session!.getNode(id)!.text,
                story.rootNodeId,
              ),
            )
            .toContain("OWNED😀EDIT");
          await page.evaluate(() => window.paginationHost.session!.undo());
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
              (
                await page.evaluate(() => window.paginationHost.session!.snapshot())
              ).metadata.word!.stories.filter((s) => s.kind === "textBox"),
            ).toHaveLength(0);
            await page.evaluate(() => window.paginationHost.session!.undo());
          }
          await page.evaluate(() => window.paginationHost.session!.undo());
          expect(await page.evaluate(() => window.paginationHost.session!.snapshot().html)).toBe(
            before.html,
          );
        },
      );
}
