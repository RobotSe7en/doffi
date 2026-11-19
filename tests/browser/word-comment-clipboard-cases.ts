import { expect, test } from "@playwright/test";
import type { Page } from "@playwright/test";
import { writeFileSync } from "node:fs";
import { createWordDocument, createDocumentSession, openDocument } from "@doffi/core";
import type { DocumentPackage, WordContentSemantics } from "@doffi/core";
import { exportDocx, importDocx } from "@doffi/office";
import { wordCommentContentFixture } from "../../packages/office/tests/word-comment-content-fixture.js";
import { wordNoteContentFixture } from "../../packages/office/tests/word-note-content-fixture.js";
import { nativeCommentFixture } from "../../packages/office/tests/native-comment-fixture.js";

type Clipboard = { html: string; text: string };
async function paste(page: Page, copied: Clipboard) {
  await page
    .frameLocator("iframe")
    .locator('[data-node-id="paragraph-1"]')
    .evaluate((paragraph, copied) => {
      (paragraph.closest("[data-office-editor]") as HTMLElement).focus();
      const range = document.createRange();
      range.selectNodeContents(paragraph);
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
}
export function registerWordCommentClipboardTests(
  mount: (page: Page, files: DocumentPackage, mode?: "edit" | "preview") => Promise<void>,
): void {
  for (const native of [false, true])
    for (const tracked of [false, true])
      test(
        "comment semantic system clipboard clips exact repeated UTF-16 anchors and rich reply bodies " +
          native +
          " " +
          tracked,
        async ({ page }, info) => {
          const source = wordCommentContentFixture(native);
          await mount(page, source.session.serialize(), tracked ? "edit" : "preview");
          const copied = await page
            .frameLocator("iframe")
            .locator('[data-node-id="' + source.paragraphId + '"]')
            .evaluate(
              (paragraph, args) => {
                const walker = document.createTreeWalker(paragraph, NodeFilter.SHOW_TEXT),
                  leaves: Text[] = [];
                let node: Node | null;
                while ((node = walker.nextNode())) leaves.push(node as Text);
                const at = (offset: number): [Node, number] => {
                  for (const leaf of leaves) {
                    if (offset <= leaf.length) return [leaf, offset];
                    offset -= leaf.length;
                  }
                  throw new Error("Missing exact source selection");
                };
                const range = document.createRange();
                range.setStart(...at(args.from));
                range.setEnd(...at(args.to));
                const selection = document.getSelection()!;
                selection.removeAllRanges();
                selection.addRange(range);
                const data = new DataTransfer(),
                  event = new ClipboardEvent("copy", { bubbles: true, cancelable: true });
                Object.defineProperty(event, "clipboardData", { value: data });
                paragraph.dispatchEvent(event);
                return { html: data.getData("text/html"), text: data.getData("text/plain") };
              },
              { from: source.from, to: source.to },
            );
          expect(copied.text).toBe(source.text.slice(source.from, source.to));
          const payload = JSON.parse(
            Buffer.from(
              /data-doffi-word-clipboard="([^"]+)"/.exec(copied.html)![1]!,
              "base64",
            ).toString("utf8"),
          ) as { semantics: WordContentSemantics };
          expect(JSON.stringify(payload) + copied.html).not.toContain("PRIVATE-COMMENT");
          expect(payload.semantics.annotations).toHaveLength(2);
          expect(payload.semantics.stories).toHaveLength(native ? 4 : 3);
          expect(payload.semantics.fields).toHaveLength(1);
          const documentId = "comment-browser-target",
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
                  text: "TARGET",
                },
                {
                  op: "addAnnotation",
                  annotation: {
                    annotationId: "thread",
                    documentId,
                    baseRevision: 0,
                    status: "open",
                    anchor: { kind: "node", nodeId: "paragraph-1" },
                    instruction: "接收方批注",
                  },
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
          await paste(page, copied);
          expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
          const snapshot = await page.evaluate(() => window.paginationHost.session!.snapshot()),
            thread = snapshot.annotations.find((a) => a.author?.name === "甲")!;
          expect(snapshot.annotations).toHaveLength(3);
          expect(thread.replies).toHaveLength(2);
          expect(thread.status).toBe("resolved");
          const anchor = thread.anchor;
          expect(
            anchor.kind === "text"
              ? anchor.quote
              : anchor.kind === "text-set"
                ? anchor.ranges.map((r) => r.quote).join("")
                : "",
          ).toBe(copied.text);
          const second = snapshot.annotations.find(
            (a) => a.annotationId !== "thread" && a.annotationId !== thread.annotationId,
          )!;
          expect(
            second.anchor.kind === "text" && [
              second.anchor.from,
              second.anchor.to,
              second.anchor.quote,
            ],
          ).toEqual([5, 7, "重复"]);
          await page.evaluate(
            (id) =>
              (window.paginationHost.view as import("@doffi/editor").DocumentTextEditor).editStory(
                id,
              ),
            snapshot.metadata.word!.stories.find((s) => s.rootNodeId === thread.bodyNodeId)!
              .storyId,
          );
          const body = page
            .frameLocator("iframe")
            .locator('[data-node-id="' + thread.bodyNodeId + '"] p');
          await body.evaluate((p) => {
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
                thread.bodyNodeId!,
              ),
            )
            .toContain("链接EDIT");
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
          const output = info.outputPath("comment-clipboard.docx");
          writeFileSync(output, bytes);
          await info.attach("comment-clipboard.docx", {
            path: output,
            contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
          });
          for (let pass = 0; pass < 2; pass++) {
            const reopened = openDocument(
                importDocx(bytes, { documentId, unsupportedContent: "reject" }),
                { kind: "document" },
              ),
              root = reopened.snapshot().annotations.find((a) => a.author?.name === "甲")!;
            expect(root.replies).toHaveLength(2);
            expect(reopened.getNode(root.bodyNodeId!)!.text).toContain("链接EDIT");
            bytes = exportDocx(reopened).bytes;
          }
          await page.evaluate(() => window.paginationHost.session!.undo());
          await page.evaluate(() => window.paginationHost.session!.undo());
          expect(await page.evaluate(() => window.paginationHost.session!.snapshot().html)).toBe(
            before.html,
          );
          expect(
            await page.evaluate(() =>
              window.paginationHost.session!.getAnnotations().map((a) => a.annotationId),
            ),
          ).toEqual(["thread"]);
          await page.evaluate(() => window.paginationHost.session!.redo());
          await page.evaluate(() =>
            (window.paginationHost.view as import("@doffi/editor").DocumentTextEditor).editStory(),
          );
          await paste(page, copied);
          expect(
            await page.evaluate(() => window.paginationHost.session!.getAnnotations().length),
          ).toBe(5);
          expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
        },
      );

  for (const inline of [false, true])
    test(
      "comment semantic system clipboard preserves opaque internal controls across colliding native sources " +
        inline,
      async ({ page }, info) => {
        const source = openDocument(
            importDocx(nativeCommentFixture("inside", inline), { documentId: "opaque-source" }),
            { kind: "document" },
          ),
          atom = source
            .snapshot()
            .metadata.word!.nativeContent!.fragments.find((f) => f.kind !== "property")!.nodeId!;
        await mount(page, source.serialize(), "preview");
        const copied = await page
          .frameLocator("iframe")
          .locator('[data-node-id="' + atom + '"]')
          .evaluate((node) => {
            const range = document.createRange();
            range.selectNode(node);
            const selection = document.getSelection()!;
            selection.removeAllRanges();
            selection.addRange(range);
            const data = new DataTransfer(),
              event = new ClipboardEvent("copy", { bubbles: true, cancelable: true });
            Object.defineProperty(event, "clipboardData", { value: data });
            node.dispatchEvent(event);
            return { html: data.getData("text/html"), text: data.getData("text/plain") };
          });
        const payload = JSON.parse(
          Buffer.from(
            /data-doffi-word-clipboard="([^"]+)"/.exec(copied.html)![1]!,
            "base64",
          ).toString("utf8"),
        ) as { semantics: WordContentSemantics };
        expect(payload.semantics.annotations).toHaveLength(1);
        expect(JSON.stringify(payload)).not.toMatch(/BEFORE|AFTER/);
        const documentId = "opaque-target",
          target = createDocumentSession(
            openDocument(createWordDocument({ documentId }), { kind: "document" }),
          );
        await mount(page, target.serialize(), "edit");
        await paste(page, copied);
        await paste(page, copied);
        expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
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
        const output = info.outputPath("comment-opaque-clipboard.docx");
        writeFileSync(output, bytes);
        await info.attach("comment-opaque-clipboard.docx", {
          path: output,
          contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        });
        for (let pass = 0; pass < 2; pass++) {
          const reopened = openDocument(importDocx(bytes, { documentId }), { kind: "document" });
          expect(reopened.snapshot().annotations).toHaveLength(2);
          expect(
            reopened
              .snapshot()
              .metadata.word!.nativeContent!.fragments.filter((f) => f.kind !== "property"),
          ).toHaveLength(2);
          bytes = exportDocx(reopened).bytes;
        }
      },
    );
  test("comment semantic system clipboard clips cross-paragraph text sets without carrying unselected messages", async ({
    page,
  }, info) => {
    const documentId = "cross-comments",
      source = createDocumentSession(
        openDocument(createWordDocument({ documentId }), { kind: "document" }),
      );
    source.commit(
      source.prepare({
        protocolVersion: 1,
        documentId,
        baseRevision: 0,
        operations: [
          {
            op: "replaceFragment",
            nodeId: "section-root-1",
            html: "<p>PRE重复X重复</p><p>😀禰\u{e0100}尾</p>",
          },
        ],
      }),
    );
    const [first, last] = source.getNode("section-root-1")!.childNodeIds;
    source.commit(
      source.prepare({
        protocolVersion: 1,
        documentId,
        baseRevision: source.snapshot().manifest.revision,
        operations: [
          {
            op: "addAnnotation",
            annotation: {
              annotationId: "cross",
              documentId,
              baseRevision: source.snapshot().manifest.revision,
              status: "open",
              anchor: {
                kind: "text-set",
                ranges: [
                  { nodeId: first!, from: 3, to: 8, quote: "重复X重复" },
                  { nodeId: last!, from: 0, to: 5, quote: "😀禰\u{e0100}" },
                ],
              },
              instruction: "跨段批注",
            },
          },
        ],
      }),
    );
    await mount(page, source.serialize(), "edit");
    const copied = await page
      .frameLocator("iframe")
      .locator('[data-node-id="' + first + '"]')
      .evaluate((p, last) => {
        const end = p.parentElement!.querySelector('[data-node-id="' + last + '"]')!;
        const at = (root: Element, offset: number): [Node, number] => {
          const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
          let node: Node | null;
          while ((node = walker.nextNode())) {
            if (node.parentElement?.closest("[data-sdk-word-autospace]")) continue;
            if (offset <= node.textContent!.length) return [node, offset];
            offset -= node.textContent!.length;
          }
          throw new Error("Missing source endpoint");
        };
        const range = document.createRange();
        range.setStart(...at(p, 6));
        range.setEnd(...at(end, 5));
        const selection = document.getSelection()!;
        selection.removeAllRanges();
        selection.addRange(range);
        const data = new DataTransfer(),
          event = new ClipboardEvent("copy", { bubbles: true, cancelable: true });
        Object.defineProperty(event, "clipboardData", { value: data });
        p.dispatchEvent(event);
        return { html: data.getData("text/html"), text: data.getData("text/plain") };
      }, last!);
    const target = openDocument(createWordDocument({ documentId: "cross-target" }), {
      kind: "document",
    });
    await mount(page, target.serialize(), "edit");
    await paste(page, copied);
    expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
    const comment = await page.evaluate(() => window.paginationHost.session!.getAnnotations()[0]!);
    expect(comment.anchor.kind === "text-set" && comment.anchor.ranges.map((r) => r.quote)).toEqual(
      ["重复", "😀禰\u{e0100}"],
    );
    const entries = await page.evaluate(() =>
      [...window.paginationHost.session!.serialize()].map(
        ([p, b]) => [p, [...b]] as [string, number[]],
      ),
    );
    const bytes = exportDocx(
      openDocument(new Map(entries.map(([p, b]) => [p, new Uint8Array(b)])), { kind: "document" }),
    ).bytes;
    const output = info.outputPath("comment-cross-clipboard.docx");
    writeFileSync(output, bytes);
    await info.attach("comment-cross-clipboard.docx", {
      path: output,
      contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    });
    const reopened = openDocument(
      importDocx(bytes, { documentId: "cross-reopen", unsupportedContent: "reject" }),
      { kind: "document" },
    );
    expect(reopened.snapshot().annotations).toHaveLength(1);
    const reopenedAnchor = reopened.snapshot().annotations[0]!.anchor;
    expect(
      reopenedAnchor.kind === "text-set" && reopenedAnchor.ranges.map((r) => r.quote).join(""),
    ).toBe("重复😀禰\u{e0100}");
  });

  test("comment semantic system clipboard follows selected note and textbox ownership", async ({
    page,
  }, info) => {
    const fixture = wordNoteContentFixture(),
      source = fixture.session,
      snapshot = source.snapshot(),
      documentId = snapshot.manifest.documentId,
      foot = snapshot.metadata.word!.notes!.find(
        (n) =>
          n.kind === "footnote" &&
          source
            .getNode(
              snapshot.metadata.word!.stories.find((s) => s.storyId === n.storyId)!.rootNodeId,
            )!
            .text.includes("FOOT-BODY"),
      )!,
      noteRoot = snapshot.metadata.word!.stories.find(
        (s) => s.storyId === foot.storyId,
      )!.rootNodeId,
      noteP = source.getNode(noteRoot)!.childNodeIds.at(-1)!,
      textStory = snapshot.metadata.word!.stories.find((s) => s.kind === "textBox")!,
      textP = source.getNode(textStory.rootNodeId)!.childNodeIds[0]!;
    source.commit(
      source.prepare({
        protocolVersion: 1,
        documentId,
        baseRevision: snapshot.manifest.revision,
        operations: [
          {
            op: "addAnnotation",
            annotation: {
              annotationId: "note-comment",
              documentId,
              baseRevision: snapshot.manifest.revision,
              status: "open",
              anchor: { kind: "text", nodeId: noteP, from: 0, to: 4, quote: "LIST" },
              instruction: "所选脚注批注",
            },
          },
          {
            op: "addAnnotation",
            annotation: {
              annotationId: "textbox-comment",
              documentId,
              baseRevision: snapshot.manifest.revision,
              status: "open",
              anchor: { kind: "node", nodeId: textP },
              instruction: "所选文本框批注",
            },
          },
          {
            op: "addAnnotation",
            annotation: {
              annotationId: "unselected-comment",
              documentId,
              baseRevision: snapshot.manifest.revision,
              status: "open",
              anchor: { kind: "node", nodeId: fixture.endId },
              instruction: "PRIVATE-UNSELECTED-COMMENT",
            },
          },
        ],
      }),
    );
    await mount(page, source.serialize(), "preview");
    const copied = await page
      .frameLocator("iframe")
      .locator('[data-node-id="' + fixture.footId + '"]')
      .evaluate((node) => {
        const range = document.createRange();
        range.selectNode(node);
        const selection = document.getSelection()!;
        selection.removeAllRanges();
        selection.addRange(range);
        const data = new DataTransfer(),
          event = new ClipboardEvent("copy", { bubbles: true, cancelable: true });
        Object.defineProperty(event, "clipboardData", { value: data });
        node.dispatchEvent(event);
        return { html: data.getData("text/html"), text: data.getData("text/plain") };
      });
    expect(copied.html).not.toContain("PRIVATE-");
    const payload = JSON.parse(
      Buffer.from(/data-doffi-word-clipboard="([^"]+)"/.exec(copied.html)![1]!, "base64").toString(
        "utf8",
      ),
    ) as { semantics: WordContentSemantics };
    expect(JSON.stringify(payload)).not.toContain("PRIVATE-");
    expect(payload.semantics.annotations).toHaveLength(2);
    await mount(page, createWordDocument({ documentId: "owned-comment-target" }), "edit");
    await paste(page, copied);
    expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
    expect(
      await page.evaluate(() =>
        window.paginationHost.session!.getAnnotations().map((a) => a.instruction),
      ),
    ).toEqual(["所选脚注批注", "所选文本框批注"]);
    const entries = await page.evaluate(() =>
      [...window.paginationHost.session!.serialize()].map(
        ([p, b]) => [p, [...b]] as [string, number[]],
      ),
    );
    const bytes = exportDocx(
      openDocument(new Map(entries.map(([p, b]) => [p, new Uint8Array(b)])), { kind: "document" }),
    ).bytes;
    const output = info.outputPath("comment-owned-clipboard.docx");
    writeFileSync(output, bytes);
    await info.attach("comment-owned-clipboard.docx", {
      path: output,
      contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    });
    const reopened = openDocument(
      importDocx(bytes, { documentId: "owned-reopen", unsupportedContent: "reject" }),
      { kind: "document" },
    );
    expect(reopened.snapshot().annotations).toHaveLength(2);
  });

  test("comment semantic system clipboard maps a complete displayed field without reviving its old cache", async ({
    page,
  }) => {
    const documentId = "field-comments",
      source = createDocumentSession(
        openDocument(createWordDocument({ documentId }), { kind: "document" }),
      );
    source.commit(
      source.prepare({
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
            text: "前后",
          },
          {
            op: "insertWordField",
            nodeId: "paragraph-1",
            offset: 1,
            instruction: " PAGE ",
            cachedText: "999",
          },
        ],
      }),
    );
    source.commit(
      source.prepare({
        protocolVersion: 1,
        documentId,
        baseRevision: source.snapshot().manifest.revision,
        operations: [
          {
            op: "addAnnotation",
            annotation: {
              annotationId: "field-comment",
              documentId,
              baseRevision: source.snapshot().manifest.revision,
              status: "open",
              anchor: { kind: "text", nodeId: "paragraph-1", from: 0, to: 5, quote: "前999后" },
              instruction: "字段批注",
            },
          },
        ],
      }),
    );
    source.commit(
      source.prepare({
        protocolVersion: 1,
        documentId,
        baseRevision: source.snapshot().manifest.revision,
        operations: [
          {
            op: "addAnnotation",
            annotation: {
              annotationId: "inner-field-comment",
              documentId,
              baseRevision: source.snapshot().manifest.revision,
              status: "open",
              anchor: { kind: "text", nodeId: "paragraph-1", from: 2, to: 3, quote: "9" },
              instruction: "字段内部待定位",
            },
          },
        ],
      }),
    );
    await mount(page, source.serialize(), "edit");
    const copied = await page
      .frameLocator("iframe")
      .locator('[data-node-id="paragraph-1"]')
      .evaluate((p) => {
        const range = document.createRange();
        range.selectNodeContents(p);
        const selection = document.getSelection()!;
        selection.removeAllRanges();
        selection.addRange(range);
        const data = new DataTransfer(),
          event = new ClipboardEvent("copy", { bubbles: true, cancelable: true });
        Object.defineProperty(event, "clipboardData", { value: data });
        p.dispatchEvent(event);
        return { html: data.getData("text/html"), text: data.getData("text/plain") };
      });
    expect(copied.text).toBe("前1后");
    expect(copied.html).not.toContain("999");
    await mount(page, createWordDocument({ documentId: "field-comment-target" }), "edit");
    await paste(page, copied);
    expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
    const comment = await page.evaluate(() => window.paginationHost.session!.getAnnotations()[0]!);
    expect(comment.anchor.kind === "text" && comment.anchor.quote).toBe("前1后");
    const recovery = await page.evaluate(() => window.paginationHost.session!.getAnnotations()[1]!);
    expect(recovery.status).toBe("needs_reanchor");
    expect(recovery.anchor.kind).toBe("node");
  });

  test("comment semantic system clipboard counts hidden deletion source before repeated selected text", async ({
    page,
  }) => {
    const documentId = "hidden-comment-source",
      source = createDocumentSession(
        openDocument(createWordDocument({ documentId }), { kind: "document" }),
      );
    const apply = (operations: import("@doffi/core").DocumentOperation[]) =>
      source.commit(
        source.prepare({
          protocolVersion: 1,
          documentId,
          baseRevision: source.snapshot().manifest.revision,
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
        text: "重复重复重复",
      },
    ]);
    apply([
      { op: "setWordReview", nodeId: "document-root", trackChanges: true, author: "Deleter" },
      { op: "replaceText", nodeId: "paragraph-1", from: 2, to: 4, expectedText: "重复", text: "" },
    ]);
    apply([
      {
        op: "addAnnotation",
        annotation: {
          annotationId: "last-comment",
          documentId,
          baseRevision: source.snapshot().manifest.revision,
          status: "open",
          anchor: { kind: "text", nodeId: "paragraph-1", from: 4, to: 6, quote: "重复" },
          instruction: "最后重复的批注",
        },
      },
    ]);
    await mount(page, source.serialize(), "edit");
    await page.evaluate(() =>
      (window.paginationHost.view as import("@doffi/editor").DocumentTextEditor).execute({
        type: "word",
        action: "review",
      }),
    );
    await page.getByLabel("修订显示", { exact: true }).selectOption("final");
    const copied = await page
      .frameLocator("iframe")
      .locator('[data-node-id="paragraph-1"]')
      .evaluate((p) => {
        const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT),
          visible: Text[] = [];
        let node: Node | null;
        while ((node = walker.nextNode()))
          if (!node.parentElement?.closest("del,[data-sdk-word-deleted]"))
            visible.push(node as Text);
        const last = visible.at(-1)!;
        const range = document.createRange();
        range.selectNodeContents(last);
        const selection = document.getSelection()!;
        selection.removeAllRanges();
        selection.addRange(range);
        const data = new DataTransfer(),
          event = new ClipboardEvent("copy", { bubbles: true, cancelable: true });
        Object.defineProperty(event, "clipboardData", { value: data });
        p.dispatchEvent(event);
        return { html: data.getData("text/html"), text: data.getData("text/plain") };
      });
    expect(copied.text).toBe("重复");
    await mount(page, createWordDocument({ documentId: "hidden-comment-target" }), "edit");
    await paste(page, copied);
    expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
    const comments = await page.evaluate(() => window.paginationHost.session!.getAnnotations());
    expect(comments).toHaveLength(1);
    expect(
      comments[0]!.anchor.kind === "text" && [
        comments[0]!.anchor.from,
        comments[0]!.anchor.to,
        comments[0]!.anchor.quote,
      ],
    ).toEqual([0, 2, "重复"]);
  });
}
