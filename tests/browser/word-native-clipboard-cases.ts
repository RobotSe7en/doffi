import { expect, test } from "@playwright/test";
import type { Page } from "@playwright/test";
import { writeFileSync } from "node:fs";
import { createWordDocument, createDocumentSession, openDocument } from "@doffi/core";
import type { DocumentPackage, WordContentSemantics } from "@doffi/core";
import { exportDocx, importDocx, readDocumentZip } from "@doffi/office";
import {
  nativeNumberingContentFixture,
  nativePreviewContentFixture,
  nativePropertyContentFixture,
  wordNativeContentFixture,
  legacyNativeContent,
} from "../../packages/office/tests/word-native-content-fixture.js";

type Clipboard = { html: string; text: string };
async function copy(page: Page, id: string): Promise<Clipboard> {
  return page
    .frameLocator("iframe")
    .locator('[data-node-id="' + id + '"]')
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
}
async function paste(page: Page, copied: Clipboard): Promise<void> {
  await page
    .frameLocator("iframe")
    .locator('[data-node-id="paragraph-1"]')
    .evaluate((node, copied) => {
      (node.closest("[data-office-editor]") as HTMLElement).focus();
      const range = document.createRange();
      range.selectNodeContents(node);
      range.collapse(true);
      const selection = document.getSelection()!;
      selection.removeAllRanges();
      selection.addRange(range);
      const data = new DataTransfer();
      data.setData("text/html", copied.html);
      data.setData("text/plain", copied.text);
      const event = new ClipboardEvent("paste", { bubbles: true, cancelable: true });
      Object.defineProperty(event, "clipboardData", { value: data });
      node.dispatchEvent(event);
    }, copied);
}
function payload(copied: Clipboard) {
  return JSON.parse(
    Buffer.from(/data-doffi-word-clipboard="([^"]+)"/.exec(copied.html)![1]!, "base64").toString(
      "utf8",
    ),
  ) as { semantics: WordContentSemantics; resources: { path: string; data: string }[] };
}
async function current(page: Page) {
  const entries = await page.evaluate(() =>
    [...window.paginationHost.session!.serialize()].map(
      ([p, b]) => [p, [...b]] as [string, number[]],
    ),
  );
  return openDocument(new Map(entries.map(([p, b]) => [p, new Uint8Array(b)])), {
    kind: "document",
  });
}
export function registerWordNativeClipboardTests(
  mount: (page: Page, files: DocumentPackage, mode?: "edit" | "preview") => Promise<void>,
): void {
  for (const legacy of [false, true])
    for (const tracked of [false, true])
      test(
        "native graph system clipboard slices legacy collections and reviews opaque blocks " +
          legacy +
          " " +
          tracked,
        async ({ page }, info) => {
          let files = importDocx(wordNativeContentFixture(), {
            documentId: "native-browser-source",
          });
          if (legacy) files = new Map(legacyNativeContent(files));
          const source = openDocument(files, { kind: "document" }),
            atom = source
              .snapshot()
              .metadata.word!.nativeContent!.fragments.find((f) => f.kind === "block")!.nodeId!;
          await mount(page, files, "preview");
          const copied = await copy(page, atom),
            data = payload(copied);
          expect(data.semantics.nativeContent!.fragments).toHaveLength(1);
          expect(data.semantics.nativeContent!.parts).toHaveLength(0);
          const transported = data.resources
            .map((r) => Buffer.from(r.data, "base64").toString("utf8"))
            .join("");
          expect(transported).toContain("SELECTED-NATIVE");
          expect(transported).not.toContain("PRIVATE-NATIVE-OUTSIDE");
          expect(transported).not.toContain("PRIVATE-UNBOUND-DATA");
          const documentId = "native-browser-receiver",
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
                  op: "setWordReview",
                  nodeId: "document-root",
                  trackChanges: tracked,
                  author: "Receiver",
                },
              ],
            }),
          );
          await mount(page, session.serialize(), "edit");
          const before = await page.evaluate(() => window.paginationHost.session!.snapshot().html);
          await paste(page, copied);
          await paste(page, copied);
          expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
          const model = await current(page),
            native = model.snapshot().metadata.word!.nativeContent!;
          expect(new Set(native.fragments.map((f) => f.nativeId)).size).toBe(2);
          expect(new Set(native.fragments.map((f) => f.sourcePart)).size).toBe(2);
          expect(
            model.snapshot().metadata.word!.revisions.filter((r) => r.kind === "nativeInsert"),
          ).toHaveLength(tracked ? 2 : 0);
          let bytes = exportDocx(model).bytes;
          const file = info.outputPath("native-block-clipboard.docx");
          writeFileSync(file, bytes);
          await info.attach("native-block-clipboard.docx", {
            path: file,
            contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
          });
          for (let pass = 0; pass < 2; pass++) {
            const reopened = openDocument(importDocx(bytes, { documentId }), { kind: "document" });
            expect(
              reopened
                .snapshot()
                .metadata.word!.nativeContent!.fragments.filter((f) => f.kind === "block"),
            ).toHaveLength(2);
            expect(
              reopened.snapshot().metadata.word!.revisions.filter((r) => r.kind === "nativeInsert"),
            ).toHaveLength(tracked ? 2 : 0);
            bytes = exportDocx(reopened).bytes;
          }
          if (tracked) {
            await page.evaluate(() =>
              (window.paginationHost.view as import("@doffi/editor").DocumentTextEditor).execute({
                type: "word",
                action: "review",
              }),
            );
            await page.getByRole("button", { name: "拒绝全部显示的修订", exact: true }).click();
            expect((await current(page)).snapshot().metadata.word!.nativeContent).toBeUndefined();
            await page.evaluate(() => window.paginationHost.session!.undo());
          }
          await page.evaluate(() => window.paginationHost.session!.undo());
          await page.evaluate(() => window.paginationHost.session!.undo());
          expect(await page.evaluate(() => window.paginationHost.session!.snapshot().html)).toBe(
            before,
          );
        },
      );
  for (const legacy of [false, true])
    test(
      "native graph system clipboard keeps only the bound datastore and isolates repeated binding identities " +
        legacy,
      async ({ page }, info) => {
        let files = importDocx(wordNativeContentFixture(true), {
          documentId: "bound-browser-source",
        });
        if (legacy) files = new Map(legacyNativeContent(files));
        const source = openDocument(files, { kind: "document" }),
          snapshot = source.snapshot(),
          bound =
            snapshot.metadata.word!.controls?.[0]?.nodeId ??
            snapshot.metadata.word!.nativeContent!.fragments.find((f) => f.kind === "block")!
              .nodeId!;
        await mount(page, files, "preview");
        const copied = await copy(page, bound),
          data = payload(copied),
          transported = data.resources
            .map((r) => Buffer.from(r.data, "base64").toString("utf8"))
            .join("");
        expect(transported).toContain("SELECTED-BOUND-DATA");
        expect(transported).not.toContain("PRIVATE-UNBOUND-DATA");
        expect(transported).not.toContain("PRIVATE-NATIVE-OUTSIDE");
        expect(data.semantics.nativeContent!.parts).toHaveLength(2);
        await mount(page, createWordDocument({ documentId: "bound-browser-target" }), "edit");
        await paste(page, copied);
        await paste(page, copied);
        expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
        let model = await current(page),
          bytes = exportDocx(model).bytes;
        const output = info.outputPath("bound-clipboard.docx");
        writeFileSync(output, bytes);
        await info.attach("bound-clipboard.docx", {
          path: output,
          contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        });
        for (let pass = 0; pass < 3; pass++) {
          const zip = readDocumentZip(bytes),
            text = [...zip]
              .filter(([p]) => p.endsWith(".xml"))
              .map(([, b]) => new TextDecoder().decode(b))
              .join("");
          expect(text).not.toContain("PRIVATE-UNBOUND-DATA");
          expect(text.match(/SELECTED-BOUND-DATA/g)).toHaveLength(2);
          const document = new TextDecoder().decode(zip.get("word/document.xml")!),
            ids = [...document.matchAll(/storeItemID="([^"]+)"/g)].map((m) => m[1]);
          expect(new Set(ids).size).toBe(2);
          const controls = [...document.matchAll(/<w:id[^>]*w:val="(\d+)"/g)].map((m) => m[1]);
          expect(controls).toHaveLength(2);
          expect(new Set(controls).size).toBe(2);
          model = openDocument(importDocx(bytes, { documentId: "bound-browser-target" }), {
            kind: "document",
          });
          const native = model.snapshot().metadata.word!.nativeContent!;
          expect(native.parts.filter((p) => p.customXml)).toHaveLength(2);
          bytes = exportDocx(model).bytes;
        }
        await page.evaluate(() => window.paginationHost.session!.undo());
        await page.evaluate(() => window.paginationHost.session!.undo());
        expect((await current(page)).snapshot().metadata.word!.nativeContent).toBeUndefined();
      },
    );
  test("native graph system clipboard preserves selected node and style properties through conflicts and further editing", async ({
    page,
  }, info) => {
    const source = openDocument(
        importDocx(nativePropertyContentFixture(), { documentId: "property-source" }),
        { kind: "document" },
      ),
      paragraph = [...source.snapshot().html.matchAll(/data-node-id="([^\"]+)"/g)]
        .map((m) => source.getNode(m[1]!)!)
        .find((n) => n.tagName === "p" && n.text === "PROPERTY-TEXT")!;
    await mount(page, source.serialize(), "preview");
    const copied = await copy(page, paragraph.nodeId),
      data = payload(copied);
    expect(data.semantics.nativeContent!.fragments.some((f) => f.styleId === "NativeStyle")).toBe(
      true,
    );
    expect(
      data.resources.map((r) => Buffer.from(r.data, "base64").toString("utf8")).join(""),
    ).not.toContain("PRIVATE-PROPERTY-OUTSIDE");
    const target = createDocumentSession(
      openDocument(createWordDocument({ documentId: "property-target" }), { kind: "document" }),
    );
    target.commit(
      target.prepare({
        protocolVersion: 1,
        documentId: "property-target",
        baseRevision: 0,
        operations: [
          {
            op: "setWordStyle",
            nodeId: "document-root",
            style: { styleId: "NativeStyle", kind: "paragraph", cssClass: "receiver-native-style" },
            runProperties: { color: "#ff0000" },
          },
        ],
      }),
    );
    await mount(page, target.serialize(), "edit");
    await paste(page, copied);
    expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
    let model = await current(page),
      native = model.snapshot().metadata.word!.nativeContent!;
    expect(native.fragments.find((f) => f.styleId)!.styleId).not.toBe("NativeStyle");
    const pasted = [...model.snapshot().html.matchAll(/data-node-id="([^\"]+)"/g)]
      .map((m) => model.getNode(m[1]!)!)
      .find((n) => n.tagName === "p" && n.text === "PROPERTY-TEXT")!;
    await page
      .frameLocator("iframe")
      .locator('[data-node-id="' + pasted.nodeId + '"]')
      .evaluate((p) => {
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
      .poll(async () => (await current(page)).getNode(pasted.nodeId)!.text)
      .toBe("PROPERTY-TEXTEDIT");
    model = await current(page);
    let bytes = exportDocx(model).bytes;
    const output = info.outputPath("native-properties-clipboard.docx");
    writeFileSync(output, bytes);
    await info.attach("native-properties-clipboard.docx", {
      path: output,
      contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    });
    for (let pass = 0; pass < 2; pass++) {
      const zip = readDocumentZip(bytes),
        document = new TextDecoder().decode(zip.get("word/document.xml")!),
        styles = new TextDecoder().decode(zip.get("word/styles.xml")!);
      expect(document).toContain("selected-property");
      expect(document).toContain("selected-run");
      expect(document).toContain("eastAsianLayout");
      expect(document).toContain("PROPERTY-TEXT");
      expect(document).toContain("EDIT");
      expect(styles).toContain("shadow");
      model = openDocument(importDocx(bytes, { documentId: "property-target" }), {
        kind: "document",
      });
      bytes = exportDocx(model).bytes;
    }
    const styleCount = (await current(page)).snapshot().metadata.word!.styles.length,
      sameDocumentCopy = await copy(page, pasted.nodeId);
    await paste(page, sameDocumentCopy);
    expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
    expect((await current(page)).snapshot().metadata.word!.styles).toHaveLength(styleCount);
    await page.evaluate(() => window.paginationHost.session!.undo());
    await page.evaluate(() => window.paginationHost.session!.undo());
    await page.evaluate(() => window.paginationHost.session!.undo());
    expect((await current(page)).snapshot().metadata.word!.nativeContent).toBeUndefined();
  });

  test("native graph system clipboard retains previews and bounded cyclic OPC closures with unique drawing identities", async ({
    page,
  }, info) => {
    const source = openDocument(
        importDocx(nativePreviewContentFixture(), { documentId: "preview-source" }),
        { kind: "document" },
      ),
      atom = source.snapshot().metadata.word!.nativeContent!.fragments.find((f) => f.preview)!;
    await mount(page, source.serialize(), "preview");
    await expect(
      page.frameLocator("iframe").locator('[data-node-id="' + atom.nodeId + '"] img'),
    ).toHaveCount(1);
    const copied = await copy(page, atom.nodeId!),
      data = payload(copied),
      native = data.semantics.nativeContent!,
      transported = data.resources
        .map((r) => Buffer.from(r.data, "base64").toString("utf8"))
        .join("");
    expect(native.fragments[0]!.preview).toBeDefined();
    expect(native.relationships.some((r) => r.external)).toBe(true);
    expect(transported).toContain("SELECTED-NATIVE-PAYLOAD");
    expect(transported).toContain("SELECTED-SHARED-DEPENDENCY");
    expect(transported).not.toContain("PRIVATE-NATIVE-PAYLOAD");
    await mount(page, createWordDocument({ documentId: "preview-target" }), "edit");
    await paste(page, copied);
    await paste(page, copied);
    expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
    await expect(
      page.frameLocator("iframe").locator('[data-sdk-native-preview="true"] img'),
    ).toHaveCount(2);
    let model = await current(page),
      bytes = exportDocx(model).bytes;
    const path = info.outputPath("native-preview-clipboard.docx");
    writeFileSync(path, bytes);
    await info.attach("native-preview-clipboard.docx", {
      path,
      contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    });
    for (let pass = 0; pass < 2; pass++) {
      const files = readDocumentZip(bytes),
        document = new TextDecoder().decode(files.get("word/document.xml")!),
        frames = [...document.matchAll(/<wp:docPr[^>]*\bid="(\d+)"/g)].map((m) => m[1]);
      expect(frames).toHaveLength(2);
      expect(new Set(frames).size).toBe(2);
      const content = [...files.values()].map((b) => new TextDecoder().decode(b)).join("");
      expect(content).not.toContain("PRIVATE-NATIVE-PAYLOAD");
      expect(content).toContain("SELECTED-SHARED-DEPENDENCY");
      model = openDocument(importDocx(bytes, { documentId: "preview-target" }), {
        kind: "document",
      });
      expect(
        model.snapshot().metadata.word!.nativeContent!.fragments.filter((f) => f.preview),
      ).toHaveLength(2);
      bytes = exportDocx(model).bytes;
    }
    await page.evaluate(() => window.paginationHost.session!.undo());
    await page.evaluate(() => window.paginationHost.session!.undo());
    expect((await current(page)).snapshot().metadata.word!.nativeContent).toBeUndefined();
  });

  test("native graph system clipboard isolates native numbering and style dependencies without unselected definitions", async ({
    page,
  }, info) => {
    const source = openDocument(
        importDocx(nativeNumberingContentFixture(), { documentId: "native-number-source" }),
        { kind: "document" },
      ),
      atom = source
        .snapshot()
        .metadata.word!.nativeContent!.fragments.find((f) => f.kind === "block")!;
    await mount(page, source.serialize(), "preview");
    const copied = await copy(page, atom.nodeId!),
      data = payload(copied),
      native = data.semantics.nativeContent!;
    expect(native.fragments.filter((f) => f.numberingKind === "num")).toHaveLength(2);
    expect(native.fragments.filter((f) => f.numberingKind === "abstractNum")).toHaveLength(1);
    expect(
      data.resources.map((r) => Buffer.from(r.data, "base64").toString("utf8")).join(""),
    ).not.toContain("PRIVATE-UNSELECTED-NUMBERING");
    await mount(page, createWordDocument({ documentId: "native-number-target" }), "edit");
    await paste(page, copied);
    await paste(page, copied);
    expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
    let model = await current(page),
      bytes = exportDocx(model).bytes;
    const path = info.outputPath("native-numbering-clipboard.docx");
    writeFileSync(path, bytes);
    await info.attach("native-numbering-clipboard.docx", {
      path,
      contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    });
    for (let pass = 0; pass < 3; pass++) {
      const files = readDocumentZip(bytes),
        document = new TextDecoder().decode(files.get("word/document.xml")!),
        numbering = new TextDecoder().decode(files.get("word/numbering.xml")!),
        ids = [...document.matchAll(/<w:numId[^>]*w:val="(\d+)"/g)].map((m) => m[1]);
      expect(ids).toHaveLength(4);
      expect(new Set(ids).size).toBe(4);
      expect(numbering).not.toContain("PRIVATE-UNSELECTED-NUMBERING");
      expect(numbering).toContain("startOverride");
      expect(new TextDecoder().decode(files.get("word/styles.xml")!)).toContain("shadow");
      model = openDocument(importDocx(bytes, { documentId: "native-number-target" }), {
        kind: "document",
      });
      bytes = exportDocx(model).bytes;
    }
    await page.evaluate(() => window.paginationHost.session!.undo());
    expect(
      (await current(page))
        .snapshot()
        .metadata.word!.nativeContent!.fragments.filter((f) => f.numberingKind === "num"),
    ).toHaveLength(2);
    await page.evaluate(() => window.paginationHost.session!.undo());
    expect((await current(page)).snapshot().metadata.word!.nativeContent).toBeUndefined();
  });
}
