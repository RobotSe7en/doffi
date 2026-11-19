import { expect, test } from "vitest";
import { createWordDocument, createDocumentSession, openDocument } from "@doffi/core";
test("creates a serializable A4 Word source with shared history", () => {
  const model = openDocument(createWordDocument({ documentId: "new-word" }), { kind: "document" });
  expect(model.snapshot().metadata.pageSetup!.width).toBeCloseTo(595.2756);
  expect(model.snapshot().metadata.locale).toBe("zh-CN");
  const session = createDocumentSession(model);
  session.commit(
    session.prepare({
      protocolVersion: 1,
      documentId: "new-word",
      baseRevision: 0,
      operations: [
        {
          op: "replaceText",
          nodeId: "paragraph-1",
          from: 0,
          to: 0,
          expectedText: "",
          text: "新文档😀",
        },
      ],
    }),
  );
  expect(session.getNode("paragraph-1")!.text).toBe("新文档😀");
  session.undo();
  expect(session.getNode("paragraph-1")!.text).toBe("");
  session.redo();
  expect(openDocument(session.serialize(), { kind: "document" }).getNode("paragraph-1")!.text).toBe(
    "新文档😀",
  );
});
test("new Word documents retain authorized page setup and safe titles", () => {
  const model = openDocument(
    createWordDocument({
      documentId: "setup",
      title: "<季度>&报告😀",
      pageSetup: {
        unit: "pt",
        width: 800,
        height: 600,
        margins: { top: 20, right: 30, bottom: 40, left: 50 },
      },
    }),
    { kind: "document" },
  );
  expect(model.snapshot().metadata.title).toBe("<季度>&报告😀");
  expect(model.snapshot().metadata.pageSetup!.margins.left).toBe(50);
  expect(() => createWordDocument({ documentId: "bad", title: "\ud800" })).toThrow();
});
