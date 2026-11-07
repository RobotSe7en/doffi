import { openContentFixture as openDocument } from "../../../tests/content-fixture.js";
import { readFileSync } from "node:fs";
import { expect, test } from "vitest";
import { createDocumentSession } from "@doffi/core";
import type { EditProposal } from "@doffi/core";
import { importDocx, readDocumentZip, writeDocumentZip } from "@doffi/office";

const directory = new URL("../../../examples/docx/", import.meta.url);
const cases = JSON.parse(readFileSync(new URL("expected.json", directory), "utf8")) as { file: string }[];
for (const file of cases.flatMap((item) => [item.file, "producer/" + item.file])) {
  test("request/proposal application preserves imported DOCX semantics and ZIP reopening: " + file, () => {
    const session = createDocumentSession(openDocument(importDocx(readFileSync(new URL(file, directory)), { documentId: "proposal-docx" })));
    const source = session.snapshot();
    const ids = [...source.html.matchAll(/data-node-id="([^"]+)"/g)].map((match) => match[1]!);
    const target = ids.map((id) => session.getNode(id)!).find((node) => node.tagName === "p" && node.text.trim())!;
    session.commit(session.prepare({ protocolVersion: 1, documentId: "proposal-docx", baseRevision: 0, operations: [{ op: "addAnnotation",
      annotation: { annotationId: "style-note", documentId: "proposal-docx", baseRevision: 0, status: "open",
        instruction: "文字使用深蓝色", anchor: { kind: "node", nodeId: target.nodeId } } }] }));
    const request = session.buildInteractionRequest({ annotationIds: ["style-note"] });
    const proposal: EditProposal = { protocolVersion: 1, requestId: request.requestId, proposalId: "fixed-style", documentId: "proposal-docx",
      baseRevision: 1, annotationIds: ["style-note"], summary: "调整文字颜色", operations: [
        { op: "setStyle", nodeId: target.nodeId, property: "color", value: "#1f4b99" },
      ] };
    session.commit(session.prepareProposal(request, proposal));
    expect(session.getAnnotations().find((annotation) => annotation.annotationId === "style-note")!.status).toBe("resolved");
    expect(session.snapshot().metadata).toEqual(source.metadata);
    expect(session.snapshot().styles).toEqual(source.styles);
    expect(openDocument(readDocumentZip(writeDocumentZip(session.serialize()))).snapshot()).toEqual(session.snapshot());
    session.undo(); expect(session.snapshot().html).toBe(source.html);
    expect(session.getAnnotations().find((annotation) => annotation.annotationId === "style-note")!.status).toBe("open");
    session.redo(); expect(session.snapshot().html).toContain("color:#1f4b99");
  });
}

test("native rich-comment requests include the complete body and associated Word metadata", () => {
  const session = createDocumentSession(openDocument(importDocx(readFileSync(new URL("08-comments-revisions.docx", directory)), { documentId: "comments" })));
  const annotation = session.getAnnotations()[0]!, request = session.buildInteractionRequest({ annotationIds: [annotation.annotationId] });
  expect(annotation.bodyNodeId).toBeDefined();
  expect(request.context.blocks.some((block) => block.html.includes('data-node-id="' + annotation.bodyNodeId + '"'))).toBe(true);
  expect(request.annotations[0]!.author!.name).toBe("审阅者");
  expect(request.context.metadata.word!.revisions).toHaveLength(2);
});
