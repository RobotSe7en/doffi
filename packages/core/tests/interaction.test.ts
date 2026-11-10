import { openContentFixture as openDocument } from "../../../tests/content-fixture.js";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { expect, test } from "vitest";
import { createDocumentSession, DocumentError, parseContract } from "@doffi/core";
import type { Annotation, DocumentInteractionRequest, DocumentOperation, DocumentSession, EditProposal } from "@doffi/core";
import { proposal as exampleProposal, request as exampleRequest, repository } from "./fixtures.js";

function session(): DocumentSession {
  return createDocumentSession(openDocument(new Map(["manifest.json", "content.html", "styles.css", "metadata.json", "annotations.json"]
    .map((path) => [path, new Uint8Array(readFileSync(join(repository, "examples/report", path)))]))));
}
function request(s: DocumentSession, annotationIds = ["a-summary", "a-heading"]): DocumentInteractionRequest {
  return s.buildInteractionRequest({ annotationIds });
}
function proposal(r: DocumentInteractionRequest, operations = exampleProposal.operations): EditProposal {
  return { ...structuredClone(exampleProposal), requestId: r.requestId, documentId: r.documentId,
    baseRevision: r.baseRevision, annotationIds: r.annotations.map((a) => a.annotationId), operations };
}
function commit(s: DocumentSession, operations: DocumentOperation[]): void {
  s.commit(s.prepare({ protocolVersion: 1, documentId: s.snapshot().manifest.documentId,
    baseRevision: s.snapshot().manifest.revision, operations }));
}
function fails(action: () => unknown, code: string): void {
  try { action(); } catch (error) {
    expect(error).toBeInstanceOf(DocumentError);
    expect((error as DocumentError).code).toBe(code);
    return;
  }
  throw new Error("Expected " + code);
}

test("the public report request and proposal execute through the same collaboration entry", () => {
  const s = session();
  s.commit(s.prepareProposal(exampleRequest, exampleProposal));
  expect(s.getNode("n-summary")!.text).toContain("第四季度");
});

test("requests freeze complete related subtrees and do not change document, history or events", () => {
  const s = session(), before = s.serialize(), events: string[] = [];
  s.subscribe((event) => events.push(event.type));
  const r = request(s);
  expect(r.context.blocks.map((block) => block.nodeId)).toEqual(["n-title", "n-summary"]);
  for (const block of r.context.blocks) expect(block.html).toBe(s.getNode(block.nodeId)!.html);
  expect(r.context.styles).toEqual(s.snapshot().styles);
  expect(r.context.metadata).toEqual(s.snapshot().metadata);
  expect(r.context.assets).toEqual(s.snapshot().manifest.assets);
  expect(parseContract("interaction-request", r)).toEqual(r);
  expect(Reflect.set(r.annotations[0]!, "instruction", "altered")).toBe(false);
  expect(Reflect.set(r.context.blocks[0]!, "html", "altered")).toBe(false);
  expect(request(s).requestId).not.toBe(r.requestId);
  expect(s.serialize()).toEqual(before);
  expect(s.history().undoSteps).toBe(0);
  expect(events).toEqual([]);
  commit(s, [{ op: "updateAnnotation", annotationId: "a-summary", changes: { instruction: "changed later" } }]);
  expect(r.annotations[0]!.instruction).not.toBe("changed later");
});

test("overlapping parent and child contexts merge without truncating descendants", () => {
  const s = session();
  commit(s, [{ op: "addAnnotation", annotation: { annotationId: "root-note", documentId: "report-demo", baseRevision: 0,
    status: "open", anchor: { kind: "node", nodeId: "root" }, instruction: "whole report" } }]);
  const r = request(s, ["root-note", "a-summary"]);
  expect(r.context.blocks).toEqual([{ nodeId: "root", html: s.getNode("root")!.html }]);
  expect(r.context.blocks[0]!.html).toContain("n-status");
});

test("context budgets fail explicitly instead of dropping annotations or cutting HTML", () => {
  const s = session(), before = s.serialize();
  for (const limits of [{ maxBytes: 10, maxNodes: 100, maxAssets: 1 }, { maxBytes: 99999, maxNodes: 1, maxAssets: 1 }])
    fails(() => s.buildInteractionRequest({ annotationIds: ["a-summary", "a-heading"], limits }), "ResourceLimitExceeded");
  fails(() => s.buildInteractionRequest({ annotationIds: ["a-summary"], limits: { maxBytes: -1, maxNodes: 1, maxAssets: 1 } }), "InvalidOperation");
  expect(s.serialize()).toEqual(before);
  expect(s.history().undoSteps).toBe(0);
});

test("only a unique nonempty selection of open annotations can generate requests", () => {
  const s = session();
  for (const ids of [[], ["a-summary", "a-summary"]]) fails(() => request(s, ids), "InvalidOperation");
  fails(() => request(s, ["missing"]), "AnchorNotFound");
  commit(s, [{ op: "updateAnnotation", annotationId: "a-heading", changes: { status: "resolved" } }]);
  fails(() => request(s, ["a-heading"]), "AnchorNotFound");
  fails(() => s.buildInteractionRequest({ annotationIds: ["a-summary"], supportedOperations: ["setStyle", "setStyle"] }), "InvalidOperation");
});

test("proposal preparation is isolated; one commit applies content and statuses and one undo restores them", () => {
  const s = session(), before = s.snapshot(), r = request(s), p = proposal(r);
  const candidate = s.prepareProposal(r, p);
  expect(s.snapshot()).toEqual(before);
  expect(candidate.preview.html).toContain("第四季度");
  expect(candidate.preview.annotations.find((a) => a.annotationId === "a-heading")!.status).toBe("resolved");
  expect(candidate.preview.annotations.find((a) => a.annotationId === "a-summary")!.status).toBe("needs_reanchor");
  expect(s.commit(candidate)).toMatchObject({ revision: 1, changed: true });
  expect(s.history().undoSteps).toBe(1);
  expect(openDocument(s.serialize()).snapshot()).toEqual(s.snapshot());
  expect(s.undo().revision).toBe(2);
  expect(s.snapshot().html).toBe(before.html);
  expect(s.getAnnotations().map((a) => a.status)).toEqual(["open", "open"]);
  expect(s.redo().revision).toBe(3);
  expect(s.snapshot().html).toBe(candidate.preview.html);
});

test("unclaimed annotations remain open when their anchors remain valid", () => {
  const s = session(), r = request(s), p = proposal(r, [exampleProposal.operations[1]!]);
  p.annotationIds = ["a-heading"];
  s.commit(s.prepareProposal(r, p));
  expect(s.getAnnotations().find((a) => a.annotationId === "a-summary")!.status).toBe("open");
  expect(s.getAnnotations().find((a) => a.annotationId === "a-heading")!.status).toBe("resolved");
});

test("empty and reverted content proposals preserve annotations, revision and redo history", () => {
  const s = session();
  commit(s, [{ op: "setStyle", nodeId: "n-title", property: "color", value: "red" }]);
  s.undo();
  const before = s.snapshot(), history = s.history(), r = request(s);
  const original = s.getNode("n-summary")!.text;
  const operations: EditProposal["operations"] = [
    { op: "replaceText", nodeId: "n-summary", from: 0, to: original.length, expectedText: original, text: "暂存" },
    { op: "replaceText", nodeId: "n-summary", from: 0, to: 2, expectedText: "暂存", text: original },
  ];
  for (const ops of [[], operations]) {
    expect(s.commit(s.prepareProposal(r, proposal(r, ops))).changed).toBe(false);
    expect(s.snapshot()).toEqual(before);
    expect(s.history()).toEqual(history);
  }
  expect(s.redo().changed).toBe(true);
});

test("rejecting a prepared proposal discards it without changing the document", () => {
  const s = session(), before = s.serialize(), r = request(s), candidate = s.prepareProposal(r, proposal(r));
  s.discard(candidate);
  expect(s.serialize()).toEqual(before);
  fails(() => s.commit(candidate), "InvalidOperation");
});

test("requests and candidates from old revisions cannot apply after editing, undo or redo", () => {
  const s = session(), r = request(s), p = proposal(r), candidate = s.prepareProposal(r, p);
  commit(s, [{ op: "setStyle", nodeId: "n-title", property: "color", value: "red" }]);
  fails(() => s.prepareProposal(r, p), "RevisionConflict");
  fails(() => s.commit(candidate), "RevisionConflict");
  s.undo();
  fails(() => s.prepareProposal(r, p), "RevisionConflict");
  s.redo();
  fails(() => s.prepareProposal(r, p), "RevisionConflict");
});

test("rejects wrong request/document, duplicate or out-of-request annotation claims and undeclared operations", () => {
  const s = session(), r = request(s), before = s.serialize(), p = proposal(r);
  for (const changes of [{ requestId: "another" }, { documentId: "another" },
    { annotationIds: ["a-heading", "a-heading"] }, { annotationIds: ["unknown"] }])
    fails(() => s.prepareProposal(r, { ...p, ...changes }), "InvalidProposal");
  const limited = s.buildInteractionRequest({ annotationIds: ["a-heading"], supportedOperations: ["replaceText"] });
  fails(() => s.prepareProposal(limited, proposal(limited, [exampleProposal.operations[1]!])), "InvalidProposal");
  expect(s.serialize()).toEqual(before);
});

test("rejects truncated, tampered or overlapping request contexts and out-of-scope targets", () => {
  const s = session(), r = request(s, ["a-heading"]), p = proposal(r, [exampleProposal.operations[1]!]);
  fails(() => s.prepareProposal(r, proposal(r, [{ op: "removeNode", nodeId: "n-item" }])), "InvalidProposal");
  for (const mutate of [
    (copy: DocumentInteractionRequest) => { copy.context.blocks[0]!.html += "hidden"; },
    (copy: DocumentInteractionRequest) => { copy.context.styles[0]!.css += "p{color:red}"; },
    (copy: DocumentInteractionRequest) => { copy.annotations[0]!.instruction = "forged"; },
    (copy: DocumentInteractionRequest) => { copy.context.blocks.push({ nodeId: "root", html: s.getNode("root")!.html }); },
  ]) {
    const copy = structuredClone(r); mutate(copy);
    fails(() => s.prepareProposal(copy, p), "InvalidProposal");
  }
});

test("unsafe HTML, unknown operations and a later invalid operation leave all state unchanged", () => {
  const s = session(), r = request(s), before = s.serialize();
  fails(() => s.prepareProposal(r, proposal(r, [{ op: "replaceFragment", nodeId: "n-summary", html: "<script>bad()</script>" }])), "UnsupportedFeature");
  const bad = { ...proposal(r), operations: [{ op: "unknown", nodeId: "n-summary" }] } as unknown as EditProposal;
  fails(() => s.prepareProposal(r, bad), "InvalidProposal");
  fails(() => s.prepareProposal(r, proposal(r, [...exampleProposal.operations,
    { op: "replaceText", nodeId: "n-summary", from: 0, to: 1, expectedText: "wrong", text: "X" }])), "InvalidOperation");
  expect(s.serialize()).toEqual(before);
  expect(s.history().undoSteps).toBe(0);
});

test("blank-region requests include a complete container and layout-invalidated anchors stay unresolved", () => {
  const s = session();
  const note: Annotation = { annotationId: "blank", documentId: "report-demo", baseRevision: 0, status: "open",
    instruction: "insert here", anchor: { kind: "region", surfaceId: "root", nodeIds: [], layoutKey: "fixed",
      rect: { x: 0, y: 0, width: 10, height: 10 } } };
  commit(s, [{ op: "addAnnotation", annotation: note }]);
  const r = request(s, ["blank"]);
  expect(r.context.blocks[0]!.nodeId).toBe("root");
  s.commit(s.prepareProposal(r, proposal(r, [{ op: "insertFragment", nodeId: "n-title", position: "after", html: "<p>新段</p>" }])));
  expect(s.getAnnotations().find((a) => a.annotationId === "blank")!.status).toBe("needs_reanchor");
});
