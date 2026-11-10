import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import type {
  Annotation, ContractName, DocumentInteractionRequest, DocumentManifest,
  WordDocumentMetadata, EditProposal,
} from "@doffi/core";

export const repository = fileURLToPath(new URL("../../../", import.meta.url));
function read(path: string): unknown {
  return JSON.parse(readFileSync(new URL("../../../" + path, import.meta.url), "utf8")) as unknown;
}
export const manifest = read("examples/report/manifest.json") as DocumentManifest;
export const metadata = read("examples/report/metadata.json") as WordDocumentMetadata;
export const annotations = read("examples/report/annotations.json") as Annotation[];
export const request = read("examples/report.request.json") as DocumentInteractionRequest;
export const proposal = read("examples/report.proposal.json") as EditProposal;

export const validCases: [ContractName, unknown][] = [
  ["asset-descriptor", { path: "assets/image.png", mimeType: "image/png", sha256: "a".repeat(64) }],
  ["document-manifest", manifest],
  ["document-metadata", metadata],
  ["document-metadata", { ...metadata, word: {
    stories: [{ storyId: "main", kind: "main", rootNodeId: "root" }],
    styles: [], numbering: [], hyperlinks: [], bookmarks: [], fields: [], tabs: [], revisions: [], tables: [],
  } }],
  ["annotation", { annotationId: "native-comment", documentId: manifest.documentId, baseRevision: 0,
    status: "open", anchor: { kind: "node", nodeId: "n-summary" }, bodyNodeId: "comment-root",
    author: { name: "审阅者", initials: "SY" }, date: "2026-01-01T08:00:00Z", sourceId: "0" }],
  ["anchor", annotations[0]!.anchor],
  ["annotation", annotations[0]],
  ["annotations", annotations],
  ["content-operation", proposal.operations[0]],
  ["content-operation", { op: "reconcileFragment", nodeId: "root", html: '<p data-node-id="n">新</p>',
    textEdits: [{ nodeId: "n", from: 0, to: 1, expectedText: "旧", text: "新" }] }],
  ["document-operation", { op: "addAnnotation", annotation: annotations[0] }],
  ["document-transaction", {
    protocolVersion: 1, documentId: manifest.documentId, baseRevision: 0, operations: proposal.operations,
  }],
  ["interaction-request", request],
  ["presentation-interaction-request", read("examples/presentation.request.json")],
  ["presentation-edit-proposal", read("examples/presentation.proposal.json")],
  ["workbook-interaction-request", read("examples/workbook.request.json")],
  ["workbook-edit-proposal", read("examples/workbook.proposal.json")],
  ["edit-proposal", proposal],
  ["document-error", { code: "InvalidDocument", message: "Invalid document.", details: [] }],
  ["document-event", {
    protocolVersion: 1, documentId: manifest.documentId, revision: 1,
    type: "documentChanged", affectedNodeIds: ["n-summary"],
  }],
  ["anchor", { kind: "text-set", ranges: [{ nodeId: "n", from: 0, to: 2, quote: "😀" }] }],
  ["anchor", { kind: "node", nodeId: "n" }],
  ["anchor", { kind: "node-set", nodeIds: ["n1", "n2"] }],
  ["anchor", {
    kind: "region", surfaceId: "root", rect: { x: 0, y: 0, width: 10, height: 20 },
    nodeIds: [], layoutKey: "fixed", pageId: "page-1",
  }],
  ["content-operation", { op: "replaceFragment", nodeId: "n", html: "" }],
  ["content-operation", { op: "insertFragment", nodeId: "n", position: "after", html: "<p>x</p>" }],
  ["content-operation", { op: "removeNode", nodeId: "n" }],
  ["content-operation", { op: "setStyle", nodeId: "n", property: "color", value: "#1f4b99" }],
  ["content-operation", { op: "replaceAsset", nodeId: "n", assetPath: "assets/new.png" }],
  ["document-operation", { op: "updateAnnotation", annotationId: "a", changes: { status: "resolved" } }],
  ["document-operation", { op: "removeAnnotation", annotationId: "a" }],
  ["document-event", {
    protocolVersion: 1, documentId: "d", revision: 1, type: "annotationChanged", annotationIds: ["a"],
  }],
  ["document-event", {
    protocolVersion: 1, documentId: "d", revision: 1, type: "error",
    error: { code: "RevisionConflict", message: "Stale proposal.", details: [] },
  }],
];

export const invalidCases: [ContractName, unknown][] = [
  ["document-manifest", { ...manifest, revision: -1 }],
  ["document-manifest", { ...manifest, revision: 0.5 }],
  ["document-manifest", { ...manifest, revision: Number.MAX_SAFE_INTEGER + 1 }],
  ["document-manifest", { ...manifest, revision: "0" }],
  ["document-manifest", { ...manifest, documentId: "" }],
  ["document-manifest", { ...manifest, formatVersion: 2 }],
  ["document-manifest", { ...manifest, format: "another-office" }],
  ["document-manifest", { ...manifest, kind: "unknown" }],
  ["document-manifest", { ...manifest, entrypoint: "../content.html" }],
  ["document-manifest", { ...manifest, typo: true }],
  ["document-manifest", { ...manifest, entrypoint: undefined }],
  ["document-metadata", { ...metadata, pageSetup: { ...metadata.pageSetup, width: 0 } }],
  ["document-metadata", { ...metadata, pageSetup: { ...metadata.pageSetup, unit: "px" } }],
  ["document-metadata", { ...metadata, sections: [] }],
  ["document-metadata", { ...metadata, extensions: { host: true } }],
  ["anchor", { kind: "text", nodeId: "n", from: -1, to: 1, quote: "" }],
  ["anchor", { kind: "text", nodeId: "n", from: 0, to: 1, quote: "", typo: true }],
  ["anchor", { kind: "node-set", nodeIds: [] }],
  ["anchor", { kind: "region", surfaceId: "root", rect: { x: 0, y: 0, width: 0, height: 1 }, nodeIds: [], layoutKey: "fixed" }],
  ["annotation", { ...annotations[0], status: "pending" }],
  ["annotation", { ...annotations[0], bodyNodeId: "duplicate-content" }],
  ["annotations", [{ ...annotations[0], author: { userId: "untrusted" } }]],
  ["asset-descriptor", { path: "/image.png", mimeType: "image/png", sha256: "a".repeat(64) }],
  ["asset-descriptor", { path: "image.png", mimeType: "image/png", sha256: "invalid" }],
  ["asset-descriptor", { path: "image.png", mimeType: "image/png", sha256: "a".repeat(64) + "\n" }],
  ["content-operation", { op: "replaceText", nodeId: "n", from: 0, to: 1, text: "new" }],
  ["content-operation", { op: "setGeometry", nodeId: "n", x: 1 }],
  ["content-operation", { op: "reconcileFragment", nodeId: "n", html: "", textEdits: [{ nodeId: "n", from: 0, to: 0, expectedText: "", text: "x", unsafe: true }] }],
  ["content-operation", { op: "addAnnotation", annotation: annotations[0] }],
  ["content-operation", { op: "insertFragment", nodeId: "root", position: "somewhere", html: "<p>x</p>" }],
  ["document-operation", { op: "updateAnnotation", annotationId: "a", changes: { unknown: true } }],
  ["document-transaction", { protocolVersion: 1, documentId: "d", baseRevision: 0, operations: [], committed: true }],
  ["interaction-request", { ...request, protocolVersion: 2 }],
  ["interaction-request", { ...request, annotations: [{ ...annotations[0], status: "resolved" }] }],
  ["interaction-request", { ...request, supportedOperations: ["addAnnotation"] }],
  ["interaction-request", { ...request, context: { ...request.context, blocks: [] } }],
  ["interaction-request", { ...request, context: { ...request.context, kind: "workbook" } }],
  ["interaction-request", { ...request, context: { ...request.context, formatVersion: 2 } }],
  ["interaction-request", {
    ...request,
    context: { ...request.context, styles: [{ ...request.context.styles[0], extra: true }] },
  }],
  ["edit-proposal", { ...proposal, requestId: "" }],
  ["edit-proposal", { ...proposal, operations: [{ op: "removeAnnotation", annotationId: "a-summary" }] }],
  ["edit-proposal", { ...proposal, accepted: true }],
  ["presentation-edit-proposal", { ...read("examples/presentation.proposal.json") as object, operations: [{op:"hostExecute", command:"x"}] }],
  ["presentation-interaction-request", { ...read("examples/presentation.request.json") as object, protocolVersion:2 }],
  ["workbook-edit-proposal", { ...read("examples/workbook.proposal.json") as object, operations: [{ op: "hostExecute", command: "x" }] }],
  ["workbook-interaction-request", { ...read("examples/workbook.request.json") as object, protocolVersion: 2 }],
  ["document-error", { code: "UnspecifiedError", message: "No.", details: [] }],
  ["document-event", { protocolVersion: 1, documentId: "d", revision: 1, type: "documentChanged" }],
  ["document-event", { protocolVersion: 1, documentId: "d", revision: 1, type: "modelRunning" }],
];
