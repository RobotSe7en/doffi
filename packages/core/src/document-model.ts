import type { ContentDocumentModel, DocumentModel, DocumentPackage, OpenDocumentOptions, WorkbookModel } from "./runtime.js";
import { copyPackage } from "./document/package.js";
import { buildDocumentState } from "./document/state.js";
import { readLimits } from "./document/validation.js";
import { readManifest } from "./document/manifest.js";
import { DocumentError } from "./errors.js";
import { openWorkbookModel } from "./workbook/model.js";
import { takeSerializedWorkbook } from "./workbook/source.js";

export { DEFAULT_DOCUMENT_READ_LIMITS } from "./document/validation.js";

/** Open authorized in-memory files. Performs no filesystem, network or model I/O. */
export function openDocument(input: DocumentPackage, options: OpenDocumentOptions & { readonly kind: "workbook" }): WorkbookModel;
export function openDocument(input: DocumentPackage, options: OpenDocumentOptions & { readonly kind: "document" | "presentation" }): ContentDocumentModel;
export function openDocument(input: DocumentPackage, options?: OpenDocumentOptions): DocumentModel;
export function openDocument(input: DocumentPackage, options: OpenDocumentOptions = {}): DocumentModel {
  const limits = readLimits(options.limits);
  const files = copyPackage(input, limits), manifest = readManifest(files);
  if (options.kind !== undefined && options.kind !== manifest.kind) throw new DocumentError("UnsupportedFormat", "Document kind does not match the expected format.");
  if (manifest.kind === "workbook") return openWorkbookModel(files, manifest, limits, options.workbookLimits, takeSerializedWorkbook(input));
  const state = buildDocumentState(files, limits);
  return Object.freeze({
    kind: state.snapshot.manifest.kind,
    snapshot: () => state.snapshot,
    getNode: (nodeId: string) => state.html.getNode(nodeId),
    serialize: () => new Map([...state.files].map(([path, bytes]) => [path, new Uint8Array(bytes)])),
  });
}
