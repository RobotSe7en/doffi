import type {
  ContentDocumentModel,
  DocumentModel,
  DocumentPackage,
  OpenDocumentOptions,
  WorkbookModel,
} from "./runtime.js";
import { copyPackage } from "./document/package.js";
import { buildDocumentState } from "./document/state.js";
import { DEFAULT_DOCUMENT_READ_LIMITS, readLimits } from "./document/validation.js";
import { readManifest } from "./document/manifest.js";
import { DocumentError } from "./errors.js";
import { openWorkbookModel } from "./workbook/model.js";
import { takeSerializedWorkbook } from "./workbook/source.js";
import { contentModel, contentSource } from "./document/model-source.js";
import type { DocumentReadLimits } from "./runtime.js";

export { DEFAULT_DOCUMENT_READ_LIMITS } from "./document/validation.js";

/** Open authorized in-memory files. Performs no filesystem, network or model I/O. */
export function openDocument(
  input: DocumentPackage,
  options: OpenDocumentOptions & { readonly kind: "workbook" },
): WorkbookModel;
export function openDocument(
  input: DocumentPackage,
  options: OpenDocumentOptions & { readonly kind: "document" | "presentation" },
): ContentDocumentModel;
export function openDocument(input: DocumentPackage, options?: OpenDocumentOptions): DocumentModel;
export function openDocument(
  input: DocumentPackage,
  options: OpenDocumentOptions = {},
): DocumentModel {
  const limits = readLimits(options.limits);
  const files = copyPackage(input, limits),
    manifest = readManifest(files);
  if (options.kind !== undefined && options.kind !== manifest.kind)
    throw new DocumentError(
      "UnsupportedFormat",
      "Document kind does not match the expected format.",
    );
  if (manifest.kind === "workbook")
    return openWorkbookModel(
      files,
      manifest,
      limits,
      options.workbookLimits,
      takeSerializedWorkbook(input),
    );
  return contentModel(buildDocumentState(files, limits));
}

/** Capture a validated revision for consumers that require default read limits. */
export function captureContentModel(
  input: ContentDocumentModel,
  kind: ContentDocumentModel["kind"] = input.kind,
): ContentDocumentModel {
  const state = contentSource(input);
  if (state?.snapshot.manifest.kind === kind &&
    (Object.keys(DEFAULT_DOCUMENT_READ_LIMITS) as (keyof DocumentReadLimits)[]).every(
      key => state.limits[key] <= DEFAULT_DOCUMENT_READ_LIMITS[key],
    ))
    return contentModel(state);
  return openDocument(input.serialize(), { kind });
}
