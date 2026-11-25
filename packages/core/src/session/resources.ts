import type { TransactionPrepareOptions } from "../runtime.js";
import type { DocumentState } from "../document/state.js";
import type { DocumentManifest, ContentDocumentMetadata } from "../contracts/document.js";
import type { DeepReadonly, DocumentPackage } from "../runtime.js";
import type { DocumentReadLimits } from "../runtime.js";
import type { ParsedHtml } from "../document/html.js";
import { buildDocumentState, materializeDocumentFiles } from "../document/state.js";
import { encode, resolveReference } from "../document/package.js";
import { validateAsset } from "../document/assets.js";
import { checkLimit, invalid } from "../document/validation.js";
import { parseContract } from "../schema-registry.js";
import { checkTransactionSize } from "./limits.js";

/** Copy authorized bytes before they can be mutated by the caller. No provider or network. */
export function stageResources(
  base: DocumentState,
  options: TransactionPrepareOptions | undefined,
  limits: DocumentReadLimits,
): { state: DocumentState; added: ReadonlySet<string> } {
  const source = base.textEntry && options?.assets?.length ? materializeDocumentFiles(base) : base.files;
  const staged = stageResourceFiles(source, base.snapshot.manifest, options, limits);
  return {
    state: staged.added.size ? buildDocumentState(staged.files, limits, base) : base,
    added: staged.added,
  };
}

export function stageResourceFiles(
  baseFiles: DocumentPackage,
  baseManifest: DeepReadonly<DocumentManifest>,
  options: TransactionPrepareOptions | undefined,
  limits: DocumentReadLimits,
): {
  files: DocumentPackage;
  manifest: DeepReadonly<DocumentManifest>;
  added: ReadonlySet<string>;
} {
  const added = new Set<string>();
  if (options === undefined) return { files: baseFiles, manifest: baseManifest, added };
  if (
    !options ||
    typeof options !== "object" ||
    Array.isArray(options) ||
    Object.keys(options).some((key) => key !== "assets") ||
    (options.assets !== undefined && !Array.isArray(options.assets))
  )
    invalid("Invalid transaction preparation options.");
  if (!options.assets?.length) return { files: baseFiles, manifest: baseManifest, added };
  checkLimit(options.assets.length, limits.maxFiles);
  const files = new Map(baseFiles),
    manifest = parseContract("document-manifest", baseManifest);
  let total = [...files.values()].reduce((sum, bytes) => sum + bytes.byteLength, 0);
  const seen = new Set<string>();
  for (const input of options.assets) {
    if (
      !input ||
      typeof input !== "object" ||
      Array.isArray(input) ||
      Object.keys(input).some((key) => key !== "descriptor" && key !== "bytes") ||
      !(input.bytes instanceof Uint8Array)
    )
      invalid("Transaction resources require descriptors and authorized bytes.");
    checkTransactionSize(input.descriptor, limits.maxFileBytes, limits.maxDepth);
    const descriptor = parseContract("asset-descriptor", input.descriptor);
    if (seen.has(descriptor.path)) invalid("Transaction resource paths must be unique.");
    seen.add(descriptor.path);
    checkLimit(input.bytes.byteLength, limits.maxFileBytes);
    const previous = manifest.assets.find((item) => item.path === descriptor.path);
    const bytes = new Uint8Array(input.bytes);
    validateAsset(descriptor, bytes);
    if (previous) {
      if (previous.sha256 !== descriptor.sha256 || previous.mimeType !== descriptor.mimeType)
        invalid("Resource paths are immutable; use a new path for new bytes.");
      continue;
    }
    for (const path of files.keys())
      if (
        path === descriptor.path ||
        path.startsWith(descriptor.path + "/") ||
        descriptor.path.startsWith(path + "/")
      )
        invalid("Resource paths cannot overlap existing package files.");
    checkLimit(files.size + 1, limits.maxFiles);
    total += bytes.byteLength;
    checkLimit(total, limits.maxPackageBytes);
    files.set(descriptor.path, bytes);
    manifest.assets.push(descriptor);
    added.add(descriptor.path);
  }
  if (!added.size) return { files: baseFiles, manifest: baseManifest, added };
  files.set("manifest.json", encode(JSON.stringify(manifest)));
  return { files, manifest, added };
}

export function checkResourceUse(
  added: ReadonlySet<string>,
  after: ParsedHtml,
  entrypoint: string,
  metadata?: ContentDocumentMetadata,
): void {
  if (!added.size) return;
  const unused = new Set(added);
  for (const id of after.nodeIds()) {
    const node = after.peekNode(id)!;
    if (node.tagName === "img") unused.delete(resolveReference(node.attributes.src!, entrypoint));
  }
  for (const drawing of metadata?.word?.drawings ?? [])
    if (drawing.kind === "chart" && after.hasNode(drawing.nodeId))
      for (const part of drawing.chart.native?.parts ?? []) unused.delete(part.assetPath);
  const native = metadata?.word?.nativeContent;
  if (native) {
    const pending: string[] = [],
      selected = new Set<string>(),
      parts = new Map(native.parts.map((p) => [p.sourcePart, p])),
      keptParts = new Set<string>();
    const requirePart = (path: string): void => {
      if (keptParts.has(path)) return;
      keptParts.add(path);
      const part = parts.get(path);
      if (part) unused.delete(part.assetPath);
      for (const relation of native.relationships)
        if (relation.sourcePart === path) pending.push(path + "\n" + relation.id);
    };
    for (const fragment of native.fragments) {
      if (
        fragment.nodeId
          ? !after.hasNode(fragment.nodeId)
          : !fragment.numberingKind &&
            !metadata?.word?.styles.some((s) => s.styleId === fragment.styleId)
      )
        continue;
      unused.delete(fragment.assetPath);
      if (fragment.preview) unused.delete(fragment.preview.assetPath);
      for (const part of fragment.dependencyParts ?? []) requirePart(part);
      for (const id of fragment.relationshipIds ?? [])
        pending.push(fragment.sourcePart + "\n" + id);
    }
    const relations = new Map(native.relationships.map((r) => [r.sourcePart + "\n" + r.id, r]));
    for (let index = 0; index < pending.length; index++) {
      const key = pending[index]!;
      if (selected.has(key)) continue;
      selected.add(key);
      const relation = relations.get(key);
      if (!relation || relation.external) continue;
      requirePart(relation.target);
    }
  }
  if (unused.size) invalid("New transaction resources must be referenced by the final source.");
}
