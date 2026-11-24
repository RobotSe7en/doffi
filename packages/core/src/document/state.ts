import { parseContract } from "../schema-registry.js";
import type { DocumentPackage, DocumentReadLimits, ContentDocumentSnapshot } from "../runtime.js";
import { normalizeCss } from "./css.js";
import { parseHtml } from "./html.js";
import { validateAssets } from "./assets.js";
import { validateWord } from "./word.js";
import { validatePresentation } from "./presentation.js";
import type { ContentDocumentMetadata } from "../contracts/document.js";
import { validateAnnotations } from "./annotations.js";
import { readManifest } from "./manifest.js";
import { encode, readJson, readText } from "./package.js";
import { checkLimit, freeze, invalid, unsupported } from "./validation.js";

/** Private bytes are never exposed; unchanged assets may be shared across states. */
export interface DocumentState {
  readonly limits: DocumentReadLimits;
  /** Text-only revisions keep the shared original entrypoint in these buffers. */
  readonly files: DocumentPackage;
  readonly textEntry?: { readonly byteLength: number; readonly patches: readonly Uint8Array[] };
  readonly snapshot: ContentDocumentSnapshot;
  readonly html: ReturnType<typeof parseHtml>;
  readonly stylesheetNodes: ReadonlyMap<string, number>;
}

export function materializeDocumentFiles(state: DocumentState): Map<string, Uint8Array> {
  const files = new Map(state.files);
  if (state.textEntry) files.set(state.snapshot.manifest.entrypoint, encode(state.snapshot.html));
  return files;
}

export function snapshotWithText(
  { manifest, styles, metadata, annotations }: Omit<ContentDocumentSnapshot, "html">,
  html: ReturnType<typeof parseHtml>,
): ContentDocumentSnapshot {
  // Keep the immutable base and its replacement runs in history. Joining a
  // whole HTML file is an on-demand projection, like serializing the package.
  return Object.freeze({ manifest, get html() { return html.html; }, styles, metadata, annotations });
}

export function buildDocumentState(
  input: DocumentPackage,
  limits: DocumentReadLimits,
  previous?: DocumentState,
): DocumentState {
  const files = new Map(input);
  const manifest = readManifest(files);
  if (manifest.kind !== "document" && manifest.kind !== "presentation")
    unsupported("This document kind is not implemented.");

  const previousAssets = new Map(previous?.snapshot.manifest.assets.map(asset => [asset.path, asset]));
  validateAssets(manifest.assets.filter(asset => {
    const old = previousAssets.get(asset.path);
    return !old || old.sha256 !== asset.sha256 || old.mimeType !== asset.mimeType ||
      files.get(asset.path) !== previous!.files.get(asset.path);
  }), files);
  const parsedMetadata = parseContract("document-metadata", readJson(files, manifest.metadata));
  if ("workbook" in parsedMetadata) invalid("Manifest kind and metadata kind must match.");
  const metadata: ContentDocumentMetadata = parsedMetadata;
  if ((manifest.kind === "presentation") !== (metadata.presentation !== undefined))
    invalid("Manifest kind and metadata kind must match.");
  const annotations = parseContract("annotations", readJson(files, manifest.annotations));
  const cssBudget = { nodes: 0 };
  const oldStyles = new Map(previous?.snapshot.styles.map(style => [style.path, style]));
  const stylesheetNodes = new Map<string, number>();
  const reusedStyles = new Set<string>();
  const styles = manifest.styles.map(path => {
    const old = oldStyles.get(path);
    // Only the owner can supply the identical private stylesheet bytes and limits.
    if (old && previous!.limits === limits && files.get(path) === previous!.files.get(path)) {
      const nodes = previous!.stylesheetNodes.get(path)!;
      stylesheetNodes.set(path, nodes);
      checkLimit(cssBudget.nodes += nodes, limits.maxCssNodes, "maxCssNodes", "CSS");
      reusedStyles.add(path);
      return old;
    }
    const before = cssBudget.nodes;
    const css = normalizeCss(readText(files, path), "stylesheet", limits, cssBudget);
    stylesheetNodes.set(path, cssBudget.nodes - before);
    return { path, css };
  });
  const html = parseHtml(readText(files, manifest.entrypoint), manifest, files, limits, cssBudget);

  if (metadata.presentation === undefined) {
    const sectionIds = new Set<string>();
    const sectionRoots = new Set<string>();
    for (const section of metadata.sections) {
      if (
        sectionIds.has(section.sectionId) ||
        sectionRoots.has(section.rootNodeId) ||
        !html.isContainer(section.rootNodeId)
      ) {
        invalid("Document sections require unique identities and existing root containers.");
      }
      sectionIds.add(section.sectionId);
      sectionRoots.add(section.rootNodeId);
    }
    validateWord(metadata, annotations, html, manifest.assets);
  } else validatePresentation(metadata.presentation, annotations, html, styles, limits, manifest.assets);
  validateAnnotations(annotations, manifest, html);
  const snapshot: ContentDocumentSnapshot = freeze({
    manifest,
    html: html.html,
    styles,
    metadata,
    annotations,
  });
  files.set("manifest.json", encode(JSON.stringify(manifest, null, 2) + "\n"));
  files.set(manifest.entrypoint, encode(snapshot.html));
  files.set(manifest.metadata, encode(JSON.stringify(metadata, null, 2) + "\n"));
  files.set(manifest.annotations, encode(JSON.stringify(annotations, null, 2) + "\n"));
  for (const style of styles) if (!reusedStyles.has(style.path)) files.set(style.path, encode(style.css));
  let total = 0;
  for (const [path, bytes] of files) {
    checkLimit(bytes.byteLength, limits.maxFileBytes, "maxFileBytes", `normalized ${path}`);
    total += bytes.byteLength;
  }
  checkLimit(total, limits.maxPackageBytes, "maxPackageBytes", "the normalized package");

  return { limits, files, snapshot, html, stylesheetNodes };
}
