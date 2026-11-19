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
  readonly files: DocumentPackage;
  readonly snapshot: ContentDocumentSnapshot;
  readonly html: ReturnType<typeof parseHtml>;
}

export function buildDocumentState(
  input: DocumentPackage,
  limits: DocumentReadLimits,
): DocumentState {
  const files = new Map(input);
  const manifest = readManifest(files);
  if (manifest.kind !== "document" && manifest.kind !== "presentation")
    unsupported("This document kind is not implemented.");

  validateAssets(manifest.assets, files);
  const parsedMetadata = parseContract("document-metadata", readJson(files, manifest.metadata));
  if ("workbook" in parsedMetadata) invalid("Manifest kind and metadata kind must match.");
  const metadata: ContentDocumentMetadata = parsedMetadata;
  if ((manifest.kind === "presentation") !== (metadata.presentation !== undefined))
    invalid("Manifest kind and metadata kind must match.");
  const annotations = parseContract("annotations", readJson(files, manifest.annotations));
  const cssBudget = { nodes: 0 };
  const styles = manifest.styles.map((path) => ({
    path,
    css: normalizeCss(readText(files, path), "stylesheet", limits, cssBudget),
  }));
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
  } else validatePresentation(metadata.presentation, annotations, html, styles, limits);
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
  for (const style of styles) files.set(style.path, encode(style.css));
  let total = 0;
  for (const [path, bytes] of files) {
    checkLimit(bytes.byteLength, limits.maxFileBytes, "maxFileBytes", `normalized ${path}`);
    total += bytes.byteLength;
  }
  checkLimit(total, limits.maxPackageBytes, "maxPackageBytes", "the normalized package");

  return { files, snapshot, html };
}
