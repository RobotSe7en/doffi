import {
  createAssetDescriptor,
  DocumentError,
  WordNativeFragmentCollectionSchema,
} from "@doffi/core";
import type { ContentDocumentModel } from "@doffi/core";
import { copyWordClipboardNative } from "./word-clipboard-native.js";

/** Compare the native part of a named style along with its modeled definition.
 * Collection keys, transport paths and source scopes are not style semantics. */
export function nativeClipboardStyleSignature(
  model: ContentDocumentModel,
  styleId: string,
): string | undefined {
  if (
    !model
      .snapshot()
      .metadata.word?.nativeContent?.fragments.some(
        (f) => f.styleId === styleId && !f.revisionSourceId,
      )
  )
    return;
  const graph = copyWordClipboardNative([], model, new Set([styleId])),
    native = graph.nativeContent!;
  const resources = new Map(
    graph.resources.map((r) => {
      const bytes = Uint8Array.from(atob(r.data), (c) => c.charCodeAt(0));
      return [
        r.path,
        {
          mimeType: r.mimeType,
          bytes,
          hash: createAssetDescriptor(r.path, r.mimeType, bytes).sha256,
        },
      ] as const;
    }),
  );
  const entries = new Set(
    native.fragments.map((f) => {
      const resource = resources.get(f.assetPath);
      if (!resource)
        throw new DocumentError(
          "AssetUnavailable",
          "A native style comparison requires authorized resources.",
        );
      const collection = WordNativeFragmentCollectionSchema.parse(
        JSON.parse(new TextDecoder().decode(resource.bytes)),
      );
      return JSON.stringify({
        kind: f.kind,
        propertyPath: f.propertyPath,
        propertyIndices: f.propertyIndices,
        propertyIndex: f.propertyIndex,
        propertyName: f.propertyName,
        styleRegion: f.styleRegion,
        xml: collection.fragments[f.fragmentKey!],
      });
    }),
  );
  const parts = native.parts
      .map((p) => ({ contentType: p.contentType, hash: resources.get(p.assetPath)!.hash }))
      .sort((a, b) => a.hash.localeCompare(b.hash)),
    relations = native.relationships.map((r) => ({
      id: r.id,
      type: r.type,
      external: r.external,
      target: r.external ? r.target : native.parts.findIndex((p) => p.sourcePart === r.target),
    }));
  const bytes = new TextEncoder().encode(
    JSON.stringify({ entries: [...entries].sort(), parts, relations }),
  );
  return createAssetDescriptor("assets/style-signature.bin", "application/octet-stream", bytes)
    .sha256;
}
