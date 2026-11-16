import {
  createAssetDescriptor,
  DocumentError,
  WordNativeFragmentCollectionSchema,
} from "@doffi/core";
import type { ContentDocumentModel, WordContentSemantics, WordNativeContent } from "@doffi/core";
import type { WordClipboardFormat } from "./word-clipboard-format.js";
import { WORD_CLIPBOARD_BYTES } from "./word-clipboard-constants.js";

/** The R07 opaque comment endpoint carries its selected indexed atom. General
 * legacy collections, property-only graphs and indirect bindings remain R08. */
export function copyWordClipboardNative(
  roots: readonly HTMLElement[],
  model: ContentDocumentModel,
): {
  nativeContent?: WordContentSemantics["nativeContent"];
  resources: NonNullable<WordClipboardFormat["resources"]>;
} {
  const snapshot = model.snapshot(),
    native = snapshot.metadata.word?.nativeContent,
    nodes = new Set(
      roots.flatMap((root) =>
        [...root.querySelectorAll<HTMLElement>("[data-word-native-object]")].map(
          (n) => n.dataset.nodeId!,
        ),
      ),
    ),
    resources: NonNullable<WordClipboardFormat["resources"]> = [];
  if (!nodes.size || !native) return { resources };
  let files: ReturnType<ContentDocumentModel["serialize"]> | undefined,
    total = 0;
  const read = (path: string): Uint8Array => {
    const bytes = model.getAsset
      ? model.getAsset(path, WORD_CLIPBOARD_BYTES)
      : (files ??= model.serialize()).get(path);
    if (!bytes)
      throw new DocumentError("AssetUnavailable", "复制原生批注端点需要已授权的片段资源。");
    return bytes;
  };
  const add = (path: string, mimeType: string, bytes: Uint8Array): void => {
    if (resources.some((item) => item.path === path)) return;
    total += bytes.length;
    if (total > WORD_CLIPBOARD_BYTES)
      throw new DocumentError("ResourceLimitExceeded", "原生批注端点超过剪贴板限制。");
    let binary = "";
    for (let at = 0; at < bytes.length; at += 32768)
      binary += String.fromCharCode(...bytes.subarray(at, at + 32768));
    resources.push({ path, mimeType, data: btoa(binary) });
  };
  const collections = new Map<
      string,
      ReturnType<typeof WordNativeFragmentCollectionSchema.parse>
    >(),
    entries: Record<string, string> = {},
    fragments = native.fragments
      .filter((f) => f.nodeId && nodes.has(f.nodeId))
      .map((fragment) => {
        if (
          fragment.fragmentFormat !== "json" ||
          !fragment.fragmentKey ||
          !fragment.relationshipIds ||
          fragment.revisionSourceId
        )
          throw new DocumentError("UnsupportedFeature", "此原生片段尚不能独立迁移，复制已取消。");
        let collection = collections.get(fragment.assetPath);
        if (!collection) {
          collection = WordNativeFragmentCollectionSchema.parse(
            JSON.parse(new TextDecoder("utf-8", { fatal: true }).decode(read(fragment.assetPath))),
          );
          collections.set(fragment.assetPath, collection);
        }
        const xml = Object.hasOwn(collection.fragments, fragment.fragmentKey)
          ? collection.fragments[fragment.fragmentKey]
          : undefined;
        if (xml === undefined)
          throw new DocumentError("InvalidDocument", "原生批注端点缺少对应片段。");
        // nativeId is unique even when several collections use the same key.
        entries[fragment.nativeId] = xml;
        const copy = JSON.parse(JSON.stringify(fragment)) as WordNativeContent["fragments"][number];
        copy.fragmentKey = fragment.nativeId;
        return { ...copy, relationshipIds: [...fragment.relationshipIds] };
      });
  const bytes = new TextEncoder().encode(
      JSON.stringify({ format: "doffi-word-native-fragments", version: 1, fragments: entries }),
    ),
    descriptor = createAssetDescriptor(
      "assets/clipboard-fragments.json",
      "application/json",
      bytes,
    ),
    path = "assets/clipboard-fragments-" + descriptor.sha256 + ".json";
  add(path, "application/json", bytes);
  for (const fragment of fragments) fragment.assetPath = path;
  const relations = new Map(native.relationships.map((r) => [r.sourcePart + "\n" + r.id, r])),
    parts = new Map(native.parts.map((p) => [p.sourcePart, p])),
    selected = new Set<string>(),
    keptParts = new Set<string>(),
    pending: string[] = [];
  const select = (key: string): void => {
    if (selected.has(key)) return;
    const relation = relations.get(key);
    if (!relation) throw new DocumentError("InvalidDocument", "原生批注端点缺少声明的关系。");
    selected.add(key);
    if (!relation.external && !keptParts.has(relation.target)) {
      if (!parts.has(relation.target))
        throw new DocumentError("UnsupportedFeature", "此原生依赖尚不能独立迁移。");
      keptParts.add(relation.target);
      pending.push(relation.target);
    }
  };
  for (const fragment of fragments)
    for (const id of fragment.relationshipIds) select(fragment.sourcePart + "\n" + id);
  for (let at = 0; at < pending.length; at++)
    for (const [key, relation] of relations) if (relation.sourcePart === pending[at]) select(key);
  const copiedParts = native.parts.filter((p) => keptParts.has(p.sourcePart));
  for (const assetPath of new Set([
    ...copiedParts.map((p) => p.assetPath),
    ...fragments.flatMap((f) => (f.preview ? [f.preview.assetPath] : [])),
  ])) {
    const asset = snapshot.manifest.assets.find((a) => a.path === assetPath);
    if (!asset) throw new DocumentError("AssetUnavailable", "原生批注端点的资源未声明。");
    add(assetPath, asset.mimeType, read(assetPath));
  }
  return {
    nativeContent: {
      mainPart: native.mainPart,
      fragments,
      parts: copiedParts.map((p) => ({ ...p })),
      relationships: native.relationships
        .filter((r) => selected.has(r.sourcePart + "\n" + r.id))
        .map(({ root: _root, ...r }) => r),
      ...(native.identities
        ? {
            identities: native.identities.filter((i) => nodes.has(i.nodeId)).map((i) => ({ ...i })),
          }
        : {}),
    },
    resources,
  };
}
