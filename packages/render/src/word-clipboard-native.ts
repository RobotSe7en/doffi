import {
  createAssetDescriptor,
  DocumentError,
  WordNativeFragmentCollectionSchema,
} from "@doffi/core";
import type { ContentDocumentModel, WordContentSemantics, WordNativeContent } from "@doffi/core";
import type { WordClipboardFormat } from "./word-clipboard-format.js";
import { ClipboardNativeXml, nativeBindingKey } from "./word-clipboard-native-xml.js";
import { WORD_CLIPBOARD_BYTES } from "./word-clipboard-constants.js";

/** Copy selected atoms, bound residual properties and their explicit OPC closure. */
export function copyWordClipboardNative(
  roots: readonly HTMLElement[],
  model: ContentDocumentModel,
  styleIds: ReadonlySet<string> = new Set(),
): {
  nativeContent?: WordContentSemantics["nativeContent"];
  resources: NonNullable<WordClipboardFormat["resources"]>;
} {
  const snapshot = model.snapshot(),
    native = snapshot.metadata.word?.nativeContent,
    nodes = new Set(
      roots.flatMap((root) =>
        [...root.querySelectorAll<HTMLElement>("[data-node-id]")].map((n) => n.dataset.nodeId!),
      ),
    ),
    resources: NonNullable<WordClipboardFormat["resources"]> = [];
  if ((!nodes.size && !styleIds.size) || !native) return { resources };
  let files: ReturnType<ContentDocumentModel["serialize"]> | undefined,
    total = 0;
  const read = (path: string): Uint8Array => {
    const bytes = model.getAsset
      ? model.getAsset(path, WORD_CLIPBOARD_BYTES)
      : (files ??= model.serialize()).get(path);
    if (!bytes) throw new DocumentError("AssetUnavailable", "复制原生内容需要已授权的片段资源。");
    return bytes;
  };
  const add = (path: string, mimeType: string, bytes: Uint8Array): void => {
    if (resources.some((item) => item.path === path)) return;
    total += bytes.length;
    if (total > WORD_CLIPBOARD_BYTES)
      throw new DocumentError("ResourceLimitExceeded", "原生内容超过剪贴板限制。");
    let binary = "";
    for (let at = 0; at < bytes.length; at += 32768)
      binary += String.fromCharCode(...bytes.subarray(at, at + 32768));
    resources.push({ path, mimeType, data: btoa(binary) });
  };
  const collections = new Map<
      string,
      ReturnType<typeof WordNativeFragmentCollectionSchema.parse>
    >(),
    legacy = new Map<string, Element>(),
    parser = new ClipboardNativeXml(),
    entries: Record<string, string> = Object.create(null) as Record<string, string>,
    copiedBindings = new Map<string, WordNativeContent["parts"][number]>(),
    decode = (path: string): string => new TextDecoder("utf-8", { fatal: true }).decode(read(path));
  // Older packages did not carry the indirect index. Reconstruct it from only
  // declared item/property edges, without retaining unrelated root resources.
  const customXml = new Map<
    string,
    { sourcePart: string; propertiesPart: string; storeItemId: string }
  >();
  const indexBindings = (): void => {
    for (const p of native.parts) {
      if (p.customXml) {
        customXml.set(nativeBindingKey(p.customXml.storeItemId), {
          ...p.customXml,
          sourcePart: p.sourcePart,
        });
        continue;
      }
      const props = native.relationships.find(
        (r) => r.sourcePart === p.sourcePart && !r.external && r.type.endsWith("/customXmlProps"),
      );
      if (!props) continue;
      const property = native.parts.find((item) => item.sourcePart === props.target);
      if (!property)
        throw new DocumentError("InvalidDocument", "A native custom XML property part is missing.");
      const id = parser.storeItemId(decode(property.assetPath));
      if (!id)
        throw new DocumentError("InvalidDocument", "A native custom XML datastore ID is missing.");
      customXml.set(nativeBindingKey(id), {
        sourcePart: p.sourcePart,
        propertiesPart: props.target,
        storeItemId: id,
      });
    }
  };
  let bindingsIndexed = false;
  const selectedFragments = native.fragments.filter(
      (f) =>
        !f.revisionSourceId &&
        ((f.nodeId && nodes.has(f.nodeId)) || (f.styleId && styleIds.has(f.styleId))),
    ),
    byNativeId = new Map(native.fragments.map((f) => [f.nativeId, f])),
    chosen = new Set(selectedFragments.map((f) => f.nativeId));
  for (let at = 0; at < selectedFragments.length; at++) {
    const f = selectedFragments[at]!;
    for (const id of Object.values(f.numberingReferences ?? {})) {
      const definition = byNativeId.get(id);
      if (definition && !chosen.has(id)) {
        chosen.add(id);
        selectedFragments.push(definition);
      }
    }
  }
  const fragments: WordNativeContent["fragments"] = [],
    wantedStyles = new Set(styleIds);
  const selectStyle = (id: string): void => {
    if (wantedStyles.has(id)) return;
    wantedStyles.add(id);
    const style = snapshot.metadata.word!.styles.find((s) => s.styleId === id);
    if (style)
      for (const dependency of [style.basedOn, style.nextStyleId, style.linkedStyleId])
        if (dependency) selectStyle(dependency);
    for (const f of native.fragments)
      if (f.styleId === id && !f.revisionSourceId && !chosen.has(f.nativeId)) {
        chosen.add(f.nativeId);
        selectedFragments.push(f);
      }
  };
  for (const id of [...wantedStyles]) {
    wantedStyles.delete(id);
    selectStyle(id);
  }
  for (const fragment of selectedFragments) {
    let xml: string | undefined;
    if (fragment.fragmentFormat === "json") {
      let collection = collections.get(fragment.assetPath);
      if (!collection) {
        collection = WordNativeFragmentCollectionSchema.parse(
          JSON.parse(decode(fragment.assetPath)),
        );
        collections.set(fragment.assetPath, collection);
      }
      xml =
        fragment.fragmentKey && Object.hasOwn(collection.fragments, fragment.fragmentKey)
          ? collection.fragments[fragment.fragmentKey]
          : undefined;
    } else {
      let root = legacy.get(fragment.assetPath);
      if (!root) {
        root = parser.parse(decode(fragment.assetPath));
        legacy.set(fragment.assetPath, root);
      }
      xml = parser.fragment(root, fragment.fragmentKey);
    }
    if (xml === undefined) throw new DocumentError("InvalidDocument", "原生内容缺少对应片段。");
    entries[fragment.nativeId] = xml;
    const copy = JSON.parse(JSON.stringify(fragment)) as WordNativeContent["fragments"][number];
    copy.fragmentKey = fragment.nativeId;
    copy.fragmentFormat = "json";
    const index =
      copy.relationshipIds === undefined ||
      copy.dependencyParts === undefined ||
      copy.styleReferences === undefined ||
      copy.numberingReferences === undefined
        ? parser.index(xml)
        : undefined;
    copy.relationshipIds ??= index!.relationshipIds;
    copy.dependencyParts ??= [];
    copy.styleReferences ??= index!.styleReferences;
    copy.numberingReferences ??= index!.numberingReferences;
    if (index?.bindings.length) {
      if (!bindingsIndexed) {
        indexBindings();
        bindingsIndexed = true;
      }
      for (const id of index.bindings) {
        const binding = customXml.get(id);
        if (!binding) continue;
        if (!copy.dependencyParts.includes(binding.sourcePart))
          copy.dependencyParts.push(binding.sourcePart);
        const original = native.parts.find((p) => p.sourcePart === binding.sourcePart)!;
        copiedBindings.set(binding.sourcePart, {
          ...original,
          customXml: { storeItemId: binding.storeItemId, propertiesPart: binding.propertiesPart },
        });
      }
    }
    for (const id of Object.values(copy.styleReferences ?? {})) selectStyle(id);
    for (const id of Object.values(copy.numberingReferences ?? {})) {
      const definition = byNativeId.get(id);
      if (definition && !chosen.has(id)) {
        chosen.add(id);
        selectedFragments.push(definition);
      }
    }
    fragments.push(copy);
  }
  const bytes = new TextEncoder().encode(
      JSON.stringify({ format: "doffi-word-native-fragments", version: 1, fragments: entries }),
    ),
    descriptor = createAssetDescriptor(
      "assets/clipboard-fragments.json",
      "application/json",
      bytes,
    ),
    path = "assets/clipboard-fragments-" + descriptor.sha256 + ".json";
  if (fragments.length) add(path, "application/json", bytes);
  for (const fragment of fragments) fragment.assetPath = path;
  const relations = new Map(native.relationships.map((r) => [r.sourcePart + "\n" + r.id, r])),
    parts = new Map(native.parts.map((p) => [p.sourcePart, p])),
    selected = new Set<string>(),
    keptParts = new Set<string>(),
    pending: string[] = [];
  const selectPart = (path: string): void => {
    if (!parts.has(path)) throw new DocumentError("InvalidDocument", "原生内容缺少依赖部件。");
    if (!keptParts.has(path)) {
      keptParts.add(path);
      pending.push(path);
    }
  };
  const select = (key: string): void => {
    if (selected.has(key)) return;
    const relation = relations.get(key);
    if (!relation) throw new DocumentError("InvalidDocument", "原生内容缺少声明的关系。");
    selected.add(key);
    if (!relation.external) selectPart(relation.target);
  };
  for (const fragment of fragments) {
    for (const id of fragment.relationshipIds!) select(fragment.sourcePart + "\n" + id);
    for (const path of fragment.dependencyParts ?? []) selectPart(path);
  }
  for (let at = 0; at < pending.length; at++)
    for (const [key, relation] of relations) if (relation.sourcePart === pending[at]) select(key);
  const copiedParts = native.parts.filter((p) => keptParts.has(p.sourcePart));
  for (const assetPath of new Set([
    ...copiedParts.map((p) => p.assetPath),
    ...fragments.flatMap((f) => (f.preview ? [f.preview.assetPath] : [])),
  ])) {
    const asset = snapshot.manifest.assets.find((a) => a.path === assetPath);
    if (!asset) throw new DocumentError("AssetUnavailable", "原生内容的资源未声明。");
    add(assetPath, asset.mimeType, read(assetPath));
  }
  return {
    nativeContent: {
      mainPart: native.mainPart,
      fragments,
      parts: copiedParts.map((p) => ({ ...(copiedBindings.get(p.sourcePart) ?? p) })),
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
