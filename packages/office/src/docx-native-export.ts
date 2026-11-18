import { nativeRangeControl, visitNativeRanges } from "./docx-native-ranges.js";
import { NativeWordComments } from "./docx-native-comments.js";
import type { NativeCommentBinding } from "./docx-native-comments.js";
import { wordCompatibilityKey, wordCompatibilityNames } from "./ooxml/word-compatibility.js";
import { wordTypographySettings } from "./ooxml/word-typography.js";
import { sortWordProperties } from "./ooxml/word-property-order.js";
import { sortWordNoteProperties } from "./ooxml/word-note-order.js";
import { NativeDrawingIdentities } from "./docx-native-identities.js";
import { bindingKey, CUSTOM_XML, isolatedStoreItemId } from "./docx-native-custom-xml.js";
import { acceptNativeBlockInsertion, insertNativeBlock } from "./docx-native-revisions.js";
import { sortWordSettings } from "./ooxml/word-settings-order.js";
import { DocumentError, WordNativeFragmentCollectionSchema } from "@doffi/core";
import type {
  ContentDocumentModel,
  DocumentReadLimits,
  DeepReadonly,
  WordMetadata,
} from "@doffi/core";
import { W, R, MC, PKG, xml } from "./ooxml/xml.js";
import type { XmlNode } from "./ooxml/xml.js";
import { serializeXml } from "./ooxml/serialize.js";
import { declaration, escapeXml } from "./export/xml.js";
export const NATIVE_MARKER = "urn:doffi:transient:native";
export const nativeMarker = (nodeId: string | undefined): string =>
  nodeId ? ' xmlns:dn="' + NATIVE_MARKER + '" dn:node="' + escapeXml(nodeId) + '"' : "";
const CT = "http://schemas.openxmlformats.org/package/2006/content-types";
const codecs = globalThis as unknown as {
  TextEncoder: new () => { encode(text: string): Uint8Array };
  TextDecoder: new (
    label: string,
    options: { fatal: boolean },
  ) => { decode(bytes: Uint8Array): string };
};
function relPath(part: string): string {
  const slash = part.lastIndexOf("/");
  return part
    ? part.slice(0, slash + 1) + "_rels/" + part.slice(slash + 1) + ".rels"
    : "_rels/.rels";
}
function name(node: XmlNode): string {
  return node.uri + " " + node.name;
}
function merge(target: XmlNode, source: XmlNode): void {
  Object.assign(target.namespaces, source.namespaces);
  for (const [key, value] of source.attrs)
    (target.attrs as Map<string, string>).set(
      key,
      key === MC + " Ignorable"
        ? [
            ...new Set(
              [...(target.attrs.get(key) ?? "").split(/\s+/), ...value.split(/\s+/)].filter(
                Boolean,
              ),
            ),
          ].join(" ")
        : value,
    );
  if (source.content.length) target.content.splice(0, target.content.length, ...source.content);
}
/** Reattach residual native data to the just-compiled current source. */
export function restoreDocxNative(
  model: ContentDocumentModel,
  generated: ReadonlyMap<string, Uint8Array>,
  limits: DocumentReadLimits,
  commentBindings: readonly NativeCommentBinding[] = [],
  unexportedNativeRevisionIds: string[] = [],
): ReadonlyMap<string, Uint8Array> {
  const native = model.snapshot().metadata.word?.nativeContent,
    mappedTypography = model.snapshot().metadata.word?.typography !== undefined,
    mappedCompatibility = new Set<string>(
      wordCompatibilityNames.filter((name) => model.snapshot().metadata.word?.[name] !== undefined),
    );
  if (model.snapshot().metadata.word?.compatibilityMode !== undefined)
    mappedCompatibility.add("compatibilityMode");
  if (!native) return generated;
  const files = new Map(generated),
    sourceFiles = model.serialize(),
    encoder = new codecs.TextEncoder();
  const budget = { nodes: 0 },
    parsed = new Map<string, XmlNode>();
  const read = (bytes: Uint8Array, path: string): XmlNode => xml(bytes, path, limits, budget);
  const fragmentAssets = new Map<string, XmlNode>();
  const fragmentEntries = new Map<string, Map<string, XmlNode>>();
  const jsonFragments = new Map<string, Record<string, string>>();
  const readFragment = (fragment: (typeof native.fragments)[number]): XmlNode => {
    if (fragment.fragmentFormat === "json") {
      let entries = jsonFragments.get(fragment.assetPath);
      if (!entries) {
        const bytes = sourceFiles.get(fragment.assetPath);
        if (!bytes)
          throw new DocumentError("AssetUnavailable", "A native fragment collection is missing.");
        let value: unknown;
        try {
          value = JSON.parse(new codecs.TextDecoder("utf-8", { fatal: true }).decode(bytes));
        } catch {
          throw new DocumentError(
            "InvalidDocument",
            "The native fragment collection is malformed.",
          );
        }
        const collection = WordNativeFragmentCollectionSchema.safeParse(value);
        if (!collection.success || Object.keys(collection.data.fragments).length > limits.maxNodes)
          throw new DocumentError("InvalidDocument", "The native fragment collection is invalid.");
        entries = collection.data.fragments;
        jsonFragments.set(fragment.assetPath, entries);
      }
      const source =
        fragment.fragmentKey && Object.hasOwn(entries, fragment.fragmentKey)
          ? entries[fragment.fragmentKey]
          : undefined;
      if (source === undefined)
        throw new DocumentError(
          "InvalidDocument",
          "The native fragment collection entry is missing.",
        );
      const key = fragment.assetPath + "\n" + fragment.fragmentKey;
      let root = fragmentAssets.get(key);
      if (!root) {
        root = read(encoder.encode(source), fragment.assetPath);
        fragmentAssets.set(key, root);
      }
      return root;
    }
    let root = fragmentAssets.get(fragment.assetPath);
    if (!root) {
      const bytes = sourceFiles.get(fragment.assetPath);
      if (!bytes) throw new DocumentError("AssetUnavailable", "A native fragment is missing.");
      root = read(bytes, fragment.assetPath);
      fragmentAssets.set(fragment.assetPath, root);
    }
    if (!fragment.fragmentKey) return root;
    let entries = fragmentEntries.get(fragment.assetPath);
    if (!entries) {
      entries = new Map();
      for (const item of root.content)
        if (
          typeof item !== "string" &&
          item.uri === "urn:doffi:word:native-fragments" &&
          item.name === "item"
        ) {
          const key = item.attrs.get(" id"),
            child = item.content.find((node): node is XmlNode => typeof node !== "string");
          if (!key || !child || entries.has(key))
            throw new DocumentError(
              "InvalidDocument",
              "The native fragment collection is invalid.",
            );
          entries.set(key, child);
        }
      fragmentEntries.set(fragment.assetPath, entries);
    }
    const node = entries.get(fragment.fragmentKey);
    if (!node)
      throw new DocumentError("InvalidDocument", "A native fragment collection entry is missing.");
    return node;
  };
  const nodeParts = new Map<string, { path: string; name: string }[]>();
  const compiledNodeIds = new WeakMap<XmlNode, string>();
  for (const [path, bytes] of files)
    if (path.endsWith(".xml") && path !== "[Content_Types].xml") {
      const root = read(bytes, path);
      parsed.set(path, root);
      const scan = (node: XmlNode): void => {
        for (const id of [
          node.attrs.get(NATIVE_MARKER + " node"),
          node.attrs.get(NATIVE_MARKER + " container"),
        ])
          if (id) {
            const locations = nodeParts.get(id) ?? [];
            locations.push({ path, name: node.name });
            nodeParts.set(id, locations);
            compiledNodeIds.set(node, id);
          }
        for (const item of node.content) if (typeof item !== "string") scan(item);
      };
      scan(root);
    }
  const rawDefinitions = native.fragments.filter((f) => f.numberingKind),
    numberingIds = new Map<string, string>();
  let numberingPart = [...parsed].find(
    ([, root]) => root.uri === W && root.name === "numbering",
  )?.[0];
  if (rawDefinitions.length && !numberingPart) {
    numberingPart = "word/numbering.xml";
    parsed.set(numberingPart, {
      name: "numbering",
      uri: W,
      attrs: new Map(),
      namespaces: {},
      content: [],
      part: numberingPart,
    });
  }
  for (const kind of ["num", "abstractNum", "numPicBullet"] as const) {
    const attribute =
        W +
        " " +
        (kind === "num" ? "numId" : kind === "abstractNum" ? "abstractNumId" : "numPicBulletId"),
      root = numberingPart ? parsed.get(numberingPart)! : undefined,
      used = new Set(
        root?.content.flatMap((n) =>
          typeof n === "string" || n.name !== kind ? [] : [n.attrs.get(attribute)],
        ) ?? [],
      );
    let next = kind === "num" ? 1 : 0;
    for (const f of rawDefinitions.filter((f) => f.numberingKind === kind)) {
      while (used.has(String(next))) next++;
      numberingIds.set(f.nativeId, String(next));
      used.add(String(next++));
    }
  }
  const partMap = new Map<string, string>([[native.mainPart, "word/document.xml"]]);
  for (const fragment of native.fragments) {
    const locations = fragment.nodeId ? nodeParts.get(fragment.nodeId) : undefined;
    const mapped =
      locations?.find((location) => location.name === fragment.ownerName)?.path ??
      locations?.[0]?.path;
    if (mapped) partMap.set(fragment.sourcePart, mapped);
    else if (fragment.kind === "property" && !fragment.nodeId) {
      const rootNames: Record<string, string> = {
        comments: "comments",
        footnotes: "footnotes",
        endnotes: "endnotes",
        settings: "settings",
        style: "styles",
        styles: "styles",
        numbering: "numbering",
        fonts: "fonts",
        commentsEx: "commentsEx",
        commentsExtensible: "commentsExtensible",
        commentsIds: "commentsIds",
        people: "people",
      };
      const rootName = rootNames[fragment.ownerName];
      const entry = rootName ? [...parsed].find(([, root]) => root.name === rootName) : undefined;
      if (entry) partMap.set(fragment.sourcePart, entry[0]);
    }
  }
  const requiredParts = new Set<string>(),
    requiredRelations = new Set<string>();
  const requirePart = (part: string): void => {
    if (requiredParts.has(part)) return;
    requiredParts.add(part);
    for (const relation of native.relationships)
      if (relation.sourcePart === part) requireRelation(part, relation.id);
  };
  const requireRelation = (part: string, id: string): void => {
    const key = part + "\n" + id;
    if (requiredRelations.has(key)) return;
    const relation = native.relationships.find((r) => r.sourcePart === part && r.id === id);
    if (!relation) return;
    requiredRelations.add(key);
    if (!relation.external) {
      requirePart(relation.target);
    }
  };
  for (const relation of native.relationships)
    if (relation.root) requireRelation(relation.sourcePart, relation.id);
  for (const fragment of native.fragments) {
    for (const part of fragment.dependencyParts ?? []) requirePart(part);
    const root = readFragment(fragment);
    const collect = (node: XmlNode): void => {
      for (const [key, value] of node.attrs)
        if (key.startsWith(R + " ")) requireRelation(fragment.sourcePart, value);
      for (const child of node.content) if (typeof child !== "string") collect(child);
    };
    collect(root);
  }
  const retained = native.parts.filter(
    (part) => requiredParts.has(part.sourcePart) && !partMap.has(part.sourcePart),
  );
  retained.forEach((part, index) =>
    partMap.set(
      part.sourcePart,
      "word/preserved/part-" + (index + 1) + (part.sourcePart.match(/\.[^.\/]+$/)?.[0] ?? ".bin"),
    ),
  );
  const relationMaps = new Map<string, Map<string, string>>();
  const relationRoots = new Map<string, XmlNode>();
  const relationshipRoot = (part: string): XmlNode => {
    let root = relationRoots.get(part);
    if (!root) {
      const bytes = files.get(relPath(part));
      root = bytes
        ? read(bytes, relPath(part))
        : {
            name: "Relationships",
            uri: PKG,
            attrs: new Map(),
            namespaces: {},
            content: [],
            part: relPath(part),
          };
      relationRoots.set(part, root);
    }
    return root;
  };
  for (const rel of native.relationships) {
    if (!requiredRelations.has(rel.sourcePart + "\n" + rel.id)) continue;
    const outputPart = partMap.get(rel.sourcePart);
    if (!outputPart) continue;
    const target = rel.external ? rel.target : partMap.get(rel.target);
    if (!target) continue;
    const root = relationshipRoot(outputPart),
      relations = root.content.filter((n): n is XmlNode => typeof n !== "string");
    const resolved = rel.external ? target : "/" + target;
    const same = relations.find(
      (n) =>
        n.attrs.get(" Type") === rel.type &&
        n.attrs.get(" Target") === resolved &&
        (n.attrs.get(" TargetMode") === "External") === rel.external,
    );
    let id = same?.attrs.get(" Id");
    if (!id) {
      const used = new Set(relations.map((n) => n.attrs.get(" Id")));
      id = rel.id;
      for (let serial = 1; used.has(id); serial++) id = "nativeRel" + serial;
      root.content.push({
        name: "Relationship",
        uri: PKG,
        attrs: new Map([
          [" Id", id],
          [" Type", rel.type],
          [" Target", resolved],
          ...(rel.external ? ([[" TargetMode", "External"]] as [string, string][]) : []),
        ]),
        namespaces: {},
        content: [],
        part: relPath(outputPart),
      });
    }
    let mapping = relationMaps.get(rel.sourcePart);
    if (!mapping) {
      mapping = new Map();
      relationMaps.set(rel.sourcePart, mapping);
    }
    mapping.set(rel.id, id);
  }
  for (const part of retained) {
    if (!part.customXml) continue;
    const root = relationshipRoot("word/document.xml"),
      target = "/" + partMap.get(part.sourcePart)!,
      type = R + "/customXml";
    if (
      root.content.some(
        (n) =>
          typeof n !== "string" &&
          n.attrs.get(" Type") === type &&
          n.attrs.get(" Target") === target,
      )
    )
      continue;
    const ids = new Set(
      root.content.flatMap((n) => (typeof n === "string" ? [] : [n.attrs.get(" Id")])),
    );
    let id = "nativeCustomXml";
    for (let at = 1; ids.has(id); at++) id = "nativeCustomXml" + at;
    root.content.push({
      name: "Relationship",
      uri: PKG,
      attrs: new Map([
        [" Id", id],
        [" Type", type],
        [" Target", target],
      ]),
      namespaces: {},
      content: [],
      part: relPath("word/document.xml"),
    });
  }
  if (rawDefinitions.length && numberingPart) {
    const root = relationshipRoot("word/document.xml"),
      type = R + "/numbering";
    if (!root.content.some((n) => typeof n !== "string" && n.attrs.get(" Type") === type)) {
      const ids = new Set(
        root.content.flatMap((n) => (typeof n === "string" ? [] : [n.attrs.get(" Id")])),
      );
      let id = "nativeNumbering";
      for (let at = 1; ids.has(id); at++) id = "nativeNumbering" + at;
      root.content.push({
        name: "Relationship",
        uri: PKG,
        attrs: new Map([
          [" Id", id],
          [" Type", type],
          [" Target", "/" + numberingPart],
        ]),
        namespaces: {},
        content: [],
        part: relPath("word/document.xml"),
      });
    }
  }
  const nativeComments = new NativeWordComments(commentBindings);
  const drawingIdentities = new NativeDrawingIdentities(
    native,
    parsed.values(),
    model.snapshot().metadata.word!.revisions.map((_, i) => String(i)),
  );
  const bindingId = (fragment: (typeof native.fragments)[number], id: string): string => {
    const item = native.parts.find(
      (part) =>
        part.customXml &&
        bindingKey(part.customXml.storeItemId) === bindingKey(id) &&
        fragment.dependencyParts?.includes(part.sourcePart),
    );
    return (
      item?.customXml?.copiedStoreItemId ??
      item?.customXml?.storeItemId ??
      // Packages predating the indirect index retain their original binding.
      (fragment.dependencyParts === undefined ? id : isolatedStoreItemId(fragment.sourcePart, id))
    );
  };
  const nativeRangeIds = new Map<string, string>();
  const fragmentXml = (
    fragment: DeepReadonly<NonNullable<WordMetadata["nativeContent"]>["fragments"][number]>,
  ): XmlNode => {
    const copy = (node: XmlNode): XmlNode => ({
      ...node,
      attrs: new Map(node.attrs),
      content: node.content.map((item) => (typeof item === "string" ? item : copy(item))),
    });
    let root = copy(readFragment(fragment));
    const mapping = relationMaps.get(fragment.sourcePart);
    if (fragment.rangeBoundaries) {
      let at = 0;
      const removed = new WeakSet<XmlNode>(),
        handled = new WeakSet<XmlNode>();
      visitNativeRanges(root, (node) => {
        const control = nativeRangeControl(node);
        if (!control) return;
        handled.add(node);
        const boundary = fragment.rangeBoundaries![at++];
        if (
          !boundary ||
          boundary.family !== control.family ||
          boundary.sourceId !== control.sourceId ||
          boundary.side !== control.side
        )
          throw new DocumentError(
            "InvalidDocument",
            "A native range index does not match its fragment.",
          );
        if (boundary.active === false) removed.add(node);
        else {
          let id = nativeRangeIds.get(boundary.rangeId);
          if (id === undefined) {
            id = String(nativeRangeIds.size);
            nativeRangeIds.set(boundary.rangeId, id);
          }
          (node.attrs as Map<string, string>).set(W + " id", id);
        }
      });
      // Unselected compatibility branches can repeat the same logical control.
      // Keep their outside pairing consistent without counting them twice.
      visitNativeRanges(
        root,
        (node) => {
          if (handled.has(node)) return;
          const control = nativeRangeControl(node);
          if (!control) return;
          const boundary = fragment.rangeBoundaries!.find(
            (b) =>
              b.family === control.family &&
              b.sourceId === control.sourceId &&
              b.side === control.side,
          );
          if (!boundary) return;
          if (boundary.active === false) removed.add(node);
          else
            (node.attrs as Map<string, string>).set(
              W + " id",
              nativeRangeIds.get(boundary.rangeId)!,
            );
        },
        true,
      );
      if (at !== fragment.rangeBoundaries.length)
        throw new DocumentError("InvalidDocument", "A native range index is incomplete.");
      const prune = (node: XmlNode): void => {
        for (let i = node.content.length - 1; i >= 0; i--) {
          const child = node.content[i]!;
          if (typeof child === "string") continue;
          if (removed.has(child)) node.content.splice(i, 1);
          else prune(child);
        }
      };
      if (removed.has(root)) root = { ...root, uri: W, name: "r", attrs: new Map(), content: [] };
      else prune(root);
    }
    if (fragment.insertionSourceId) acceptNativeBlockInsertion(root, fragment.insertionSourceId);
    drawingIdentities.mark(root, fragment.sourcePart, fragment.nativeId);
    if (fragment.kind !== "property") nativeComments.restore(root, fragment.nodeId);
    if (fragment.currentRunPropertiesOnly && root.uri === W && root.name === "rPr")
      root.content.splice(
        0,
        root.content.length,
        ...root.content.filter(
          (child) =>
            typeof child === "string" ||
            child.uri !== W ||
            !["ins", "del", "moveFrom", "moveTo", "rPrChange"].includes(child.name),
        ),
      );
    const rewrite = (node: XmlNode): void => {
      if (fragment.numberingKind && node === root) {
        const attr =
          W +
          " " +
          (fragment.numberingKind === "num"
            ? "numId"
            : fragment.numberingKind === "abstractNum"
              ? "abstractNumId"
              : "numPicBulletId");
        (node.attrs as Map<string, string>).set(attr, numberingIds.get(fragment.nativeId)!);
      }
      if (node.uri === W) {
        const value = node.attrs.get(W + " val"),
          word = model.snapshot().metadata.word!;
        if (
          value &&
          ["pStyle", "rStyle", "tblStyle"].includes(node.name) &&
          fragment.styleReferences?.[value]
        )
          (node.attrs as Map<string, string>).set(W + " val", fragment.styleReferences[value]!);
        const referenceKey =
            value &&
            (node.name === "abstractNumId"
              ? "abstractNum:" + value
              : node.name === "lvlPicBulletId"
                ? "numPicBullet:" + value
                : node.name === "numId"
                  ? value
                  : undefined),
          reference = referenceKey ? fragment.numberingReferences?.[referenceKey] : undefined;
        if (reference) {
          const rawId = numberingIds.get(reference);
          const index = word.numbering.findIndex((n) => n.numberingId === reference);
          if (rawId !== undefined) (node.attrs as Map<string, string>).set(W + " val", rawId);
          else {
            if (index < 0)
              throw new DocumentError(
                "ExportFailed",
                "A preserved native numbering definition is missing.",
              );
            (node.attrs as Map<string, string>).set(W + " val", String(index + 1));
          }
        }
      }
      if (node.uri === W && node.name === "dataBinding") {
        const id = node.attrs.get(W + " storeItemID");
        if (id)
          (node.attrs as Map<string, string>).set(W + " storeItemID", bindingId(fragment, id));
      }
      for (const [key, value] of node.attrs)
        if (key.startsWith(R + " ") && mapping?.has(value))
          (node.attrs as Map<string, string>).set(key, mapping.get(value)!);
      for (const item of node.content) if (typeof item !== "string") rewrite(item);
    };
    rewrite(root);
    if (fragment.kind === "block") {
      const word = model.snapshot().metadata.word!,
        revision = word.revisions.find(
          (r) => r.nodeId === fragment.nodeId && r.kind === "nativeInsert",
        );
      if (
        revision &&
        !insertNativeBlock(root, {
          id: String(word.revisions.indexOf(revision)),
          author: revision.author ?? "",
          ...(revision.date ? { date: revision.date } : {}),
        })
      )
        unexportedNativeRevisionIds.push(revision.revisionId ?? revision.nodeId);
    }
    return root;
  };
  const restored = new Set<string>();
  const byNode = new Map<string, (typeof native.fragments)[number][]>();
  for (const fragment of native.fragments)
    if (fragment.nodeId) {
      const list = byNode.get(fragment.nodeId) ?? [];
      list.push(fragment);
      byNode.set(fragment.nodeId, list);
    }
  const mergeProperty = (target: XmlNode, original: XmlNode): void => {
    if (
      target.uri === W &&
      (wordTypographySettings.some((name) => name === target.name) ||
        ["overflowPunct", "topLinePunct", "doNotWrapTextWithPunct"].includes(target.name))
    ) {
      // Older document packages may still carry the whole declaration as a
      // residual. Keep extension attributes, with the compiled value taking precedence.
      for (const key of [W + " val", W + " lang"])
        if (target.attrs.has(key)) (original.attrs as Map<string, string>).delete(key);
    }
    if (target.uri === W && target.name === "compat" && mappedCompatibility.size)
      original.content.splice(
        0,
        original.content.length,
        ...target.content.filter(
          (child) =>
            typeof child !== "string" &&
            child.uri === W &&
            mappedCompatibility.has(wordCompatibilityKey(child)),
        ),
        ...original.content.filter(
          (child) =>
            typeof child === "string" ||
            child.uri !== W ||
            !mappedCompatibility.has(wordCompatibilityKey(child)),
        ),
      );
    merge(target, original);
  };
  const currentNoteProperties = (
    original: XmlNode,
    compiled: XmlNode | undefined,
    fragment: (typeof native.fragments)[number],
  ): void => {
    const metadata = model.snapshot().metadata,
      section = metadata.sections?.find(
        (section) =>
          section.rootNodeId === (compiled ? compiledNodeIds.get(compiled) : undefined) ||
          section.rootNodeId === fragment.nodeId,
      ),
      word = metadata.word;
    const patch = (
      raw: XmlNode,
      current: XmlNode | undefined,
      kind: "footnote" | "endnote",
    ): void => {
      const options = fragment.nodeId
        ? section?.[kind === "footnote" ? "footnoteOptions" : "endnoteOptions"]
        : word?.[kind === "footnote" ? "footnoteOptions" : "endnoteOptions"];
      if (!options) return;
      const names = new Set(
        Object.entries({
          pos: options.position,
          numFmt: options.format,
          numStart: options.start,
          numRestart: options.restart,
        })
          .filter(([, value]) => value !== undefined)
          .map(([name]) => name),
      );
      if (!names.size) return;
      if (raw.name === kind + "Pr") {
        raw.content.splice(
          0,
          raw.content.length,
          ...(current?.content ?? []).filter(
            (node) => typeof node !== "string" && node.uri === W && names.has(node.name),
          ),
          ...raw.content.filter(
            (node) => typeof node === "string" || node.uri !== W || !names.has(node.name),
          ),
        );
      } else if (names.has(raw.name)) (raw.attrs as Map<string, string>).delete(W + " val");
    };
    if (original.uri !== W) return;
    if (original.name === "sectPr") {
      for (const kind of ["footnote", "endnote"] as const) {
        const raw = original.content.find(
            (node): node is XmlNode =>
              typeof node !== "string" && node.uri === W && node.name === kind + "Pr",
          ),
          current = compiled?.content.find(
            (node): node is XmlNode =>
              typeof node !== "string" && node.uri === W && node.name === kind + "Pr",
          );
        if (raw) patch(raw, current, kind);
        else if (section?.[kind === "footnote" ? "footnoteOptions" : "endnoteOptions"] && current) {
          const predecessors = [
            "headerReference",
            "footerReference",
            ...(kind === "endnote" ? ["footnotePr"] : []),
          ];
          let index = 0;
          for (const [slot, node] of original.content.entries())
            if (typeof node !== "string" && node.uri === W && predecessors.includes(node.name))
              index = slot + 1;
          original.content.splice(index, 0, current);
        }
      }
      if (section?.noEndnote !== undefined) {
        const current =
          compiled?.content.filter(
            (node) => typeof node !== "string" && node.uri === W && node.name === "noEndnote",
          ) ?? [];
        let replaced = false;
        for (let index = original.content.length - 1; index >= 0; index--) {
          const child = original.content[index]!;
          if (typeof child !== "string" && child.uri === W && child.name === "noEndnote") {
            original.content.splice(index, 1, ...current);
            replaced = true;
          }
        }
        if (!replaced) {
          const trailing = [
              "titlePg",
              "textDirection",
              "bidi",
              "rtlGutter",
              "docGrid",
              "printerSettings",
              "sectPrChange",
            ],
            index = original.content.findIndex(
              (node) => typeof node !== "string" && node.uri === W && trailing.includes(node.name),
            );
          original.content.splice(index < 0 ? original.content.length : index, 0, ...current);
        }
      }
      const history = compiled?.content.find(
        (node): node is XmlNode =>
          typeof node !== "string" && node.uri === W && node.name === "sectPrChange",
      );
      if (
        history &&
        !original.content.some(
          (node) => typeof node !== "string" && node.uri === W && node.name === "sectPrChange",
        )
      )
        original.content.push(history);
    } else
      for (const kind of ["footnote", "endnote"] as const)
        if (original.name === kind + "Pr" || fragment.propertyPath.includes(W + " " + kind + "Pr"))
          patch(original, compiled, kind);
  };
  const mappedRunProperties = new Set([
    "rStyle",
    "rFonts",
    "b",
    "bCs",
    "i",
    "iCs",
    "caps",
    "smallCaps",
    "strike",
    "dstrike",
    "color",
    "spacing",
    "sz",
    "szCs",
    "highlight",
    "u",
    "vertAlign",
    "lang",
    "shd",
    "kern",
    "snapToGrid",
  ]);
  const applyProperty = (
    root: XmlNode,
    fragment: (typeof native.fragments)[number],
    previous = false,
    resetRun = false,
  ): void => {
    restored.add(fragment.nativeId);
    let target = root;
    for (const [index, path] of fragment.propertyPath.entries()) {
      let next = target.content.filter(
        (n): n is XmlNode =>
          typeof n !== "string" &&
          name(n) === path &&
          (path === W + " tblStylePr" && fragment.styleRegion
            ? n.attrs.get(W + " type") === fragment.styleRegion
            : true),
      )[
        path === W + " tblStylePr" && fragment.styleRegion
          ? 0
          : (fragment.propertyIndices?.[index] ?? 0)
      ];
      if (!next) {
        const split = path.lastIndexOf(" ");
        next = {
          uri: path.slice(0, split),
          name: path.slice(split + 1),
          attrs: new Map(),
          namespaces: {},
          content: [],
          part: root.part,
        };
        if (path === W + " tblStylePr" && fragment.styleRegion)
          (next.attrs as Map<string, string>).set(W + " type", fragment.styleRegion);
        target.content.unshift(next);
      }
      target = next;
    }
    const original = fragmentXml(fragment);
    if (
      resetRun &&
      !fragment.revisionPrevious &&
      fragment.propertyPath.length <= 1 &&
      original.uri === W
    ) {
      if (original.name === "rPr") {
        const compiled =
          target.name === "rPr"
            ? target
            : target.content.find(
                (child): child is XmlNode =>
                  typeof child !== "string" && child.uri === W && child.name === "rPr",
              );
        const current = compiled?.content ?? [],
          claims = new Set(
            current
              .filter(
                (child): child is XmlNode =>
                  typeof child !== "string" && child.name !== "rPrChange",
              )
              .map(name),
          );
        const unknown = original.content.filter(
          (child) =>
            typeof child === "string" ||
            ((child.uri !== W ||
              (!mappedRunProperties.has(child.name) && child.name !== "rPrChange")) &&
              !claims.has(name(child))),
        );
        const compiledHistory = current.filter(
            (child) => typeof child !== "string" && child.uri === W && child.name === "rPrChange",
          ),
          history = compiledHistory.length
            ? compiledHistory
            : original.content.filter(
                (child) =>
                  typeof child !== "string" && child.uri === W && child.name === "rPrChange",
              );
        original.content.splice(
          0,
          original.content.length,
          ...current.filter(
            (child) =>
              typeof child === "string" || !(child.uri === W && child.name === "rPrChange"),
          ),
          ...unknown,
          ...history,
        );
      } else if (fragment.propertyPath[0] === W + " rPr" && mappedRunProperties.has(original.name))
        return;
    }
    if (mappedCompatibility.size && original.uri === W) {
      if (original.name === "compat")
        original.content.splice(
          0,
          original.content.length,
          ...original.content.filter(
            (child) =>
              typeof child === "string" ||
              child.uri !== W ||
              !mappedCompatibility.has(wordCompatibilityKey(child)),
          ),
        );
      else if (
        mappedCompatibility.has(wordCompatibilityKey(original)) &&
        (original.name !== "doNotWrapTextWithPunct" || original.attrs.has(W + " val")) &&
        fragment.propertyPath.some((path) => path === W + " compat")
      )
        return;
    }
    if (original.name === "tblStylePr" && fragment.styleRegion && !original.attrs.has(W + " type"))
      (original.attrs as Map<string, string>).set(W + " type", fragment.styleRegion);
    if (!previous && !fragment.revisionPrevious)
      currentNoteProperties(
        original,
        name(original) === name(target)
          ? target
          : target.content.find(
              (node): node is XmlNode => typeof node !== "string" && name(node) === name(original),
            ),
        fragment,
      );
    if (name(original) === name(target)) mergeProperty(target, original);
    else {
      const current = target.content.filter(
        (n): n is XmlNode =>
          typeof n !== "string" &&
          name(n) === name(original) &&
          (original.name === "tblStylePr" && fragment.styleRegion
            ? n.attrs.get(W + " type") === fragment.styleRegion
            : true),
      )[original.name === "tblStylePr" && fragment.styleRegion ? 0 : (fragment.propertyIndex ?? 0)];
      if (current) mergeProperty(current, original);
      else if (
        !(
          mappedTypography &&
          target.name === "settings" &&
          original.uri === W &&
          wordTypographySettings.some((name) => name === original.name)
        )
      )
        target.content.push(original);
    }
  };
  const apply = (root: XmlNode): XmlNode => {
    const beforeSource = root.attrs.get(NATIVE_MARKER + " before-source"),
      beforeProperty = root.attrs.get(NATIVE_MARKER + " before-property"),
      beforeReset = root.attrs.get(NATIVE_MARKER + " before-reset") === "1";
    (root.attrs as Map<string, string>).delete(NATIVE_MARKER + " before-reset");
    (root.attrs as Map<string, string>).delete(NATIVE_MARKER + " before-source");
    (root.attrs as Map<string, string>).delete(NATIVE_MARKER + " before-property");
    if (beforeSource && beforeProperty) {
      const target = root.content.find(
        (node): node is XmlNode =>
          typeof node !== "string" && node.uri === W && node.name === beforeProperty,
      );
      if (!target)
        throw new DocumentError("ExportFailed", "The previous native property owner is missing.");
      for (const fragment of byNode.get(beforeSource) ?? []) {
        if (
          fragment.kind !== "property" ||
          fragment.revisionPrevious ||
          fragment.propertyPath.some((path) => path.endsWith("Change"))
        )
          continue;
        if (fragment.ownerName === beforeProperty)
          applyProperty(target, fragment, true, beforeReset);
        else if (fragment.propertyPath[0] === W + " " + beforeProperty)
          applyProperty(
            target,
            {
              ...fragment,
              propertyPath: fragment.propertyPath.slice(1),
              ...(fragment.propertyIndices
                ? { propertyIndices: fragment.propertyIndices.slice(1) }
                : {}),
            },
            true,
            beforeReset,
          );
        else if (
          !fragment.propertyPath.length &&
          fragment.propertyName === W + " " + beforeProperty
        )
          applyProperty(target, fragment, true, beforeReset);
      }
    }
    const resetRun = root.attrs.get(NATIVE_MARKER + " reset-run") === "1";
    (root.attrs as Map<string, string>).delete(NATIVE_MARKER + " reset-run");
    const id = root.attrs.get(NATIVE_MARKER + " node"),
      object = root.attrs.get(NATIVE_MARKER + " object");
    if (object) {
      const fragment = native.fragments.find((f) => f.nativeId === object && f.kind !== "property");
      if (!fragment)
        throw new DocumentError("InvalidDocument", "A native object descriptor is missing.");
      restored.add(fragment.nativeId);
      if (!fragment.runChild) return fragmentXml(fragment);
      const container = root.attrs.get(NATIVE_MARKER + " container");
      for (const key of ["node", "object", "container"])
        (root.attrs as Map<string, string>).delete(NATIVE_MARKER + " " + key);
      root.content.push(fragmentXml(fragment));
      for (const residual of container ? (byNode.get(container) ?? []) : [])
        if (residual.kind === "property") applyProperty(root, residual);
      return root;
    }
    const container = root.attrs.get(NATIVE_MARKER + " container");
    (root.attrs as Map<string, string>).delete(NATIVE_MARKER + " identity");
    (root.attrs as Map<string, string>).delete(NATIVE_MARKER + " node");
    (root.attrs as Map<string, string>).delete(NATIVE_MARKER + " container");
    for (let index = 0; index < root.content.length; index++) {
      const child = root.content[index]!;
      if (typeof child !== "string") root.content[index] = apply(child);
    }
    for (const fragment of [
      ...(id ? (byNode.get(id) ?? []) : []),
      ...(container ? (byNode.get(container) ?? []) : []),
    ])
      if (fragment.kind === "property" && fragment.ownerName === root.name)
        applyProperty(root, fragment, false, resetRun);
    return root;
  };
  for (const [path, root] of parsed) {
    apply(root);
    for (const fragment of native.fragments)
      if (
        fragment.kind === "property" &&
        !fragment.nodeId &&
        partMap.get(fragment.sourcePart) === path
      ) {
        const target = fragment.styleId
          ? root.content.find(
              (n): n is XmlNode =>
                typeof n !== "string" &&
                n.name === "style" &&
                n.attrs.get(W + " styleId") === fragment.styleId,
            )
          : root;
        if (!target)
          throw new DocumentError("ExportFailed", "A preserved native style owner is missing.");
        if (fragment.numberingKind) {
          target.content.push(fragmentXml(fragment));
          restored.add(fragment.nativeId);
        } else applyProperty(target, fragment);
      }
    sortWordSettings(root);
    sortWordProperties(root);
    sortWordNoteProperties(root);
    if (root.uri === W && root.name === "numbering") {
      const order = ["numPicBullet", "abstractNum", "num", "numIdMacAtCleanup"];
      root.content.sort(
        (a, b) =>
          (typeof a === "string" ? -1 : order.indexOf(a.name)) -
          (typeof b === "string" ? -1 : order.indexOf(b.name)),
      );
    }
  }
  nativeComments.removeDuplicates(parsed.values());
  drawingIdentities.rewrite(parsed.values());
  for (const [path, root] of parsed) {
    files.set(path, encoder.encode(declaration + serializeXml(root)));
  }
  for (const part of retained) {
    const bytes = sourceFiles.get(part.assetPath);
    if (!bytes) throw new DocumentError("AssetUnavailable", "A retained OPC part is missing.");
    const item = native.parts.find(
      (p) => p.customXml?.propertiesPart === part.sourcePart && p.customXml.copiedStoreItemId,
    );
    if (item?.customXml) {
      const root = read(bytes, part.sourcePart);
      if (root.uri !== CUSTOM_XML || root.name !== "datastoreItem")
        throw new DocumentError("InvalidDocument", "A custom XML properties root is invalid.");
      (root.attrs as Map<string, string>).set(
        CUSTOM_XML + " itemID",
        item.customXml.copiedStoreItemId!,
      );
      files.set(partMap.get(part.sourcePart)!, encoder.encode(declaration + serializeXml(root)));
    } else files.set(partMap.get(part.sourcePart)!, bytes);
  }
  for (const [part, root] of relationRoots)
    files.set(relPath(part), encoder.encode(declaration + serializeXml(root)));
  for (const fragment of native.fragments)
    if (
      !restored.has(fragment.nativeId) &&
      !retained.some((part) => part.sourcePart === fragment.sourcePart)
    )
      throw new DocumentError(
        "ExportFailed",
        "A preserved native fragment could not be attached to its current source owner.",
        [
          {
            code: "UnattachedNativeContent",
            path: [fragment.sourcePart, fragment.nativeId],
            message: fragment.label,
          },
        ],
      );
  const types = read(files.get("[Content_Types].xml")!, "[Content_Types].xml");
  if (
    rawDefinitions.length &&
    numberingPart &&
    !types.content.some(
      (n) => typeof n !== "string" && n.attrs.get(" PartName") === "/" + numberingPart,
    )
  )
    types.content.push({
      name: "Override",
      uri: CT,
      attrs: new Map([
        [" PartName", "/" + numberingPart],
        [
          " ContentType",
          "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
        ],
      ]),
      namespaces: {},
      content: [],
      part: "[Content_Types].xml",
    });
  for (const part of retained)
    types.content.push({
      name: "Override",
      uri: CT,
      attrs: new Map([
        [" PartName", "/" + partMap.get(part.sourcePart)!],
        [" ContentType", part.contentType],
      ]),
      namespaces: {},
      content: [],
      part: "[Content_Types].xml",
    });
  files.set("[Content_Types].xml", encoder.encode(declaration + serializeXml(types)));
  let total = 0;
  for (const bytes of files.values()) {
    total += bytes.length;
    if (bytes.length > limits.maxFileBytes)
      throw new DocumentError(
        "ResourceLimitExceeded",
        "A preserved Office part exceeds its byte budget.",
      );
  }
  if (files.size > limits.maxFiles || total > limits.maxPackageBytes)
    throw new DocumentError(
      "ResourceLimitExceeded",
      "Preserved Office output exceeds its package budget.",
    );
  return files;
}
/** Markers exist only between compilation and native reattachment. */
export function nativeObjectMarker(
  id: string,
  block: boolean,
  nodeId?: string,
  containerId?: string,
): string {
  return (
    "<w:" +
    (block ? "p" : "r") +
    ' xmlns:w="' +
    W +
    '" xmlns:dn="' +
    NATIVE_MARKER +
    '" dn:object="' +
    escapeXml(id) +
    '"' +
    (nodeId ? ' dn:node="' + escapeXml(nodeId) + '"' : "") +
    (containerId ? ' dn:container="' + escapeXml(containerId) + '"' : "") +
    "/>"
  );
}
