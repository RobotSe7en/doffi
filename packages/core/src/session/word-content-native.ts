import type { WordNativeContent, WordMetadata } from "../contracts/word.js";
import type { DocumentEdit } from "../document/edit.js";
import { createOpaqueId } from "../identifiers.js";
import { DocumentError } from "../errors.js";

/** Attach the selected native graph. Resource authorization/validation still
 * belongs to the transaction's asset adapter and the complete package reader. */
export function attachWordContentNative(
  word: WordMetadata,
  edit: DocumentEdit,
  ids: ReadonlyMap<string, string>,
  supplied: WordNativeContent | undefined,
  numberingIds: ReadonlyMap<string, string> = new Map(),
): void {
  if (!supplied) return;
  const prefix = "word/clipboard-" + createOpaqueId() + "/",
    parts = new Map<string, string>();
  const part = (id: string): string => {
    let mapped = parts.get(id);
    if (!mapped) {
      mapped = prefix + id;
      parts.set(id, mapped);
    }
    return mapped;
  };
  const map = (id: string): string =>
    ids.get(id) ??
    (() => {
      throw new DocumentError(
        "InvalidOperation",
        "Copied native endpoints require selected transport nodes.",
      );
    })();
  if (
    supplied.relationships.some((r) => r.root) ||
    supplied.fragments.some(
      (f) => (!f.nodeId && !f.styleId && !f.numberingKind) || f.revisionSourceId,
    )
  )
    throw new DocumentError(
      "InvalidOperation",
      "Copied native content cannot carry document root or revision ownership.",
    );
  const declaredParts = new Map(supplied.parts.map((p) => [p.sourcePart, p])),
    relations = new Map(supplied.relationships.map((r) => [r.sourcePart + "\n" + r.id, r])),
    kept = new Set<string>(),
    selected = new Set<string>(),
    pending: string[] = [];
  const requirePart = (path: string): void => {
    if (!declaredParts.has(path))
      throw new DocumentError("InvalidOperation", "A copied native dependency is missing.");
    if (!kept.has(path)) {
      kept.add(path);
      pending.push(path);
    }
  };
  const select = (key: string): void => {
    if (selected.has(key)) return;
    const relation = relations.get(key);
    if (!relation)
      throw new DocumentError("InvalidOperation", "A copied native relationship is missing.");
    selected.add(key);
    if (!relation.external) requirePart(relation.target);
  };
  for (const fragment of supplied.fragments) {
    if (fragment.relationshipIds === undefined || fragment.fragmentFormat !== "json")
      throw new DocumentError(
        "InvalidOperation",
        "Copied native fragments require an explicit dependency index.",
      );
    for (const id of fragment.relationshipIds) select(fragment.sourcePart + "\n" + id);
    for (const path of fragment.dependencyParts ?? []) requirePart(path);
  }
  for (let at = 0; at < pending.length; at++)
    for (const [key, relation] of relations) if (relation.sourcePart === pending[at]) select(key);
  if (kept.size !== supplied.parts.length || selected.size !== supplied.relationships.length)
    throw new DocumentError(
      "InvalidOperation",
      "Copied native content contains unowned package parts or relationships.",
    );
  const definitions = new Map(
      supplied.fragments.filter((f) => f.numberingKind).map((f) => [f.nativeId, f]),
    ),
    definitionQueue = supplied.fragments.filter((f) => !f.numberingKind),
    selectedDefinitions = new Set<string>();
  for (let at = 0; at < definitionQueue.length; at++)
    for (const id of Object.values(definitionQueue[at]!.numberingReferences ?? {})) {
      const definition = definitions.get(id);
      if (definition && !selectedDefinitions.has(id)) {
        selectedDefinitions.add(id);
        definitionQueue.push(definition);
      }
    }
  if (selectedDefinitions.size !== definitions.size)
    throw new DocumentError(
      "InvalidOperation",
      "Copied native numbering definitions must belong to selected content.",
    );
  const native = (word.nativeContent ??= {
    mainPart: "word/document.xml",
    fragments: [],
    parts: [],
    relationships: [],
  });
  const nativeIds = new Map(
    supplied.fragments.map((f) => [f.nativeId, "word-content-native-" + createOpaqueId()]),
  );
  const rangeIds = new Map<string, string>();
  native.fragments.push(
    ...supplied.fragments.map((original) => {
      const copy = JSON.parse(JSON.stringify(original)) as typeof original;
      if (copy.nodeId) copy.nodeId = map(copy.nodeId);
      else if (!copy.numberingKind && !word.styles.some((style) => style.styleId === copy.styleId))
        throw new DocumentError(
          "InvalidOperation",
          "A copied native property requires its imported style.",
        );
      copy.nativeId = nativeIds.get(original.nativeId)!;
      copy.sourcePart = part(copy.sourcePart);
      for (const boundary of copy.rangeBoundaries ?? []) {
        let id = rangeIds.get(boundary.rangeId);
        if (!id) {
          id = "native-range-" + createOpaqueId();
          rangeIds.set(boundary.rangeId, id);
        }
        boundary.rangeId = id;
      }
      if (copy.dependencyParts) copy.dependencyParts = copy.dependencyParts.map(part);
      if (copy.numberingReferences)
        copy.numberingReferences = Object.fromEntries(
          Object.entries(copy.numberingReferences).map(([id, definition]) => {
            const mapped = nativeIds.get(definition) ?? numberingIds.get(definition);
            if (!mapped)
              throw new DocumentError(
                "InvalidOperation",
                "A native numbering reference requires its copied definition.",
              );
            return [id, mapped];
          }),
        );
      if (copy.kind !== "property" && copy.nodeId)
        edit.attributes(copy.nodeId, { "data-word-native-object": copy.nativeId });
      return copy;
    }),
  );
  native.parts.push(
    ...supplied.parts.map((p) => ({
      ...p,
      sourcePart: part(p.sourcePart),
      ...(p.customXml
        ? {
            customXml: {
              ...p.customXml,
              propertiesPart: part(p.customXml.propertiesPart),
              copiedStoreItemId: "{" + createOpaqueId().toUpperCase() + "}",
            },
          }
        : {}),
    })),
  );
  native.relationships.push(
    ...supplied.relationships.map((r) => ({
      ...r,
      sourcePart: part(r.sourcePart),
      target: r.external ? r.target : part(r.target),
    })),
  );
  if (supplied.identities?.length)
    (native.identities ??= []).push(
      ...supplied.identities.map((i) => ({
        ...i,
        nodeId: map(i.nodeId),
        sourcePart: part(i.sourcePart),
      })),
    );
}
