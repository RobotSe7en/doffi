import type { WordNativeContent, WordMetadata } from "../contracts/word.js";
import type { DocumentEdit } from "../document/edit.js";
import { createOpaqueId } from "../identifiers.js";
import { DocumentError } from "../errors.js";

/** Attach only selected opaque atoms. Resource authorization/validation still
 * belongs to the transaction's asset adapter and the complete package reader. */
export function attachWordContentNative(
  word: WordMetadata,
  edit: DocumentEdit,
  ids: ReadonlyMap<string, string>,
  supplied: WordNativeContent | undefined,
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
    supplied.fragments.some((f) => !f.nodeId || f.revisionSourceId)
  )
    throw new DocumentError(
      "InvalidOperation",
      "Copied native endpoints cannot carry document root or revision ownership.",
    );
  const native = (word.nativeContent ??= {
    mainPart: "word/document.xml",
    fragments: [],
    parts: [],
    relationships: [],
  });
  native.fragments.push(
    ...supplied.fragments.map((original) => {
      const copy = JSON.parse(JSON.stringify(original)) as typeof original;
      copy.nodeId = map(copy.nodeId!);
      copy.nativeId = "word-content-native-" + createOpaqueId();
      copy.sourcePart = part(copy.sourcePart);
      if (copy.kind !== "property")
        edit.attributes(copy.nodeId, { "data-word-native-object": copy.nativeId });
      return copy;
    }),
  );
  native.parts.push(...supplied.parts.map((p) => ({ ...p, sourcePart: part(p.sourcePart) })));
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
