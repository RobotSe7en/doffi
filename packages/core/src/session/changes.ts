import type { ChangeMapping, DocumentNodeChange, TextOffsetMapping } from "../runtime.js";
import type { DocumentState } from "../document/state.js";
import { freeze } from "../document/validation.js";
import { affectedPresentationSlides } from "./presentation.js";

/** ReadonlyMap is a runtime boundary too: Object.freeze(new Map()) still exposes set(). */
export function readonlyMap<Key, Value>(input: Map<Key, Value>): ReadonlyMap<Key, Value> {
  const result: ReadonlyMap<Key, Value> = Object.freeze({
    get size() { return input.size; },
    has: (key: Key) => input.has(key),
    get: (key: Key) => input.get(key),
    keys: () => input.keys(),
    values: () => input.values(),
    entries: () => input.entries(),
    [Symbol.iterator]: () => input[Symbol.iterator](),
    forEach: (callback: (value: Value, key: Key, map: ReadonlyMap<Key, Value>) => void, thisArg?: unknown) => {
      input.forEach((value, key) => callback.call(thisArg, value, key, result));
    },
  });
  return result;
}

export function compareStates(before: DocumentState, after: DocumentState, textEdits: readonly TextOffsetMapping[] = []) {
  const changes: DocumentNodeChange[] = [];
  const nodes = new Map<string, string | null>();
  const affected = new Set<string>();
  const ids = new Set([...before.html.nodeIds(), ...after.html.nodeIds()]);
  for (const nodeId of ids) {
    const previous = before.html.signature(nodeId);
    const next = after.html.signature(nodeId);
    if (previous === next) continue;
    changes.push({ nodeId, kind: previous === undefined ? "added" : next === undefined ? "removed" : "modified" });
    if (previous !== undefined) nodes.set(nodeId, next === undefined ? null : nodeId);
    for (const state of [before, after]) {
      for (let id: string | null = nodeId; id !== null; id = state.html.peekNode(id)?.parentNodeId ?? null) affected.add(id);
    }
  }
  const previousAnnotations = new Map(before.snapshot.annotations.map((item) => [item.annotationId, JSON.stringify(item)]));
  const nextAnnotations = new Map(after.snapshot.annotations.map((item) => [item.annotationId, JSON.stringify(item)]));
  const changedAnnotationIds = [...new Set([...previousAnnotations.keys(), ...nextAnnotations.keys()])]
    .filter((id) => previousAnnotations.get(id) !== nextAnnotations.get(id));
  const mapping: ChangeMapping = Object.freeze({ nodes: readonlyMap(nodes), textEdits: freeze(textEdits.map((edit) => ({ ...edit }))) });
  const affectedSlideIds = before.snapshot.metadata.presentation ? affectedPresentationSlides(before, after) : undefined;
  if (affectedSlideIds) for (const slide of [...(before.snapshot.metadata.presentation?.slides ?? []), ...(after.snapshot.metadata.presentation?.slides ?? [])]) if (affectedSlideIds.includes(slide.slideId)) affected.add(slide.rootNodeId);
  return freeze({ changes, affectedNodeIds: [...affected], changedAnnotationIds, mapping, ...(affectedSlideIds ? { affectedSlideIds } : {}) });
}

/** Retained serialized bytes, counting shared immutable buffers only once. */
export function retainedBytes(states: Iterable<DocumentState>): number {
  const buffers = new Set<Uint8Array>();
  let total = 0;
  for (const state of states) for (const bytes of state.files.values()) {
    if (!buffers.has(bytes)) { buffers.add(bytes); total += bytes.byteLength; }
  }
  return total;
}
