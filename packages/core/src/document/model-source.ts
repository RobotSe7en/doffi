import type { ContentDocumentModel } from "../runtime.js";
import { copyContentAsset } from "./assets.js";
import type { DocumentState } from "./state.js";
import { materializeDocumentFiles } from "./state.js";

const sources = new WeakMap<ContentDocumentModel, () => DocumentState>();

export function registerContentSource(model: ContentDocumentModel, source: () => DocumentState): void {
  sources.set(model, source);
}

export function contentSource(model: ContentDocumentModel): DocumentState | undefined {
  return sources.get(model)?.();
}

export function contentModel(source: DocumentState): ContentDocumentModel {
  const model = Object.freeze<ContentDocumentModel>({
    kind: source.snapshot.manifest.kind,
    snapshot: () => source.snapshot,
    getNode: id => source.html.getNode(id),
    getAsset: (path, maxBytes = source.limits.maxFileBytes) =>
      copyContentAsset(source.files, source.snapshot.manifest.assets, path, maxBytes),
    serialize: () => new Map([...materializeDocumentFiles(source)].map(([path, bytes]) => [path, new Uint8Array(bytes)])),
  });
  registerContentSource(model, () => source);
  return model;
}
