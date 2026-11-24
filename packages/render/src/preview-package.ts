import { captureContentModel } from "@doffi/core";
import type { ContentDocumentModel } from "@doffi/core";

export interface PreparedPreview {
  readonly model: ContentDocumentModel;
  readonly document: Document;
  readonly imageUrls: ReadonlyMap<string, string>;
}
export function packagePath(reference: string, entrypoint: string): string {
  const parts = entrypoint.split("/").slice(0, -1);
  for (const token of reference.split("/")) { const value = decodeURIComponent(token); if (value === "..") parts.pop(); else if (value !== ".") parts.push(value); }
  return parts.join("/");
}
export function safeStyleText(source: string): string { return source.replace(/<\/style/gi, value => "<\\" + value.slice(1)); }
/** Shared scriptless package preparation. Object URLs are scoped to this view,
 * including repeated references to one asset, and are revoked by its owner. */
export function releasePreviewUrls(previous: PreparedPreview, next?: PreparedPreview): void {
  for (const [path, url] of previous.imageUrls)
    if (next?.imageUrls.get(path) !== url) URL.revokeObjectURL(url);
}

export function preparePreviewPackage(input: ContentDocumentModel, previous?: PreparedPreview): PreparedPreview {
  const model = captureContentModel(input), snapshot = model.snapshot();
  const document = new DOMParser().parseFromString(snapshot.html, "text/html");
  for (const link of document.querySelectorAll("link")) link.remove();
  const policy = document.createElement("meta"); policy.httpEquiv = "Content-Security-Policy";
  policy.content = "default-src 'none'; img-src blob:; style-src 'unsafe-inline'; font-src 'none'; base-uri 'none'; form-action 'none'"; document.head.prepend(policy);
  const style = document.createElement("style"); style.dataset.sdkSourceStyles = "";
  style.textContent = safeStyleText(snapshot.styles.map(s => s.css).join("\n")); document.head.append(style);
  const assets = new Map(snapshot.manifest.assets.map(asset => [asset.path, asset]));
  const previousAssets = new Map(previous?.model.snapshot().manifest.assets.map(asset => [asset.path, asset]));
  const imageUrls = new Map<string, string>(), prepared = { model, document, imageUrls };
  try {
    for (const image of document.querySelectorAll("img")) {
      const path = packagePath(image.getAttribute("src")!, snapshot.manifest.entrypoint);
      let url = imageUrls.get(path);
      if (!url) {
        const asset = assets.get(path)!, prior = previousAssets.get(path);
        if (prior?.sha256 === asset.sha256 && prior.mimeType === asset.mimeType)
          url = previous?.imageUrls.get(path);
        if (!url) {
          const bytes = model.getAsset!(path)!, buffer = new ArrayBuffer(bytes.length);
          new Uint8Array(buffer).set(bytes);
          url = URL.createObjectURL(new Blob([buffer], { type: asset.mimeType }));
        }
        imageUrls.set(path, url);
      }
      image.src = url;
    }
  } catch (error) {
    releasePreviewUrls(prepared, previous);
    throw error;
  }
  return prepared;
}
