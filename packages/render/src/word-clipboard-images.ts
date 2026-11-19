import { DocumentError } from "@doffi/core";
import type { ContentDocumentModel } from "@doffi/core";
import { WORD_CLIPBOARD_BYTES } from "./word-clipboard-constants.js";
const raster = /^image\/(png|jpeg|webp)$/;
/** Inline only the bytes already authorized by this model; never read a URL. */
export function copyWordClipboardImages(root: ParentNode, model: ContentDocumentModel): void {
  const snapshot = model.snapshot(),
    byNode = new Map(snapshot.metadata.word?.drawings?.map((item) => [item.nodeId, item]));
  let files: ReturnType<ContentDocumentModel["serialize"]> | undefined,
    total = 0;
  for (const image of root.querySelectorAll<HTMLImageElement>("img")) {
    const url = new URL(
        image.getAttribute("src")!,
        "https://doffi.invalid/" + snapshot.manifest.entrypoint,
      ),
      path = decodeURIComponent(url.pathname.slice(1)),
      asset = snapshot.manifest.assets.find((asset) => asset.path === path);
    if (url.origin !== "https://doffi.invalid" || !asset || !raster.test(asset.mimeType))
      throw new DocumentError("AssetUnavailable", "复制图片需要当前文档已授权的栅格图片资源。");
    const bytes = model.getAsset
      ? model.getAsset(path, WORD_CLIPBOARD_BYTES)
      : (files ??= model.serialize()).get(path);
    if (!bytes) throw new DocumentError("AssetUnavailable", "复制图片的资源字节不可用。");
    total += bytes.byteLength;
    if (total > WORD_CLIPBOARD_BYTES)
      throw new DocumentError("ResourceLimitExceeded", "剪贴板图片超过 64 MiB。");
    let binary = "";
    for (let offset = 0; offset < bytes.length; offset += 32768)
      binary += String.fromCharCode(...bytes.subarray(offset, offset + 32768));
    image.setAttribute("src", "data:" + asset.mimeType + ";base64," + btoa(binary));
    const drawing = byNode.get(image.dataset.nodeId ?? "");
    if (drawing)
      image.setAttribute(
        "style",
        (image.getAttribute("style") ?? "") +
          ";width:" +
          drawing.geometry.width +
          "pt;height:" +
          drawing.geometry.height +
          "pt",
      );
  }
}
