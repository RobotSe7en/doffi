import { createAssetDescriptor, createOpaqueId, DocumentError } from "@doffi/core";
import type { TransactionPrepareOptions } from "@doffi/core";
import type { WordClipboardFormat } from "@doffi/render";
import { WORD_CLIPBOARD_BYTES } from "@doffi/render";

export function mergeClipboardResources(
  ...options: (TransactionPrepareOptions | undefined)[]
): TransactionPrepareOptions | undefined {
  const assets = new Map<string, NonNullable<TransactionPrepareOptions["assets"]>[number]>();
  for (const option of options)
    for (const asset of option?.assets ?? []) {
      const previous = assets.get(asset.descriptor.path);
      if (
        previous &&
        (previous.descriptor.sha256 !== asset.descriptor.sha256 ||
          previous.descriptor.mimeType !== asset.descriptor.mimeType)
      )
        throw new DocumentError("InvalidOperation", "剪贴板资源身份冲突。");
      assets.set(asset.descriptor.path, asset);
    }
  return assets.size ? { assets: [...assets.values()] } : undefined;
}
/** Decode the selected chart graph only; package bytes never grant access to
 * an existing receiver path. All descriptors are regenerated from actual bytes. */
export function pasteClipboardResources(
  format: WordClipboardFormat,
): TransactionPrepareOptions | undefined {
  const parts = (format.semantics?.drawings ?? []).flatMap((d) =>
      d.kind === "chart" ? (d.chart.native?.parts ?? []) : [],
    ),
    required = new Set(parts.map((p) => p.assetPath)),
    paths = new Map<string, string>(),
    assets: NonNullable<TransactionPrepareOptions["assets"]>[number][] = [],
    prefix = createOpaqueId();
  if (
    format.resources !== undefined &&
    (!Array.isArray(format.resources) || format.resources.length > 1024)
  )
    throw new DocumentError("InvalidOperation", "剪贴板原生资源列表无效。");
  let total = 0;
  for (const item of format.resources ?? []) {
    if (
      !item ||
      typeof item.path !== "string" ||
      typeof item.mimeType !== "string" ||
      typeof item.data !== "string" ||
      !required.has(item.path) ||
      paths.has(item.path)
    )
      throw new DocumentError("InvalidOperation", "剪贴板资源必须属于所选图表的依赖图。");
    if (item.data.length > Math.ceil(WORD_CLIPBOARD_BYTES / 3) * 4)
      throw new DocumentError("ResourceLimitExceeded", "剪贴板资源过大。");
    let binary: string;
    try {
      binary = atob(item.data);
    } catch {
      throw new DocumentError("InvalidOperation", "剪贴板资源字节无效。");
    }
    const bytes = Uint8Array.from(binary, (c) => c.charCodeAt(0));
    total += bytes.length;
    if (total > WORD_CLIPBOARD_BYTES)
      throw new DocumentError("ResourceLimitExceeded", "剪贴板资源过大。");
    const descriptor = createAssetDescriptor(
      "assets/clipboard-native-" + prefix + "-" + assets.length + ".bin",
      item.mimeType,
      bytes,
    );
    paths.set(item.path, descriptor.path);
    assets.push({ descriptor, bytes });
  }
  for (const part of parts)
    part.assetPath =
      paths.get(part.assetPath) ??
      (() => {
        throw new DocumentError("AssetUnavailable", "剪贴板图表缺少依赖资源。");
      })();
  return assets.length ? { assets } : undefined;
}
