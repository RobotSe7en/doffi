import { sha256 } from "@noble/hashes/sha2.js";
import { bytesToHex } from "@noble/hashes/utils.js";
import type { AssetDescriptor } from "../contracts/document.js";
import { DocumentError } from "../errors.js";
import type { DeepReadonly, DocumentPackage } from "../runtime.js";
import { invalid, unsupported, checkLimit } from "./validation.js";

export function copyContentAsset(
  files: DocumentPackage,
  assets: DeepReadonly<readonly AssetDescriptor[]>,
  path: string,
  maxBytes: number,
): Uint8Array | undefined {
  if (!Number.isSafeInteger(maxBytes) || maxBytes < 1)
    invalid("Asset byte limits must be positive safe integers.");
  if (!assets.some((asset) => asset.path === path)) return undefined;
  const bytes = files.get(path);
  if (!bytes) return undefined;
  checkLimit(bytes.byteLength, maxBytes);
  return new Uint8Array(bytes);
}

function signature(bytes: Uint8Array, offset: number, expected: readonly number[]): boolean {
  return expected.every((byte, index) => bytes[offset + index] === byte);
}

/** Identify supported raster containers without decoding pixels or rewriting image bytes. */
function imageMimeType(bytes: Uint8Array): string | undefined {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  if (signature(bytes, 0, [137, 80, 78, 71, 13, 10, 26, 10])) {
    if (
      bytes.length < 45 ||
      !signature(bytes, 8, [0, 0, 0, 13, 73, 72, 68, 82]) ||
      !view.getUint32(16) ||
      !view.getUint32(20)
    )
      return undefined;
    let hasData = false;
    for (let offset = 8; offset + 12 <= bytes.length; ) {
      const size = view.getUint32(offset);
      const end = offset + size + 12;
      if (end > bytes.length) return undefined;
      if (signature(bytes, offset + 4, [73, 68, 65, 84])) hasData = true;
      if (signature(bytes, offset + 4, [73, 69, 78, 68]))
        return hasData && !size && end === bytes.length ? "image/png" : undefined;
      offset = end;
    }
  }
  if (signature(bytes, 0, [255, 216, 255]) && signature(bytes, bytes.length - 2, [255, 217])) {
    let hasFrame = false;
    for (let offset = 2; offset + 4 <= bytes.length; ) {
      if (bytes[offset++] !== 255) return undefined;
      while (bytes[offset] === 255) offset++;
      const marker = bytes[offset++];
      if (marker === 0xda) return hasFrame ? "image/jpeg" : undefined;
      if (offset + 2 > bytes.length) return undefined;
      const size = view.getUint16(offset);
      if (size < 2 || offset + size > bytes.length) return undefined;
      if (
        marker !== undefined &&
        marker >= 0xc0 &&
        marker <= 0xcf &&
        ![0xc4, 0xc8, 0xcc].includes(marker)
      ) {
        if (size < 8 || !view.getUint16(offset + 3) || !view.getUint16(offset + 5))
          return undefined;
        hasFrame = true;
      }
      offset += size;
    }
  }
  if (
    bytes.length >= 20 &&
    signature(bytes, 0, [82, 73, 70, 70]) &&
    signature(bytes, 8, [87, 69, 66, 80]) &&
    view.getUint32(4, true) === bytes.length - 8
  ) {
    let offset = 12;
    let hasImage = false;
    while (offset + 8 <= bytes.length) {
      const size = view.getUint32(offset + 4, true);
      if (
        (signature(bytes, offset, [86, 80, 56, 32]) && size >= 10) ||
        (signature(bytes, offset, [86, 80, 56, 76]) && size >= 5) ||
        (signature(bytes, offset, [65, 78, 77, 70]) && size >= 16)
      )
        hasImage = true;
      offset += 8 + size + (size & 1);
      if (offset > bytes.length) return undefined;
    }
    if (offset === bytes.length && hasImage) return "image/webp";
  }
  return undefined;
}

export function assetDigest(bytes: Uint8Array): string {
  return bytesToHex(sha256(bytes));
}

export function validateAsset(asset: DeepReadonly<AssetDescriptor>, bytes: Uint8Array): void {
  if (assetDigest(bytes) !== asset.sha256)
    invalid("A package resource does not match its declared SHA-256 digest.");
  if (
    ["image/png", "image/jpeg", "image/webp"].includes(asset.mimeType) &&
    imageMimeType(bytes) !== asset.mimeType
  )
    invalid("An image does not match its declared raster format.");
}

export function validateAssets(
  assets: readonly DeepReadonly<AssetDescriptor>[],
  files: DocumentPackage,
): void {
  for (const asset of assets) {
    const bytes = files.get(asset.path);
    if (!bytes)
      throw new DocumentError(
        "AssetUnavailable",
        "A declared resource is not supplied in the package.",
      );
    validateAsset(asset, bytes);
  }
}

export function validateImageAsset(asset: DeepReadonly<AssetDescriptor>): void {
  if (!["image/png", "image/jpeg", "image/webp"].includes(asset.mimeType)) {
    unsupported("Only PNG, JPEG and WebP resources can be used as document images.");
  }
}
