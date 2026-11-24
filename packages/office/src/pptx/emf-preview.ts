import { zlibSync } from "fflate";
import { DocumentError } from "@doffi/core";
import { fail } from "../ooxml/xml.js";
import { crc32 } from "../zip-format.js";

const signature = Uint8Array.of(137, 80, 78, 71, 13, 10, 26, 10);
// Framing, state and comments can surround a legacy raster fallback.
const fallbackRecords = new Set([1, 13, 14, 18, 19, 20, 21, 22, 24, 25, 27, 28, 33, 34, 35, 36, 37, 40, 48, 58, 70, 76, 81, 82, 98, 115]);

function chunk(name: string, data: Uint8Array): Uint8Array {
  const bytes = new Uint8Array(data.length + 12);
  const fields = new DataView(bytes.buffer);
  fields.setUint32(0, data.length);
  for (let index = 0; index < name.length; index++) bytes[index + 4] = name.charCodeAt(index);
  bytes.set(data, 8);
  fields.setUint32(bytes.length - 4, crc32(bytes.subarray(4, bytes.length - 4)));
  return bytes;
}

/** Extract a full-frame SRCCOPY bitmap. The native metafile remains opaque. */
export function emfPreview(bytes: Uint8Array, part: string, maxBytes: number, count: () => void): Uint8Array {
  const source = { part, name: "image" };
  const malformed = (): never => fail(source, "The EMF bitmap or record framing is invalid.");
  const unsupported = (): never => fail(source, "Compatible preview needs one full-frame EMF bitmap.", "UnsupportedFeature");
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const u32 = (offset: number) => view.getUint32(offset, true);
  const i32 = (offset: number) => view.getInt32(offset, true);
  if (bytes.length < 88 || u32(0) !== 1 || u32(4) < 88 || u32(40) !== 0x464d4520 || u32(48) !== bytes.length)
    return malformed();
  let bitmap: { offset: number; width: number; height: number; bottomUp: boolean } | undefined;
  let records = 0, finished = false;
  for (let start = 0; start < bytes.length;) {
    count();
    if (finished || bytes.length - start < 8) return malformed();
    const type = u32(start), size = u32(start + 4);
    if (size < 8 || size % 4 || size > bytes.length - start) return malformed();
    if (!fallbackRecords.has(type)) return unsupported();
    // Some dual metafiles end their fallback with a destination-only blit.
    if (type === 76 && (size < 100 || u32(start + 40) !== 0x00aa0029)) return unsupported();
    records++;
    if (type === 14) {
      if (size < 20 || start + size !== bytes.length) return malformed();
      finished = true;
    }
    if (type === 81) {
      if (size < 80) return malformed();
      const header = u32(start + 48), headerLength = u32(start + 52);
      const pixels = u32(start + 56), pixelsLength = u32(start + 60);
      if (header < 80 || headerLength < 40 || header + headerLength > size ||
          pixels < header + headerLength || pixels + pixelsLength > size) return malformed();
      const dib = start + header, width = i32(dib + 4), signedHeight = i32(dib + 8), height = Math.abs(signedHeight);
      const compression = u32(dib + 16);
      if (width <= 0 || height === 0 || view.getUint16(dib + 12, true) !== 1) return malformed();
      if (bitmap || u32(start + 64) !== 0 || u32(start + 68) !== 0x00cc0020 ||
          i32(start + 32) !== 0 || i32(start + 36) !== 0 || i32(start + 40) !== width || i32(start + 44) !== height ||
          i32(start + 72) <= 0 || i32(start + 76) <= 0 ||
          [8, 12, 16, 20].some(offset => Math.abs(i32(start + offset) - i32(offset)) > 1) ||
          u32(dib) !== 40 || view.getUint16(dib + 14, true) !== 32 || (compression !== 0 && compression !== 3))
        return unsupported();
      if (compression === 3 && (headerLength < 52 || u32(dib + 40) !== 0xff0000 || u32(dib + 44) !== 0xff00 || u32(dib + 48) !== 0xff))
        return unsupported();
      if (width * height > maxBytes / 4) throw new DocumentError("ResourceLimitExceeded", "The EMF bitmap exceeds the image byte budget.");
      if (pixelsLength !== width * height * 4) return malformed();
      bitmap = { offset: start + pixels, width, height, bottomUp: signedHeight > 0 };
    }
    start += size;
  }
  if (!finished || records !== u32(52)) return malformed();
  if (!bitmap) return unsupported();

  const { offset, width, height, bottomUp } = bitmap;
  const stride = width * 3 + 1;
  const scanlines = new Uint8Array(stride * height);
  for (let y = 0; y < height; y++) {
    const row = offset + (bottomUp ? height - y - 1 : y) * width * 4;
    for (let x = 0; x < width; x++) {
      const source = row + x * 4, target = y * stride + x * 3 + 1;
      scanlines[target] = bytes[source + 2]!;
      scanlines[target + 1] = bytes[source + 1]!;
      scanlines[target + 2] = bytes[source]!;
    }
  }
  const header = new Uint8Array(13), dimensions = new DataView(header.buffer);
  dimensions.setUint32(0, width);
  dimensions.setUint32(4, height);
  header.set([8, 2], 8);
  const chunks = [signature, chunk("IHDR", header), chunk("IDAT", zlibSync(scanlines)), chunk("IEND", new Uint8Array())];
  const png = new Uint8Array(chunks.reduce((length, item) => length + item.length, 0));
  let cursor = 0;
  for (const item of chunks) {
    png.set(item, cursor);
    cursor += item.length;
  }
  return png;
}
