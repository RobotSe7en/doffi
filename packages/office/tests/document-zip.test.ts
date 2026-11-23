import { readFileSync } from "node:fs";
import { crc32, deflateRawSync } from "node:zlib";
import { describe, expect, test } from "vitest";
import { DocumentError, openDocument, parseContract } from "@doffi/core";
import type { DocumentErrorCode } from "@doffi/core";
import { DEFAULT_MAX_ZIP_BYTES, readDocumentZip, writeDocumentZip, writeDocumentZipAsync } from "@doffi/office";

const encoder = new TextEncoder();
function files(): Map<string, Uint8Array> {
  return new Map(["manifest.json", "content.html", "styles.css", "metadata.json", "annotations.json",
    "assets/sample.png", "assets/sample.jpg", "assets/sample.webp"].map((path) =>
    [path, new Uint8Array(readFileSync(new URL("../../../examples/fidelity/" + path, import.meta.url)))]));
}
function error(code: DocumentErrorCode) {
  return (value: unknown) => {
    expect(value).toBeInstanceOf(DocumentError);
    const data = (value as DocumentError).toJSON();
    expect(parseContract("document-error", data).code).toBe(code);
    expect(JSON.stringify(data)).not.toContain("private text");
    return true;
  };
}

interface Entry {
  name: string;
  data?: Uint8Array;
  method?: number;
  size?: number;
  flags?: number;
  localName?: string;
  attrs?: number;
  crc?: number;
  rawName?: Uint8Array;
}

/** Independent classic ZIP fixtures permit duplicate names and dishonest metadata. */
function rawZip(entries: Entry[]): Uint8Array {
  const localParts: Uint8Array[] = [];
  const centralParts: Uint8Array[] = [];
  let localSize = 0;
  let centralSize = 0;
  for (const entry of entries) {
    const name = entry.rawName ?? encoder.encode(entry.name);
    const localName = entry.localName ? encoder.encode(entry.localName) : name;
    const data = entry.data ?? encoder.encode("private text");
    const method = entry.method ?? 0;
    const compressed = method === 8 ? deflateRawSync(data) : data;
    const crc = entry.crc ?? crc32(data);
    const header = new Uint8Array(30 + localName.length + compressed.length);
    const central = new Uint8Array(46 + name.length);
    const localView = new DataView(header.buffer);
    const centralView = new DataView(central.buffer);
    localView.setUint32(0, 0x04034b50, true);
    localView.setUint16(4, 20, true);
    localView.setUint16(6, entry.flags ?? 0x800, true);
    localView.setUint16(8, method, true);
    localView.setUint32(14, crc, true);
    localView.setUint32(18, compressed.length, true);
    localView.setUint32(22, entry.size ?? data.length, true);
    localView.setUint16(26, localName.length, true);
    header.set(localName, 30);
    header.set(compressed, 30 + localName.length);
    centralView.setUint32(0, 0x02014b50, true);
    centralView.setUint16(4, 0x314, true);
    centralView.setUint16(6, 20, true);
    centralView.setUint16(8, entry.flags ?? 0x800, true);
    centralView.setUint16(10, method, true);
    centralView.setUint32(16, crc, true);
    centralView.setUint32(20, compressed.length, true);
    centralView.setUint32(24, entry.size ?? data.length, true);
    centralView.setUint16(28, name.length, true);
    centralView.setUint32(38, entry.attrs ?? 0, true);
    centralView.setUint32(42, localSize, true);
    central.set(name, 46);
    localParts.push(header); centralParts.push(central);
    localSize += header.length; centralSize += central.length;
  }
  const result = new Uint8Array(localSize + centralSize + 22);
  let cursor = 0;
  for (const part of [...localParts, ...centralParts]) { result.set(part, cursor); cursor += part.length; }
  const end = new DataView(result.buffer, cursor);
  end.setUint32(0, 0x06054b50, true);
  end.setUint16(8, entries.length, true); end.setUint16(10, entries.length, true);
  end.setUint32(12, centralSize, true); end.setUint32(16, localSize, true);
  return result;
}

describe("document ZIP roundtrips", () => {
  test("opens an archive produced independently by Python zipfile", () => {
    const archive = new Uint8Array(readFileSync(new URL("./fixtures/python-fidelity.zip", import.meta.url)));
    expect(readDocumentZip(archive)).toEqual(files());
    expect(openDocument(readDocumentZip(archive)).snapshot()).toEqual(openDocument(files()).snapshot());
  });

  test("archives complete resources and preserves source bytes, metadata and styles", () => {
    const input = files();
    const before = new Map([...input].map(([path, bytes]) => [path, new Uint8Array(bytes)]));
    input.set("sources/original.bin", new Uint8Array([0, 255, 7]));
    const archive = writeDocumentZip(input);
    const decoded = readDocumentZip(archive);
    expect(decoded).toEqual(input);
    expect(openDocument(decoded).snapshot()).toEqual(openDocument(input).snapshot());
    for (const [path, bytes] of before) expect(input.get(path)).toEqual(bytes);
    input.get("assets/sample.png")!.fill(0);
    archive.fill(0);
    expect(() => openDocument(decoded)).not.toThrow();
    expect(decoded.get("sources/original.bin")).toEqual(new Uint8Array([0, 255, 7]));
  });

  test("writes deterministic archives regardless of Map order", () => {
    const input = files();
    expect(writeDocumentZip(input)).toEqual(writeDocumentZip(new Map([...input].reverse())));
    expect(DEFAULT_MAX_ZIP_BYTES).toBe(64 * 1024 * 1024);
  });

  test("cooperative ZIP output preserves captured bytes and cancels without returning a partial archive", async () => {
    const input = files();
    input.set("sources/large.bin", new Uint8Array(2 * 1024 * 1024).fill(7));
    const before = writeDocumentZip(input);
    const pending = writeDocumentZipAsync(input);
    input.get("sources/large.bin")!.fill(0);
    expect(await pending).toEqual(before);

    const controller = new AbortController();
    const cancelled = writeDocumentZipAsync(input, { signal: controller.signal });
    controller.abort();
    await expect(cancelled).rejects.toMatchObject({ code: "OperationCancelled" });
    expect(await writeDocumentZipAsync(input)).toEqual(writeDocumentZip(input));
  });

  test.each([0, 8])("reads independently compressed entries with method %i and literal Unicode paths", (method) => {
    const data = encoder.encode("中文与 é😀");
    const archive = rawZip([{ name: "正文/", data: new Uint8Array() }, { name: "正文/图 片.txt", data, method }]);
    const decoded = readDocumentZip(archive);
    expect([...decoded.keys()]).toEqual(["正文/图 片.txt"]);
    expect(decoded.get("正文/图 片.txt")).toEqual(data);
  });

  test("preserves filename case, Unicode normalization and a literal leading U+FEFF", () => {
    const names = ["A.txt", "a.txt", "é.txt", "é.txt", "\ufeffname.txt"];
    expect([...readDocumentZip(rawZip(names.map((name) => ({ name })))).keys()]).toEqual(names);
  });

  test("reads deeply nested paths without rebuilding every ancestor", () => {
    const path = "a/".repeat(32_000) + "b";
    const archive = rawZip([{ name: path, data: new Uint8Array() }]);
    expect([...readDocumentZip(archive)]).toEqual([[path, new Uint8Array()]]);
  }, 2000);

  test("preserves sibling path prefixes and explicit parent directories", () => {
    const names = ["a-copy", "a/b", "ab", "a/c", "a/"];
    const archive = rawZip(names.map((name) => ({ name, data: new Uint8Array() })));
    expect([...readDocumentZip(archive).keys()]).toEqual(names.filter((name) => !name.endsWith("/")));
  });

  test("rejects missing declared resources when exporting and never rewrites unknown-version input", () => {
    const input = files();
    input.delete("assets/sample.png");
    expect(() => writeDocumentZip(input)).toThrow(expect.toSatisfy(error("AssetUnavailable")));
    const source = JSON.parse(new TextDecoder().decode(input.get("manifest.json")));
    const changed = encoder.encode(JSON.stringify({ ...source, formatVersion: 99 }));
    input.set("manifest.json", changed);
    expect(() => writeDocumentZip(input)).toThrow(expect.toSatisfy(error("UnsupportedVersion")));
    expect(input.get("manifest.json")).toBe(changed);
  });

  test.each([false, true])("core and ZIP writing reject file ancestors before persistence in either order: %s", (reverse) => {
    const input = files();
    const manifest = JSON.parse(new TextDecoder().decode(input.get("manifest.json")));
    input.set("content.html/metadata.json", input.get("metadata.json")!);
    input.set("manifest.json", encoder.encode(JSON.stringify({ ...manifest, metadata: "content.html/metadata.json" })));
    const ordered = reverse ? new Map([...input].reverse()) : input;
    const before = new Map([...ordered].map(([path, bytes]) => [path, new Uint8Array(bytes)]));
    expect(() => openDocument(ordered)).toThrow(expect.toSatisfy(error("InvalidDocument")));
    expect(() => writeDocumentZip(ordered)).toThrow(expect.toSatisfy(error("InvalidDocument")));
    expect(ordered).toEqual(before);
  });
});

describe("ZIP validation before path normalization or full decompression", () => {
  test.each(["../private text", "./file", "/absolute", "a\\b", "C:/file", "a//b", "a/../b", "a\u0000b", "a?b", "a#b"])
    ("rejects noncanonical source path %j", (name) => {
      expect(() => readDocumentZip(rawZip([{ name }]))).toThrow(expect.toSatisfy(error("InvalidDocument")));
    });

  test("rejects duplicate entries before one can overwrite another", () => {
    const archive = rawZip([{ name: "same" }, { name: "same", data: new Uint8Array() }]);
    expect(() => readDocumentZip(archive)).toThrow(expect.toSatisfy(error("InvalidDocument")));
  });

  test.each([
    { entries: [{ name: "a" }, { name: "a/b" }] },
    { entries: [{ name: "a" }, { name: "a/", data: new Uint8Array() }] },
    { entries: [{ name: "a/", data: new Uint8Array() }, { name: "a" }] },
    { entries: [{ name: "a/b" }, { name: "a-copy" }, { name: "a" }] },
    { entries: [{ name: "a/b/", data: new Uint8Array() }, { name: "a" }] },
    { entries: [{ name: "safe", localName: "../private text" }] },
    { entries: [{ name: "bad", rawName: new Uint8Array([0xff]) }] },
  ])("rejects ambiguous or conflicting entries %j", ({ entries }) => {
    expect(() => readDocumentZip(rawZip(entries))).toThrow(expect.toSatisfy(error("InvalidDocument")));
  });

  test.each([
    { name: "encrypted", flags: 0x801 },
    { name: "unsupported-compression", method: 99 },
    { name: "link", attrs: (0xa1ff << 16) >>> 0 },
    { name: "legacy-encoding", flags: 0, rawName: new Uint8Array([0x82]) },
  ])("reports unsupported container features %j", (entry) => {
    expect(() => readDocumentZip(rawZip([entry]))).toThrow(expect.toSatisfy(error("UnsupportedFeature")));
  });

  test.each([0, 8])("verifies CRC and actual size for method %i", (method) => {
    expect(() => readDocumentZip(rawZip([{ name: "damaged", method, crc: 123 }])))
      .toThrow(expect.toSatisfy(error("InvalidDocument")));
    expect(() => readDocumentZip(rawZip([{ name: "wrong-size", method, size: 999 }])))
      .toThrow(expect.toSatisfy(error("InvalidDocument")));
  });

  test("checks declared expanded sizes before touching invalid compressed data", () => {
    const archive = rawZip([{ name: "large", data: new Uint8Array([255]), method: 8, size: 1_000_000 }]);
    archive[30 + encoder.encode("large").length] = 0xff; // Invalid DEFLATE block type.
    expect(() => readDocumentZip(archive, { limits: { maxFileBytes: 64 } }))
      .toThrow(expect.toSatisfy(error("ResourceLimitExceeded")));
  });

  test("stops bounded inflation when the advertised size lies", () => {
    const archive = rawZip([{ name: "bomb", data: new Uint8Array(4 * 1024 * 1024), method: 8, size: 1 }]);
    expect(() => readDocumentZip(archive, { limits: { maxFileBytes: 128 } }))
      .toThrow(expect.toSatisfy(error("ResourceLimitExceeded")));
  });

  test("checks compressed bytes, entry count including directories, per-file and total expanded bytes", () => {
    const archive = rawZip([{ name: "a", data: new Uint8Array(5) }, { name: "b", data: new Uint8Array(5) }]);
    for (const options of [
      { maxZipBytes: archive.length - 1 }, { limits: { maxFiles: 1 } },
      { limits: { maxFileBytes: 4 } }, { limits: { maxPackageBytes: 9 } },
    ]) expect(() => readDocumentZip(archive, options)).toThrow(expect.toSatisfy(error("ResourceLimitExceeded")));
    expect(readDocumentZip(archive, { maxZipBytes: archive.length, limits: { maxFiles: 2, maxFileBytes: 5, maxPackageBytes: 10 } }).size).toBe(2);
    expect(() => readDocumentZip(rawZip([{ name: "empty/", data: new Uint8Array() }, { name: "a" }]), { limits: { maxFiles: 1 } }))
      .toThrow(expect.toSatisfy(error("ResourceLimitExceeded")));
  });

  test.each([0, -1, 1.5, Infinity])("rejects invalid ZIP byte limit %s", (maxZipBytes) => {
    expect(() => readDocumentZip(new Uint8Array(), { maxZipBytes })).toThrow(expect.toSatisfy(error("InvalidDocument")));
  });

  test("bounds ZIP output and permits its exact byte limit", () => {
    const input = files();
    const archive = writeDocumentZip(input);
    expect(() => writeDocumentZip(input, { maxZipBytes: archive.length - 1 })).toThrow(expect.toSatisfy(error("ResourceLimitExceeded")));
    expect(writeDocumentZip(input, { maxZipBytes: archive.length })).toEqual(archive);
  });

  test("returns bounded errors for random, truncated and inconsistent ZIP bytes", () => {
    const archive = rawZip([{ name: "private text" }]);
    for (const bytes of [new Uint8Array(), new Uint8Array([1, 2, 3]), archive.subarray(0, archive.length - 1), archive.subarray(1)]) {
      expect(() => readDocumentZip(bytes)).toThrow(expect.toSatisfy(error("InvalidDocument")));
    }
    const overlap = archive.slice();
    new DataView(overlap.buffer).setUint32(archive.length - 6, 0, true);
    expect(() => readDocumentZip(overlap)).toThrow(expect.toSatisfy(error("InvalidDocument")));
  });

  test("reports classic-ZIP boundaries for ZIP64 and multi-volume metadata", () => {
    const archive = rawZip([{ name: "a" }]);
    const zip64 = archive.slice();
    new DataView(zip64.buffer).setUint32(zip64.length - 10, 0xffffffff, true);
    expect(() => readDocumentZip(zip64)).toThrow(expect.toSatisfy(error("UnsupportedFeature")));
    const multiVolume = archive.slice();
    new DataView(multiVolume.buffer).setUint16(multiVolume.length - 18, 1, true);
    expect(() => readDocumentZip(multiVolume)).toThrow(expect.toSatisfy(error("UnsupportedFeature")));
  });
});
