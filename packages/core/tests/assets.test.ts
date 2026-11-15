import { createHash } from "node:crypto";
import { readFileSync } from "node:fs";
import { describe, expect, test } from "vitest";
import { collectDocumentAssets, createAssetDescriptor, DocumentError, openDocument, parseContract, createDocumentSession } from "@doffi/core";
import type { AssetDescriptor, AssetProvider, DocumentErrorCode, DocumentPackage } from "@doffi/core";

const encoder = new TextEncoder();
function files(): Map<string, Uint8Array> {
  return new Map(["manifest.json", "content.html", "styles.css", "metadata.json", "annotations.json",
    "assets/sample.png", "assets/sample.jpg", "assets/sample.webp"].map((path) =>
    [path, new Uint8Array(readFileSync(new URL("../../../examples/fidelity/" + path, import.meta.url)))]));
}
function manifest(input: DocumentPackage) {
  return parseContract("document-manifest", JSON.parse(new TextDecoder().decode(input.get("manifest.json"))));
}
function changeAsset(input: Map<string, Uint8Array>, changed: Partial<AssetDescriptor>): void {
  const source = manifest(input);
  input.set("manifest.json", encoder.encode(JSON.stringify({ ...source, assets: [{ ...source.assets[0], ...changed }, ...source.assets.slice(1)] })));
}
function error(code: DocumentErrorCode) {
  return (value: unknown) => {
    expect(value).toBeInstanceOf(DocumentError);
    expect((value as DocumentError).code).toBe(code);
    expect(parseContract("document-error", (value as DocumentError).toJSON()).code).toBe(code);
    return true;
  };
}

describe("resource integrity and preservation", () => {
  test("describes bytes using standard SHA-256 without granting storage access", () => {
    const bytes = encoder.encode("abc");
    const asset = createAssetDescriptor("assets/opaque.bin", "application/octet-stream", bytes);
    expect(asset.sha256).toBe("ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad");
    expect(Object.isFrozen(asset)).toBe(true);
  });

  test("preserves supported styles, node identities and PNG/JPEG/WebP bytes across reopen", () => {
    const input = files();
    const document = openDocument(input);
    const saved = document.serialize();
    const reopened = openDocument(saved);
    expect(reopened.snapshot()).toEqual(document.snapshot());
    for (const asset of manifest(input).assets) {
      expect(createHash("sha256").update(input.get(asset.path)!).digest("hex")).toBe(asset.sha256);
      expect(saved.get(asset.path)).toEqual(input.get(asset.path));
      expect(saved.get(asset.path)).not.toBe(input.get(asset.path));
    }
    expect(reopened.getNode("paragraph")!.attributes.style).toContain("text-indent:24pt");
    expect(reopened.getNode("table")!.html).toContain('rowspan="2"');
    expect(reopened.getNode("list")!.attributes.start).toBe("3");
    expect(reopened.getNode("unicode")!.text).toBe("é😀");
    expect(reopened.snapshot().styles[0]!.css).toContain('"Noto Serif CJK SC"');
  });

  test.each(["image/jpeg", "image/webp"])("rejects declared MIME %s on PNG bytes", (mimeType) => {
    const input = files();
    changeAsset(input, { mimeType });
    expect(() => openDocument(input)).toThrow(expect.toSatisfy(error("InvalidDocument")));
  });

  test("rejects altered bytes or digest, even on an unreferenced declared resource", () => {
    const input = files();
    const image = input.get("assets/sample.png")!;
    image[30] = image[30]! ^ 1;
    expect(() => openDocument(input)).toThrow(expect.toSatisfy(error("InvalidDocument")));
    const unused = createAssetDescriptor("assets/unused.bin", "application/octet-stream", new Uint8Array([1]));
    const unreferenced = files();
    const source = manifest(unreferenced);
    unreferenced.set("manifest.json", encoder.encode(JSON.stringify({ ...source, assets: [...source.assets, unused] })));
    unreferenced.set(unused.path, new Uint8Array([2]));
    expect(() => openDocument(unreferenced)).toThrow(expect.toSatisfy(error("InvalidDocument")));
    unreferenced.delete(unused.path);
    expect(() => openDocument(unreferenced)).toThrow(expect.toSatisfy(error("AssetUnavailable")));
  });

  test.each(["image/png", "image/jpeg", "image/webp"])("rejects truncated containers declared as %s", (mimeType) => {
    expect(() => createAssetDescriptor("assets/truncated", mimeType, new Uint8Array([137, 80, 78, 71])))
      .toThrow(expect.toSatisfy(error("InvalidDocument")));
  });

  test("retains opaque attachments as data and rejects using them as images", () => {
    const input = files();
    const bytes = encoder.encode('<svg onload="private text"></svg>');
    const asset = createAssetDescriptor("assets/opaque.svg", "image/svg+xml", bytes);
    const source = manifest(input);
    input.set(asset.path, bytes);
    input.set("manifest.json", encoder.encode(JSON.stringify({ ...source, assets: [...source.assets, asset] })));
    expect(openDocument(input).serialize().get(asset.path)).toEqual(bytes);
    const html = new TextDecoder().decode(input.get("content.html")).replace("assets/sample.png", asset.path);
    input.set("content.html", encoder.encode(html));
    expect(() => openDocument(input)).toThrow(expect.toSatisfy(error("UnsupportedFeature")));
  });
});

describe("host-authorized resource collection", () => {
  test("fetches only missing declarations, freezes descriptors and copies both input and returned bytes", async () => {
    const input = files();
    const image = input.get("assets/sample.png")!;
    input.delete("assets/sample.png");
    const before = new Map(input);
    const called: string[] = [];
    const collected = await collectDocumentAssets(input, { async readAsset(asset) {
      expect(Object.isFrozen(asset)).toBe(true);
      called.push(asset.path);
      return image;
    } });
    expect(called).toEqual(["assets/sample.png"]);
    expect(input).toEqual(before);
    expect(openDocument(collected).getNode("png")).toBeDefined();
    image.fill(0);
    input.get("styles.css")!.fill(0);
    expect(() => openDocument(collected)).not.toThrow();
  });

  test.each([undefined, "throw", new Uint8Array([1, 2, 3])])("reports unavailable or invalid host bytes: %s", async (value) => {
    const input = files();
    input.delete("assets/sample.png");
    const promise = collectDocumentAssets(input, { async readAsset() {
      if (value === "throw") throw new Error("private host credentials");
      return value instanceof Uint8Array ? value : undefined;
    } });
    await expect(promise).rejects.toSatisfy(error(value instanceof Uint8Array ? "InvalidDocument" : "AssetUnavailable"));
    await promise.catch((value: DocumentError) => expect(JSON.stringify(value.toJSON())).not.toContain("private host credentials"));
    expect(input.has("assets/sample.png")).toBe(false);
  });

  test("checks file counts before requesting the host and byte budgets before retaining responses", async () => {
    const input = files();
    const original = input.get("assets/sample.png")!;
    input.delete("assets/sample.png");
    let calls = 0;
    const provider: AssetProvider = { async readAsset() { calls++; return original; } };
    await expect(collectDocumentAssets(input, provider, { limits: { maxFiles: input.size } }))
      .rejects.toSatisfy(error("ResourceLimitExceeded"));
    expect(calls).toBe(0);
    const size = [...input.values()].reduce((sum, bytes) => sum + bytes.length, 0);
    await expect(collectDocumentAssets(input, provider, { limits: { maxPackageBytes: size } }))
      .rejects.toSatisfy(error("ResourceLimitExceeded"));
    expect(calls).toBe(1);
    await expect(collectDocumentAssets(input, { async readAsset() { return new Uint8Array(9 * 1024 * 1024); } }))
      .rejects.toSatisfy(error("ResourceLimitExceeded"));
  });

  test("keeps two providers with identical package paths isolated", async () => {
    const input = files();
    const image = input.get("assets/sample.png")!;
    input.delete("assets/sample.png");
    const first = collectDocumentAssets(input, { async readAsset() { return image; } });
    const second = collectDocumentAssets(input, { async readAsset() { return undefined; } });
    await expect(second).rejects.toSatisfy(error("AssetUnavailable"));
    expect(openDocument(await first).getNode("png")).toBeDefined();
  });
});


test("content models, candidates and sessions copy only declared assets within the requested byte budget",()=>{
 const document=openDocument(files(),{kind:'document'}),session=createDocumentSession(document),candidate=session.prepare({protocolVersion:1,documentId:session.snapshot().manifest.documentId,baseRevision:session.snapshot().manifest.revision,operations:[{op:'setStyle',nodeId:'paragraph',property:'color',value:'#123456'}]});
 for(const model of [document,session,candidate.model]){const bytes=model.getAsset!('assets/sample.png')!,original=bytes[0];bytes[0]=0;expect(model.getAsset!('assets/sample.png')![0]).toBe(original);expect(model.getAsset!('content.html')).toBeUndefined();expect(model.getAsset!('missing.png')).toBeUndefined();expect(()=>model.getAsset!('assets/sample.png',1)).toThrowError(expect.objectContaining({code:'ResourceLimitExceeded'}));expect(()=>model.getAsset!('assets/sample.png',0)).toThrowError(expect.objectContaining({code:'InvalidDocument'}));}
 session.discard(candidate);
});
