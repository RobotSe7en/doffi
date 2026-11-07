import { PackagePathSchema } from "../contracts/primitives.js";
import type { DocumentPackage, DocumentReadLimits } from "../runtime.js";
import { checkLimit, invalid } from "./validation.js";

// Standard UTF-8 codecs exist in supported browsers and Node. Keep their platform
// declarations private so public core types need neither DOM nor Node ambient types.
const codecs = globalThis as unknown as {
  TextDecoder: new (label: string, options: { fatal: boolean; ignoreBOM?: boolean }) => { decode(bytes: Uint8Array): string };
  TextEncoder: new () => { encode(text: string): Uint8Array };
};

export function encode(text: string): Uint8Array {
  return new codecs.TextEncoder().encode(text);
}

/** Detach already valid document text. Each owner gets private, non-streaming
 * codecs; copying many small cells need not create codecs for every cell. */
export function createUtf8TextCopy(): (text: string) => string {
  const encoder = new codecs.TextEncoder(), decoder = new codecs.TextDecoder("utf-8", { fatal: true, ignoreBOM: true });
  return text => decoder.decode(encoder.encode(text));
}

export function copyPackage(input: DocumentPackage, limits: DocumentReadLimits): Map<string, Uint8Array> {
  const files = validatePackage(input, limits);
  return new Map([...files].map(([path, bytes]) => [path, new Uint8Array(bytes)]));
}

/** Validate already-owned immutable source without copying complete workbooks. */
export function validatePackage(input: DocumentPackage, limits: DocumentReadLimits): Map<string, Uint8Array> {
  const files = new Map<string, Uint8Array>();
  let total = 0;
  for (const [path, bytes] of input) {
    if (!PackagePathSchema.safeParse(path).success || files.has(path) || !(bytes instanceof Uint8Array)) {
      invalid("Package files must have unique canonical paths and byte contents.");
    }
    checkLimit(files.size + 1, limits.maxFiles);
    checkLimit(bytes.byteLength, limits.maxFileBytes);
    total += bytes.byteLength;
    checkLimit(total, limits.maxPackageBytes);
    files.set(path, bytes);
  }
  // A trailing slash groups a file with its descendants even when sibling names
  // contain punctuation that sorts before '/'. No ancestor strings are rebuilt.
  const paths = [...files.keys()].map((path) => path + "/").sort();
  for (let index = 1; index < paths.length; index++) {
    if (paths[index]!.startsWith(paths[index - 1]!)) {
      invalid("Package file paths must not overlap with ancestor files.");
    }
  }
  return files;
}

export function readText(files: DocumentPackage, path: string): string {
  const bytes = files.get(path);
  if (!bytes) invalid("A referenced document file is missing.");
  return decodeText(bytes);
}

export function decodeText(bytes: Uint8Array): string {
  try {
    return new codecs.TextDecoder("utf-8", { fatal: true }).decode(bytes);
  } catch {
    return invalid("Document text files must contain valid UTF-8.");
  }
}

export function readJson(files: DocumentPackage, path: string): unknown {
  const text = readText(files, path);
  try {
    return JSON.parse(text) as unknown;
  } catch {
    return invalid("A document JSON file is malformed.");
  }
}

/** Resolve one HTML URL decoding step against a literal package path, never a host URL. */
export function resolveReference(reference: string, from: string): string {
  if (!reference || reference !== reference.trim() || /^[\/]/.test(reference)
    || /[\\:\u0000-\u001f\u007f?#]/.test(reference)) {
    invalid("Resource references must resolve inside the supplied package.");
  }
  const parts = from.split("/").slice(0, -1);
  for (const encoded of reference.split("/")) {
    let part: string;
    try {
      part = decodeURIComponent(encoded);
    } catch {
      return invalid("A resource reference has invalid URL encoding.");
    }
    if (part === ".") continue;
    if (part === "..") {
      if (!parts.length) invalid("A resource reference escapes the package.");
      parts.pop();
    } else {
      if (!part || /[/\\:\u0000-\u001f\u007f?#]/.test(part)) {
        invalid("A resource reference contains an invalid path segment.");
      }
      parts.push(part);
    }
  }
  const path = parts.join("/");
  if (!PackagePathSchema.safeParse(path).success) invalid("A resource reference is not a package file.");
  return path;
}
