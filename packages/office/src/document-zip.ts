import { DEFAULT_DOCUMENT_READ_LIMITS, DocumentError, openDocument } from "@doffi/core";
import type { DocumentPackage, DocumentReadLimits, OpenDocumentOptions } from "@doffi/core";
import { Inflate, Zip, ZipDeflate } from "fflate/browser";
import { inspectZip, invalidZip, updateCrc32, zipLimit } from "./zip-format.js";
import type { ZipEntry } from "./zip-format.js";
import { ImportControl } from "./import-control.js";

export interface DocumentZipOptions extends OpenDocumentOptions {
  readonly maxZipBytes?: number;
}

export const DEFAULT_MAX_ZIP_BYTES: number = 64 * 1024 * 1024;

function zipOptions(options: DocumentZipOptions): { limits: DocumentReadLimits; maxZipBytes: number } {
  const limits = { ...DEFAULT_DOCUMENT_READ_LIMITS, ...options.limits };
  for (const [name, value] of Object.entries(options.limits ?? {})) {
    if (!Object.hasOwn(DEFAULT_DOCUMENT_READ_LIMITS, name) || !Number.isSafeInteger(value) || value <= 0) {
      throw new DocumentError("InvalidDocument", "Document limits must be positive safe integers with supported names.");
    }
  }
  const maxZipBytes = options.maxZipBytes ?? DEFAULT_MAX_ZIP_BYTES;
  if (!Number.isSafeInteger(maxZipBytes) || maxZipBytes <= 0) throw new DocumentError("InvalidDocument", "The ZIP byte limit must be a positive safe integer.");
  return { limits, maxZipBytes };
}

function* extractSteps(bytes: Uint8Array, entry: ZipEntry, limits: DocumentReadLimits): Generator<number, Uint8Array> {
  const compressed = bytes.subarray(entry.dataOffset, entry.dataOffset + entry.compressedSize);
  let result: Uint8Array, crc = 0xffffffff;
  if (entry.method === 0) {
    if (entry.compressedSize !== entry.size) invalidZip();
    result = new Uint8Array(entry.size);
    for (let cursor = 0; cursor < compressed.length; cursor += 65536) {
      const chunk = compressed.subarray(cursor, cursor + 65536);
      result.set(chunk, cursor); crc = updateCrc32(chunk, crc);
      yield Math.min(cursor + 65536, compressed.length);
    }
  } else {
    result = new Uint8Array(entry.size);
    let written = 0;
    const inflater = new Inflate((chunk) => {
      written += chunk.length;
      zipLimit(written, limits.maxFileBytes);
      if (written > entry.size) invalidZip();
      result.set(chunk, written - chunk.length);
      crc = updateCrc32(chunk, crc);
    });
    // Feed bounded compressed chunks; forged sizes cannot trigger an unbounded full inflate.
    if (!compressed.length) invalidZip();
    for (let cursor = 0; cursor < compressed.length; cursor += 128) {
      inflater.push(compressed.subarray(cursor, cursor + 128), cursor + 128 >= compressed.length);
      yield written;
    }
    if (written !== entry.size) invalidZip();
  }
  if (((crc ^ 0xffffffff) >>> 0) !== entry.crc) invalidZip();
  return result;
}
function extract(bytes: Uint8Array, entry: ZipEntry, limits: DocumentReadLimits): Uint8Array {
  const steps = extractSteps(bytes, entry, limits);
  while (true) { const step = steps.next(); if (step.done) return step.value; }
}

/** Internal asynchronous variant with identical ZIP validation and bounded inflate. */
export async function readDocumentZipAsync(input: Uint8Array, options: DocumentZipOptions, control: ImportControl): Promise<DocumentPackage> {
  const { limits, maxZipBytes } = zipOptions(options);
  control.check();
  if (!(input instanceof Uint8Array)) invalidZip();
  zipLimit(input.byteLength, maxZipBytes);
  const bytes = new Uint8Array(input), files = new Map<string, Uint8Array>();
  let entries: ZipEntry[];
  try { entries = inspectZip(bytes, limits); }
  catch (error) { if (error instanceof DocumentError) throw error; return invalidZip(); }
  const total = entries.reduce((sum, entry) => sum + entry.size, 0);
  let completed = 0;
  await control.checkpoint({ phase: "zip", completed, total });
  try {
    for (const entry of entries) {
      const steps = extractSteps(bytes, entry, limits); let lastYield = 0, stepsSinceYield = 0;
      while (true) {
        const step = steps.next();
        if (step.done) { if (!entry.directory) files.set(entry.path, step.value); break; }
        control.check();
        if (step.value - lastYield >= 1024 * 1024 || ++stepsSinceYield >= 512) {
          await control.checkpoint({ phase: "zip", completed: completed + step.value, total });
          lastYield = step.value; stepsSinceYield = 0;
        }
      }
      completed += entry.size;
      await control.checkpoint({ phase: "zip", completed, total });
    }
    return files;
  } catch (error) {
    if (error instanceof DocumentError) throw error;
    return invalidZip();
  }
}

/** Read a bounded ZIP into independent package files. Use core.openDocument to validate the document. */
export function readDocumentZip(input: Uint8Array, options: DocumentZipOptions = {}): DocumentPackage {
  const { limits, maxZipBytes } = zipOptions(options);
  if (!(input instanceof Uint8Array)) invalidZip();
  zipLimit(input.byteLength, maxZipBytes);
  try {
    const entries = inspectZip(input, limits);
    const files = new Map<string, Uint8Array>();
    for (const entry of entries) {
      const bytes = extract(input, entry, limits);
      if (!entry.directory) files.set(entry.path, bytes);
    }
    return files;
  } catch (error) {
    if (error instanceof DocumentError) throw error;
    return invalidZip();
  }
}

/** Write a complete validated document package, retaining source files and resource bytes. */
export function writeDocumentZip(input: DocumentPackage, options: DocumentZipOptions = {}): Uint8Array {
  return writeZip(input, options, true);
}
/** Write the same validated package with cooperative cancellation during compression. */
export async function writeDocumentZipAsync(input: DocumentPackage, options: DocumentZipOptions & { readonly signal?: { readonly aborted: boolean } } = {}): Promise<Uint8Array> {
  const { signal, ...zip } = options;
  const scheduler = globalThis as unknown as { setTimeout(callback: () => void, delay: number): unknown };
  const check = (): void => {
    if (signal?.aborted) throw new DocumentError("OperationCancelled", "Document package export was cancelled.");
  };
  const checkpoint = async (): Promise<void> => {
    await new Promise<void>(resolve => scheduler.setTimeout(resolve, 0));
    check();
  };
  check();
  const bytes = await compressZip(writeZipSteps(input, zip, true), checkpoint);
  await checkpoint();
  return bytes;
}
/** Internal OPC output primitive; shares all raw ZIP/path and byte limits. */
export function writeOpcZip(input: DocumentPackage, options: DocumentZipOptions = {}): Uint8Array {
  return writeZip(input, options, false);
}
export function writeOpcZipAsync(input: DocumentPackage, options: DocumentZipOptions, checkpoint: () => Promise<void>): Promise<Uint8Array> {
  return compressZip(writeZipSteps(input, options, false), checkpoint);
}
async function compressZip(steps: Generator<void, Uint8Array>, checkpoint: () => Promise<void>): Promise<Uint8Array> {
  try {
    for (let next = steps.next(); ; next = steps.next()) {
      if (next.done) return next.value;
      await checkpoint();
    }
  } finally { steps.return(new Uint8Array(0)); }
}
function writeZip(input: DocumentPackage, options: DocumentZipOptions, document: boolean): Uint8Array {
  const steps = writeZipSteps(input, options, document);
  for (let next = steps.next(); ; next = steps.next()) if (next.done) return next.value;
}
function* writeZipSteps(input: DocumentPackage, options: DocumentZipOptions, document: boolean): Generator<void, Uint8Array> {
  const { limits, maxZipBytes } = zipOptions(options);
  const files = new Map<string, Uint8Array>();
  let inputSize = 0;
  for (const [path, bytes] of input) {
    if (files.has(path) || !(bytes instanceof Uint8Array)) invalidZip();
    zipLimit(files.size + 1, limits.maxFiles);
    zipLimit(bytes.length, limits.maxFileBytes);
    inputSize += bytes.length;
    zipLimit(inputSize, limits.maxPackageBytes);
    files.set(path, new Uint8Array(bytes));
  }
  // Validate stable copied bytes, including every declared resource. Normalization
  // remains an explicit document-model serialization step rather than an archive side effect.
  if (document) openDocument(files, { limits, ...(options.workbookLimits ? { workbookLimits: options.workbookLimits } : {}), ...(options.kind ? { kind: options.kind } : {}) });
  const chunks: Uint8Array[] = [];
  let length = 0;
  try {
    const archive = new Zip((error, chunk) => {
      if (error) throw error;
      length += chunk.length;
      zipLimit(length, maxZipBytes);
      chunks.push(chunk);
    });
    let blocks = 0;
    for (const [path, bytes] of [...files].sort(([left], [right]) => left < right ? -1 : left > right ? 1 : 0)) {
      const file = new ZipDeflate(path, { level: 6 });
      file.mtime = new Date(1980, 0, 1);
      archive.add(file);
      if (!bytes.length) file.push(bytes, true);
      for (let cursor = 0; cursor < bytes.length; cursor += 65536) {
        file.push(bytes.subarray(cursor, cursor + 65536), cursor + 65536 >= bytes.length);
        if (++blocks % 16 === 0) yield;
      }
    }
    archive.end();
    const result = new Uint8Array(length);
    let offset = 0;
    for (const chunk of chunks) { result.set(chunk, offset); offset += chunk.length; }
    // Enforce the same raw-path/container rules on output, including path collisions.
    inspectZip(result, limits);
    return result;
  } catch (error) {
    if (error instanceof DocumentError) throw error;
    throw new DocumentError("ExportFailed", "The document package could not be written as ZIP.");
  }
}
