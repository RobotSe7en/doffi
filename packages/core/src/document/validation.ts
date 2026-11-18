import { DocumentError } from "../errors.js";
import type { DocumentReadLimits } from "../runtime.js";

export function invalid(message: string): never {
  throw new DocumentError("InvalidDocument", message);
}

export function unsupported(message: string): never {
  throw new DocumentError("UnsupportedFeature", message);
}

export function checkLimit(value: number, maximum: number, limit?: keyof DocumentReadLimits, context?: string): void {
  if (value > maximum) {
    throw new DocumentError("ResourceLimitExceeded", limit
      ? `Document input exceeds ${limit} (${value} > ${maximum})${context ? ` in ${context}` : ""}.`
      : "Document input exceeds the configured limit.");
  }
}

export const DEFAULT_DOCUMENT_READ_LIMITS: DocumentReadLimits = Object.freeze({
  maxFiles: 1024,
  maxPackageBytes: 32 * 1024 * 1024,
  maxFileBytes: 8 * 1024 * 1024,
  maxNodes: 100_000,
  maxDepth: 128,
  maxCssNodes: 100_000,
});

export function readLimits(input: Partial<DocumentReadLimits> = {}): DocumentReadLimits {
  for (const [key, value] of Object.entries(input)) {
    if (!Object.hasOwn(DEFAULT_DOCUMENT_READ_LIMITS, key)
      || !Number.isSafeInteger(value) || value <= 0) {
      invalid("Document limits must be positive safe integers with supported names.");
    }
  }
  return Object.freeze({ ...DEFAULT_DOCUMENT_READ_LIMITS, ...input });
}

export function freeze<Value>(value: Value): Value {
  if (value !== null && typeof value === "object") {
    for (const child of Object.values(value)) freeze(child);
    Object.freeze(value);
  }
  return value;
}
