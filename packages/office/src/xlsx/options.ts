import { DEFAULT_DOCUMENT_READ_LIMITS, DEFAULT_WORKBOOK_READ_LIMITS, DocumentError } from "@doffi/core";
import type { DocumentReadLimits, WorkbookReadLimits } from "@doffi/core";
import type { OfficeImportControl } from "../import-control.js";

/** Native archive/XML budgets are separate from the generated HTML package.
 * These protective caps are not statements of tested workbook capacity. */
export interface XlsxNativeLimits {
  readonly maxZipBytes: number;
  readonly maxFiles: number;
  readonly maxPackageBytes: number;
  readonly maxPartBytes: number;
  readonly maxMetadataBytes: number;
  readonly maxXmlNodes: number;
  readonly maxXmlDepth: number;
  readonly maxRetainedNodes: number;
  readonly maxSharedStrings: number;
  readonly maxSharedStringCodeUnits: number;
  readonly maxCssCodeUnits: number;
}
export const DEFAULT_XLSX_NATIVE_LIMITS: XlsxNativeLimits = Object.freeze({
  maxZipBytes: 128 * 1024 * 1024, maxFiles: 2048, maxPackageBytes: 512 * 1024 * 1024, maxPartBytes: 256 * 1024 * 1024,
  maxMetadataBytes: 8 * 1024 * 1024, maxXmlNodes: 2_000_000, maxXmlDepth: 128, maxRetainedNodes: 100_000,
  maxSharedStrings: 250_000, maxSharedStringCodeUnits: 16 * 1024 * 1024, maxCssCodeUnits: 4 * 1024 * 1024,
});
export interface XlsxReadOptions extends OfficeImportControl {
  readonly limits?: Partial<DocumentReadLimits>;
  readonly workbookLimits?: Partial<WorkbookReadLimits>;
  readonly nativeLimits?: Partial<XlsxNativeLimits>;
}
function limits<T extends object>(defaults: T, input: Partial<T> | undefined): T {
  if (input !== undefined && (!input || typeof input !== "object" || Array.isArray(input))) throw new DocumentError("InvalidOperation", "Invalid XLSX limit options.");
  for (const [name, value] of Object.entries(input ?? {})) if (!Object.hasOwn(defaults, name) || !Number.isSafeInteger(value) || (value as number) <= 0)
    throw new DocumentError("InvalidOperation", "XLSX limits must be positive safe integers with supported names.");
  return Object.freeze({ ...defaults, ...input });
}
export function xlsxLimits(options: Pick<XlsxReadOptions, "limits" | "workbookLimits" | "nativeLimits">): { source: DocumentReadLimits; workbook: WorkbookReadLimits; native: XlsxNativeLimits; xml: DocumentReadLimits } {
  const source = limits(DEFAULT_DOCUMENT_READ_LIMITS, options.limits), workbook = limits(DEFAULT_WORKBOOK_READ_LIMITS, options.workbookLimits), native = limits(DEFAULT_XLSX_NATIVE_LIMITS, options.nativeLimits);
  return { source, workbook, native, xml: { ...source, maxFiles: native.maxFiles, maxPackageBytes: native.maxPackageBytes, maxFileBytes: native.maxPartBytes, maxNodes: native.maxXmlNodes, maxDepth: native.maxXmlDepth } };
}
