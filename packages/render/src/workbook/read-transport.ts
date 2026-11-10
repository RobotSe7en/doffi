import { DocumentError } from "@doffi/core";
import type { WorkbookSnapshot } from "@doffi/core";

export const WORKBOOK_READ_PROTOCOL: string = "doffi/workbook-read/1";
export interface WorkbookReadTransportLimits {
  readonly maxCells: number;
  readonly maxCodeUnits: number;
  readonly maxAssetBytes: number;
  readonly maxCharts: number;
  readonly maxReferences: number;
  readonly maxPoints: number;
  readonly maxChartFormulaCells: number;
  readonly maxChartFormulaSteps: number;
  readonly maxRules: number;
  readonly maxRuleValues: number;
  readonly maxRuleSteps: number;
}
export interface WorkbookReadTransportOptions {
  readonly maxPendingRequests?: number;
  readonly requestTimeoutMs?: number;
  readonly maxSnapshotNodes?: number;
  readonly maxSnapshotCodeUnits?: number;
  readonly readLimits?: Partial<WorkbookReadTransportLimits>;
}
export interface WorkbookReadTransportConfig {
  readonly maxPendingRequests: number;
  readonly requestTimeoutMs: number;
  readonly maxSnapshotNodes: number;
  readonly maxSnapshotCodeUnits: number;
  readonly readLimits: WorkbookReadTransportLimits;
}
const readDefaults: WorkbookReadTransportLimits = Object.freeze({ maxCells: 25_000, maxCodeUnits: 8 * 1024 * 1024, maxAssetBytes: 8 * 1024 * 1024, maxCharts: 128, maxReferences: 5000, maxPoints: 50_000,
  maxRules: 256, maxRuleValues: 250_000, maxRuleSteps: 5_000_000, maxChartFormulaCells: 250_000, maxChartFormulaSteps: 5_000_000 });
export function transportConfig(input: WorkbookReadTransportOptions): WorkbookReadTransportConfig {
  const { readLimits, ...own } = input;
  const config = { maxPendingRequests: 8, requestTimeoutMs: 60_000, maxSnapshotNodes: 500_000, maxSnapshotCodeUnits: 32 * 1024 * 1024, ...own };
  const reads = { ...readDefaults, ...readLimits };
  if (Object.keys(own).some(key => !["maxPendingRequests", "requestTimeoutMs", "maxSnapshotNodes", "maxSnapshotCodeUnits"].includes(key))
    || Object.keys(reads).some(key => !Object.hasOwn(readDefaults, key))
    || ![...Object.values(config), ...Object.values(reads)].every(value => Number.isSafeInteger(value) && value > 0)
    || config.requestTimeoutMs > 2_147_483_647)
    throw new DocumentError("InvalidOperation", "Workbook transport limits must be supported positive safe integers.");
  return Object.freeze({ ...config, readLimits: Object.freeze(reads) });
}
export function isRecord(value: unknown): value is Record<string, unknown> { return typeof value === "object" && value !== null && !Array.isArray(value); }
export function requestId(value: unknown): value is number { return typeof value === "number" && Number.isSafeInteger(value) && value > 0; }
export function keys(value: Record<string, unknown>, allowed: readonly string[]): boolean { return Object.keys(value).every(key => allowed.includes(key)); }
export function transportError(error: unknown): DocumentError {
  return error instanceof DocumentError ? error : new DocumentError("InvalidOperation", "The authorized workbook transport could not complete the read.");
}

/** Both sides budget the metadata graph without stringifying it or cloning the
 * full workbook. Source models are already validated by core; this check does
 * not claim to validate a workbook supplied by an untrusted message sender. */
export function transportSnapshot(value: unknown, config: WorkbookReadTransportConfig, freeze = false): WorkbookSnapshot {
  if (!isRecord(value) || !isRecord(value.manifest) || value.manifest.kind !== "workbook" || typeof value.manifest.documentId !== "string"
    || !Number.isSafeInteger(value.manifest.revision) || (value.manifest.revision as number) < 0 || !isRecord(value.metadata)
    || !isRecord(value.metadata.workbook) || !Array.isArray(value.metadata.workbook.sheets) || !Array.isArray(value.styles) || !Array.isArray(value.annotations))
    throw new DocumentError("InvalidDocument", "Workbook transport returned an invalid snapshot envelope.");
  const stack: { value: unknown; leave?: boolean }[] = [{ value }], visiting = new Set<object>(), done = new Set<object>(); let nodes = 1, units = 0;
  while (stack.length) {
    const entry = stack.pop()!, item = entry.value;
    if (typeof item === "string") units += item.length;
    else if (typeof item === "object" && item !== null) {
      if (entry.leave) { visiting.delete(item); done.add(item); if (freeze) Object.freeze(item); continue; }
      if (visiting.has(item)) throw new DocumentError("InvalidDocument", "Workbook snapshot data must not contain cycles.");
      if (done.has(item)) continue;
      if (!Array.isArray(item) && Object.getPrototypeOf(item) !== Object.prototype && Object.getPrototypeOf(item) !== null)
        throw new DocumentError("InvalidDocument", "Workbook snapshot data must contain only plain source values.");
      const names = Object.keys(item); nodes += names.length;
      if (nodes > config.maxSnapshotNodes) throw new DocumentError("ResourceLimitExceeded", "Workbook snapshot exceeds its transport node budget.");
      visiting.add(item); stack.push({ value: item, leave: true });
      for (const key of names) { units += key.length; stack.push({ value: (item as Record<string, unknown>)[key] }); }
    } else if (item !== null && typeof item !== "boolean" && (typeof item !== "number" || !Number.isFinite(item)))
      throw new DocumentError("InvalidDocument", "Workbook snapshot data must contain finite JSON values.");
    if (units > config.maxSnapshotCodeUnits) throw new DocumentError("ResourceLimitExceeded", "Workbook snapshot exceeds its transport text budget.");
  }
  return value as unknown as WorkbookSnapshot;
}
export type WorkbookReadOperation = "cells" | "asset" | "charts" | "rules";
/** Reject oversized client budgets before invoking any model adapter. The
 * adapter then validates exact coordinates/references against its own model. */
export function transportReadRequest(operation: WorkbookReadOperation, value: unknown, config: WorkbookReadTransportConfig): Record<string, unknown> {
  if (!isRecord(value) || typeof value.documentId !== "string" || !Number.isSafeInteger(value.revision) || (value.revision as number) < 0)
    throw new DocumentError("InvalidOperation", "Workbook transport requests require a source identity.");
  const limit = (name: string, maximum: number): void => {
    const number = value[name];
    if (typeof number !== "number" || !Number.isSafeInteger(number) || number <= 0) throw new DocumentError("InvalidOperation", "Workbook transport read budgets must be positive safe integers.");
    if (number > maximum) throw new DocumentError("ResourceLimitExceeded", "Workbook transport read exceeds its server budget.");
  };
  const limits = config.readLimits;
  if (operation === "cells") {
    if (!keys(value, ["documentId", "revision", "windows", "maxCells", "maxCodeUnits"]) || !Array.isArray(value.windows)) throw new DocumentError("InvalidOperation", "Invalid workbook window transport request.");
    limit("maxCells", limits.maxCells); limit("maxCodeUnits", limits.maxCodeUnits);
    if (value.windows.length > limits.maxCells) throw new DocumentError("ResourceLimitExceeded", "Too many transported workbook windows.");
    for (const window of value.windows) if (!isRecord(window) || !keys(window, ["sheetId", "row", "column", "rows", "columns"])) throw new DocumentError("InvalidOperation", "Invalid transported workbook window.");
  } else if (operation === "asset") {
    if (!keys(value, ["documentId", "revision", "path", "maxBytes"]) || typeof value.path !== "string" || value.path.length > 1024) throw new DocumentError("InvalidOperation", "Invalid workbook resource transport request.");
    limit("maxBytes", limits.maxAssetBytes);
  } else if (operation === "rules") {
    if (!keys(value, ["documentId", "revision", "sheetId", "coordinates", "includeCurrent", "maxCells", "maxRules", "maxValues", "maxSteps", "maxCodeUnits"])
      || value.includeCurrent !== undefined && typeof value.includeCurrent !== "boolean" || typeof value.sheetId !== "string" || value.sheetId.length > 1024 || !Array.isArray(value.coordinates)) throw new DocumentError("InvalidOperation", "Invalid workbook rule transport request.");
    limit("maxCells", limits.maxCells); limit("maxRules", limits.maxRules); limit("maxValues", limits.maxRuleValues); limit("maxSteps", limits.maxRuleSteps); limit("maxCodeUnits", limits.maxCodeUnits);
    if (value.coordinates.length > (value.maxCells as number)) throw new DocumentError("ResourceLimitExceeded", "Too many transported workbook rule coordinates.");
    for (const coordinate of value.coordinates) if (!isRecord(coordinate) || !keys(coordinate, ["row", "column"]) || !Number.isSafeInteger(coordinate.row)
      || !Number.isSafeInteger(coordinate.column) || (coordinate.row as number) < 0 || (coordinate.row as number) >= 1_048_576 || (coordinate.column as number) < 0 || (coordinate.column as number) >= 16_384)
      throw new DocumentError("InvalidOperation", "Invalid transported workbook rule coordinate.");
  } else {
    if (!keys(value, ["documentId", "revision", "chartIds", "maxCharts", "maxReferences", "maxPoints", "maxCodeUnits", "maxFormulaCells", "maxFormulaSteps"]) || !Array.isArray(value.chartIds)
      || value.chartIds.some(id => typeof id !== "string" || id.length > 1024)) throw new DocumentError("InvalidOperation", "Invalid workbook chart transport request.");
    limit("maxCharts", limits.maxCharts); limit("maxReferences", limits.maxReferences); limit("maxPoints", limits.maxPoints); limit("maxCodeUnits", limits.maxCodeUnits);
    limit("maxFormulaCells", limits.maxChartFormulaCells); limit("maxFormulaSteps", limits.maxChartFormulaSteps);
    if (value.chartIds.length > limits.maxCharts) throw new DocumentError("ResourceLimitExceeded", "Too many transported workbook charts.");
  }
  return value;
}
