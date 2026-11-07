import { DocumentError } from "@doffi/core";
import type { ChangeMapping, DocumentHistoryState, DocumentTransaction, TransactionPreparationProgress, TransactionPrepareOptions, TransactionReceipt,
  WorkbookCandidateChange, WorkbookCoordinate, WorkbookRange, WorkbookScalarContext } from "@doffi/core";
import type { WorkbookReadConnection, WorkbookReadTransportOptions } from "@doffi/render";
import { resolveWorkbookReadTransportOptions } from "@doffi/render";
import type { WorkbookFindQuery, WorkbookFindResult, WorkbookReplaceOptions } from "./find-contracts.js";
import type { WorkbookRangeCopyOptions } from "./copy-range.js";
import type { WorkbookFillOptions } from "./fill.js";
import type { WorkbookPasteOptions, WorkbookTextInputMode } from "./text-input.js";

export const WORKBOOK_SESSION_PROTOCOL = "doffi/workbook-session/1";
export interface WorkbookSessionIdentity { readonly documentId: string; readonly revision: number }
export interface WorkbookSessionState extends WorkbookSessionIdentity { readonly history: DocumentHistoryState }
export interface WorkbookSessionInput extends WorkbookSessionIdentity { readonly sheetId: string; readonly at: WorkbookCoordinate; readonly text: string }
/** JSON mapping, rather than a cloned runtime Map with mutable methods. */
export interface WorkbookSessionMapping extends Omit<ChangeMapping, "nodes"> { readonly nodes: readonly (readonly [string, string | null])[] }
export interface WorkbookSessionReceipt extends Omit<TransactionReceipt, "mapping"> { readonly mapping: WorkbookSessionMapping }
export interface WorkbookSessionCandidate extends Pick<WorkbookCandidateChange, "documentId" | "baseRevision" | "affectedNodeIds" | "affectedSheetIds" | "changedRanges" | "changes" | "changedAnnotationIds"> {
  readonly mapping: WorkbookSessionMapping;
}
export type WorkbookSessionCommand =
  | { readonly kind: "transaction"; readonly transaction: DocumentTransaction }
  | (WorkbookSessionIdentity & { readonly kind: "input"; readonly sheetId: string; readonly at: WorkbookCoordinate; readonly text: string; readonly mode?: WorkbookTextInputMode })
  | (WorkbookSessionIdentity & { readonly kind: "paste"; readonly sheetId: string; readonly at: WorkbookCoordinate; readonly text: string; readonly options?: WorkbookPasteOptions })
  | (WorkbookSessionIdentity & { readonly kind: "copy"; readonly sheetId: string; readonly source: WorkbookRange; readonly target: WorkbookCoordinate;
      readonly mode?: WorkbookRangeCopyOptions["mode"]; readonly skipBlanks?: boolean })
  | (WorkbookSessionIdentity & { readonly kind: "fill"; readonly sheetId: string; readonly source: WorkbookRange; readonly target: WorkbookRange;
      readonly mode?: WorkbookFillOptions["mode"]; readonly axis?: WorkbookFillOptions["axis"] })
  | (WorkbookSessionIdentity & { readonly kind: "replace"; readonly query: WorkbookFindQuery; readonly replacement: string; readonly valueMode?: WorkbookReplaceOptions["valueMode"] });
export interface WorkbookSessionProgress { readonly phase: TransactionPreparationProgress["phase"] | "copy" | "fill" | "find"; readonly completed: number; readonly total: number }
export interface WorkbookSessionExecution {
  readonly signal?: AbortSignal;
  readonly onProgress?: (value: WorkbookSessionProgress) => void;
}
export interface WorkbookSessionViewOptions extends WorkbookSessionExecution {
  /** Default follows the source's automatic/manual calculation mode. */
  readonly calculate?: boolean;
  readonly calculationContext?: WorkbookScalarContext;
  /** A still-retained view from this connection. Its exact private model and
   * calculation can supply an optional incremental preview invalidation. */
  readonly previous?: WorkbookReadConnection;
}
export interface WorkbookSessionFindOptions extends WorkbookSessionExecution {
  readonly calculationContext?: WorkbookScalarContext;
  readonly locale?: string;
}
export interface WorkbookSessionTransportOptions {
  readonly maxPendingRequests?: number;
  readonly maxMessageBytes?: number;
  readonly maxMessageDepth?: number;
  readonly maxViews?: number;
  readonly requestTimeoutMs?: number;
  readonly reads?: WorkbookReadTransportOptions;
}
export interface WorkbookSessionConnection {
  state(): WorkbookSessionState;
  prepare(command: WorkbookSessionCommand, options?: TransactionPrepareOptions, execution?: WorkbookSessionExecution): Promise<WorkbookSessionCandidate>;
  /** A view is immutable and belongs to its caller. It transfers metadata and
   * requested windows, never the package or a second editing session. */
  view(candidate?: WorkbookSessionCandidate, options?: WorkbookSessionViewOptions): Promise<WorkbookReadConnection>;
  /** One bounded reversible input spelling, read inside the owning Worker. */
  readInput(sheetId: string, at: WorkbookCoordinate, candidate?: WorkbookSessionCandidate, execution?: WorkbookSessionExecution): Promise<WorkbookSessionInput>;
  commit(candidate: WorkbookSessionCandidate): Promise<WorkbookSessionReceipt>;
  discard(candidate: WorkbookSessionCandidate): Promise<void>;
  undo(): Promise<WorkbookSessionReceipt>;
  redo(): Promise<WorkbookSessionReceipt>;
  find(query: WorkbookFindQuery, options?: WorkbookSessionFindOptions): Promise<WorkbookFindResult>;
  subscribe(listener: (state: WorkbookSessionState) => void): () => void;
  dispose(): void;
}
export interface WorkbookSessionServer { readonly disposed: boolean; dispose(): void }
export interface SessionTransportConfig {
  readonly maxPendingRequests: number; readonly maxMessageBytes: number; readonly maxMessageDepth: number; readonly maxViews: number; readonly requestTimeoutMs: number;
  readonly reads: WorkbookReadTransportOptions;
}
export function sessionTransportConfig(input: WorkbookSessionTransportOptions): SessionTransportConfig {
  if (!record(input) || !keys(input, ["maxPendingRequests", "maxMessageBytes", "maxMessageDepth", "maxViews", "requestTimeoutMs", "reads"])) invalidSessionTransport("Invalid workbook session transport options.");
  const { reads = {}, ...own } = input, config = { maxPendingRequests: 8, maxMessageBytes: 32 * 1024 * 1024, maxMessageDepth: 128, maxViews: 16, requestTimeoutMs: 60_000, ...own };
  if (!Object.values(config).every(value => Number.isSafeInteger(value) && value > 0) || config.requestTimeoutMs > 2_147_483_647 || !record(reads)) invalidSessionTransport("Workbook session budgets must be positive safe integers.");
  return Object.freeze({ ...config, reads: resolveWorkbookReadTransportOptions(reads) });
}
export function record(value: unknown): value is Record<string, unknown> { return value !== null && typeof value === "object" && !Array.isArray(value); }
export function keys(value: Record<string, unknown>, allowed: readonly string[]): boolean { return Object.keys(value).every(key => allowed.includes(key)); }
export function positiveId(value: unknown): value is number { return typeof value === "number" && Number.isSafeInteger(value) && value > 0; }
export function invalidSessionTransport(message: string): never { throw new DocumentError("InvalidOperation", message); }
/** Budget before any schema parser or JSON serialization. Binary resources
 * count by bytes; unsupported prototypes and cyclic messages fail explicitly. */
export function checkSessionMessage(value: unknown, config: SessionTransportConfig): void {
  const stack: { value: unknown; depth: number; leave?: boolean }[] = [{ value, depth: 0 }], active = new Set<object>(); let bytes = 0;
  while (stack.length) {
    const entry = stack.pop()!, item = entry.value;
    if (entry.leave) { active.delete(item as object); continue; }
    bytes += typeof item === "string" ? item.length * 6 + 2 : 32;
    if (entry.depth > config.maxMessageDepth || bytes > config.maxMessageBytes) throw new DocumentError("ResourceLimitExceeded", "Workbook session message exceeds its budget.");
    if (item === null || typeof item === "string" || typeof item === "boolean" || typeof item === "number" && Number.isFinite(item)) continue;
    if (item instanceof Uint8Array) { bytes += item.byteLength; if (bytes > config.maxMessageBytes) throw new DocumentError("ResourceLimitExceeded", "Workbook session resource exceeds its budget."); continue; }
    if (typeof item !== "object" || !Array.isArray(item) && Object.getPrototypeOf(item) !== Object.prototype && Object.getPrototypeOf(item) !== null) invalidSessionTransport("Workbook session messages require finite plain values.");
    if (Array.isArray(item) && item.length > config.maxMessageBytes / 32) throw new DocumentError("ResourceLimitExceeded", "Workbook session array exceeds its budget.");
    if (active.has(item)) invalidSessionTransport("Workbook session messages must be acyclic.");
    active.add(item); stack.push({ value: item, depth: entry.depth, leave: true });
    const names = Object.keys(item); bytes += names.reduce((sum, name) => sum + name.length * 6 + 3, 0);
    if (bytes > config.maxMessageBytes) throw new DocumentError("ResourceLimitExceeded", "Workbook session message exceeds its budget.");
    for (let index = names.length - 1; index >= 0; index--) stack.push({ value: (item as Record<string, unknown>)[names[index]!], depth: entry.depth + 1 });
  }
}
export function freezeSessionValue<Value>(value: Value): Value {
  if (value && typeof value === "object" && !Object.isFrozen(value)) { for (const child of Object.values(value)) freezeSessionValue(child); Object.freeze(value); }
  return value;
}
export function sessionMapping(mapping: ChangeMapping): WorkbookSessionMapping { return { nodes: [...mapping.nodes], textEdits: mapping.textEdits, ...(mapping.workbook ? { workbook: mapping.workbook } : {}) }; }
export function sessionReceipt(receipt: TransactionReceipt): WorkbookSessionReceipt { return { ...receipt, mapping: sessionMapping(receipt.mapping) }; }
export function calculationOptions(context: WorkbookScalarContext | undefined): { context?: WorkbookScalarContext } { return context === undefined ? {} : { context }; }
