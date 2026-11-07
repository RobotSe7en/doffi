import type { WorkbookModel, DocumentReadLimits, WorkbookReadLimits, DocumentPackage } from "../runtime.js";
import type { ChunkIndex } from "./chunk-index.js";
import type { ParsedHtml } from "../document/html.js";

export interface WorkbookSource {
  readonly files: DocumentPackage;
  readonly chunks: ReadonlyMap<string, ChunkIndex>;
  readonly classes: ReadonlySet<string>;
  readonly bindingKey: string;
  readonly limits: DocumentReadLimits;
  readonly workbookLimits: WorkbookReadLimits;
  readonly content?: ParsedHtml;
}
/** Private ownership registry; public serialization always copies bytes. */
const sources = new WeakMap<WorkbookModel, WorkbookSource>();
const providers = new WeakMap<WorkbookModel, () => WorkbookModel>();
export function registerWorkbookSource(model: WorkbookModel, source: WorkbookSource): void { sources.set(model, source); }
export function workbookSource(model: WorkbookModel): WorkbookSource | undefined { return sources.get(model); }
export function registerWorkbookModelProvider(model: WorkbookModel, provider: () => WorkbookModel): void { providers.set(model, provider); }
export function currentWorkbookModel(model: WorkbookModel): WorkbookModel { return providers.get(model)?.() ?? model; }
/** Capture a built-in editing session's immutable current revision without
 * serializing its package. Host-provided models must already be immutable. */
export function captureWorkbookModel(model: WorkbookModel): WorkbookModel { return currentWorkbookModel(model); }
