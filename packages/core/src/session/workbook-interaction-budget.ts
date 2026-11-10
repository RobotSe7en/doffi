import { defaultTreeAdapter, html as htmlNamespace } from "parse5";
import type { DefaultTreeAdapterTypes } from "parse5";
import { DocumentError } from "../errors.js";
import type { DocumentReadLimits, TransactionPreparationExecution } from "../runtime.js";
import { WorkbookInteractionLimitsSchema } from "../contracts/workbook-interaction.js";
import type { WorkbookInteractionLimits } from "../contracts/workbook-interaction.js";
import { bodyTags } from "../document/html.js";
import { parseHtmlFragment } from "../document/html-parser.js";
import { encode } from "../document/package.js";
import { checkLimit } from "../document/validation.js";
import { checkTransactionSize } from "./limits.js";

export const DEFAULT_WORKBOOK_INTERACTION_LIMITS: Readonly<WorkbookInteractionLimits> = Object.freeze({
  maxBytes: 4 * 1024 * 1024, maxCells: 10_000, maxNodes: 10_000, maxDependencies: 250_000, maxSteps: 1_000_000, maxAssets: 256,
});
export function workbookInteractionLimits(input: Partial<WorkbookInteractionLimits> = {}): WorkbookInteractionLimits {
  const parsed = WorkbookInteractionLimitsSchema.partial().safeParse(input);
  if (!parsed.success || !Object.values(parsed.data).every(Number.isSafeInteger)) throw new DocumentError("InvalidOperation", "Workbook context limits must be supported positive safe integers.");
  const result = { ...DEFAULT_WORKBOOK_INTERACTION_LIMITS };
  for (const key of Object.keys(result) as (keyof WorkbookInteractionLimits)[]) { const value = parsed.data[key]; if (value !== undefined) result[key] = value; }
  return result;
}
/** One request owns all source, dependency, parsing and evaluation work. Limits
 * never restart for another chunk, expression, name, table or native object. */
export class WorkbookInteractionBudget {
  private bytes = 0;
  private nodes = 0;
  private work = 0;
  private cells = 0;
  private derivedUnits = 0;
  private dependencies = 0;
  constructor(readonly limits: Readonly<WorkbookInteractionLimits>, private readonly sourceLimits: DocumentReadLimits, private readonly signal?: TransactionPreparationExecution["signal"]) {}
  step = (): void => {
    if (this.signal?.aborted) throw new DocumentError("OperationCancelled", "Workbook context preparation was cancelled.");
    checkLimit(++this.work, this.limits.maxSteps);
  };
  cell(): void { this.step(); checkLimit(++this.cells, this.limits.maxCells); }
  node(): void { this.step(); checkLimit(++this.nodes, this.limits.maxNodes); }
  dependency(): void { this.step(); checkLimit(++this.dependencies, this.limits.maxDependencies); }
  account = (units: number): void => { this.derivedUnits += units; checkLimit(this.derivedUnits, this.limits.maxBytes); };
  take(value: unknown): void {
    this.step(); const remaining = this.limits.maxBytes - this.bytes;
    // Preflight the escape upper bound before allocating JSON or UTF-8 bytes.
    checkTransactionSize(value, Math.min(Number.MAX_SAFE_INTEGER, remaining * 8), this.sourceLimits.maxDepth + 16);
    this.bytes += encode(JSON.stringify(value)).byteLength; checkLimit(this.bytes, this.limits.maxBytes);
  }
  html(source: string, cell = false): void {
    // Admitted plain cells need no second HTML tree. Rich cells and auxiliary
    // blocks use the existing inert, bounded parser and count every text node.
    const plain = cell ? /^<td\b[^<>]*>([^<]*)<\/td>$/.exec(source) : undefined;
    if (plain) { this.node(); if (plain[1]) this.node(); return; }
    const remaining = this.limits.maxNodes - this.nodes; checkLimit(1, remaining);
    const context = defaultTreeAdapter.createElement(cell ? "tr" : "div", htmlNamespace.NS.HTML, []);
    const fragment = parseHtmlFragment(source, context, { ...this.sourceLimits, maxNodes: remaining + 16 }, bodyTags), pending: DefaultTreeAdapterTypes.Node[] = [...fragment.childNodes];
    while (pending.length) { const node = pending.pop()!; this.node(); if ("childNodes" in node) pending.push(...node.childNodes); }
  }
  progress(): { readonly completed: number; readonly total: number } { return { completed: this.cells, total: 0 }; }
}
