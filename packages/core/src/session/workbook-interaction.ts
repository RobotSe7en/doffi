import type { WorkbookEditProposal, WorkbookInteractionRequest, WorkbookInteractionValue, WorkbookProposalReview } from "../contracts/workbook-interaction.js";
import type { TransactionPreparationExecution, TransactionPreparationProgress, TransactionPrepareOptions, WorkbookCandidateChange, WorkbookModel } from "../runtime.js";
import type { WorkbookReferenceArea, WorkbookReferenceLocation } from "../workbook/reference.js";
import { DocumentError } from "../errors.js";
import { checkLimit, freeze } from "../document/validation.js";
import { encode } from "../document/package.js";
import { parseContract } from "../schema-registry.js";
import { copyWorkbookScalarContext, createWorkbookScalarKernel } from "../workbook/scalar.js";
import { workbookSource } from "../workbook/source.js";
import { buildWorkbookRequestSteps } from "./workbook-interaction-context.js";
import { WorkbookInteractionBudget, workbookInteractionLimits } from "./workbook-interaction-budget.js";
import { workbookContextContains, workbookContextCovers } from "./workbook-interaction-scope.js";
import { createWorkbookInteractionRuleReader, workbookInteractionExpression } from "./workbook-interaction-rules.js";
import { validateWorkbookProposalSource, validateWorkbookProposalTargets } from "./workbook-proposal-scope.js";
import { checkTransactionSize } from "./limits.js";
import { stageResourceFiles } from "./resources.js";

const equal = (a: unknown, b: unknown) => JSON.stringify(a) === JSON.stringify(b);
const calculationContext = (request: WorkbookInteractionRequest) => copyWorkbookScalarContext({ ...(request.context.calculationContext.now ? { now: request.context.calculationContext.now } : {}), ...(request.context.calculationContext.locale !== undefined ? { locale: request.context.calculationContext.locale } : {}) });
function invalid(message: string): never { throw new DocumentError("InvalidProposal", message); }
export function finishWorkbookInteraction<Value>(steps: Generator<TransactionPreparationProgress, Value>): Value {
  try { for (;;) { const next = steps.next(); if (next.done) return next.value; } } finally { steps.return(undefined as Value); }
}
export function workbookInteractionExecution(input: TransactionPreparationExecution = {}): TransactionPreparationExecution {
  if (!input || typeof input !== "object" || Array.isArray(input) || Object.keys(input).some(key => !["signal", "yield", "onProgress"].includes(key))
    || input.signal !== undefined && (!input.signal || typeof input.signal.aborted !== "boolean") || input.yield !== undefined && typeof input.yield !== "function" || input.onProgress !== undefined && typeof input.onProgress !== "function") throw new DocumentError("InvalidOperation", "Invalid workbook interaction execution options.");
  return { ...input };
}
export async function finishWorkbookInteractionAsync<Value>(steps: Generator<TransactionPreparationProgress, Value>, execution: TransactionPreparationExecution, guard: () => void): Promise<Value> {
  try {
    for (;;) {
      guard(); const next = steps.next(); guard(); if (next.done) return next.value;
      try { execution.onProgress?.(freeze({ ...next.value })); } catch { /* Host progress is reporting only. */ }
      guard(); await execution.yield?.(); guard();
    }
  } finally { steps.return(undefined as Value); }
}
export interface WorkbookProposalCapture {
  readonly request: WorkbookInteractionRequest;
  readonly proposal: WorkbookEditProposal;
  readonly budget: WorkbookInteractionBudget;
  readonly options: TransactionPrepareOptions;
}
/** Capture all caller data, including authorized resource bytes, before yielding.
 * Rebuilding checks completeness and provenance rather than trusting an ID. */
export function* validateWorkbookProposalSteps(model: WorkbookModel, requestInput: WorkbookInteractionRequest, proposalInput: WorkbookEditProposal, optionsInput: TransactionPrepareOptions | undefined,
  maxTransactionBytes: number, execution: TransactionPreparationExecution = {}): Generator<TransactionPreparationProgress, WorkbookProposalCapture> {
  const source = workbookSource(model)!, limits = source.limits;
  checkTransactionSize(requestInput, Math.min(Number.MAX_SAFE_INTEGER, limits.maxPackageBytes * 8), limits.maxDepth + 16);
  checkLimit(encode(JSON.stringify(requestInput)).byteLength, limits.maxPackageBytes);
  checkTransactionSize(proposalInput, maxTransactionBytes, limits.maxDepth + 16);
  const request = parseContract("workbook-interaction-request", requestInput), proposal = parseContract("workbook-edit-proposal", proposalInput), manifest = model.snapshot().manifest;
  if (request.documentId !== manifest.documentId || proposal.documentId !== request.documentId || proposal.requestId !== request.requestId) invalid("The workbook proposal does not match this request and document.");
  if (request.baseRevision !== manifest.revision || proposal.baseRevision !== request.baseRevision) throw new DocumentError("RevisionConflict", "The workbook request or proposal is based on a different revision.");
  const budget = new WorkbookInteractionBudget(workbookInteractionLimits(request.context.limits), limits, execution.signal);
  if (optionsInput !== undefined && (!optionsInput || typeof optionsInput !== "object" || Array.isArray(optionsInput) || Object.keys(optionsInput).some(key => !["assets", "calculationContext", "validation", "measureWorkbookLayout"].includes(key)))) invalid("Invalid workbook proposal preparation options.");
  if (optionsInput?.measureWorkbookLayout !== undefined && typeof optionsInput.measureWorkbookLayout !== "function") invalid("Workbook layout measurement requires an authorized measurer.");
  // The existing resource adapter validates/copies each byte vector before a
  // host callback can mutate it. The same immutable copies enter the candidate.
  const staged = stageResourceFiles(source.files, manifest, optionsInput?.assets !== undefined ? { assets: optionsInput.assets } : undefined, limits);
  const options: TransactionPrepareOptions = {
    ...(optionsInput?.measureWorkbookLayout ? { measureWorkbookLayout: optionsInput.measureWorkbookLayout } : {}),
    ...(optionsInput?.assets ? { assets: optionsInput.assets.map(resource => ({ descriptor: parseContract("asset-descriptor", resource.descriptor), bytes: new Uint8Array(staged.files.get(resource.descriptor.path)!) })) } : {}),
    calculationContext: calculationContext(request),
    ...(optionsInput?.validation ? { validation: { acceptedRuleIds: [...optionsInput.validation.acceptedRuleIds] } } : {}),
  };
  validateWorkbookProposalTargets(model, request, proposal, optionsInput, budget);
  const canonical = yield* buildWorkbookRequestSteps(model, { kind: "workbook", annotationIds: request.annotations.map(annotation => annotation.annotationId), scope: request.context.scope,
    readRanges: request.context.requestedReadRanges, supportedOperations: request.supportedOperations, assets: request.context.providedAssets, limits: request.context.limits, calculationContext: calculationContext(request) }, execution, request.requestId, budget);
  if (!equal(request, canonical)) invalid("The workbook request differs from its complete frozen source context.");
  return { request, proposal, budget, options };
}

/** Verify current candidate reads against the frozen sparse coverage and expose
 * derived value differences without turning consumers into writable cells. */
export function* reviewWorkbookProposalSteps(model: WorkbookModel, capture: WorkbookProposalCapture, candidate: WorkbookCandidateChange): Generator<TransactionPreparationProgress, WorkbookProposalReview> {
  const { request, proposal, budget } = capture, context = request.context, coverage = [...context.writableRanges, ...context.dependencies];
  validateWorkbookProposalSource(model, request, proposal, candidate, budget);
  const next = candidate.model, global = context.scope.kind === "global", knownNames = new Set(context.metadata.workbook.names?.map(name => name.nameId));
  const check = (area: import("../runtime.js").DeepReadonly<WorkbookReferenceArea>): void => {
    budget.step(); if (!global && !workbookContextCovers(coverage, { sheetId: area.sheetId, range: area }, budget.step)) invalid("The workbook candidate needs inputs outside the frozen context; expand read scope and rebuild the request.");
  };
  const name = (definition: { readonly nameId: string }): void => { budget.step(); if (!global && !knownNames.has(definition.nameId)) invalid("The workbook candidate needs an undisclosed defined name; rebuild the request with its read context."); };
  const kernel = createWorkbookScalarKernel(next, { maxCells: context.limits.maxCells, maxSteps: context.limits.maxSteps, maxCodeUnits: context.limits.maxBytes }, calculationContext(request),
    { step: budget.step, account: budget.account, reference: check, read: at => check({ sheetId: at.sheetId, from: at, to: at }), name });
  const discover = (expression: string, origin: WorkbookReferenceLocation, at = origin): void => workbookInteractionExpression(next, expression, origin, at, budget.step, check, name, kernel);
  const changed: WorkbookProposalReview["calculatedChanges"] = [], seen = new Set<string>(), before = new Map(context.cells.map(cell => [JSON.stringify([cell.sheetId, cell.row, cell.column]), cell]));
  const sheets = new Set(next.snapshot().metadata.workbook.sheets.map(sheet => sheet.sheetId));
  function result(at: WorkbookReferenceLocation): WorkbookInteractionValue {
    if (!sheets.has(at.sheetId)) return { kind: "value", value: { type: "blank" } };
    const value = kernel.cell(at); return value.kind === "value" ? { kind: "value", value: value.value } : { kind: "unavailable", reason: value.reason, ...(value.functionName ? { functionName: value.functionName } : {}) };
  }
  try {
    let completed = 0;
    const at = function* (): Generator<WorkbookReferenceLocation> {
      for (const cell of context.cells) yield { sheetId: cell.sheetId, row: cell.row, column: cell.column };
      for (const entry of candidate.changedRanges ?? []) {
        const range = entry.range;
        // Only stored candidate cells need expression validation. Removed and
        // implicit blank values are already compared through captured cells.
        for (const chunk of next.snapshot().manifest.chunks) {
          budget.step(); if (chunk.sheetId !== entry.sheetId) continue;
          const row = Math.max(chunk.row, range.from.row), column = Math.max(chunk.column, range.from.column), rows = Math.min(chunk.row + 255, range.to.row) - row + 1, columns = Math.min(chunk.column + 31, range.to.column) - column + 1;
          if (rows > 0 && columns > 0) for (const cell of workbookSource(next)!.chunks.get(chunk.path)!.readWindow({ row, column, rows, columns })) yield { sheetId: entry.sheetId, row: cell.row, column: cell.column };
        }
      }
    };
    for (const location of at()) {
      budget.step(); const key = JSON.stringify([location.sheetId, location.row, location.column]); if (seen.has(key)) continue; seen.add(key);
      const cell = sheets.has(location.sheetId) ? next.getCell(location.sheetId, location.row, location.column) : undefined, previous = before.get(key);
      if (cell?.formula) discover(cell.formula.expression, location);
      const after = result(location), old: WorkbookInteractionValue = previous?.current ?? { kind: "value", value: { type: "blank" } };
      if (!equal(old, after)) changed.push({ ...location, access: workbookContextContains(context.writableRanges, location.sheetId, location.row, location.column, budget.step) ? "editable" : "read-only", before: old, after });
      if (++completed % 128 === 0) { yield { phase: "calculation", completed, total: 0 }; budget.step(); }
    }
    const readRules = createWorkbookInteractionRuleReader(next, budget, kernel, discover, check);
    for (const entry of coverage) if (sheets.has(entry.sheetId)) yield* readRules(entry);
    for (const object of next.snapshot().metadata.workbook.objects ?? []) {
      budget.step(); if (object.kind !== "chart" || !global && !context.writableObjectIds.includes(object.nodeId) && !context.writableNewObjectSheetIds.includes(object.sheetId)) continue;
      const chart = next.snapshot().metadata.workbook.charts!.find(chart => chart.chartId === object.chartId)!;
      const references = [chart.categories, ...chart.series.flatMap(series => [series.values, series.name?.kind === "reference" ? series.name.reference : undefined]), chart.titleReference, chart.categoryAxisTitle?.reference, chart.valueAxisTitle?.reference];
      for (const reference of references) if (reference && "expression" in reference) {
        const origin = { sheetId: object.sheetId, row: 0, column: 0 }; discover(reference.expression, origin); const bound = kernel.referenceAreas(reference.expression, origin); if (bound.kind === "ranges") for (const area of bound.areas) check(area);
      }
    }
    const review = { requestId: request.requestId, proposalId: proposal.proposalId, summary: proposal.summary, calculatedChanges: changed };
    budget.take(review); return freeze(review);
  } finally { kernel.dispose(); }
}
