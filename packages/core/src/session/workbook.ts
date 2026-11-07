import type { AnnotationOperation, DocumentOperation, DocumentTransaction } from "../contracts/operations.js";
import { WorkbookOperationSchema } from "../contracts/workbook-operations.js";
import type { WorkbookOperation } from "../contracts/workbook-operations.js";
import type { DocumentSessionOptions, TransactionPrepareOptions, WorkbookModel, WorkbookSession, WorkbookStructureMapping } from "../runtime.js";
import { parseContract } from "../schema-registry.js";
import { DocumentError } from "../errors.js";
import { openDocument } from "../document-model.js";
import { readManifest } from "../document/manifest.js";
import { encode, validatePackage } from "../document/package.js";
import { readLimits } from "../document/validation.js";
import { openWorkbookModel, openWorkbookModelSteps } from "../workbook/model.js";
import { createWorkbookCalculator } from "../workbook/calculation.js";
import { copyWorkbookScalarContext } from "../workbook/scalar.js";
import { workbookSource, registerWorkbookModelProvider } from "../workbook/source.js";
import { createSessionLifecycle } from "./lifecycle.js";
import type { SessionPreparationControl, SessionPrepared } from "./lifecycle.js";
import { finishWorkbookPreparation, finishWorkbookPreparationAsync } from "./workbook-preparation.js";
import type { WorkbookPreparation } from "./workbook-preparation.js";
import { compareWorkbooks, compareWorkbookSteps, retainedWorkbookBytes, sameWorkbookSource } from "../workbook/source-changes.js";
import { WorkbookCellEdit } from "./workbook-cells.js";
import { checkResourceUse, stageResourceFiles } from "./resources.js";
import { validateWorkbookCandidateSteps } from "./workbook-validation.js";
import { checkTransactionSize } from "./limits.js";
import { WorkbookContentEdit } from "./workbook-content.js";
import { editWorkbookSheet } from "./workbook-sheets.js";
import { editWorkbookHyperlink, sameWorkbookHyperlinks } from "./workbook-hyperlinks.js";
import { compileNumberFormat } from "../workbook/number-format.js";
import { WorkbookRegionIndex } from "../workbook/region-index.js";

const names = new Set<string>(WorkbookOperationSchema.options.map(schema => schema.shape.op.value));
export function isWorkbookOperation(operation: DocumentOperation): operation is WorkbookOperation { return names.has(operation.op); }
function invalid(message: string): never { throw new DocumentError("InvalidOperation", message); }

export function createWorkbookSession(document: WorkbookModel, options: DocumentSessionOptions): WorkbookSession {
  const limits = readLimits(options.limits);
  const reopen = (previous: WorkbookModel, files: Map<string, Uint8Array>): WorkbookModel => {
    validatePackage(files, limits);
    const manifest = readManifest(files); if (manifest.kind !== "workbook") invalid("Workbook transactions cannot change the document kind.");
    return openWorkbookModel(files, manifest, limits, options.workbookLimits, previous);
  };
  const source = workbookSource(document);
  const initial = source ? reopen(document, new Map(source.files)) : openDocument(document.serialize(), { ...options, kind: "workbook" });
  function rebase(model: WorkbookModel, revision: number): WorkbookModel {
    const files = new Map(workbookSource(model)!.files), manifest = { ...model.snapshot().manifest, revision };
    const annotations = model.snapshot().annotations.map(annotation => annotation.status === "needs_reanchor" ? annotation : { ...annotation, baseRevision: revision });
    files.set("manifest.json", encode(JSON.stringify(manifest))); files.set(manifest.annotations, encode(JSON.stringify(annotations)));
    return reopen(model, files);
  }
  function* preparationSteps(base: WorkbookModel, transaction: DocumentTransaction, allocate: () => string, resolvedIds?: readonly string[], preparation?: TransactionPrepareOptions): WorkbookPreparation<SessionPrepared<WorkbookModel>> {
    if (preparation !== undefined && (!preparation || typeof preparation !== "object" || Array.isArray(preparation) || Object.keys(preparation).some(name => !["assets", "calculationContext", "validation"].includes(name)))) invalid("Invalid workbook preparation options.");
    const context = preparation?.calculationContext === undefined ? undefined : copyWorkbookScalarContext(preparation.calculationContext);
    if (preparation?.validation !== undefined) {
      checkTransactionSize(preparation.validation, limits.maxFileBytes, limits.maxDepth);
      if (!preparation.validation || typeof preparation.validation !== "object" || Object.keys(preparation.validation).some(name => name !== "acceptedRuleIds")
        || !Array.isArray(preparation.validation.acceptedRuleIds) || preparation.validation.acceptedRuleIds.some(id => typeof id !== "string" || !id)) invalid("Invalid validation acknowledgment options.");
    }
    const captured: TransactionPrepareOptions = { ...(context ? { calculationContext: context } : {}),
      ...(preparation?.validation ? { validation: { acceptedRuleIds: [...preparation.validation.acceptedRuleIds] } } : {}) };
    const staged = stageResourceFiles(workbookSource(base)!.files, base.snapshot().manifest, preparation?.assets !== undefined ? { assets: preparation.assets } : undefined, limits);
    const parsed = parseContract("document-manifest", staged.manifest); if (parsed.kind !== "workbook") invalid("Expected a workbook manifest.");
    const manifest = parsed, files = new Map(staged.files), metadata = parseContract("document-metadata", base.snapshot().metadata);
    if (!("workbook" in metadata)) invalid("Expected workbook metadata.");
    const workbook = metadata.workbook, annotations = parseContract("annotations", base.snapshot().annotations);
    const baseSheets = new Set(workbook.sheets.map(sheet => sheet.sheetId));
    for (const operation of transaction.operations) {
      if (!isWorkbookOperation(operation) && !["addAnnotation", "updateAnnotation", "removeAnnotation"].includes(operation.op)) invalid("Content and presentation operations cannot modify workbook source.");
      if ("sheetId" in operation && !baseSheets.has(operation.sheetId)) invalid("Worksheet targets must belong to the base snapshot.");
      if ("beforeSheetId" in operation && operation.beforeSheetId !== undefined && !baseSheets.has(operation.beforeSheetId)) invalid("Worksheet references must belong to the base snapshot.");
    }
    const cells = new WorkbookCellEdit(base, manifest, files, limits, allocate);
    const richContent = new WorkbookContentEdit(manifest, files, limits, allocate), mapping: WorkbookStructureMapping[] = [], hyperlinkSheets = new Set<string>();
    function annotation(operation: AnnotationOperation): void {
      if (operation.op === "addAnnotation") {
        const value = operation.annotation;
        if (value.documentId !== manifest.documentId || value.baseRevision !== manifest.revision || annotations.some(item => item.annotationId === value.annotationId)) invalid("Invalid workbook annotation identity or revision.");
        if (value.anchor.kind !== "cells" || !baseSheets.has(value.anchor.sheetId)) invalid("New workbook annotations require an existing worksheet range.");
        annotations.push(value); return;
      }
      const index = annotations.findIndex(item => item.annotationId === operation.annotationId); if (index < 0) invalid("The annotation does not exist.");
      if (operation.op === "removeAnnotation") { annotations.splice(index, 1); return; }
      const previous = annotations[index]!, value = parseContract("annotation", { ...previous, ...operation.changes });
      if (previous.status === "needs_reanchor" && value.status !== "needs_reanchor" && !operation.changes.anchor) invalid("Reactivating an annotation requires an explicit anchor.");
      if (operation.changes.anchor && (value.anchor.kind !== "cells" || !baseSheets.has(value.anchor.sheetId))) invalid("Workbook anchors require an existing worksheet.");
      if (operation.changes.anchor && value.anchor.kind === "cells") {
        for (const target of [value, ...value.replies ?? []]) if (target.bodyNodeId) {
          const story = workbook.stories?.find(story => story.rootNodeId === target.bodyNodeId); if (story) story.sheetId = value.anchor.sheetId;
        }
      }
      if (value.status !== "needs_reanchor") {
        const binding = workbook.comments?.find(comment => comment.annotationId === value.annotationId);
        if (binding) { delete binding.recoveryStatus; delete binding.recoveryReasons; }
      }
      annotations[index] = value;
    }
    let operationCount = 0;
    yield { kind: "progress", value: { phase: "operations", completed: 0, total: transaction.operations.length } };
    for (const operation of transaction.operations) {
      if (!isWorkbookOperation(operation)) annotation(operation as AnnotationOperation);
      else if (operation.op === "setWorkbookCells" || operation.op === "setWorkbookCellFormats") {
        const sheet = workbook.sheets.find(item => item.sheetId === operation.sheetId) ?? invalid("An earlier operation removed the worksheet.");
        const seen = new Set<string>(), merges = new WorkbookRegionIndex(sheet.merges);
        let completed = 0;
        for (const cell of operation.cells) {
          const key = cell.row + ":" + cell.column; if (seen.has(key)) invalid("A cell batch cannot repeat coordinates."); seen.add(key);
          const merge = merges.at(cell.row, cell.column);
          if (merge && (merge.from.row !== cell.row || merge.from.column !== cell.column)) invalid("Only the leading cell of a merged region can be edited.");
          if (operation.op === "setWorkbookCells" && "source" in cell) cells.set(operation.sheetId, cell);
          else cells.format(operation.sheetId, cell);
          if (++completed % 256 === 0) yield { kind: "progress", value: { phase: "cells", completed, total: operation.cells.length } };
        }
        if (completed % 256 !== 0) yield { kind: "progress", value: { phase: "cells", completed, total: operation.cells.length } };
      } else if (operation.op === "setWorkbookNumberFormat") {
        if ("code" in operation.format) compileNumberFormat(operation.format.code);
        const index = workbook.numberFormats.findIndex(format => format.formatId === operation.format.formatId);
        if (index < 0) workbook.numberFormats.push(operation.format); else workbook.numberFormats[index] = operation.format;
      } else if (operation.op === "setWorkbookHyperlink") {
        const sheet = workbook.sheets.find(item => item.sheetId === operation.sheetId) ?? invalid("An earlier operation removed the worksheet.");
        editWorkbookHyperlink(sheet, operation, limits.maxNodes);
        hyperlinkSheets.add(sheet.sheetId);
      } else if (operation.op === "clearWorkbookRanges") {
        if (!workbook.sheets.some(sheet => sheet.sheetId === operation.sheetId)) invalid("An earlier operation removed the worksheet.");
        for (const value of cells.clearSteps(operation.sheetId, operation.ranges, operation.mode)) yield { kind: "progress", value: { phase: "cells", ...value } };
      }
      else if (operation.op === "setWorkbookSettings") Object.assign(workbook, operation.changes);
      else editWorkbookSheet(workbook, cells, richContent, annotations, operation, allocate, mapping);
      yield { kind: "progress", value: { phase: "operations", completed: ++operationCount, total: transaction.operations.length } };
    }
    for (const sheet of workbook.sheets.filter(sheet => hyperlinkSheets.has(sheet.sheetId))) {
      const previous = base.snapshot().metadata.workbook.sheets.find(previous => previous.sheetId === sheet.sheetId);
      if (previous && sameWorkbookHyperlinks(previous.hyperlinks ?? [], sheet.hyperlinks ?? [])) {
        if (previous.hyperlinks === undefined) delete sheet.hyperlinks;
        else sheet.hyperlinks = previous.hyperlinks.map(link => ({ ...link, range: { from: { ...link.range.from }, to: { ...link.range.to } } }));
      }
    }
    if (resolvedIds) for (const value of annotations) if (resolvedIds.includes(value.annotationId) && value.status === "open") value.status = "resolved";
    for (const value of cells.flushSteps()) yield { kind: "progress", value: { phase: "chunks", ...value } };
    richContent.flush();
    files.set("manifest.json", encode(JSON.stringify(manifest))); files.set(manifest.metadata, encode(JSON.stringify(metadata)));
    files.set(manifest.annotations, encode(JSON.stringify(annotations)));
    yield { kind: "progress", value: { phase: "validation", completed: 0, total: 0 } };
    validatePackage(files, limits);
    const loadedManifest = readManifest(files); if (loadedManifest.kind !== "workbook") invalid("Workbook transactions cannot change the document kind.");
    const loader = openWorkbookModelSteps(files, loadedManifest, limits, options.workbookLimits, base);
    let next: WorkbookModel;
    try {
      for (;;) {
        const step = loader.next();
        if (step.done) { next = step.value ?? invalid("Workbook loading did not complete."); break; }
        yield { kind: "progress", value: { phase: "validation", ...step.value } };
      }
    } finally { loader.return(undefined); }
    const content = workbookSource(next)!.content;
    if (staged.added.size) { if (!content || !manifest.content) invalid("New resources must be referenced by workbook content."); checkResourceUse(staged.added, content, manifest.content); }
    const comparison = compareWorkbookSteps(base, next, { nodes: new Map<string, string | null>(), textEdits: [], workbook: mapping });
    let diff: ReturnType<typeof compareWorkbooks>;
    try {
      for (;;) {
        const step = comparison.next();
        if (step.done) { diff = step.value ?? invalid("Workbook comparison did not complete."); break; }
        yield { kind: "progress", value: { phase: "validation", ...step.value } };
      }
    } finally { comparison.return(undefined); }
    yield* validateWorkbookCandidateSteps(base, next, diff.changedRanges, captured);
    yield { kind: "progress", value: { phase: "validation", completed: 1, total: 1 } };
    return { state: next, mapping: diff.mapping, diff };
  }
  const prepare = (base: WorkbookModel, transaction: DocumentTransaction, allocate: () => string, resolvedIds?: readonly string[], preparation?: TransactionPrepareOptions) => finishWorkbookPreparation(preparationSteps(base, transaction, allocate, resolvedIds, preparation));
  const prepareAsync = (base: WorkbookModel, transaction: DocumentTransaction, allocate: () => string, resolvedIds: readonly string[] | undefined, preparation: TransactionPrepareOptions | undefined, control: SessionPreparationControl) => finishWorkbookPreparationAsync(preparationSteps(base, transaction, allocate, resolvedIds, preparation), control);
  const boundary = createSessionLifecycle<WorkbookModel, WorkbookModel>(initial, { model: state => state, prepare, prepareAsync, rebase, compare: compareWorkbooks,
    same: sameWorkbookSource, retained: retainedWorkbookBytes, identities: model => workbookSource(model)?.content?.nodeIds() ?? [] }, options);
  const current = boundary.state;
  const calculator = createWorkbookCalculator(current, options.calculationLimits);
  const session = Object.freeze<WorkbookSession>({ kind: "workbook", snapshot: () => current().snapshot(), getNode: id => current().getNode(id),
    getCell: (sheetId, row, column) => current().getCell(sheetId, row, column), readWindow: window => current().readWindow(window),
    getAsset: (path, maxBytes) => current().getAsset(path, maxBytes), resolveReference: (expression, origin) => current().resolveReference(expression, origin),
    serialize: () => current().serialize(), ...boundary.lifecycle,
    prepareAsync: (transaction, preparation, execution) => boundary.prepareAsyncTransaction(transaction, undefined, preparation, execution),
    calculate: calculator.calculate, calculateSync: calculator.calculateSync,
    getCalculatedCell: (sheetId, row, column) => calculator.getCell({ sheetId, row, column }),
    dispose: () => { boundary.lifecycle.dispose(); calculator.dispose(); },
  });
  registerWorkbookModelProvider(session, current); return session;
}
