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
import { copyWorkbookSheet } from "./workbook-copy-sheet.js";
import { copyWorkbookComments, workbookCommentCopyPlan, workbookCommentCopySheets } from "./workbook-copy-comments.js";
import { editWorkbookAxis, restoreWorkbookAxisSource } from "./workbook-axis.js";
import { editWorkbookHyperlink, replaceWorkbookHyperlinks, sameWorkbookHyperlinks } from "./workbook-hyperlinks.js";
import { compileNumberFormat } from "../workbook/number-format.js";
import { WorkbookRegionIndex } from "../workbook/region-index.js";
import { editWorkbookMerge, sameWorkbookMerges } from "./workbook-merges.js";
import { editWorkbookDefinition, isWorkbookDefinitionOperation, restoreWorkbookDefinitionSource } from "./workbook-definitions.js";
import { replaceWorkbookRules } from "./workbook-rule-replace.js";
import { reorderWorkbookRows } from "./workbook-reorder.js";
import { moveWorkbookRange } from "./workbook-move.js";
import { SourceStyleEdit } from "./source-styles.js";
import { editWorkbookStyle, isWorkbookStyleOperation } from "./workbook-styles.js";
import { editWorkbookDrawing, isWorkbookDrawingOperation, validateWorkbookDrawingTarget } from "./workbook-drawings.js";
import { editWorkbookComment, isWorkbookCommentOperation, reanchorWorkbookCommentSources, removeWorkbookCommentSources, restoreWorkbookCommentSource, validateWorkbookCommentTarget } from "./workbook-comments.js";
import { WorkbookCommentMapping } from "./workbook-comment-mapping.js";
import { buildWorkbookRequestSteps } from "./workbook-interaction-context.js";
import { finishWorkbookInteraction, finishWorkbookInteractionAsync, reviewWorkbookProposalSteps, validateWorkbookProposalSteps, workbookInteractionExecution } from "./workbook-interaction.js";
import { sessionLimits } from "./limits.js";
import type { WorkbookCandidateChange, TransactionPreparationExecution } from "../runtime.js";
import type { WorkbookProposalReview } from "../contracts/workbook-interaction.js";
import { compensateWorkbookLayout, measureWorkbookCandidate, workbookLayoutBounds, workbookLayoutOwners } from "./workbook-layout.js";

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
    if (preparation !== undefined && (!preparation || typeof preparation !== "object" || Array.isArray(preparation) || Object.keys(preparation).some(name => !["assets", "calculationContext", "validation", "measureWorkbookLayout"].includes(name)))) invalid("Invalid workbook preparation options.");
    const measureLayout = preparation?.measureWorkbookLayout;
    if (measureLayout !== undefined && typeof measureLayout !== "function") invalid("Workbook layout measurement requires an authorized measurer.");
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
      if ("targetSheetId" in operation && !baseSheets.has(operation.targetSheetId)) invalid("Worksheet destinations must belong to the base snapshot.");
      if ("beforeSheetId" in operation && operation.beforeSheetId !== undefined && !baseSheets.has(operation.beforeSheetId)) invalid("Worksheet references must belong to the base snapshot.");
      if (isWorkbookOperation(operation) && isWorkbookDrawingOperation(operation)) validateWorkbookDrawingTarget(base, operation);
      if (isWorkbookCommentOperation(operation)) validateWorkbookCommentTarget(base, operation);
    }
    const cells = new WorkbookCellEdit(base, manifest, files, limits, allocate);
    const styles = new SourceStyleEdit(files, manifest.styles, limits);
    const richContent = new WorkbookContentEdit(manifest, files, limits, allocate, workbook), mapping: WorkbookStructureMapping[] = [], hyperlinkSheets = new Set<string>(), mergeSheets = new Set<string>();
    const commentMapping = new WorkbookCommentMapping();
    let measuredModel = base;
    function* projectLayoutModel(): WorkbookPreparation<WorkbookModel> {
      const projectedManifest = parseContract("document-manifest", manifest); if (projectedManifest.kind !== "workbook") invalid("Expected a workbook candidate.");
      const projectedFiles = new Map(files);
      for (const value of cells.projectSteps(projectedFiles, projectedManifest)) yield { kind: "progress", value: { phase: "chunks", ...value } };
      richContent.project(projectedFiles);
      projectedFiles.set("manifest.json", encode(JSON.stringify(projectedManifest)));
      projectedFiles.set(manifest.metadata, encode(JSON.stringify(metadata))); projectedFiles.set(manifest.annotations, encode(JSON.stringify(annotations)));
      const loader = openWorkbookModelSteps(projectedFiles, projectedManifest, limits, options.workbookLimits, measuredModel);
      try {
        for (;;) {
          const step = loader.next();
          if (step.done) { measuredModel = step.value ?? invalid("Workbook measurement source did not load."); return measuredModel; }
          yield { kind: "progress", value: { phase: "validation", ...step.value } };
        }
      } finally { loader.return(undefined); }
    }
    function annotation(operation: AnnotationOperation): void {
      if (operation.op === "addAnnotation") {
        const value = operation.annotation;
        if (value.documentId !== manifest.documentId || value.baseRevision !== manifest.revision || annotations.some(item => item.annotationId === value.annotationId)) invalid("Invalid workbook annotation identity or revision.");
        if (value.anchor.kind !== "cells" || !baseSheets.has(value.anchor.sheetId)) invalid("New workbook annotations require an existing worksheet range.");
        annotations.push(value); return;
      }
      const index = annotations.findIndex(item => item.annotationId === operation.annotationId); if (index < 0) invalid("The annotation does not exist.");
      if (operation.op === "removeAnnotation") { removeWorkbookCommentSources(workbook, richContent, annotations[index]!); annotations.splice(index, 1); return; }
      const previous = annotations[index]!, value = parseContract("annotation", { ...previous, ...operation.changes });
      if (previous.status === "needs_reanchor" && value.status !== "needs_reanchor" && !operation.changes.anchor) invalid("Reactivating an annotation requires an explicit anchor.");
      if (operation.changes.anchor && (value.anchor.kind !== "cells" || !baseSheets.has(value.anchor.sheetId))) invalid("Workbook anchors require an existing worksheet.");
      if (operation.changes.anchor && value.anchor.kind === "cells") {
        reanchorWorkbookCommentSources(workbook, value);
      }
      if (value.status !== "needs_reanchor") {
        const binding = workbook.comments?.find(comment => comment.annotationId === value.annotationId);
        if (binding) { delete binding.recoveryStatus; delete binding.recoveryReasons; }
      } else if (previous.status !== "needs_reanchor") {
        const binding = workbook.comments?.find(comment => comment.annotationId === value.annotationId); if (binding) binding.recoveryStatus = previous.status;
      }
      annotations[index] = value;
    }
    let operationCount = 0, changedAxis = false, changedDefinition = false;
    yield { kind: "progress", value: { phase: "operations", completed: 0, total: transaction.operations.length } };
    for (const operation of transaction.operations) {
      let layoutBefore: import("./workbook-layout.js").WorkbookMeasuredLayout | undefined;
      let boundsBefore: ReturnType<typeof workbookLayoutBounds> | undefined;
      const copySheets = operation.op === "copyWorkbookComments" ? workbookCommentCopySheets(workbook, annotations, operation) : [];
      if (copySheets.length && !measureLayout) invalid("Copying positioned notes requires a workbook layout measurer.");
      if (measureLayout) {
        const owners = (yield* workbookLayoutOwners(workbook, annotations))!;
        if (owners.length || copySheets.length) {
          const model = (yield* projectLayoutModel())!;
          layoutBefore = (yield* measureWorkbookCandidate(model, [...new Set([...owners.map(owner => owner.sheetId), ...copySheets])], measureLayout, limits.maxNodes, context))!;
          boundsBefore = workbookLayoutBounds(owners, layoutBefore);
        }
      }
      if ((operation.op === "updateAnnotation" || operation.op === "updateWorkbookComment") && (operation.changes.anchor !== undefined || operation.changes.status !== undefined))
        commentMapping.explicit(workbook, annotations, mapping, operation.annotationId);
      if (operation.op === "copyWorkbookComments") {
        const plan = workbookCommentCopyPlan(operation);
        if (!plan.unchanged) for (const annotation of annotations) if (plan.affected(annotation, operation.targetSheetId, plan.footprint)) commentMapping.explicit(workbook, annotations, mapping, annotation.annotationId);
      }
      if (operation.op === "editWorkbookAxis" || operation.op === "reorderWorkbookRows" || operation.op === "moveWorkbookRange") commentMapping.capture(workbook, annotations, mapping.length);
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
      } else if (isWorkbookStyleOperation(operation)) {
        for (const value of editWorkbookStyle(workbook, cells, styles, operation, allocate)) yield { kind: "progress", value: { phase: "cells", ...value } };
      } else if (operation.op === "setWorkbookNumberFormat") {
        if ("code" in operation.format) compileNumberFormat(operation.format.code);
        const index = workbook.numberFormats.findIndex(format => format.formatId === operation.format.formatId);
        if (index < 0) workbook.numberFormats.push(operation.format); else workbook.numberFormats[index] = operation.format;
      } else if (operation.op === "setWorkbookHyperlink" || operation.op === "replaceWorkbookHyperlinks") {
        const sheet = workbook.sheets.find(item => item.sheetId === operation.sheetId) ?? invalid("An earlier operation removed the worksheet.");
        if (operation.op === "setWorkbookHyperlink") editWorkbookHyperlink(sheet, operation, limits.maxNodes);
        else for (const value of replaceWorkbookHyperlinks(sheet, operation, limits.maxNodes)) yield { kind: "progress", value: { phase: "cells", ...value } };
        hyperlinkSheets.add(sheet.sheetId);
      } else if (operation.op === "replaceWorkbookRules") {
        const sheet = workbook.sheets.find(item => item.sheetId === operation.sheetId) ?? invalid("An earlier operation removed the worksheet.");
        for (const value of replaceWorkbookRules(sheet, operation, limits.maxNodes)) yield { kind: "progress", value: { phase: "cells", ...value } };
        changedDefinition = true;
      } else if (operation.op === "clearWorkbookRanges") {
        if (!workbook.sheets.some(sheet => sheet.sheetId === operation.sheetId)) invalid("An earlier operation removed the worksheet.");
        for (const value of cells.clearSteps(operation.sheetId, operation.ranges, operation.mode)) yield { kind: "progress", value: { phase: "cells", ...value } };
      } else if (operation.op === "mergeWorkbookCells" || operation.op === "unmergeWorkbookCells") {
        const sheet = workbook.sheets.find(sheet => sheet.sheetId === operation.sheetId) ?? invalid("An earlier operation removed the worksheet.");
        for (const value of editWorkbookMerge(sheet, cells, operation)) yield { kind: "progress", value: { phase: "cells", ...value } };
        mergeSheets.add(sheet.sheetId);
      }
      else if (isWorkbookDefinitionOperation(operation)) {
        for (const value of editWorkbookDefinition(workbook, cells, operation, allocate)) yield { kind: "progress", value: { phase: "cells", ...value } };
        changedDefinition = true;
      }
      else if (isWorkbookDrawingOperation(operation)) editWorkbookDrawing(workbook, richContent, base, operation, allocate, styles);
      else if (isWorkbookCommentOperation(operation)) editWorkbookComment(workbook, annotations, richContent, base, operation, allocate);
      else if (operation.op === "copyWorkbookComments") {
        for (const value of copyWorkbookComments(workbook, annotations, richContent, base, operation, allocate, styles, layoutBefore)) yield { kind: "progress", value: { phase: "cells", ...value } };
      }
      else if (operation.op === "moveWorkbookRange") {
        for (const value of moveWorkbookRange(workbook, cells, annotations, operation, mapping)) yield { kind: "progress", value: { phase: "cells", ...value } };
        changedAxis = true;
        for (const sheet of workbook.sheets) { mergeSheets.add(sheet.sheetId); hyperlinkSheets.add(sheet.sheetId); }
      }
      else if (operation.op === "reorderWorkbookRows") {
        for (const value of reorderWorkbookRows(workbook, cells, annotations, operation, mapping)) yield { kind: "progress", value: { phase: "cells", ...value } };
        changedAxis = true; mergeSheets.add(operation.sheetId); hyperlinkSheets.add(operation.sheetId);
      }
      else if (operation.op === "editWorkbookAxis") {
        for (const value of editWorkbookAxis(workbook, cells, annotations, operation, mapping, allocate, !!measureLayout)) yield { kind: "progress", value: { phase: "cells", ...value } };
        changedAxis = true;
      }
      else if (operation.op === "copyWorkbookSheet") {
        for (const value of copyWorkbookSheet(workbook, cells, richContent, annotations, base, operation, allocate, styles)) yield { kind: "progress", value: { phase: "cells", ...value } };
      }
      else if (operation.op === "setWorkbookSettings") {
        const { defaultTableStyle, ...changes } = operation.changes; Object.assign(workbook, changes);
        if (defaultTableStyle === null) delete workbook.defaultTableStyle;
        else if (defaultTableStyle !== undefined) workbook.defaultTableStyle = defaultTableStyle;
      }
      else editWorkbookSheet(workbook, cells, richContent, annotations, operation, allocate, mapping);
      if (measureLayout && boundsBefore && layoutBefore) {
        const owners = (yield* workbookLayoutOwners(workbook, annotations))!.filter(owner => boundsBefore!.has(owner.key));
        if (owners.length || copySheets.length) {
          const model = (yield* projectLayoutModel())!;
          const layoutAfter = (yield* measureWorkbookCandidate(model, [...new Set([...owners.map(owner => owner.sheetId), ...copySheets])], measureLayout, limits.maxNodes, context))!;
          if (layoutAfter.fontEnvironment !== layoutBefore.fontEnvironment) invalid("Available workbook fonts changed during candidate layout measurement.");
          for (const value of compensateWorkbookLayout(owners, boundsBefore, layoutAfter, operation)) yield { kind: "progress", value: { phase: "validation", ...value } };
        }
      }
      yield { kind: "progress", value: { phase: "operations", completed: ++operationCount, total: transaction.operations.length } };
    }
    for (const value of commentMapping.finish(workbook, annotations, mapping)) yield { kind: "progress", value: { phase: "cells", ...value } };
    if (changedAxis || mapping.some(step => step.kind === "move" || step.kind === "sort")) restoreWorkbookAxisSource(workbook, base.snapshot().metadata.workbook, annotations, base.snapshot().annotations);
    if (changedDefinition || changedAxis) restoreWorkbookDefinitionSource(workbook, base.snapshot().metadata.workbook);
    for (const sheet of workbook.sheets.filter(sheet => mergeSheets.has(sheet.sheetId))) {
      const previous = base.snapshot().metadata.workbook.sheets.find(previous => previous.sheetId === sheet.sheetId);
      if (previous && sameWorkbookMerges(previous.merges, sheet.merges)) sheet.merges = previous.merges.map(range => ({ from: { ...range.from }, to: { ...range.to } }));
    }
    for (const sheet of workbook.sheets.filter(sheet => hyperlinkSheets.has(sheet.sheetId))) {
      const previous = base.snapshot().metadata.workbook.sheets.find(previous => previous.sheetId === sheet.sheetId);
      if (previous && sameWorkbookHyperlinks(previous.hyperlinks ?? [], sheet.hyperlinks ?? [])) {
        if (previous.hyperlinks === undefined) delete sheet.hyperlinks;
        else sheet.hyperlinks = previous.hyperlinks.map(link => ({ ...link, range: { from: { ...link.range.from }, to: { ...link.range.to } } }));
      }
    }
    restoreWorkbookCommentSource(workbook, annotations, base.snapshot());
    for (const value of cells.flushSteps()) yield { kind: "progress", value: { phase: "chunks", ...value } };
    richContent.flush();
    if (resolvedIds?.length) {
      const previous = base.snapshot(), original = workbookSource(base)!.files;
      const equal = (left: unknown, right: unknown) => JSON.stringify(left) === JSON.stringify(right);
      const reserved = new Set(["manifest.json", manifest.metadata, manifest.annotations]);
      const sourceChanged = !equal(metadata, previous.metadata) || !equal(annotations, previous.annotations) || !equal(manifest, previous.manifest)
        || [...new Set([...original.keys(), ...files.keys()])].some(path => {
          if (reserved.has(path)) return false;
          const a = original.get(path), b = files.get(path);
          return a !== b && (!a || !b || a.byteLength !== b.byteLength || a.some((byte, index) => byte !== b[index]));
        });
      // A no-op proposal cannot claim instructions as completed. Resolution is
      // part of the same candidate only after an actual source/annotation edit.
      if (sourceChanged) for (const value of annotations) if (resolvedIds.includes(value.annotationId) && value.status === "open") value.status = "resolved";
    }
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
  const interactionBudgets = sessionLimits(options.sessionLimits), reviews = new WeakMap<WorkbookCandidateChange, WorkbookProposalReview>();
  let interactionDisposed = false;
  const interactionGuard = (base: WorkbookModel, execution: TransactionPreparationExecution) => (): void => {
    if (interactionDisposed || execution.signal?.aborted) throw new DocumentError("OperationCancelled", "Workbook interaction was cancelled or disposed.");
    if (current() !== base) throw new DocumentError("RevisionConflict", "The workbook changed while preparing its interaction.");
  };
  const discardInteraction = (candidate: WorkbookCandidateChange): void => { if (!interactionDisposed) boundary.lifecycle.discard(candidate); };
  const session = Object.freeze<WorkbookSession>({ kind: "workbook", snapshot: () => current().snapshot(), getNode: id => current().getNode(id),
    getCell: (sheetId, row, column) => current().getCell(sheetId, row, column), readWindow: window => current().readWindow(window),
    getAsset: (path, maxBytes) => current().getAsset(path, maxBytes), resolveReference: (expression, origin) => current().resolveReference(expression, origin),
    serialize: () => current().serialize(), ...boundary.lifecycle,
    prepareAsync: (transaction, preparation, execution) => boundary.prepareAsyncTransaction(transaction, undefined, preparation, execution),
    buildInteractionRequest: input => boundary.execute(() => finishWorkbookInteraction(buildWorkbookRequestSteps(current(), input))),
    async buildInteractionRequestAsync(input, execution) {
      const base = current(), config = workbookInteractionExecution(execution), guard = interactionGuard(base, config);
      return finishWorkbookInteractionAsync(buildWorkbookRequestSteps(base, input, config), config, guard);
    },
    prepareProposal: (request, proposal, preparation) => boundary.execute(() => {
      const base = current(), capture = finishWorkbookInteraction(validateWorkbookProposalSteps(base, request, proposal, preparation, interactionBudgets.maxTransactionBytes));
      const candidate = boundary.prepareTransaction({ protocolVersion: 1, documentId: capture.proposal.documentId, baseRevision: capture.proposal.baseRevision, operations: capture.proposal.operations }, capture.proposal.annotationIds, capture.options);
      try { reviews.set(candidate, finishWorkbookInteraction(reviewWorkbookProposalSteps(base, capture, candidate))); return candidate; }
      catch (error) { discardInteraction(candidate); throw error; }
    }),
    async prepareProposalAsync(request, proposal, preparation, execution) {
      const base = current(), config = workbookInteractionExecution(execution), guard = interactionGuard(base, config);
      const capture = await finishWorkbookInteractionAsync(validateWorkbookProposalSteps(base, request, proposal, preparation, interactionBudgets.maxTransactionBytes, config), config, guard); guard();
      const candidate = await boundary.prepareAsyncTransaction({ protocolVersion: 1, documentId: capture.proposal.documentId, baseRevision: capture.proposal.baseRevision, operations: capture.proposal.operations }, capture.proposal.annotationIds, capture.options, config);
      try { guard(); reviews.set(candidate, await finishWorkbookInteractionAsync(reviewWorkbookProposalSteps(base, capture, candidate), config, guard)); guard(); return candidate; }
      catch (error) { discardInteraction(candidate); throw error; }
    },
    getProposalReview: candidate => { current(); return reviews.get(candidate); },
    calculate: calculator.calculate, calculateSync: calculator.calculateSync,
    getCalculatedCell: (sheetId, row, column) => calculator.getCell({ sheetId, row, column }),
    dispose: () => { boundary.lifecycle.dispose(); interactionDisposed = true; calculator.dispose(); },
  });
  registerWorkbookModelProvider(session, current); return session;
}
