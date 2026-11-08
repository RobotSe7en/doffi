import { AssetDescriptorSchema } from "../contracts/document.js";
import { WorkbookInteractionRangeSchema, WorkbookInteractionRequestSchema, WorkbookProposalOperationNameSchema } from "../contracts/workbook-interaction.js";
import type { WorkbookInteractionCell, WorkbookInteractionContext, WorkbookInteractionRange, WorkbookInteractionRequest } from "../contracts/workbook-interaction.js";
import type { DeepReadonly, TransactionPreparationExecution, TransactionPreparationProgress, WorkbookInteractionRequestInput, WorkbookModel } from "../runtime.js";
import type { WorkbookDocumentMetadata } from "../contracts/workbook.js";
import { DocumentError } from "../errors.js";
import { createOpaqueId } from "../identifiers.js";
import { encode, resolveReference } from "../document/package.js";
import { checkLimit, freeze } from "../document/validation.js";
import { CHUNK_COLUMNS, CHUNK_ROWS } from "../workbook/chunk-index.js";
import { copyWorkbookScalarContext, createWorkbookScalarKernel } from "../workbook/scalar.js";
import type { ScalarResolved } from "../workbook/scalar.js";
import type { WorkbookReferenceArea, WorkbookReferenceLocation } from "../workbook/reference.js";
import { workbookSource } from "../workbook/source.js";
import { WorkbookInteractionBudget, workbookInteractionLimits } from "./workbook-interaction-budget.js";
import { workbookContextContains, workbookContextCovers, workbookContextIntersects, workbookWritableContext } from "./workbook-interaction-scope.js";
import { checkTransactionSize } from "./limits.js";
import { workbookInteractionDependencies } from "./workbook-interaction-dependencies.js";
import { createWorkbookInteractionRuleReader, workbookInteractionExpression } from "./workbook-interaction-rules.js";
import { checkRange } from "../workbook/ranges.js";

export const workbookGlobalProposalOperations: ReadonlySet<string> = new Set([
  "setWorkbookNumberFormat", "setWorkbookStyleClass", "setWorkbookCellStyleDefinition", "setWorkbookCellStyleName", "removeWorkbookCellStyleName",
  "setWorkbookDefinedName", "removeWorkbookDefinedName", "setWorkbookTableStyle", "removeWorkbookTableStyle", "formatWorkbookTableStyleElement", "setWorkbookSettings",
  "insertWorkbookSheet", "copyWorkbookSheet", "removeWorkbookSheet", "moveWorkbookSheet",
]);
type Preparation = Generator<TransactionPreparationProgress, WorkbookInteractionRequest>;
const equal = (left: unknown, right: unknown) => JSON.stringify(left) === JSON.stringify(right);
const cellKey = (at: WorkbookReferenceLocation) => JSON.stringify([at.sheetId, at.row, at.column]);

/** Read only the declared sparse coverage and its complete source dependencies.
 * No serialization, current revision mutation or imported cache substitution. */
export function* buildWorkbookRequestSteps(model: WorkbookModel, input: WorkbookInteractionRequestInput, control: TransactionPreparationExecution = {}, requestId = createOpaqueId(), sharedBudget?: WorkbookInteractionBudget): Preparation {
  const source = workbookSource(model); if (!source) throw new DocumentError("InvalidOperation", "Workbook context requires an owned immutable source.");
  checkTransactionSize(input, source.limits.maxPackageBytes, source.limits.maxDepth + 16);
  if (!input || input.kind !== "workbook" || Object.keys(input).some(key => !["kind", "annotationIds", "scope", "readRanges", "supportedOperations", "assets", "limits", "calculationContext"].includes(key))
    || !Array.isArray(input.annotationIds) || !input.annotationIds.length || new Set(input.annotationIds).size !== input.annotationIds.length
    || input.readRanges !== undefined && !Array.isArray(input.readRanges) || input.assets !== undefined && !Array.isArray(input.assets) || input.supportedOperations !== undefined && !Array.isArray(input.supportedOperations))
    throw new DocumentError("InvalidOperation", "Select a unique nonempty list of current workbook annotations.");
  const limits = workbookInteractionLimits(input.limits);
  if (input.limits?.maxBytes === undefined) limits.maxBytes = Math.min(limits.maxBytes, source.limits.maxPackageBytes);
  const budget = sharedBudget ?? new WorkbookInteractionBudget(limits, source.limits, control.signal), snapshot = model.snapshot(), workbook = snapshot.metadata.workbook;
  checkLimit(limits.maxBytes, source.limits.maxPackageBytes);
  checkLimit(input.annotationIds.length, limits.maxNodes); checkLimit(input.assets?.length ?? 0, limits.maxAssets);
  checkLimit(input.readRanges?.length ?? 0, limits.maxDependencies);
  checkLimit(input.supportedOperations?.length ?? 0, WorkbookProposalOperationNameSchema.options.length);
  if (input.scope?.kind === "ranges" && Array.isArray(input.scope.ranges)) checkLimit(input.scope.ranges.length, limits.maxDependencies);
  if (input.scope?.kind === "tables" && Array.isArray(input.scope.tableIds)) checkLimit(input.scope.tableIds.length, limits.maxNodes);
  if (input.scope && input.scope.kind !== "global" && Array.isArray(input.scope.objectIds)) checkLimit(input.scope.objectIds.length, limits.maxNodes);
  if (input.scope && input.scope.kind !== "global" && Array.isArray(input.scope.newObjectSheetIds)) checkLimit(input.scope.newObjectSheetIds.length, limits.maxNodes);
  const sourceAnnotations = new Map(snapshot.annotations.map(annotation => { budget.step(); return [annotation.annotationId, annotation] as const; }));
  const annotations = input.annotationIds.map(id => {
    budget.step(); const annotation = sourceAnnotations.get(id);
    if (!annotation || annotation.status !== "open" || annotation.baseRevision !== snapshot.manifest.revision) throw new DocumentError("AnchorNotFound", "Workbook requests require open annotations at the current revision.");
    return annotation;
  });
  const writable = workbookWritableContext(model, annotations, input.scope, budget.step), global = writable.scope.kind === "global";
  const providedAssets = (input.assets ?? []).map(asset => AssetDescriptorSchema.parse(asset)), declaredOperations = input.supportedOperations ? [...input.supportedOperations] : undefined;
  const context = copyWorkbookScalarContext(input.calculationContext), dependencies: WorkbookInteractionRange[] = [], queue: WorkbookInteractionRange[] = [...writable.ranges];
  const seenAreas = new Set(writable.ranges.map(range => JSON.stringify(range))), cells = new Map<string, WorkbookInteractionCell>(), current = new Map<string, ScalarResolved>(), usedNames = new Set<string>();
  const sheetIds = new Set(workbook.sheets.map(sheet => sheet.sheetId));
  const requestedReadRanges = (input.readRanges ?? []).map(value => { budget.step(); return WorkbookInteractionRangeSchema.parse(value); });
  function add(area: DeepReadonly<WorkbookReferenceArea>): void {
    budget.step(); const value: WorkbookInteractionRange = { sheetId: area.sheetId, range: { from: { row: area.from.row, column: area.from.column }, to: { row: area.to.row, column: area.to.column } } };
    checkRange(value.range);
    const key = JSON.stringify(value); if (seenAreas.has(key)) return; seenAreas.add(key);
    if (!sheetIds.has(area.sheetId)) throw new DocumentError("InvalidDocument", "A workbook context dependency has no source worksheet.");
    if (workbookContextCovers([...writable.ranges, ...dependencies], value, budget.step)) return;
    budget.dependency(); dependencies.push(value); queue.push(value);
  }
  const evaluator = createWorkbookScalarKernel(model, { maxCells: limits.maxCells, maxSteps: limits.maxSteps, maxCodeUnits: limits.maxBytes }, context, {
    step: budget.step, account: budget.account, reference: add,
    read: at => add({ sheetId: at.sheetId, from: at, to: at }), cachedCell: at => current.get(cellKey(at)),
    evaluated: (at, value) => current.set(cellKey(at), value), name: definition => usedNames.add(definition.nameId),
  });
  const discover = (expression: string, origin: WorkbookReferenceLocation, at = origin) => workbookInteractionExpression(model, expression, origin, at, budget.step, add, definition => usedNames.add(definition.nameId), evaluator);
  const readRules = createWorkbookInteractionRuleReader(model, budget, evaluator, discover, add);
  try {
    budget.take(annotations); budget.take(writable.scope); budget.take(context); budget.take(limits);
    yield { phase: "operations", completed: 0, total: 0 }; budget.step();
    for (const entry of requestedReadRanges) add({ sheetId: entry.sheetId, ...entry.range });
    const formulaDependencies = global ? undefined : yield* workbookInteractionDependencies(model, writable.ranges, context, budget, add);
    const selectedObjects = new Set(writable.objectIds), charts = (workbook.charts ?? []).filter(chart => selectedObjects.has(chart.nodeId));
    for (const chart of charts) {
      const object = workbook.objects!.find(object => object.nodeId === chart.nodeId)!;
      const references = [chart.categories, ...chart.series.flatMap(series => [series.name?.kind === "reference" ? series.name.reference : undefined, series.values]), chart.titleReference, chart.categoryAxisTitle?.reference, chart.valueAxisTitle?.reference];
      for (const reference of references) {
        budget.step(); if (!reference) continue;
        if ("expression" in reference) {
          const at = { sheetId: object.sheetId, row: 0, column: 0 }; discover(reference.expression, at);
          const bound = evaluator.referenceAreas(reference.expression, at); if (bound.kind === "ranges") for (const area of bound.areas) add(area);
        } else for (const _point of reference.points) budget.node();
      }
    }
    let scanned = 0;
    for (let cursor = 0; cursor < queue.length; cursor++) {
      const entry = queue[cursor]!, area = entry.range;
      yield* readRules(entry);
      for (const descriptor of snapshot.manifest.chunks) {
        budget.step(); if (descriptor.sheetId !== entry.sheetId || !workbookContextIntersects(area, { from: descriptor, to: { row: descriptor.row + CHUNK_ROWS - 1, column: descriptor.column + CHUNK_COLUMNS - 1 } })) continue;
        const row = Math.max(area.from.row, descriptor.row), column = Math.max(area.from.column, descriptor.column);
        const values = source.chunks.get(descriptor.path)!.readWindow({ row, column, rows: Math.min(area.to.row, descriptor.row + CHUNK_ROWS - 1) - row + 1, columns: Math.min(area.to.column, descriptor.column + CHUNK_COLUMNS - 1) - column + 1 });
        for (const indexed of values) {
          budget.step(); const at = { sheetId: entry.sheetId, row: indexed.row, column: indexed.column }, key = cellKey(at); if (cells.has(key)) continue;
          budget.cell(); const sourceCell = model.getCell(at.sheetId, at.row, at.column)!;
          const editable = workbookContextContains(writable.ranges, at.sheetId, at.row, at.column, budget.step);
          let value: WorkbookInteractionCell["current"];
          if (sourceCell.formula) {
            const recorded = formulaDependencies?.get(key);
            if (recorded) { for (const area of recorded.areas) add(area); for (const id of recorded.nameIds) usedNames.add(id); }
            else discover(sourceCell.formula.expression, at);
            const result = evaluator.cell(at, { ...(control.signal ? { signal: control.signal } : {}) }); value = result.kind === "value" ? { kind: "value", value: result.value } : { kind: "unavailable", reason: result.reason, ...(result.functionName ? { functionName: result.functionName } : {}) };
          }
          else value = { kind: "value", value: sourceCell.value };
          const cell: WorkbookInteractionCell = { ...sourceCell, access: editable ? "editable" : "read-only", current: value };
          budget.take(cell); budget.html(cell.html, true); cells.set(key, cell);
          if (++scanned % 128 === 0) { yield { phase: "cells", ...budget.progress() }; budget.step(); }
        }
      }
      yield { phase: "chunks", completed: cursor + 1, total: queue.length }; budget.step();
    }
    const coverage = [...writable.ranges, ...dependencies], sheets = new Set([...coverage.map(entry => entry.sheetId), ...writable.newObjectSheetIds, workbook.activeSheetId]);
    for (const name of workbook.names ?? []) if (usedNames.has(name.nameId) && name.sheetId) sheets.add(name.sheetId);
    const touches = (sheetId: string, range: DeepReadonly<WorkbookInteractionRange["range"]>) => coverage.some(entry => { budget.step(); return entry.sheetId === sheetId && workbookContextIntersects(entry.range, range); });
    const supportingAnnotations = snapshot.annotations.filter(annotation => global || annotation.anchor.kind === "cells" && annotation.anchor.ranges.some(range => touches(annotation.anchor.kind === "cells" ? annotation.anchor.sheetId : "", range)));
    const writableAnnotationIds = supportingAnnotations.filter(annotation => global || annotation.anchor.kind === "cells" && annotation.anchor.ranges.every(range => workbookContextCovers(writable.ranges, { sheetId: annotation.anchor.kind === "cells" ? annotation.anchor.sheetId : "", range }, budget.step))).map(annotation => annotation.annotationId);
    const editableAnnotations = new Set(writableAnnotationIds), annotationIds = new Set(supportingAnnotations.map(annotation => annotation.annotationId)), bodyAccess = new Map(supportingAnnotations.flatMap(annotation => [annotation, ...annotation.replies ?? []].flatMap(body => body.bodyNodeId ? [[body.bodyNodeId, editableAnnotations.has(annotation.annotationId)] as const] : []))), bodyIds = new Set(bodyAccess.keys());
    const objectIds = new Set(writable.objectIds), objects = (workbook.objects ?? []).filter(object => objectIds.has(object.nodeId)), chartIds = new Set(objects.flatMap(object => object.kind === "chart" ? [object.chartId] : [])), objectSheets = new Set([...objects.map(object => object.sheetId), ...writable.newObjectSheetIds]);
    for (const sheetId of objectSheets) sheets.add(sheetId);
    const projected: DeepReadonly<WorkbookDocumentMetadata> = { ...snapshot.metadata, workbook: { ...workbook,
      sheets: workbook.sheets.filter(sheet => global || sheets.has(sheet.sheetId)).map(sheet => { const { print, autoFilter, ...sourceSheet } = sheet; return { ...sourceSheet,
        rowDimensions: sheet.rowDimensions.filter(span => global || objectSheets.has(sheet.sheetId) || coverage.some(entry => entry.sheetId === sheet.sheetId && entry.range.from.row <= span.to && span.from <= entry.range.to.row)),
        columnDimensions: sheet.columnDimensions.filter(span => global || objectSheets.has(sheet.sheetId) || coverage.some(entry => entry.sheetId === sheet.sheetId && entry.range.from.column <= span.to && span.from <= entry.range.to.column)),
        ...(global && print ? { print } : {}), ...(autoFilter && (global || touches(sheet.sheetId, autoFilter.range)) ? { autoFilter } : {}),
        merges: sheet.merges.filter(range => global || touches(sheet.sheetId, range)),
        ...(sheet.tables ? { tables: sheet.tables.filter(table => global || touches(sheet.sheetId, table.range)) } : {}),
        ...(sheet.conditionalRules ? { conditionalRules: sheet.conditionalRules.filter(rule => global || rule.ranges.some(range => touches(sheet.sheetId, range))) } : {}),
        ...(sheet.dataValidations ? { dataValidations: sheet.dataValidations.filter(rule => global || rule.ranges.some(range => touches(sheet.sheetId, range))) } : {}),
        ...(sheet.hyperlinks ? { hyperlinks: sheet.hyperlinks.filter(link => global || touches(sheet.sheetId, link.range)) } : {}),
      }; }), objects, charts: (workbook.charts ?? []).filter(chart => chartIds.has(chart.chartId)),
      names: (workbook.names ?? []).filter(name => global || usedNames.has(name.nameId)),
      comments: (workbook.comments ?? []).filter(comment => annotationIds.has(comment.annotationId)), stories: (workbook.stories ?? []).filter(story => bodyIds.has(story.rootNodeId)),
      objectHyperlinks: (workbook.objectHyperlinks ?? []).filter(link => objectIds.has(link.nodeId)),
    } };
    const blocks: WorkbookInteractionContext["blocks"] = [], assets = new Set<string>(), blockIds = [...objectIds, ...bodyIds], emitted = new Set<string>(), blockNodes = new Set<string>();
    for (const id of blockIds) {
      budget.step(); if (emitted.has(id)) continue; emitted.add(id); const node = model.getNode(id);
      if (!node) throw new DocumentError("AnchorNotFound", "A workbook context body or object no longer exists.");
      const block = { nodeId: id, html: node.html, access: objectIds.has(id) || bodyAccess.get(id) ? "editable" as const : "read-only" as const }; budget.take(block); budget.html(block.html); blocks.push(block);
      const pending = [id];
      while (pending.length) { budget.step(); const current = source.content?.peekNode(pending.pop()!); if (!current) continue; blockNodes.add(current.nodeId); pending.push(...current.childNodeIds); if (current.tagName === "img") assets.add(resolveReference(current.attributes.src!, snapshot.content!.path)); }
    }
    if (new Set(providedAssets.map(asset => asset.path)).size !== providedAssets.length) throw new DocumentError("InvalidOperation", "Workbook context resources must have unique paths.");
    const descriptors = new Map(snapshot.manifest.assets.filter(asset => assets.has(asset.path)).map(asset => [asset.path, asset]));
    for (const descriptor of providedAssets) { budget.step(); const original = snapshot.manifest.assets.find(asset => asset.path === descriptor.path); if (original && !equal(original, descriptor)) throw new DocumentError("InvalidOperation", "Workbook context resources cannot replace immutable source paths."); descriptors.set(descriptor.path, descriptor); }
    checkLimit(descriptors.size, limits.maxAssets);
    const supportedOperations = declaredOperations ?? WorkbookProposalOperationNameSchema.options.filter(name => global || !workbookGlobalProposalOperations.has(name));
    if (!Array.isArray(supportedOperations) || new Set(supportedOperations).size !== supportedOperations.length || supportedOperations.some(name => !WorkbookProposalOperationNameSchema.safeParse(name).success || !global && workbookGlobalProposalOperations.has(name))) throw new DocumentError("InvalidOperation", "Workbook proposal operations must match the declared scope.");
    const authors = new Set<string>();
    for (const comment of projected.workbook.comments ?? []) { if (comment.authorId) authors.add(comment.authorId); if (comment.format === "modern") { for (const mention of comment.mentions ?? []) authors.add(mention.authorId); for (const reply of comment.replies) { if (reply.authorId) authors.add(reply.authorId); for (const mention of reply.mentions ?? []) authors.add(mention.authorId); } } }
    const metadata: DeepReadonly<WorkbookDocumentMetadata> = { ...projected, workbook: { ...projected.workbook,
      commentAuthors: (workbook.commentAuthors ?? []).filter(author => authors.has(author.authorId)),
      objectHyperlinks: (workbook.objectHyperlinks ?? []).filter(link => blockNodes.has(link.nodeId)),
    } };
    budget.take(metadata); budget.take(supportingAnnotations); budget.take(snapshot.styles); budget.take([...descriptors.values()]); budget.take(providedAssets); budget.take(supportedOperations);
    const sheetOrder = new Map(workbook.sheets.map((sheet, index) => [sheet.sheetId, index]));
    const orderedCells = [...cells.values()].sort((left, right) => { budget.step(); return sheetOrder.get(left.sheetId)! - sheetOrder.get(right.sheetId)! || left.row - right.row || left.column - right.column; });
    const request = { protocolVersion: 1 as const, requestId, documentId: snapshot.manifest.documentId, baseRevision: snapshot.manifest.revision, annotations,
      context: { formatVersion: 1 as const, kind: "workbook" as const, scope: writable.scope, writableRanges: writable.ranges, writableObjectIds: writable.objectIds, writableNewObjectSheetIds: writable.newObjectSheetIds, writableAnnotationIds, requestedReadRanges, dependencies,
        cells: orderedCells, annotations: supportingAnnotations, blocks, styles: snapshot.styles, metadata, assets: [...descriptors.values()], providedAssets, calculationContext: context, limits }, supportedOperations };
    checkTransactionSize(request, Math.min(Number.MAX_SAFE_INTEGER, limits.maxBytes * 8), source.limits.maxDepth + 16);
    checkLimit(encode(JSON.stringify(request)).byteLength, limits.maxBytes); budget.step();
    return freeze(WorkbookInteractionRequestSchema.parse(request));
  } finally { evaluator.dispose(); current.clear(); }
}
