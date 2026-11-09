import { generate, ident, parse, walk } from "css-tree";
import type { DeepReadonly, TransactionPrepareOptions, WorkbookCandidateChange, WorkbookModel } from "../runtime.js";
import type { WorkbookEditProposal, WorkbookInteractionRequest } from "../contracts/workbook-interaction.js";
import type { WorkbookConditionalRule, WorkbookDataValidation, WorkbookHyperlink, WorkbookRange } from "../contracts/workbook.js";
import { DocumentError } from "../errors.js";
import { resolveReference } from "../document/package.js";
import { workbookSource } from "../workbook/source.js";
import { workbookGlobalProposalOperations } from "./workbook-interaction-context.js";
import type { WorkbookInteractionBudget } from "./workbook-interaction-budget.js";
import { workbookContextCovers, workbookContextIntersects } from "./workbook-interaction-scope.js";
import { subtractWorkbookRange } from "../workbook/range-mapping.js";
import { sameWorkbookHyperlinks } from "./workbook-hyperlinks.js";
import { sameWorkbookRangeCoverage } from "../workbook/ranges.js";
import { workbookCommentCopyPlan } from "./workbook-copy-comments.js";
import { planWorkbookRegionCopy } from "./workbook-copy-region.js";
import { workbookObjectCopies } from "./workbook-copy-objects.js";

const equal = (a: unknown, b: unknown) => JSON.stringify(a) === JSON.stringify(b);
function invalid(message: string): never { throw new DocumentError("InvalidProposal", message); }
const fullSheet: WorkbookRange = { from: { row: 0, column: 0 }, to: { row: 1_048_575, column: 16_383 } };
export function validateWorkbookProposalTargets(model: WorkbookModel, request: WorkbookInteractionRequest, proposal: WorkbookEditProposal, options: TransactionPrepareOptions | undefined, budget: WorkbookInteractionBudget): void {
  const context = request.context, global = context.scope.kind === "global", workbook = model.snapshot().metadata.workbook;
  const proposedRules = new Map<string, Extract<WorkbookEditProposal["operations"][number], { op: "setWorkbookConditionalRule" }>["rule"]>();
  for (const op of proposal.operations) { budget.step(); if (op.op === "setWorkbookConditionalRule") proposedRules.set(JSON.stringify([op.sheetId, op.rule.ruleId]), op.rule); }
  const objects = new Set(context.writableObjectIds), annotations = new Set(context.writableAnnotationIds), supported = new Set(request.supportedOperations);
  const objectSheets = new Map<string, string>();
  if (proposal.operations.some(op => op.op === "duplicateWorkbookObjects")) for (const object of context.metadata.workbook.objects ?? []) { budget.step(); objectSheets.set(object.nodeId, object.sheetId); }
  const requireRange = (sheetId: string, range: DeepReadonly<WorkbookRange>): void => { if (!workbookContextCovers(context.writableRanges, { sheetId, range }, budget.step)) invalid("A workbook proposal targets cells outside its declared write scope."); };
  const anchor = (value: { readonly kind: string; readonly sheetId: string; readonly ranges: readonly DeepReadonly<WorkbookRange>[] }): void => {
    if (value.kind !== "cells") invalid("Workbook proposal comments require worksheet anchors.");
    for (const range of value.ranges) requireRange(value.sheetId, range);
  };
  for (const resource of options?.assets ?? []) { budget.step(); if (!resource?.descriptor || !context.providedAssets.some(asset => equal(asset, resource.descriptor))) invalid("Workbook proposal bytes lack a matching host-authorized request descriptor."); }
  if (options?.calculationContext !== undefined && !equal(options.calculationContext, context.calculationContext)) invalid("A proposal cannot replace its frozen calculation context.");
  if (new Set(proposal.annotationIds).size !== proposal.annotationIds.length || proposal.annotationIds.some(id => !request.annotations.some(annotation => annotation.annotationId === id))) invalid("A workbook proposal claims annotations outside this request.");
  for (const op of proposal.operations) {
    budget.step(); if (!supported.has(op.op) || !global && workbookGlobalProposalOperations.has(op.op)) invalid("A workbook proposal uses an operation outside its declared scope.");
    for (const id of [...("nodeId" in op ? [op.nodeId] : []), ...("nodeIds" in op ? op.nodeIds : [])]) if (!objects.has(id)) invalid("A workbook proposal targets a drawing outside its declared write scope.");
    if ("annotationId" in op && !annotations.has(op.annotationId)) invalid("A workbook proposal targets a read-only or undisclosed comment.");
    if (op.op === "addWorkbookComment") anchor(op.anchor);
    if (op.op === "updateWorkbookComment" && op.changes.anchor) anchor(op.changes.anchor);
    if (op.op === "copyWorkbookComments" && !global) {
      const plan = workbookCommentCopyPlan(op), disclosed = new Set(context.annotations.map(annotation => annotation.annotationId));
      for (const range of plan.source) requireRange(op.sheetId, range);
      for (const range of plan.footprint) requireRange(op.targetSheetId, range);
      for (const annotation of model.snapshot().annotations) if (plan.affected(annotation, op.sheetId, plan.source) && !disclosed.has(annotation.annotationId)) invalid("A workbook proposal cannot copy an undisclosed comment body.");
    }
    if (global) continue;
    switch (op.op) {
      case "copyWorkbookObjects": {
        const plan = planWorkbookRegionCopy(op);
        const readable = [...context.writableRanges, ...context.requestedReadRanges, ...context.dependencies];
        for (const range of plan.source) if (!workbookContextCovers(readable, { sheetId: op.sheetId, range }, budget.step)) invalid("A workbook proposal cannot copy drawings from undisclosed source cells.");
        for (const range of plan.footprint) requireRange(op.targetSheetId, range);
        if (!context.writableNewObjectSheetIds.includes(op.targetSheetId)) invalid("Copying workbook drawings requires explicit destination creation scope.");
        break;
      }
      case "copyWorkbookComments": break; // Exact read/write masks and disclosed bodies were checked above.
      case "setWorkbookCells": case "setWorkbookCellFormats": for (const at of op.cells) requireRange(op.sheetId, { from: at, to: at }); break;
      case "formatWorkbookText": requireRange(op.sheetId, { from: op.at, to: op.at }); break;
      case "formatWorkbookConditionalRule": {
        const original = workbook.sheets.find(sheet => sheet.sheetId === op.sheetId)?.conditionalRules?.find(rule => rule.ruleId === op.ruleId), proposed = proposedRules.get(JSON.stringify([op.sheetId, op.ruleId]));
        for (const range of [...original?.ranges ?? [], ...proposed?.ranges ?? []]) requireRange(op.sheetId, range); break;
      }
      case "formatWorkbookRanges": case "clearWorkbookRanges": case "replaceWorkbookRules": for (const range of op.ranges) requireRange(op.sheetId, range); break;
      case "setWorkbookHyperlink": case "replaceWorkbookHyperlinks": case "mergeWorkbookCells": case "unmergeWorkbookCells": case "reorderWorkbookRows": requireRange(op.sheetId, op.range); break;
      case "moveWorkbookRange": requireRange(op.sheetId, op.range); requireRange(op.targetSheetId, { from: op.target, to: { row: op.target.row + op.range.to.row - op.range.from.row, column: op.target.column + op.range.to.column - op.range.from.column } }); break;
      case "editWorkbookAxis": requireRange(op.sheetId, { from: { row: op.axis === "row" ? op.index : 0, column: op.axis === "column" ? op.index : 0 }, to: fullSheet.to }); break;
      case "updateWorkbookSheet": requireRange(op.sheetId, fullSheet); break;
      case "setWorkbookAutoFilter": {
        const original = workbook.sheets.find(sheet => sheet.sheetId === op.sheetId)?.autoFilter;
        for (const value of [original, op.filter]) if (value) requireRange(op.sheetId, value.range); break;
      }
      case "setWorkbookTable": case "removeWorkbookTable": {
        const id = op.op === "setWorkbookTable" ? op.table.tableId : op.tableId, original = workbook.sheets.find(sheet => sheet.sheetId === op.sheetId)?.tables?.find(table => table.tableId === id);
        if (original) requireRange(op.sheetId, original.range); if (op.op === "setWorkbookTable") requireRange(op.sheetId, op.table.range); break;
      }
      case "setWorkbookConditionalRule": case "removeWorkbookConditionalRule": case "setWorkbookDataValidation": case "removeWorkbookDataValidation": {
        const conditional = op.op === "setWorkbookConditionalRule" || op.op === "removeWorkbookConditionalRule", id = "rule" in op ? op.rule.ruleId : op.ruleId;
        const sheet = workbook.sheets.find(sheet => sheet.sheetId === op.sheetId), original = (conditional ? sheet?.conditionalRules : sheet?.dataValidations)?.find(rule => rule.ruleId === id);
        for (const range of [...original?.ranges ?? [], ...("rule" in op ? op.rule.ranges : [])]) requireRange(op.sheetId, range); break;
      }
      case "insertWorkbookObject": case "insertWorkbookChart": if (!context.writableNewObjectSheetIds.includes(op.sheetId)) invalid("New workbook drawings require explicit worksheet creation scope."); break;
      case "duplicateWorkbookObjects":
        for (const id of op.nodeIds) {
          budget.step(); const sheetId = objectSheets.get(id);
          if (!sheetId || !context.writableNewObjectSheetIds.includes(sheetId)) invalid("Duplicating workbook drawings requires explicit worksheet creation scope.");
        }
        break;
      default: // Remaining operations target a selected drawing or comment;
        // the checks above run even when their typed operation would be a no-op.
        if (!("nodeId" in op) && !("nodeIds" in op) && !("annotationId" in op) && op.op !== "addWorkbookComment") invalid("This workbook operation requires explicit global scope.");
    }
  }
}

/** Compare the actual candidate's source complement. Typed target validation
 * cannot by itself catch formula rewrites or indirect metadata changes. */
export function validateWorkbookProposalSource(model: WorkbookModel, request: WorkbookInteractionRequest, candidate: WorkbookCandidateChange, budget: WorkbookInteractionBudget): void {
  const before = model.snapshot(), after = candidate.preview, context = request.context, previous = workbookSource(model)!, next = workbookSource(candidate.model)!;
  const assets = new Set(context.assets.map(asset => asset.path));
  for (const change of candidate.changes) {
    budget.step(); const node = next.content?.peekNode(change.nodeId);
    if (node?.tagName === "img" && !assets.has(resolveReference(node.attributes.src!, after.content!.path))) invalid("The workbook candidate introduces an image outside its frozen resources.");
  }
  if (context.scope.kind === "global") return;
  for (const entry of candidate.changedRanges ?? []) if (!workbookContextCovers(context.writableRanges, entry, budget.step)) invalid("The workbook candidate indirectly rewrites cell source outside its frozen write scope.");
  const objectIds = new Set(context.writableObjectIds), annotationIds = new Set(context.writableAnnotationIds);
  const originalObjects = new Set(before.metadata.workbook.objects?.map(object => object.nodeId));
  const copied = workbookObjectCopies(candidate.model);
  for (const object of after.metadata.workbook.objects ?? []) {
    budget.step(); let source = copied?.get(object.nodeId); const seen = new Set<string>();
    while (source !== undefined && copied?.has(source)) {
      budget.step(); if (seen.has(source)) invalid("A workbook drawing copy has invalid source provenance."); seen.add(source); source = copied.get(source);
    }
    if (source !== undefined && originalObjects.has(source) && !context.writableObjectIds.includes(source)) invalid("A workbook proposal cannot copy an undisclosed drawing source.");
  }
  for (const object of after.metadata.workbook.objects ?? []) if (!originalObjects.has(object.nodeId) && context.writableNewObjectSheetIds.includes(object.sheetId)) objectIds.add(object.nodeId);
  const originalAnnotations = new Set(before.annotations.map(annotation => annotation.annotationId));
  for (const annotation of after.annotations) if (!originalAnnotations.has(annotation.annotationId) && annotation.anchor.kind === "cells" && annotation.anchor.ranges.every(range => workbookContextCovers(context.writableRanges, { sheetId: annotation.anchor.kind === "cells" ? annotation.anchor.sheetId : "", range }, budget.step))) annotationIds.add(annotation.annotationId);
  const roots = new Set<string>(objectIds);
  for (const snapshot of [before, after]) for (const annotation of snapshot.annotations) if (annotationIds.has(annotation.annotationId)) for (const body of [annotation, ...annotation.replies ?? []]) if (body.bodyNodeId) roots.add(body.bodyNodeId);
  const owned = new Set<string>();
  for (const source of [previous, next]) for (const root of roots) {
    const pending = [root]; while (pending.length) { budget.step(); const id = pending.pop()!, node = source.content?.peekNode(id); if (node) { owned.add(id); pending.push(...node.childNodeIds); } }
  }
  for (const change of candidate.changes) {
    budget.step(); if (owned.has(change.nodeId)) continue;
    const a = previous.content?.peekNode(change.nodeId), b = next.content?.peekNode(change.nodeId);
    const projection = (node: typeof a, source: typeof previous) => node && ({ ...node, childNodeIds: node.childNodeIds.filter(id => !owned.has(id)), text: source.content?.directText(node.nodeId) });
    // The first authorized object/comment creates its inert common container.
    if (!previous.content && b?.tagName === "article" && b.parentNodeId === null && b.attributes["data-office-kind"] === "workbook-content" && b.childNodeIds.every(id => owned.has(id)) && !next.content?.directText(b.nodeId)) continue;
    if (!equal(projection(a, previous), projection(b, next))) invalid("The workbook candidate changes rich source outside its selected objects or comment bodies.");
  }
  const freshClasses = new Set([...next.classes].filter(name => !previous.classes.has(name)));
  const freshOwnedNodes = new Set<string>(); for (const id of owned) { budget.step(); if (!previous.content?.peekNode(id)) freshOwnedNodes.add(id); }
  const cssProjection = (css: string): string => {
    const ast = parse(css, { context: "stylesheet", parseCustomProperty: true });
    if (ast.type !== "StyleSheet") invalid("Invalid workbook candidate stylesheet.");
    ast.children.forEach((rule, item, list) => {
      budget.step(); if (rule.type !== "Rule" || rule.prelude.type !== "SelectorList") return;
      rule.prelude.children.forEach((selector, at, selectors) => { let fresh = false; walk(selector, node => {
        budget.step(); if (node.type === "ClassSelector" && freshClasses.has(ident.decode(node.name))) fresh = true;
        if (node.type === "AttributeSelector" && ident.decode(node.name.name) === "data-node-id" && node.matcher === "=" && node.value
          && freshOwnedNodes.has(node.value.type === "String" ? node.value.value : ident.decode(node.value.name))) fresh = true;
      }); if (fresh) selectors.remove(at); });
      if (!rule.prelude.children.size) list.remove(item);
    });
    return generate(ast);
  };
  if (!equal(before.styles.map(style => ({ path: style.path, css: cssProjection(style.css) })), after.styles.map(style => ({ path: style.path, css: cssProjection(style.css) })))) invalid("A scoped workbook proposal changes shared source styles.");
  const project = (snapshot: typeof before) => {
    const workbook = snapshot.metadata.workbook;
    const rangeOwned = (sheetId: string, range: DeepReadonly<WorkbookRange>) => workbookContextCovers(context.writableRanges, { sheetId, range }, budget.step);
    const chartIds = new Set(workbook.objects?.flatMap(object => objectIds.has(object.nodeId) && object.kind === "chart" ? [object.chartId] : []));
    return { ...snapshot.metadata, workbook: { ...workbook,
      sheets: workbook.sheets.map(sheet => rangeOwned(sheet.sheetId, fullSheet) ? { sheetId: sheet.sheetId } : { ...sheet,
        merges: sheet.merges.filter(range => !rangeOwned(sheet.sheetId, range)),
        tables: (sheet.tables ?? []).filter(table => !rangeOwned(sheet.sheetId, table.range)),
        conditionalRules: undefined,
        dataValidations: undefined,
        hyperlinks: undefined,
        autoFilter: sheet.autoFilter && !rangeOwned(sheet.sheetId, sheet.autoFilter.range) ? sheet.autoFilter : undefined,
      }),
      objects: (workbook.objects ?? []).filter(object => !objectIds.has(object.nodeId)), charts: (workbook.charts ?? []).filter(chart => !chartIds.has(chart.chartId)),
      comments: (workbook.comments ?? []).filter(comment => !annotationIds.has(comment.annotationId)),
      stories: (workbook.stories ?? []).filter(story => !roots.has(story.rootNodeId)), objectHyperlinks: (workbook.objectHyperlinks ?? []).filter(link => !owned.has(link.nodeId)),
      commentAuthors: (workbook.commentAuthors ?? []).filter(author => before.metadata.workbook.commentAuthors?.some(original => original.authorId === author.authorId)),
      cellStyles: workbook.cellStyles && { ...workbook.cellStyles, bindings: workbook.cellStyles.bindings.filter(binding => !freshClasses.has(binding.cssClass)) },
    } };
  };
  if (!equal(project(before), project(after))) invalid("The workbook candidate changes metadata outside its frozen write scope.");
  // A link crossing the write boundary can split without changing any outside
  // target. Compare exact outside coverage and attributes, rather than its
  // original bounding rectangle or the ordering of equivalent fragments.
  const outsideRanges = (sheetId: string, ranges: readonly DeepReadonly<WorkbookRange>[]): DeepReadonly<WorkbookRange>[] => {
    if (workbookContextCovers(context.writableRanges, { sheetId, range: fullSheet }, budget.step)) return [];
    const result: DeepReadonly<WorkbookRange>[] = [];
    for (const range of ranges) {
      budget.step(); let remaining: DeepReadonly<WorkbookRange>[] = [range];
      for (const cover of context.writableRanges) {
        budget.step(); if (cover.sheetId !== sheetId) continue;
        const next: DeepReadonly<WorkbookRange>[] = [];
        for (const range of remaining) { budget.step(); if (workbookContextIntersects(range, cover.range)) next.push(...subtractWorkbookRange(range, cover.range)); else next.push(range); }
        remaining = next; if (!remaining.length) break;
      }
      for (const range of remaining) { budget.step(); result.push(range); }
    }
    return result;
  };
  const outsideLinks = (sheetId: string, links: readonly DeepReadonly<WorkbookHyperlink>[]): DeepReadonly<WorkbookHyperlink>[] => links.flatMap(link => outsideRanges(sheetId, [link.range]).map(range => range === link.range ? link : { ...link, range }));
  const outsideRules = (sheetId: string, rules: readonly DeepReadonly<WorkbookConditionalRule | WorkbookDataValidation>[]) => {
    const result = new Map<string, { template: unknown; ranges: DeepReadonly<WorkbookRange>[] }>();
    for (const rule of rules) {
      budget.step(); const ranges = outsideRanges(sheetId, rule.ranges);
      if (ranges.length) result.set(rule.ruleId, { template: { ...rule, ranges: undefined }, ranges });
    }
    return result;
  };
  for (const sheet of before.metadata.workbook.sheets) {
    const current = after.metadata.workbook.sheets.find(value => value.sheetId === sheet.sheetId);
    if (!current || !sameWorkbookHyperlinks(outsideLinks(sheet.sheetId, sheet.hyperlinks ?? []), outsideLinks(sheet.sheetId, current.hyperlinks ?? []))) invalid("The workbook candidate changes hyperlinks outside its frozen write scope.");
    for (const key of ["conditionalRules", "dataValidations"] as const) {
      const original = outsideRules(sheet.sheetId, sheet[key] ?? []), proposed = outsideRules(sheet.sheetId, current[key] ?? []);
      if (original.size !== proposed.size || [...original].some(([id, value]) => {
        budget.step(); const other = proposed.get(id);
        return !other || !equal(value.template, other.template) || !sameWorkbookRangeCoverage(value.ranges as WorkbookRange[], other.ranges as WorkbookRange[]);
      })) invalid("The workbook candidate changes rules outside its frozen write scope.");
    }
  }
  const annotationSource = (snapshot: typeof before) => snapshot.annotations.filter(annotation => !annotationIds.has(annotation.annotationId));
  if (!equal(annotationSource(before), annotationSource(after))) invalid("The workbook candidate remaps an annotation outside its frozen write scope.");
}
