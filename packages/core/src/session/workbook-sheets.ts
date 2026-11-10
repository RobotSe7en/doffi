import type { Annotation } from "../contracts/annotations.js";
import type { WorkbookMetadata, WorkbookSheet } from "../contracts/workbook.js";
import type { WorkbookOperation } from "../contracts/workbook-operations.js";
import type { WorkbookStructureMapping } from "../runtime.js";
import { DocumentError } from "../errors.js";
import { WorkbookCellEdit } from "./workbook-cells.js";
import { WorkbookContentEdit } from "./workbook-content.js";
import { rewriteWorkbookReferences } from "./workbook-references.js";

function invalid(message: string): never { throw new DocumentError("InvalidOperation", message); }
export function editWorkbookSheet(workbook: WorkbookMetadata, cells: WorkbookCellEdit, content: WorkbookContentEdit,
  annotations: Annotation[], operation: WorkbookOperation, allocate: () => string, mapping: WorkbookStructureMapping[]): void {
  const sheet = (id: string): WorkbookSheet => workbook.sheets.find(sheet => sheet.sheetId === id) ?? invalid("An earlier operation removed the worksheet.");
  const checkName = (name: string, target?: WorkbookSheet): void => {
    if (name.startsWith("'") || name.endsWith("'") || workbook.sheets.some(sheet => sheet !== target && sheet.name.toUpperCase() === name.toUpperCase()))
      invalid("Worksheet names must remain valid and unique throughout a reference-editing batch.");
  };
  if (operation.op === "insertWorkbookSheet") {
    checkName(operation.name);
    const before = operation.beforeSheetId ? sheet(operation.beforeSheetId) : undefined, template = workbook.sheets[0]!;
    const added: WorkbookSheet = { sheetId: allocate(), name: operation.name, visibility: "visible", defaultRowHeightPt: 15, defaultColumnWidth: 8.43,
      defaultCssClass: template.defaultCssClass, ...(template.defaultNumberFormatId ? { defaultNumberFormatId: template.defaultNumberFormatId } : {}),
      rowDimensions: [], columnDimensions: [], merges: [], frozen: { rows: 0, columns: 0 }, showGridLines: true, rightToLeft: false };
    workbook.sheets.splice(before ? workbook.sheets.indexOf(before) : workbook.sheets.length, 0, added); return;
  }
  if (!("sheetId" in operation)) invalid("Expected a worksheet operation.");
  const target = sheet(operation.sheetId);
  if (operation.op === "setWorkbookAutoFilter") {
    if (operation.filter === null) delete target.autoFilter; else target.autoFilter = operation.filter;
    return;
  }
  if (operation.op === "updateWorkbookSheet") {
    if (operation.changes.name !== undefined) checkName(operation.changes.name, target);
    if (operation.changes.name !== undefined && operation.changes.name !== target.name) rewriteWorkbookReferences(workbook, cells, { kind: "renameSheet", from: target.name, to: operation.changes.name });
    Object.assign(target, operation.changes);
    if (target.visibility !== "visible" && workbook.activeSheetId === target.sheetId) workbook.activeSheetId = workbook.sheets.find(sheet => sheet.visibility === "visible")?.sheetId ?? target.sheetId;
    return;
  }
  if (operation.op === "moveWorkbookSheet") {
    const before = operation.beforeSheetId ? sheet(operation.beforeSheetId) : undefined;
    if (before === target) return;
    workbook.sheets.splice(workbook.sheets.indexOf(target), 1);
    workbook.sheets.splice(before ? workbook.sheets.indexOf(before) : workbook.sheets.length, 0, target); return;
  }
  if (operation.op !== "removeWorkbookSheet") invalid("Expected a worksheet structure operation.");
  if (workbook.sheets.length === 1) invalid("A workbook must retain at least one worksheet.");
  rewriteWorkbookReferences(workbook, cells, { kind: "deleteSheet", sheet: target.name });
  cells.removeSheet(target.sheetId);
  workbook.sheets = workbook.sheets.filter(sheet => sheet !== target);
  if (workbook.activeSheetId === target.sheetId) workbook.activeSheetId = (workbook.sheets.find(sheet => sheet.visibility === "visible") ?? workbook.sheets[0]!).sheetId;
  if (workbook.names) workbook.names = workbook.names.filter(name => name.sheetId !== target.sheetId);
  const removed = new Set<string>();
  for (const object of workbook.objects ?? []) if (object.sheetId === target.sheetId) {
    const html = content.parse(), pending = [object.nodeId];
    while (pending.length) { const id = pending.pop()!; removed.add(id); pending.push(...html.peekNode(id)!.childNodeIds); }
    content.edit().removeSource(object.nodeId);
  }
  if (workbook.objects) workbook.objects = workbook.objects.filter(object => !removed.has(object.nodeId));
  if (workbook.charts) workbook.charts = workbook.charts.filter(chart => !removed.has(chart.nodeId));
  if (workbook.objectHyperlinks) workbook.objectHyperlinks = workbook.objectHyperlinks.filter(link => !removed.has(link.nodeId));
  for (const annotation of annotations) if (annotation.anchor.kind === "cells" && annotation.anchor.sheetId === target.sheetId && annotation.status !== "needs_reanchor") {
    const binding = workbook.comments?.find(comment => comment.annotationId === annotation.annotationId);
    if (binding) binding.recoveryStatus = annotation.status;
    annotation.status = "needs_reanchor";
  }
  mapping.push({ kind: "deleteSheet", sheetId: target.sheetId });
}
