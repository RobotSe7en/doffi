import type { ChangeMapping, DocumentNodeChange, WorkbookModel, WorkbookRangeChange } from "../runtime.js";
import { freeze, invalid } from "../document/validation.js";
import { readonlyMap } from "../session/changes.js";
import { workbookSource } from "./source.js";
import { CHUNK_COLUMNS, CHUNK_ROWS } from "./chunk-index.js";

interface WorkbookSourceDiff {
  readonly changes: readonly DocumentNodeChange[];
  readonly affectedNodeIds: readonly string[];
  readonly affectedSheetIds: readonly string[];
  readonly changedRanges: readonly WorkbookRangeChange[];
  readonly changedAnnotationIds: readonly string[];
  readonly mapping: ChangeMapping;
}
export function compareWorkbooks(before: WorkbookModel, after: WorkbookModel, mapping?: ChangeMapping): WorkbookSourceDiff {
  const steps = compareWorkbookSteps(before, after, mapping);
  for (;;) { const next = steps.next(); if (next.done) return next.value ?? invalid("Workbook comparison did not complete."); }
}
export function* compareWorkbookSteps(before: WorkbookModel, after: WorkbookModel, mapping?: ChangeMapping): Generator<{ completed: number; total: number }, WorkbookSourceDiff | undefined> {
  const previous = workbookSource(before)!, next = workbookSource(after)!, changedRanges: WorkbookRangeChange[] = [], sheets = new Set<string>();
  const oldDescriptors = new Map(before.snapshot().manifest.chunks.map(chunk => [chunk.path, chunk]));
  const newDescriptors = new Map(after.snapshot().manifest.chunks.map(chunk => [chunk.path, chunk]));
  const paths = new Set([...oldDescriptors.keys(), ...newDescriptors.keys()]);
  let completed = 0;
  for (const path of paths) {
    const old = oldDescriptors.get(path), current = newDescriptors.get(path);
    if (previous.files.get(path) === next.files.get(path) && old?.sheetId === current?.sheetId && old?.row === current?.row && old?.column === current?.column) continue;
    // Each immutable chunk is read at its own position. A stable resource path
    // does not imply a stable worksheet or coordinate after a host replacement.
    const oldCells = new Map(old && previous.chunks.get(path)?.readWindow({ ...old, rows: CHUNK_ROWS, columns: CHUNK_COLUMNS }).map(cell => [cell.row * 16_384 + cell.column, cell.html]));
    const newCells = new Map(current && next.chunks.get(path)?.readWindow({ ...current, rows: CHUNK_ROWS, columns: CHUNK_COLUMNS }).map(cell => [cell.row * 16_384 + cell.column, cell.html]));
    for (const sheetId of new Set([...(old ? [old.sheetId] : []), ...(current ? [current.sheetId] : [])])) {
      const left = old?.sheetId === sheetId ? oldCells : new Map<number, string>(), right = current?.sheetId === sheetId ? newCells : new Map<number, string>();
      for (const key of [...new Set([...left.keys(), ...right.keys()])].sort((a, b) => a - b)) {
        if (left.get(key) === right.get(key)) continue;
        const row = Math.floor(key / 16_384), column = key % 16_384, last = changedRanges.at(-1);
        if (last?.sheetId === sheetId && last.range.from.row === row && last.range.to.column + 1 === column)
          changedRanges[changedRanges.length - 1] = { sheetId, range: { from: last.range.from, to: { row, column } } };
        else changedRanges.push({ sheetId, range: { from: { row, column }, to: { row, column } } });
        sheets.add(sheetId);
      }
    }
    yield { completed: ++completed, total: 0 };
  }
  const a = before.snapshot(), b = after.snapshot();
  for (const sheetId of new Set([...a.metadata.workbook.sheets, ...b.metadata.workbook.sheets].map(sheet => sheet.sheetId))) {
    if (JSON.stringify(a.metadata.workbook.sheets.find(sheet => sheet.sheetId === sheetId)) !== JSON.stringify(b.metadata.workbook.sheets.find(sheet => sheet.sheetId === sheetId))) sheets.add(sheetId);
  }
  const globals = (model: WorkbookModel): string => { const { sheets: _sheets, ...rest } = model.snapshot().metadata.workbook; return JSON.stringify(rest); };
  if (globals(before) !== globals(after) || a.metadata.locale !== b.metadata.locale || JSON.stringify(a.styles) !== JSON.stringify(b.styles)) for (const sheet of [...a.metadata.workbook.sheets, ...b.metadata.workbook.sheets]) sheets.add(sheet.sheetId);
  const nodes = new Map<string, string | null>(), changes: DocumentNodeChange[] = [], affected = new Set<string>();
  const contentOwners = ([[a, previous.content], [b, next.content]] as const).map(([snapshot, html]) => ({ html,
    owners: new Map([...(snapshot.metadata.workbook.objects ?? []), ...(snapshot.metadata.workbook.stories ?? [])].map(owner => ["nodeId" in owner ? owner.nodeId : owner.rootNodeId, owner.sheetId])) }));
  for (const id of new Set([...(previous.content?.nodeIds() ?? []), ...(next.content?.nodeIds() ?? [])])) {
    const left = previous.content?.signature(id), right = next.content?.signature(id); if (left === right) continue;
    changes.push({ nodeId: id, kind: left === undefined ? "added" : right === undefined ? "removed" : "modified" });
    if (left !== undefined) nodes.set(id, right === undefined ? null : id);
    for (const html of [previous.content, next.content]) for (let node: string | null = id; node !== null; node = html?.peekNode(node)?.parentNodeId ?? null) affected.add(node);
    for (const { html, owners } of contentOwners) {
      for (let node: string | null = id; node !== null; node = html?.peekNode(node)?.parentNodeId ?? null) if (owners.has(node)) { sheets.add(owners.get(node)!); break; }
    }
  }
  const leftAnnotations = new Map(a.annotations.map(item => [item.annotationId, JSON.stringify(item)]));
  const rightAnnotations = new Map(b.annotations.map(item => [item.annotationId, JSON.stringify(item)]));
  const changedAnnotationIds = [...new Set([...leftAnnotations.keys(), ...rightAnnotations.keys()])].filter(id => leftAnnotations.get(id) !== rightAnnotations.get(id));
  const annotationIds = new Set(changedAnnotationIds);
  for (const annotation of [...a.annotations, ...b.annotations]) if (annotationIds.has(annotation.annotationId) && annotation.anchor.kind === "cells") sheets.add(annotation.anchor.sheetId);
  return freeze({ changes, affectedNodeIds: [...affected], affectedSheetIds: [...sheets], changedRanges, changedAnnotationIds,
    mapping: { nodes: readonlyMap(nodes), textEdits: mapping?.textEdits ?? [], workbook: mapping?.workbook ?? [] } });
}
export function sameWorkbookSource(before: WorkbookModel, after: WorkbookModel): boolean {
  const left = workbookSource(before)!.files, right = workbookSource(after)!.files;
  if (left.size !== right.size) return false;
  for (const [path, bytes] of left) {
    const other = right.get(path);
    if (bytes === other) continue;
    if (!other || bytes.byteLength !== other.byteLength || bytes.some((byte, index) => byte !== other[index])) return false;
  }
  return true;
}
export function retainedWorkbookBytes(models: Iterable<WorkbookModel>): number {
  const buffers = new Set<Uint8Array>(); let total = 0;
  for (const model of models) for (const bytes of workbookSource(model)!.files.values()) if (!buffers.has(bytes)) { buffers.add(bytes); total += bytes.byteLength; }
  return total;
}
