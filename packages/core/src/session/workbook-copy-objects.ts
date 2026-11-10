import type { WorkbookChartReference, WorkbookDrawingAnchor, WorkbookMetadata, WorkbookObject, WorkbookRange } from "../contracts/workbook.js";
import type { WorkbookOperation } from "../contracts/workbook-operations.js";
import type { WorkbookModel } from "../runtime.js";
import { DocumentError } from "../errors.js";
import { encode } from "../document/package.js";
import { WorkbookDependencyIndex } from "../workbook/dependency-index.js";
import { subtractWorkbookRange } from "../workbook/range-mapping.js";
import { parseFormula, type FormulaExpression } from "../workbook/formula.js";
import { rewriteFormula } from "../workbook/formula-rewrite.js";
import { WorkbookReferences } from "../workbook/reference.js";
import { planWorkbookRegionCopy } from "./workbook-copy-region.js";
import { cloneWorkbookDrawing } from "./workbook-drawings.js";
import { translateWorkbookAnchor } from "./workbook-anchor-translation.js";
import type { WorkbookMeasuredLayout } from "./workbook-layout.js";
import type { WorkbookContentEdit } from "./workbook-content.js";
import type { SourceStyleEdit } from "./source-styles.js";

type Copy = Extract<WorkbookOperation, { op: "copyWorkbookObjects" }>;
interface Rect { x: number; y: number; width: number; height: number }
const copies = new WeakMap<WorkbookModel, ReadonlyMap<string, string>>();
interface WorkbookObjectCopyState { sources: Map<string, string>; bytes: number; nodes: number }
export function createWorkbookObjectCopyState(): WorkbookObjectCopyState { return { sources: new Map<string, string>(), bytes: 0, nodes: 0 }; }
export function bindWorkbookObjectCopies(model: WorkbookModel, sources: ReadonlyMap<string, string>): void { if (sources.size) copies.set(model, new Map(sources)); }
export function workbookObjectCopies(model: WorkbookModel): ReadonlyMap<string, string> | undefined { return copies.get(model); }
const invalid = (message: string): never => { throw new DocumentError("InvalidOperation", message); };
function plan(workbook: WorkbookMetadata, operation: Copy) {
  if (![operation.sheetId, operation.targetSheetId].every(id => workbook.sheets.some(sheet => sheet.sheetId === id))) invalid("Drawing copying requires existing worksheets.");
  const objects = new Map(workbook.objects?.map(object => [object.nodeId, object]));
  if (new Set(operation.nodeIds).size !== operation.nodeIds.length || operation.nodeIds.some(id => objects.get(id)?.sheetId !== operation.sheetId)) invalid("Drawing copying requires distinct current object identities from its source worksheet.");
  const region = planWorkbookRegionCopy(operation); let work = 0;
  const step = () => { if (++work > 10_000_000) throw new DocumentError("ResourceLimitExceeded", "Drawing copying exceeds its range-work budget."); };
  const index = new WorkbookDependencyIndex(region.source.map((range, at) => ({ area: { ...range, sheetId: operation.sheetId }, owner: String(at) })), step);
  const contains = (range: WorkbookRange): boolean => {
    let pieces = [range];
    for (const key of index.query({ ...range, sheetId: operation.sheetId }, step)) {
      pieces = pieces.flatMap(piece => { step(); return subtractWorkbookRange(piece, region.source[Number(key)]!); });
      if (!pieces.length) return true;
      if (pieces.length > 4096) throw new DocumentError("ResourceLimitExceeded", "Drawing copying exceeds its footprint-fragment budget.");
    }
    return false;
  };
  return { ...region, step, contains };
}
export function workbookObjectCopySheets(workbook: WorkbookMetadata, operation: Copy): readonly string[] {
  const region = plan(workbook, operation); if (region.unchanged || !region.footprint.length) return [];
  const selected = new Set(operation.nodeIds);
  return (workbook.objects ?? []).some(object => selected.has(object.nodeId) && object.sheetId === operation.sheetId && (object.anchor.kind === "absolute"
    || region.contains({ from: object.anchor.from, to: object.anchor.from }))) ? [...new Set([operation.sheetId, operation.targetSheetId])] : [];
}
function subtractRect(value: Rect, mask: Rect): Rect[] {
  const left = Math.max(value.x, mask.x), right = Math.min(value.x + value.width, mask.x + mask.width), top = Math.max(value.y, mask.y), bottom = Math.min(value.y + value.height, mask.y + mask.height);
  if (right <= left || bottom <= top) return [value];
  const result: Rect[] = [];
  if (value.y < top) result.push({ ...value, height: top - value.y });
  if (bottom < value.y + value.height) result.push({ ...value, y: bottom, height: value.y + value.height - bottom });
  if (value.x < left) result.push({ x: value.x, y: top, width: left - value.x, height: bottom - top });
  if (right < value.x + value.width) result.push({ x: right, y: top, width: value.x + value.width - right, height: bottom - top });
  return result;
}
function directReference(value: FormulaExpression): boolean {
  return ["cell", "range", "axisRange"].includes(value.kind) || value.kind === "binary" && [",", " ", ":"].includes(value.operator) && directReference(value.left) && directReference(value.right);
}
/** Cell markers select their logical owner; an absolute box must be completely
 * covered by the actual measured source footprint. Existing target drawings
 * are independent and remain intact. All copies use fresh source identities. */
export function* copyWorkbookObjects(workbook: WorkbookMetadata, content: WorkbookContentEdit, operation: Copy, allocate: () => string, styles: SourceStyleEdit,
  layout: WorkbookMeasuredLayout | undefined, state: ReturnType<typeof createWorkbookObjectCopyState>, limits: { readonly maxNodes: number; readonly maxFileBytes: number }): Generator<{ completed: number; total: number }, void> {
  const region = plan(workbook, operation); if (region.unchanged || !region.footprint.length) return;
  const selectedIds = new Set(operation.nodeIds), originals = [...workbook.objects ?? []].filter(object => selectedIds.has(object.nodeId) && object.sheetId === operation.sheetId && (object.anchor.kind === "absolute" || region.contains({ from: object.anchor.from, to: object.anchor.from })));
  if (!originals.length) return;
  const measured = layout ?? invalid("Copying drawings requires measured source and destination worksheet layouts."), axes = measured.sheets.get(operation.sheetId) ?? invalid("The copied drawing source lacks its measured worksheet.");
  const rectangles = region.source.map(range => ({ x: axes.columns.position(range.from.column), y: axes.rows.position(range.from.row),
    width: axes.columns.position(range.to.column + 1) - axes.columns.position(range.from.column), height: axes.rows.position(range.to.row + 1) - axes.rows.position(range.from.row) }));
  const covered = (anchor: Extract<WorkbookDrawingAnchor, { kind: "absolute" }>) => {
    let pieces: Rect[] = [{ x: anchor.xPt, y: anchor.yPt, width: anchor.widthPt, height: anchor.heightPt }];
    for (const mask of rectangles) {
      pieces = pieces.flatMap(piece => { region.step(); return subtractRect(piece, mask); }); if (!pieces.length) return true;
      if (pieces.length > 4096) throw new DocumentError("ResourceLimitExceeded", "Drawing copying exceeds its physical-footprint budget.");
    }
    return false;
  };
  const selected = (object: WorkbookObject) => object.anchor.kind === "absolute" ? covered(object.anchor) : region.contains({ from: object.anchor.from, to: object.anchor.from });
  const sheets = new Map(workbook.sheets.map(sheet => [sheet.sheetId, sheet])), references = new WorkbookReferences(workbook), charts = new Map(workbook.charts?.map(chart => [chart.nodeId, chart]));
  const identities = new Map<string, string>(); let completed = 0, units = 0;
  for (const original of originals) {
    if (selected(original)) {
      if (state.nodes >= limits.maxNodes) throw new DocumentError("ResourceLimitExceeded", "Drawing copying exceeds its cumulative node budget.");
      const sourceChart = charts.get(original.nodeId);
      const html = content.parse().getNode(original.nodeId)!.html, metadata = JSON.stringify([original, sourceChart]);
      units += html.length + metadata.length;
      if (units > (operation.maxSourceCodeUnits ?? limits.maxFileBytes)) throw new DocumentError("ResourceLimitExceeded", "Drawing copying exceeds its source-text budget.");
      state.bytes += encode(html).byteLength + encode(metadata).byteLength;
      if (state.bytes > limits.maxFileBytes) throw new DocumentError("ResourceLimitExceeded", "Drawing copying exceeds its cumulative source-byte budget.");
      const cloned = cloneWorkbookDrawing(workbook, content, original, allocate), at = original.anchor.kind === "absolute" ? operation.range.from : original.anchor.from;
      state.nodes += cloned.identities.size;
      if (state.nodes > limits.maxNodes) throw new DocumentError("ResourceLimitExceeded", "Drawing copying exceeds its cumulative node budget.");
      cloned.object.sheetId = operation.targetSheetId;
      cloned.object.anchor = translateWorkbookAnchor(original.anchor, { sheetId: operation.sheetId, at, layout: measured },
        { sheetId: operation.targetSheetId, at: { row: at.row + region.delta.row, column: at.column + region.delta.column }, layout: measured });
      if (sourceChart && cloned.object.kind === "chart") {
        const chart = workbook.charts!.at(-1)!;
        const reference = (value: WorkbookChartReference | undefined) => {
          if (!value) return;
          const originId = value.originSheetId ?? original.sheetId, origin = sheets.get(originId) ?? invalid("A copied chart expression requires its current origin worksheet."), target = sheets.get(operation.targetSheetId)!;
          const bound = directReference(parseFormula(value.expression, { row: 0, column: 0 })) ? references.resolve(value.expression, { sheetId: originId, row: 0, column: 0 }) : undefined;
          if (bound?.kind === "ranges" && bound.areas.length
            && bound.areas.every(area => area.sheetId === operation.sheetId && region.contains(area))) {
            value.expression = rewriteFormula(value.expression, { sheet: origin.name, row: 0, column: 0 }, { sheet: target.name, row: 0, column: 0 },
              { kind: "move", sheet: sheets.get(operation.sheetId)!.name, range: operation.range, target: { sheet: target.name, ...operation.target } });
            delete value.originSheetId;
          } else if (originId !== operation.targetSheetId) value.originSheetId = originId;
          if (value.cache) value.cache.state = "stale";
        };
        for (const value of [chart.categories, chart.titleReference, chart.categoryAxisTitle?.reference, chart.valueAxisTitle?.reference]) if (value && "expression" in value) reference(value);
        for (const series of chart.series) { if ("expression" in series.values) reference(series.values); if (series.name?.kind === "reference") reference(series.name.reference); }
      }
      state.sources.set(cloned.object.nodeId, original.nodeId); for (const [source, target] of cloned.identities) identities.set(source, target);
    }
    if (++completed % 64 === 0) yield { completed, total: originals.length };
  }
  styles.aliasNodeIds(identities);
}
