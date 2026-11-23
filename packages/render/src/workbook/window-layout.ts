import { DocumentError } from "@doffi/core";
import type { WorkbookNumberFormatter, WorkbookRuleReadResponse, WorkbookSnapshot } from "@doffi/core";
import type { WorkbookFormulaSource, WorkbookNumberDisplay } from "@doffi/core";
import type { FontResolution } from "../contracts.js";
import type { FontProbeCache } from "../font-probe-cache.js";
import { WorkbookCellSources } from "./cell-source.js";
import { workbookCellBorders, workbookCellDiagonals } from "./cell-borders.js";
import type { WorkbookCellBorders } from "./cell-borders.js";
import { workbookCellFill } from "./cell-fill.js";
import { workbookCellShape } from "./cell-shape.js";
import type { WorkbookTextLayout } from "./text-layout.js";
import { workbookTextRotation } from "./text-alignment.js";
import { projectWorkbookBorders } from "./border-projection.js";
import { projectWorkbookDiagonals } from "./diagonal-projection.js";
import { createWorkbookBorderView } from "./border-view.js";
import { projectWorkbookGridLines, WORKBOOK_GRID_PAINT } from "./grid-projection.js";
import type { AxisDimensions } from "./dimensions.js";
import { WorkbookMergeIndex } from "./projection.js";
import type { WorkbookGridCell } from "./projection.js";
import type { WorkbookSurface } from "./surface.js";
import type { WorkbookWindowFrame } from "./window-loader.js";
import { paintWorkbookWindow } from "./window-paint.js";
import { layoutWorkbookCellTexts } from "./cell-text-layout.js";
import { layoutCachedWorkbookCellTexts } from "./cell-text-cache.js";
import type { WorkbookCellTextCache } from "./cell-text-cache.js";
import { WorkbookTextFlow } from "./text-flow.js";
import { intersectRect } from "./projection.js";
import type { WorkbookRect } from "./projection.js";

export interface WorkbookWindowLayoutOptions {
  readonly normalSpaceWidthPx: number;
  readonly maxNodes?: number;
  readonly maxSourceNodes?: number;
  readonly maxTextCodeUnits?: number;
  readonly maxCellFragments?: number;
  readonly printGridLines?: boolean;
}
export interface WorkbookWindowLayout {
  readonly element: HTMLElement;
  readonly fonts: readonly FontResolution[];
  readonly usedFallbackLineMetrics: boolean;
  readonly cells: readonly {
    readonly projection: WorkbookGridCell;
    readonly text?: WorkbookTextLayout;
    readonly cacheState?: WorkbookFormulaSource["cacheState"];
    readonly numberOverflow?: Extract<WorkbookNumberDisplay, { kind: "overflow" }>["reason"];
  }[];
}

/** Internal cell compositor. Source CSS is measured in the source surface;
 * derived text is measured in the isolated display surface. This prepares a
 * whole detached grid and never replaces the caller's published window.
 * Objects/rules/automatic row heights belong to the enclosing preview stage. */
export function layoutWorkbookWindow(surface: Pick<WorkbookSurface, "source" | "view">, snapshot: WorkbookSnapshot, frame: WorkbookWindowFrame,
  dimensions: { readonly rows: AxisDimensions; readonly columns: AxisDimensions }, formatter: WorkbookNumberFormatter,
  options: WorkbookWindowLayoutOptions, rules?: WorkbookRuleReadResponse, cache?: WorkbookCellTextCache, probes?: FontProbeCache): WorkbookWindowLayout {
  const maxNodes = options.maxNodes ?? 50_000, maxSourceNodes = options.maxSourceNodes ?? 50_000, maxTextUnits = options.maxTextCodeUnits ?? 8 * 1024 * 1024, maxCellFragments = options.maxCellFragments ?? 5000;
  if (Object.keys(options).some(key => !["normalSpaceWidthPx", "maxNodes", "maxSourceNodes", "maxTextCodeUnits", "maxCellFragments", "printGridLines"].includes(key))
    || ![maxNodes, maxSourceNodes, maxTextUnits, maxCellFragments].every(n => Number.isSafeInteger(n) && n > 0) || !Number.isFinite(options.normalSpaceWidthPx) || options.normalSpaceWidthPx < 0
    || surface.source.ownerDocument !== surface.view.ownerDocument || !surface.view.isConnected) throw new DocumentError("InvalidOperation", "Workbook window layout requires one connected surface and positive budgets.");
  const document = surface.source.ownerDocument, sheet = snapshot.metadata.workbook.sheets.find(sheet => sheet.sheetId === frame.sheetId);
  if (!sheet) throw new DocumentError("InvalidOperation", "The workbook window has no source worksheet.");
  const sources = new WorkbookCellSources(surface.source, snapshot, frame, maxSourceNodes, rules, Boolean(cache));
  const current = new Map(rules?.cells.map(cell => [cell.row * 16_384 + cell.column, cell.current]));
  try {
    const styles = new Map<number, CSSStyleDeclaration>();
    const classStyles = new Map<string, CSSStyleDeclaration>();
    const style = (row: number, column: number): CSSStyleDeclaration => {
      const key = row * 16_384 + column, existing = styles.get(key); if (existing) return existing;
      const source = sources.get(row, column);
      if (!source) throw new DocumentError("InvalidOperation", "A workbook appearance source is missing from its window.");
      const shared = cache ? classStyles.get(source.format.cssClass) : undefined;
      const value = shared ?? document.defaultView!.getComputedStyle(source.element);
      if (cache && !shared) classStyles.set(source.format.cssClass, value);
      styles.set(key, value); return value;
    };
    const borders = new WeakMap<CSSStyleDeclaration, WorkbookCellBorders>();
    const bordersAt = (row: number, column: number): WorkbookCellBorders => {
      const computed = style(row, column), known = borders.get(computed); if (known) return known;
      const value = workbookCellBorders(document, computed, sheet.rightToLeft); borders.set(computed, value); return value;
    };
    const bordered = (cell: WorkbookGridCell): boolean => {
      const { from, to } = cell.range, anchor = bordersAt(from.row, from.column);
      return Boolean(anchor.top || anchor.before || bordersAt(from.row, to.column).after || bordersAt(to.row, from.column).bottom);
    };
    const merges = new WorkbookMergeIndex(sheet.merges), flow = new WorkbookTextFlow(frame, sheet, dimensions, merges, sources, style, Math.min(1_000_000, maxNodes * 8));
    const valueAt = (row: number, column: number) => {
      const calculated = current.get(row * 16_384 + column); if (calculated?.kind === "value") return calculated.value;
      if (calculated && options.printGridLines !== undefined) return undefined;
      return frame.getCell(row, column)?.value;
    };
    const projections = options.printGridLines === undefined ? [...frame.projections, ...flow.candidates(valueAt)] : frame.projections, textRects = new Map<WorkbookGridCell, WorkbookRect>();
    const textInputs = projections.map(projection => {
      const source = sources.get(projection.row, projection.column)!, cell = frame.getCell(projection.row, projection.column);
      const value = valueAt(projection.row, projection.column), rect = value && value.type !== "blank" ? flow.textRect(projection) : projection.rect;
      textRects.set(projection, rect);
      return { source, value, formula: Boolean(cell?.formula), widthPt: rect.width, heightPt: rect.height, bordered: bordered(projection) };
    });
    const context = [document.defaultView!.innerWidth, document.defaultView!.innerHeight];
    const keys = cache ? projections.map((projection, index) => JSON.stringify([frame.sheetId, projection.row, projection.column,
      textInputs[index]!.widthPt, textInputs[index]!.heightPt, textInputs[index]!.bordered, textInputs[index]!.value, textInputs[index]!.formula,
      textInputs[index]!.source.format, frame.getCell(projection.row, projection.column)?.html, context])) : [];
    const textLayout = cache ? layoutCachedWorkbookCellTexts(surface.view, snapshot, textInputs, keys, formatter, options, maxNodes, maxTextUnits, cache)
      : layoutWorkbookCellTexts(surface.view, snapshot, textInputs, formatter, options, maxNodes, maxTextUnits, probes);
    let textNodes = textLayout.nodes;
    const remaining = (): number => {
      if (textNodes >= maxNodes) throw new DocumentError("ResourceLimitExceeded", "Workbook window text exceeds its cumulative node budget.");
      return maxNodes - textNodes;
    };
    const textIndices = new Map(projections.map((projection, index) => [projection, index]));
    const fills = new WeakMap<CSSStyleDeclaration, ReturnType<typeof workbookCellFill>>(), rotations = new WeakMap<CSSStyleDeclaration, number>();
    const rotation = (computed: CSSStyleDeclaration): number => {
      const known = rotations.get(computed); if (known !== undefined) return known;
      const value = workbookTextRotation(computed).rotationDegrees; rotations.set(computed, value); return value;
    };
    const all = projections.map((projection, index) => {
      const computed = style(projection.row, projection.column), text = textLayout.cells[index]!.text;
      const shape = workbookCellShape(projection.rect.width, projection.rect.height,
        rotation(computed), bordered(projection));
      const textRect = textRects.get(projection) ?? projection.rect, value = valueAt(projection.row, projection.column);
      const textClip = text && value ? flow.clip(projection, textRect, text, value.type) : undefined;
      let fill = fills.get(computed); if (!fill) { fill = workbookCellFill(document, computed); fills.set(computed, fill); }
      return { projection, shape, fill, ...(text ? { text, textRect, ...(textClip ? { textClip } : {}) } : {}) };
    });
    const prepared = all.slice(0, frame.projections.length), overflow = all.slice(frame.projections.length).filter(cell => {
      const pane = frame.layout.panes.find(pane => pane.kind === cell.projection.pane)!;
      const paneBounds = { x: sheet.rightToLeft ? frame.layout.widthPt - pane.inlineStartPt - pane.widthPt : pane.inlineStartPt, y: pane.blockStartPt, width: pane.widthPt, height: pane.heightPt };
      const x = sheet.rightToLeft ? frame.layout.widthPt - cell.projection.rect.x - cell.projection.rect.width : cell.projection.rect.x;
      if (cell.shape.topOffsetPt && intersectRect({ x: x + Math.min(0, cell.shape.topOffsetPt), y: cell.projection.rect.y,
        width: cell.projection.rect.width + Math.abs(cell.shape.topOffsetPt), height: cell.projection.rect.height }, paneBounds)) return true;
      if (!cell.text || !cell.textRect || !cell.textClip) return false;
      const bodyX = sheet.rightToLeft ? frame.layout.widthPt - cell.textRect.x - cell.textRect.width : cell.textRect.x;
      const clipX = sheet.rightToLeft ? frame.layout.widthPt - cell.textClip.x - cell.textClip.width : cell.textClip.x;
      return Boolean(intersectRect({ x: bodyX + cell.text.ink.xPt, y: cell.textRect.y + cell.text.ink.yPt, width: cell.text.ink.widthPt, height: cell.text.ink.heightPt }, { ...cell.textClip, x: clipX }));
    });
    if (prepared.length + overflow.length > maxCellFragments) throw new DocumentError("ResourceLimitExceeded", "All workbook panes and offscreen cell fragments exceed their combined cell budget.");
    const projectedBorders = projectWorkbookBorders(dimensions, frame.layout, merges, bordersAt, sheet.rightToLeft,
      maxNodes, (row, column) => rotation(style(row, column)), new Map(frame.skewColumns.map(item => [item.pane, item.columns])));
    const diagonals = projectWorkbookDiagonals(frame.projections, frame.layout, (row, column) => workbookCellDiagonals(document, style(row, column)), sheet.rightToLeft, maxNodes);
    const materialize = (panes: ReturnType<typeof projectWorkbookBorders>): SVGSVGElement[] => panes.filter(pane => pane.edges.length).map(pane => {
      const element = createWorkbookBorderView(document, pane.bounds, pane.edges, remaining()); textNodes += 1 + element.querySelectorAll("*").length; return element;
    });
    const views = materialize([...projectedBorders, ...diagonals]);
    const gridPaint = WORKBOOK_GRID_PAINT;
    const gridLines = (options.printGridLines ?? sheet.showGridLines) ? materialize(sheet.merges.length ? projectWorkbookBorders(dimensions, frame.layout, merges,
      () => ({ before: gridPaint, after: gridPaint, top: gridPaint, bottom: gridPaint }), sheet.rightToLeft, maxNodes) : projectWorkbookGridLines(frame.layout, sheet.rightToLeft, maxNodes)) : [];
    for (const grid of gridLines) grid.dataset.officeView = "workbook-gridlines";
    const element = paintWorkbookWindow(document, frame, prepared, sheet.rightToLeft, views, maxNodes, gridLines, overflow);
    return Object.freeze({ element, fonts: textLayout.fonts, usedFallbackLineMetrics: textLayout.usedFallbackLineMetrics,
      cells: Object.freeze([...prepared, ...overflow].map(cell => {
        const stored = frame.getCell(cell.projection.row, cell.projection.column);
        const overflow = textLayout.cells[textIndices.get(cell.projection)!]?.numberOverflow;
        return Object.freeze({ projection: cell.projection, ...(cell.text ? { text: cell.text } : {}), ...(stored?.formula ? { cacheState: stored.formula.cacheState } : {}),
          ...(overflow ? { numberOverflow: overflow } : {}) });
      })) });
  } finally { sources.dispose(); }
}
