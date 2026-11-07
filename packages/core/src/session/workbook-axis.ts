import type { Annotation } from "../contracts/annotations.js";
import type { WorkbookAutoFilter, WorkbookCoordinate, WorkbookMetadata, WorkbookRange, WorkbookSort } from "../contracts/workbook.js";
import type { WorkbookOperation } from "../contracts/workbook-operations.js";
import type { DeepReadonly, WorkbookCell, WorkbookStructureMapping } from "../runtime.js";
import { DocumentError } from "../errors.js";
import { MAX_SHEET_COLUMNS, MAX_SHEET_ROWS } from "../workbook/chunk-index.js";
import { copyFormula } from "../workbook/formula-rewrite.js";
import { parseFormula, type FormulaExpression } from "../workbook/formula.js";
import { mapWorkbookRanges } from "../workbook/range-mapping.js";
import { sameWorkbookRangeCoverage } from "../workbook/ranges.js";
import { WorkbookCellEdit } from "./workbook-cells.js";
import { rewriteWorkbookReferenceSteps } from "./workbook-references.js";

type AxisEdit = Extract<WorkbookOperation, { op: "editWorkbookAxis" }>;
function invalid(message: string): never { throw new DocumentError("InvalidOperation", message); }
function contextualName(node: FormulaExpression, global: boolean): boolean {
  if (node.kind === "cell") return !node.absoluteRow || !node.absoluteColumn || global && node.sheet === undefined;
  if (node.kind === "range") return contextualName(node.from, global) || contextualName(node.to, global);
  if (node.kind === "axisRange") return !node.from.absolute || !node.to.absolute || global && node.sheet === undefined;
  if (node.kind === "unary") return contextualName(node.operand, global);
  if (node.kind === "binary") return contextualName(node.left, global) || contextualName(node.right, global);
  if (node.kind === "call") return node.arguments.some(argument => argument !== null && contextualName(argument, global));
  return false;
}

/** Reversed structure changes must not publish canonicalized formula syntax or
 * split equivalent dimension spans as a new authoritative revision. */
export function restoreWorkbookAxisSource(workbook: WorkbookMetadata, before: DeepReadonly<WorkbookMetadata>, annotations: Annotation[], priorAnnotations: readonly DeepReadonly<Annotation>[]): void {
  const sameFormula = (a: string, b: string, origin: WorkbookCoordinate): boolean =>
    a === b || JSON.stringify(parseFormula(a, origin)) === JSON.stringify(parseFormula(b, origin));
  function dimensionKey(input: readonly { readonly from: number; readonly to: number }[]): string {
    const result: { from: number; to: number; settings: string }[] = [];
    for (const span of input) {
      const { from, to, ...settings } = span, key = JSON.stringify(settings), last = result.at(-1);
      if (last && last.to + 1 === from && last.settings === key) last.to = to;
      else result.push({ from, to, settings: key });
    }
    return JSON.stringify(result);
  }
  const priorSheets = new Map(before.sheets.map(sheet => [sheet.sheetId, sheet]));
  for (const sheet of workbook.sheets) {
    const old = priorSheets.get(sheet.sheetId); if (!old) continue;
    for (const key of ["rowDimensions", "columnDimensions"] as const) if (dimensionKey(sheet[key]) === dimensionKey(old[key])) {
      if (key === "rowDimensions") sheet.rowDimensions = old.rowDimensions.map(span => ({ ...span }));
      else sheet.columnDimensions = old.columnDimensions.map(span => ({ ...span }));
    }
    const priorRules = new Map([...old.conditionalRules ?? [], ...old.dataValidations ?? []].map(rule => [rule.ruleId, rule]));
    for (const rule of [...sheet.conditionalRules ?? [], ...sheet.dataValidations ?? []]) {
      const previous = priorRules.get(rule.ruleId);
      if (!previous || previous.origin.row !== rule.origin.row || previous.origin.column !== rule.origin.column) continue;
      if ("expression" in rule && "expression" in previous && sameFormula(rule.expression, previous.expression, rule.origin)) rule.expression = previous.expression;
      if ("formulas" in rule && "formulas" in previous) rule.formulas = rule.formulas.map((formula, index) =>
        previous.formulas[index] !== undefined && sameFormula(formula, previous.formulas[index]!, rule.origin) ? previous.formulas[index]! : formula);
      if (rule.kind === "colorScale" && previous.kind === "colorScale") for (const [index, stop] of rule.stops.entries()) {
        const prior = previous.stops[index];
        if (stop.kind === "formula" && prior?.kind === "formula" && sameFormula(stop.expression, prior.expression, rule.origin)) stop.expression = prior.expression;
      }
    }
  }
  const priorNames = new Map(before.names?.map(name => [name.nameId, name]));
  for (const name of workbook.names ?? []) {
    const old = priorNames.get(name.nameId);
    if (old && old.origin.row === name.origin.row && old.origin.column === name.origin.column && sameFormula(name.expression, old.expression, name.origin)) name.expression = old.expression;
  }
  const prior = new Map(priorAnnotations.map(annotation => [annotation.annotationId, annotation]));
  const bounds = (ranges: readonly DeepReadonly<WorkbookRange>[]): string => JSON.stringify(ranges.reduce((span, range) =>
    [Math.min(span[0]!, range.from.row), Math.min(span[1]!, range.from.column), Math.max(span[2]!, range.to.row), Math.max(span[3]!, range.to.column)], [Infinity, Infinity, -1, -1]));
  for (const annotation of annotations) {
    const old = prior.get(annotation.annotationId);
    if (old?.anchor.kind === "cells" && annotation.anchor.kind === "cells" && old.anchor.sheetId === annotation.anchor.sheetId
      && old.status === annotation.status && old.anchor.ranges.length !== annotation.anchor.ranges.length
      && bounds(old.anchor.ranges) === bounds(annotation.anchor.ranges) && sameWorkbookRangeCoverage(old.anchor.ranges, annotation.anchor.ranges))
      annotation.anchor.ranges = old.anchor.ranges.map(range => ({ from: { ...range.from }, to: { ...range.to } }));
  }
}

/** Sparse source relocation and metadata/reference mapping are one candidate.
 * Complex native object/table integration is added at its editing boundary;
 * until then those combinations fail atomically rather than drifting. */
export function* editWorkbookAxis(workbook: WorkbookMetadata, cells: WorkbookCellEdit, annotations: Annotation[], operation: AxisEdit,
  mappings: WorkbookStructureMapping[]): Generator<{ completed: number; total: number }, void> {
  const sheet = workbook.sheets.find(sheet => sheet.sheetId === operation.sheetId) ?? invalid("An earlier operation removed the worksheet.");
  const { axis, action, index, count } = operation, bound = axis === "row" ? MAX_SHEET_ROWS : MAX_SHEET_COLUMNS;
  if (index + count > bound) invalid("The worksheet axis edit exceeds its native bounds.");
  if (action === "delete" && operation.formatSource !== undefined) invalid("A deletion cannot inherit inserted-axis formatting.");
  if (sheet.tables?.length || workbook.objects?.some(object => object.sheetId === sheet.sheetId)
    || workbook.comments?.some(comment => annotations.some(annotation => annotation.annotationId === comment.annotationId && annotation.anchor.kind === "cells" && annotation.anchor.sheetId === sheet.sheetId)))
    invalid("Axis editing of native tables, drawings and comments is not yet connected.");
  if (workbook.names?.some(name => contextualName(parseFormula(name.expression, name.origin), name.sheetId === undefined)))
    invalid("Axis editing with contextual defined names is not yet connected.");
  const step: WorkbookStructureMapping = { kind: "axis", sheetId: sheet.sheetId, axis, action, index, count };
  const styleSource = action === "insert" && operation.formatSource !== undefined && operation.formatSource !== "none"
    ? operation.formatSource === "before" ? index > 0 ? index - 1 : undefined : index : undefined;
  const formats: WorkbookCell[] = [];
  if (styleSource !== undefined) {
    const band = { from: { row: 0, column: 0, [axis]: styleSource }, to: { row: MAX_SHEET_ROWS - 1, column: MAX_SHEET_COLUMNS - 1, [axis]: styleSource } };
    let scanned = 0;
    for (const cell of cells.entries(sheet.sheetId, false, [band])) {
      if (cell.cssClass || cell.formatId) formats.push(cell);
      if (++scanned % 256 === 0) yield { completed: scanned, total: 0 };
    }
  }
  const dimensionSource = styleSource === undefined ? undefined : (axis === "row" ? sheet.rowDimensions : sheet.columnDimensions).find(span => span.from <= styleSource && span.to >= styleSource);
  const point = (value: number): number | undefined => {
    if (value < index) return value;
    if (action === "delete" && value < index + count) return undefined;
    const mapped = value + (action === "insert" ? count : -count);
    if (mapped >= bound) invalid("The worksheet edit would discard source outside its native bounds.");
    return mapped;
  };
  // Definitions expand across an internal insertion. Cell identities below
  // deliberately use the separate fragment mapping which excludes new cells.
  function interval(from: number, to: number): { from: number; to: number } | undefined {
    if (action === "insert") return { from: point(from)!, to: point(to)! };
    if (to < index) return { from, to };
    if (from >= index + count) return { from: from - count, to: to - count };
    const first = from < index ? from : index, last = to >= index + count ? to - count : index - 1;
    return first <= last ? { from: first, to: last } : undefined;
  }
  const rectangle = (range: WorkbookRange): WorkbookRange | undefined => {
    const mapped = interval(range.from[axis], range.to[axis]);
    return mapped && { from: { ...range.from, [axis]: mapped.from }, to: { ...range.to, [axis]: mapped.to } };
  };
  const ranges = (input: WorkbookRange[], inherit = false): WorkbookRange[] => input.flatMap(range => {
    const mapped = rectangle(range);
    if (mapped && inherit && styleSource !== undefined && range.from[axis] <= styleSource && range.to[axis] >= styleSource) {
      mapped.from[axis] = Math.min(mapped.from[axis], index); mapped.to[axis] = Math.max(mapped.to[axis], index + count - 1);
    }
    return mapped ? [mapped] : [];
  });
  function dimensions<T extends { from: number; to: number }>(input: T[]): T[] {
    // The low-level insertion introduces default-sized/default-formatted empty
    // axes. Retain existing dimensions by identity, without expanding blanks.
    return input.flatMap(span => {
      if (action === "delete") { const mapped = interval(span.from, span.to); return mapped ? [{ ...span, ...mapped }] : []; }
      const result: T[] = [];
      if (span.from < index) result.push({ ...span, to: Math.min(span.to, index - 1) });
      if (span.to >= index) result.push({ ...span, from: point(Math.max(span.from, index))!, to: point(span.to)! });
      return result;
    });
  }
  function sort(value: WorkbookSort | undefined): WorkbookSort | undefined {
    if (!value) return undefined;
    const range = rectangle(value.range), keys = value.keys.flatMap(key => {
      const column = axis === "column" ? point(key.column) : key.column; return column === undefined ? [] : [{ ...key, column }];
    });
    return range && keys.length ? { ...value, range, keys } : undefined;
  }
  function filter(value: WorkbookAutoFilter): WorkbookAutoFilter | undefined {
    const range = rectangle(value.range); if (!range) return undefined;
    if (axis === "row" && action === "delete" && point(value.range.from.row) === undefined)
      invalid("Deleting a surviving filter's header is not yet connected.");
    const columns = value.columns.flatMap(item => {
      const column = axis === "column" ? point(item.column) : item.column; return column === undefined ? [] : [{ ...item, column }];
    });
    const sorted = sort(value.sort);
    return { range, columns, ...(sorted ? { sort: sorted } : {}) };
  }
  // Rebase rules whose definition origin is deleted onto the first surviving
  // old target before rewriting references, so relative templates retain meaning.
  const origins: { rule: { origin: WorkbookCoordinate }; origin: WorkbookCoordinate }[] = [];
  for (const key of ["conditionalRules", "dataValidations"] as const) {
    const retained = (sheet[key] ?? []).filter(rule => {
      const mapped = ranges(rule.ranges, true); if (!mapped.length) return false;
      let origin = rule.origin;
      if (point(origin[axis]) === undefined) {
        const survivor = rule.ranges.find(range => range.from[axis] < index || range.to[axis] >= index + count)!;
        const coordinate = { ...survivor.from, [axis]: survivor.from[axis] < index ? survivor.from[axis] : index + count };
        const rebase = (expression: string) => copyFormula(expression, origin, coordinate);
        if ("expression" in rule) rule.expression = rebase(rule.expression);
        if ("formulas" in rule) rule.formulas = rule.formulas.map(rebase);
        if (rule.kind === "colorScale") for (const stop of rule.stops) if (stop.kind === "formula") stop.expression = rebase(stop.expression);
        rule.origin = coordinate; origin = coordinate;
      }
      rule.ranges = mapped;
      origins.push({ rule, origin: { ...origin, [axis]: point(origin[axis])! } });
      return true;
    });
    if (key === "conditionalRules" && sheet.conditionalRules !== undefined) sheet.conditionalRules = retained as NonNullable<typeof sheet.conditionalRules>;
    if (key === "dataValidations" && sheet.dataValidations !== undefined) sheet.dataValidations = retained as NonNullable<typeof sheet.dataValidations>;
  }
  yield* rewriteWorkbookReferenceSteps(workbook, cells, { kind: "axis", sheet: sheet.name, axis, action, index, count });
  for (const item of origins) item.rule.origin = item.origin;
  const affected: WorkbookRange = { from: { row: 0, column: 0, [axis]: index }, to: { row: MAX_SHEET_ROWS - 1, column: MAX_SHEET_COLUMNS - 1 } };
  const moving: { cell: WorkbookCell; position: number }[] = [];
  let completed = 0;
  for (const cell of cells.entries(sheet.sheetId, false, [affected])) {
    const position = point(cell[axis]);
    if (position !== undefined) moving.push({ cell, position });
    if (++completed % 256 === 0) yield { completed, total: 0 };
  }
  yield* cells.clearSteps(sheet.sheetId, [affected], "all");
  completed = 0;
  for (const { cell, position } of moving) {
    cells.relocate(cell, axis === "row" ? position : cell.row, axis === "column" ? position : cell.column);
    if (++completed % 256 === 0) yield { completed, total: moving.length };
  }
  cells.checkAdditionalCells(formats.length * count);
  let formatted = 0;
  for (const cell of formats) for (let offset = 0; offset < count; offset++) {
    cells.format(sheet.sheetId, { row: axis === "row" ? index + offset : cell.row, column: axis === "column" ? index + offset : cell.column,
      ...(cell.cssClass ? { cssClass: cell.cssClass } : {}), ...(cell.formatId ? { formatId: cell.formatId } : {}) });
    if (++formatted % 256 === 0) yield { completed: formatted, total: formats.length * count };
  }
  if (axis === "row") sheet.rowDimensions = dimensions(sheet.rowDimensions);
  else sheet.columnDimensions = dimensions(sheet.columnDimensions);
  if (dimensionSource) {
    const source = { ...dimensionSource, from: index, to: index + count - 1 };
    if (axis === "row") { sheet.rowDimensions.push(source); sheet.rowDimensions.sort((a, b) => a.from - b.from); }
    else { sheet.columnDimensions.push({ ...source, unit: "characters" }); sheet.columnDimensions.sort((a, b) => a.from - b.from); }
  }
  sheet.merges = ranges(sheet.merges).filter(range => range.from.row !== range.to.row || range.from.column !== range.to.column);
  const frozenKey = axis === "row" ? "rows" : "columns", frozen = sheet.frozen[frozenKey];
  if (index < frozen) sheet.frozen[frozenKey] = action === "insert" ? frozen + count : frozen - Math.min(count, frozen - index);
  if (sheet.hyperlinks) sheet.hyperlinks = sheet.hyperlinks.flatMap(link => { const range = rectangle(link.range); return range ? [{ ...link, range }] : []; });
  if (sheet.autoFilter) { const next = filter(sheet.autoFilter); if (next) sheet.autoFilter = next; else delete sheet.autoFilter; }
  if (sheet.print) {
    const print = sheet.print;
    print.areas = ranges(print.areas);
    const key = axis === "row" ? "repeatRows" : "repeatColumns", repeat = print[key];
    if (repeat) { const mapped = interval(repeat.from, repeat.to); if (mapped) print[key] = mapped; else delete print[key]; }
    for (const breakKey of ["rowBreaks", "columnBreaks"] as const) {
      const parallel = (axis === "row") === (breakKey === "rowBreaks");
      print[breakKey] = print[breakKey].flatMap(item => {
        if (parallel) { const before = point(item.before); return before === undefined || before === 0 ? [] : [{ ...item, before }]; }
        const mapped = interval(item.from, item.to); return mapped ? [{ ...item, ...mapped }] : [];
      });
    }
  }
  for (const annotation of annotations) if (annotation.anchor.kind === "cells" && annotation.anchor.sheetId === sheet.sheetId && annotation.status !== "needs_reanchor") {
    const mapped = mapWorkbookRanges(annotation.anchor.ranges.map(range => ({ sheetId: sheet.sheetId, range })), [step]);
    if (mapped.length) annotation.anchor.ranges = mapped.map(item => ({ from: { ...item.range.from }, to: { ...item.range.to } }));
    else annotation.status = "needs_reanchor";
  }
  mappings.push(step);
  yield { completed, total: moving.length };
}
