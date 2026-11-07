import { WorkbookSheetSchema } from "../contracts/workbook.js";
import { checkLimit, invalid } from "../document/validation.js";
import { MAX_SHEET_COLUMNS, MAX_SHEET_ROWS } from "./chunk-index.js";
import { parseFormula, type FormulaCell, type FormulaExpression, type FormulaOrigin } from "./formula.js";
import { checkFormulaOrigin, printFormula } from "./formula-print.js";
import { bindStructuredReference, type FormulaTableDefinition } from "./formula-table-binding.js";

export interface FormulaLocation extends FormulaOrigin { readonly sheet: string }
export interface FormulaRectangle { readonly from: FormulaOrigin; readonly to: FormulaOrigin }
export type FormulaReferenceChange =
  | { readonly kind: "axis"; readonly sheet: string; readonly axis: "row" | "column"; readonly action: "insert" | "delete"; readonly index: number; readonly count: number }
  | { readonly kind: "move"; readonly sheet: string; readonly range: FormulaRectangle; readonly target: FormulaLocation }
  | { readonly kind: "sort"; readonly sheet: string; readonly range: FormulaRectangle }
  | { readonly kind: "renameSheet"; readonly from: string; readonly to: string }
  | { readonly kind: "deleteSheet"; readonly sheet: string }
  | { readonly kind: "renameTable"; readonly from: string; readonly to: string }
  | { readonly kind: "renameColumn"; readonly table: string; readonly from: string; readonly to: string };
export interface FormulaRewriteBindings {
  /** The table containing this formula, for unqualified structured references. */
  readonly table?: string;
  /** Required only when a bare identifier might name the table being renamed. */
  readonly isTableName?: (name: string, sheet: string) => boolean;
  readonly tables?: readonly FormulaTableDefinition[];
}
const same = (a: string, b: string): boolean => a.toLowerCase() === b.toLowerCase();
const refError = (): FormulaExpression => ({ kind: "error", value: "#REF!" });
const inside = (point: FormulaOrigin, range: FormulaRectangle): boolean => point.row >= range.from.row && point.row <= range.to.row && point.column >= range.from.column && point.column <= range.to.column;
function checkRectangle(range: FormulaRectangle): void {
  checkFormulaOrigin(range.from); checkFormulaOrigin(range.to);
  if (range.from.row > range.to.row || range.from.column > range.to.column) invalid("Invalid formula mapping rectangle.");
}
function intersect(a: FormulaRectangle, b: FormulaRectangle): FormulaRectangle | undefined {
  const from = { row: Math.max(a.from.row, b.from.row), column: Math.max(a.from.column, b.from.column) };
  const to = { row: Math.min(a.to.row, b.to.row), column: Math.min(a.to.column, b.to.column) };
  return from.row <= to.row && from.column <= to.column ? { from, to } : undefined;
}
/** At most four rectangles, independent of the number of covered cells. */
function subtract(a: FormulaRectangle, b: FormulaRectangle): FormulaRectangle[] {
  const overlap = intersect(a, b); if (!overlap) return [a];
  const pieces: FormulaRectangle[] = [];
  if (a.from.row < overlap.from.row) pieces.push({ from: a.from, to: { row: overlap.from.row - 1, column: a.to.column } });
  if (overlap.to.row < a.to.row) pieces.push({ from: { row: overlap.to.row + 1, column: a.from.column }, to: a.to });
  if (a.from.column < overlap.from.column) pieces.push({ from: { row: overlap.from.row, column: a.from.column }, to: { row: overlap.to.row, column: overlap.from.column - 1 } });
  if (overlap.to.column < a.to.column) pieces.push({ from: { row: overlap.from.row, column: overlap.to.column + 1 }, to: { row: overlap.to.row, column: a.to.column } });
  return pieces;
}

export function copyFormula(source: string, from: FormulaOrigin, to: FormulaOrigin): string {
  const ast = parseFormula(source, from); checkFormulaOrigin(to);
  if (from.row === to.row && from.column === to.column) return source;
  const output = printFormula(ast, to); parseFormula(output, to); return output;
}

/** Rewrite semantic references, never strings that happen to contain addresses. */
export function rewriteFormula(source: string, from: FormulaLocation, to: FormulaLocation, change: FormulaReferenceChange, bindings: FormulaRewriteBindings = {}): string {
  const ast = parseFormula(source, from); checkFormulaOrigin(to);
  if (change.kind === "sort" || change.kind === "move") checkRectangle(change.range);
  if (change.kind === "sort") {
    if (!same(from.sheet, change.sheet) || !inside(from, change.range)) {
      if (!same(from.sheet, to.sheet) || from.row !== to.row || from.column !== to.column) invalid("A sort cannot move an outside formula.");
      return source;
    }
    if (!same(to.sheet, change.sheet) || !inside(to, change.range)) invalid("Sorted formula moved outside its range.");
    return copyFormula(source, from, to);
  }
  if (change.kind === "move") {
    checkFormulaOrigin(change.target);
    checkFormulaOrigin({ row: change.target.row + change.range.to.row - change.range.from.row, column: change.target.column + change.range.to.column - change.range.from.column });
  }
  if (change.kind === "axis") {
    const bound = change.axis === "row" ? MAX_SHEET_ROWS : MAX_SHEET_COLUMNS;
    if (!Number.isInteger(change.index) || !Number.isInteger(change.count) || change.index < 0 || change.count <= 0 || change.index + change.count > bound) invalid("Invalid formula axis change.");
  }
  if (change.kind === "renameSheet" && !WorkbookSheetSchema.shape.name.safeParse(change.to).success) invalid("Invalid renamed sheet.");
  if (change.kind === "renameTable") {
    if (change.to.length > 255 || parseFormula("=" + change.to, from).kind !== "name") invalid("Invalid renamed table.");
  }
  if (change.kind === "renameColumn" && (!change.to || change.to.length > 255 || change.to.trim() !== change.to || /[\u0000-\u001f]/.test(change.to))) invalid("Invalid renamed table column.");
  const location = (cell: FormulaCell): FormulaLocation => ({ row: cell.row + (cell.absoluteRow ? 0 : from.row), column: cell.column + (cell.absoluteColumn ? 0 : from.column), sheet: cell.sheet ?? from.sheet });
  const cellAt = (point: FormulaLocation, template: FormulaCell): FormulaCell => ({ kind: "cell", row: point.row - (template.absoluteRow ? 0 : to.row), column: point.column - (template.absoluteColumn ? 0 : to.column),
    absoluteRow: template.absoluteRow, absoluteColumn: template.absoluteColumn, ...(template.sheet === undefined && same(point.sheet, to.sheet) ? {} : { sheet: point.sheet }) });
  function mapSheet(sheet: string): string | undefined {
    if (change.kind === "deleteSheet" && same(sheet, change.sheet)) return undefined;
    return change.kind === "renameSheet" && same(sheet, change.from) ? change.to : sheet;
  }
  function mapPosition(point: FormulaLocation): FormulaLocation | undefined {
    const sheet = mapSheet(point.sheet); if (sheet === undefined) return undefined;
    if (change.kind === "axis" && same(point.sheet, change.sheet)) {
      const value = point[change.axis];
      if (change.action === "delete" && value >= change.index && value < change.index + change.count) return undefined;
      return { ...point, sheet, [change.axis]: value < change.index ? value : value + (change.action === "insert" ? change.count : -change.count) };
    }
    if (change.kind === "move" && same(point.sheet, change.sheet) && inside(point, change.range)) return { sheet: change.target.sheet, row: point.row + change.target.row - change.range.from.row, column: point.column + change.target.column - change.range.from.column };
    return { ...point, sheet };
  }
  function interval(start: number, end: number): [number, number] | undefined {
    if (change.kind !== "axis") return [start, end];
    const low = Math.min(start, end), high = Math.max(start, end);
    let first = low, last = high;
    if (change.action === "insert") {
      if (low >= change.index) first += change.count;
      if (high >= change.index) last += change.count;
    } else {
      if (high < change.index) return [start, end];
      if (low >= change.index + change.count) { first -= change.count; last -= change.count; }
      else {
        first = low < change.index ? low : change.index;
        last = high >= change.index + change.count ? high - change.count : change.index - 1;
        if (last < first) return undefined;
      }
    }
    return start <= end ? [first, last] : [last, first];
  }
  function rectangleExpression(range: FormulaRectangle, sheet: string, first: FormulaCell, last: FormulaCell): FormulaExpression {
    const a = cellAt({ ...range.from, sheet }, first), b = cellAt({ ...range.to, sheet }, last);
    return range.from.row === range.to.row && range.from.column === range.to.column ? a : { kind: "range", from: a, to: b };
  }
  function range(node: Extract<FormulaExpression, { kind: "range" }>): FormulaExpression {
    const first = location(node.from), last = location(node.to), sheet = mapSheet(first.sheet);
    if (sheet === undefined) return refError();
    if (change.kind === "axis" && same(first.sheet, change.sheet)) {
      const mapped = interval(first[change.axis], last[change.axis]); if (!mapped) return refError();
      return { kind: "range", from: cellAt({ ...first, sheet, [change.axis]: mapped[0] }, node.from), to: cellAt({ ...last, sheet, [change.axis]: mapped[1] }, node.to) };
    }
    if (change.kind === "move" && same(first.sheet, change.sheet)) {
      const rectangle = { from: { row: Math.min(first.row, last.row), column: Math.min(first.column, last.column) }, to: { row: Math.max(first.row, last.row), column: Math.max(first.column, last.column) } };
      const overlap = intersect(rectangle, change.range);
      if (overlap) {
        const fragments = subtract(rectangle, change.range).map(value => ({ range: value, sheet }));
        const delta = { row: change.target.row - change.range.from.row, column: change.target.column - change.range.from.column };
        const moved = { from: { row: overlap.from.row + delta.row, column: overlap.from.column + delta.column }, to: { row: overlap.to.row + delta.row, column: overlap.to.column + delta.column } };
        // Overlapping destinations must not introduce a duplicated reference.
        let additions = [moved];
        for (const fragment of fragments) if (same(fragment.sheet, change.target.sheet)) additions = additions.flatMap(value => subtract(value, fragment.range));
        fragments.push(...additions.map(value => ({ range: value, sheet: change.target.sheet })));
        checkLimit(fragments.length, 32);
        const expressions = fragments.map(value => rectangleExpression(value.range, value.sheet, node.from, node.to));
        return expressions.slice(1).reduce<FormulaExpression>((left, right) => ({ kind: "binary", operator: ",", left, right }), expressions[0]!);
      }
    }
    return { kind: "range", from: cellAt({ ...first, sheet }, node.from), to: cellAt({ ...last, sheet }, node.to) };
  }
  function visit(node: FormulaExpression): FormulaExpression {
    if (node.kind === "cell") { const point = mapPosition(location(node)); return point ? cellAt(point, node) : refError(); }
    if (node.kind === "range") return range(node);
    if (node.kind === "axisRange") {
      const sourceSheet = node.sheet ?? from.sheet, sheet = mapSheet(sourceSheet); if (sheet === undefined) return refError();
      const start = node.from.index + (node.from.absolute ? 0 : from[node.axis]), end = node.to.index + (node.to.absolute ? 0 : from[node.axis]);
      if (change.kind === "move" && same(sourceSheet, change.sheet)) {
        const cell = (value: number, end: boolean): FormulaCell => {
          const endpoint = end ? node.to : node.from, absoluteRow = node.axis === "row" ? endpoint.absolute : true, absoluteColumn = node.axis === "column" ? endpoint.absolute : true;
          return { kind: "cell", row: (node.axis === "row" ? value : end ? MAX_SHEET_ROWS - 1 : 0) - (absoluteRow ? 0 : from.row),
            column: (node.axis === "column" ? value : end ? MAX_SHEET_COLUMNS - 1 : 0) - (absoluteColumn ? 0 : from.column), absoluteRow, absoluteColumn, sheet: sourceSheet };
        };
        const rectangle = { from: { row: node.axis === "row" ? Math.min(start, end) : 0, column: node.axis === "column" ? Math.min(start, end) : 0 }, to: { row: node.axis === "row" ? Math.max(start, end) : MAX_SHEET_ROWS - 1, column: node.axis === "column" ? Math.max(start, end) : MAX_SHEET_COLUMNS - 1 } };
        if (intersect(rectangle, change.range)) return range({ kind: "range", from: cell(Math.min(start, end), false), to: cell(Math.max(start, end), true) });
      }
      const positions = change.kind === "axis" && same(sourceSheet, change.sheet) && node.axis === change.axis ? interval(start, end) : [start, end];
      if (!positions) return refError();
      return { ...node, from: { ...node.from, index: positions[0]! - (node.from.absolute ? 0 : to[node.axis]) }, to: { ...node.to, index: positions[1]! - (node.to.absolute ? 0 : to[node.axis]) },
        ...(node.sheet === undefined && same(sheet, to.sheet) ? {} : { sheet }) };
    }
    if (node.kind === "name") {
      const sheet = mapSheet(node.sheet ?? from.sheet); if (sheet === undefined) return refError();
      if (change.kind === "deleteSheet" && node.sheet === undefined && bindings.isTableName?.(node.name, from.sheet)
        && bindings.tables?.some(table => same(table.name, node.name) && same(table.sheet, change.sheet))) return refError();
      if (change.kind === "renameTable" && same(node.name, change.from)) {
        if (!bindings.isTableName) invalid("Bare table rename requires resolved name bindings.");
        if (bindings.isTableName(node.name, node.sheet ?? from.sheet)) return { ...node, name: change.to };
      }
      return { ...node, ...(node.sheet === undefined && same(sheet, to.sheet) ? {} : { sheet }) };
    }
    if (node.kind === "structured") {
      const table = node.table ?? bindings.table;
      if (change.kind === "deleteSheet" && table !== undefined && bindings.tables?.some(definition => same(definition.name, table) && same(definition.sheet, change.sheet))) return refError();
      if (change.kind === "move" && (!same(from.sheet, to.sheet) || from.row !== to.row || from.column !== to.column)) {
        if (node.selectors.includes("thisRow")) {
          if (!bindings.tables) invalid("Moving a current-row table formula requires table geometry.");
          const bound = bindStructuredReference(node, from, bindings.tables, bindings.table);
          if ("error" in bound) return { kind: "error", value: bound.error };
          if (!bound.range) return refError();
          const cell = (point: FormulaOrigin): FormulaCell => ({ kind: "cell", row: point.row - from.row, column: point.column, absoluteRow: false, absoluteColumn: true,
            ...(same(bound.sheet, from.sheet) ? {} : { sheet: bound.sheet }) });
          return range({ kind: "range", from: cell(bound.range.from), to: cell(bound.range.to) });
        }
        if (node.table === undefined) {
          if (table === undefined) invalid("Moving an unqualified table formula requires its containing table.");
          return { ...node, table };
        }
      }
      if (change.kind === "renameTable" && node.table !== undefined && same(node.table, change.from)) return { ...node, table: change.to };
      if (change.kind === "renameColumn") {
        if (table === undefined) invalid("Unqualified table column rename requires a containing table binding.");
        if (same(table, change.table)) return { ...node, columns: node.columns.map(column => same(column, change.from) ? change.to : column) };
      }
      return node;
    }
    if (node.kind === "binary") return { ...node, left: visit(node.left), right: visit(node.right) };
    if (node.kind === "unary") return { ...node, operand: visit(node.operand) };
    if (node.kind === "call") return { ...node, arguments: node.arguments.map(argument => argument === null ? null : visit(argument)) };
    return node;
  }
  const rewritten = visit(ast);
  // Unaffected formulas retain their exact source bytes, so a small structural
  // edit does not rewrite every workbook chunk merely to canonicalize syntax.
  if (from.row === to.row && from.column === to.column && JSON.stringify(rewritten) === JSON.stringify(ast)) return source;
  const output = printFormula(rewritten, to);
  // The parser rechecks generated identifiers, grammar, depth and token budgets.
  parseFormula(output, to); return output;
}
