import type { WorkbookCoordinate, WorkbookDefinedName, WorkbookMetadata, WorkbookRange, WorkbookSheet } from "../contracts/workbook.js";
import type { DeepReadonly } from "../runtime.js";
import { checkLimit, freeze, invalid } from "../document/validation.js";
import { parseFormula } from "./formula.js";
import type { FormulaExpression } from "./formula.js";
import { bindStructuredReference } from "./formula-table-binding.js";
import type { FormulaTableDefinition } from "./formula-table-binding.js";
import { MAX_SHEET_COLUMNS, MAX_SHEET_ROWS } from "./chunk-index.js";
import { checkFormulaOrigin } from "./formula-print.js";

export interface WorkbookReferenceLocation extends WorkbookCoordinate { readonly sheetId: string }
export interface WorkbookReferenceArea extends WorkbookRange { readonly sheetId: string }
/** Static source binding is separate from formula evaluation. Unions preserve
 * order and multiplicity; an empty table body has an empty area list. */
export type WorkbookReferenceResult =
  | { readonly kind: "ranges"; readonly areas: readonly DeepReadonly<WorkbookReferenceArea>[] }
  | { readonly kind: "error"; readonly code: "#REF!" | "#NAME?" | "#VALUE!" | "#NULL!" }
  | { readonly kind: "unresolved"; readonly reason: "calculation-required" | "circular-name" | "ambiguous-name-scope" };
type Sheet = DeepReadonly<WorkbookSheet>;
type DefinedName = DeepReadonly<WorkbookDefinedName>;
/** A qualified global name has no implicit worksheet for bang references.
 * A local definition supplies its own worksheet. This state is derived only. */
export interface WorkbookNameContext { readonly inName?: boolean; readonly unboundSheet?: boolean }
const error = (code: Extract<WorkbookReferenceResult, { kind: "error" }>["code"]): WorkbookReferenceResult => ({ kind: "error", code });
const unresolved = (reason: Extract<WorkbookReferenceResult, { kind: "unresolved" }>["reason"]): WorkbookReferenceResult => ({ kind: "unresolved", reason });
const inside = (cell: WorkbookCoordinate, range: DeepReadonly<WorkbookRange>): boolean => range.from.row <= cell.row && cell.row <= range.to.row && range.from.column <= cell.column && cell.column <= range.to.column;
const key = (scope: string | undefined, name: string): string => JSON.stringify([scope ?? null, name.toUpperCase()]);
const ranges = (areas: readonly DeepReadonly<WorkbookReferenceArea>[]): WorkbookReferenceResult => { checkLimit(areas.length, 4096); return { kind: "ranges", areas }; };

/** One immutable workbook owns its indexes and lazy parsed-name cache. Nothing
 * here reads cell values, expands whole axes, or treats imported caches as valid. */
export class WorkbookReferences {
  private readonly sheetsByName: ReadonlyMap<string, Sheet>;
  private readonly sheetsById: ReadonlyMap<string, Sheet>;
  private readonly names: ReadonlyMap<string, DefinedName>;
  private readonly parsedNames = new Map<string, FormulaExpression>();
  private readonly tables: ReadonlyMap<string, FormulaTableDefinition>;
  constructor(workbook: DeepReadonly<WorkbookMetadata>) {
    this.sheetsByName = new Map(workbook.sheets.map(sheet => [sheet.name.toUpperCase(), sheet]));
    this.sheetsById = new Map(workbook.sheets.map(sheet => [sheet.sheetId, sheet]));
    this.names = new Map((workbook.names ?? []).map(name => [key(name.sheetId, name.name), name]));
    this.tables = new Map(workbook.sheets.flatMap(sheet => (sheet.tables ?? []).map(table => [table.name.toUpperCase(), { name: table.name, sheet: sheet.name, range: table.range, header: table.header, totals: table.totals, columns: table.columns.map(column => column.name) }] as const)));
  }
  resolve(expression: string, origin: WorkbookReferenceLocation): WorkbookReferenceResult {
    return this.resolveExpression(parseFormula(expression, origin), origin);
  }
  bindName(node: Extract<FormulaExpression, { kind: "name" }>, origin: WorkbookReferenceLocation, context: WorkbookNameContext = {}) {
    const explicit = node.sheet === undefined || node.sheet === "" ? undefined : this.sheetsByName.get(node.sheet.toUpperCase());
    if (node.sheet !== undefined && node.sheet !== "" && !explicit || node.sheet === "" && context.unboundSheet) return { kind: "error" as const, code: "#REF!" as const };
    const scope = explicit?.sheetId ?? origin.sheetId;
    // Native name formulas bind bare nested names to the workbook. Bang names
    // bind only locally; ordinary cell formulas try local then global.
    const definition = node.workbook === 0 || context.inName && node.sheet === undefined ? this.names.get(key(undefined, node.name))
      : this.names.get(key(scope, node.name)) ?? (node.sheet === "" ? undefined : this.names.get(key(undefined, node.name)));
    if (!definition) return undefined;
    return { kind: "name" as const, definition, location: { ...origin, sheetId: definition.sheetId ?? origin.sheetId },
      context: { inName: true, unboundSheet: definition.sheetId === undefined && (context.unboundSheet === true || node.workbook === 0 || explicit !== undefined) } };
  }
  resolveExpression(ast: FormulaExpression, origin: WorkbookReferenceLocation, context: WorkbookNameContext = {}): WorkbookReferenceResult {
    if (Object.keys(origin).some(name => !["sheetId", "row", "column"].includes(name)) || !this.sheetsById.has(origin.sheetId)) invalid("A reference query requires an existing worksheet and source coordinate.");
    checkFormulaOrigin(origin);
    const path = new Set<string>(); let operations = 0;
    function area(sheet: Sheet | undefined, from: WorkbookCoordinate, to = from): WorkbookReferenceResult {
      if (!sheet || [from, to].some(point => point.row < 0 || point.row >= MAX_SHEET_ROWS || point.column < 0 || point.column >= MAX_SHEET_COLUMNS)) return error("#REF!");
      return ranges([{ sheetId: sheet.sheetId, from: { row: Math.min(from.row, to.row), column: Math.min(from.column, to.column) }, to: { row: Math.max(from.row, to.row), column: Math.max(from.column, to.column) } }]);
    }
    const visit = (node: FormulaExpression, context: WorkbookReferenceLocation, depth: number, nameContext: WorkbookNameContext = {}): WorkbookReferenceResult => {
      checkLimit(++operations, 100_000); checkLimit(depth, 128);
      const sheetFor = (name: string | undefined): Sheet | undefined => name === "" && nameContext.unboundSheet ? undefined
        : name === undefined || name === "" ? this.sheetsById.get(context.sheetId) : this.sheetsByName.get(name.toUpperCase());
      const coordinate = (value: number, absolute: boolean, axis: "row" | "column"): number => {
        const result = value + (absolute ? 0 : context[axis]), bound = axis === "row" ? MAX_SHEET_ROWS : MAX_SHEET_COLUMNS;
        return nameContext.inName && !absolute ? (result % bound + bound) % bound : result;
      };
      const cell = (value: Extract<FormulaExpression, { kind: "cell" }>): WorkbookCoordinate => ({ row: coordinate(value.row, value.absoluteRow, "row"), column: coordinate(value.column, value.absoluteColumn, "column") });
      if (node.kind === "cell") return area(sheetFor(node.sheet), cell(node));
      if (node.kind === "range") {
        const first = sheetFor(node.from.sheet), last = sheetFor(node.to.sheet);
        if (first !== last) return error("#REF!");
        return area(first, cell(node.from), cell(node.to));
      }
      if (node.kind === "axisRange") {
        const from = coordinate(node.from.index, node.from.absolute, node.axis), to = coordinate(node.to.index, node.to.absolute, node.axis);
        return area(sheetFor(node.sheet), node.axis === "row" ? { row: from, column: 0 } : { row: 0, column: from }, node.axis === "row" ? { row: to, column: MAX_SHEET_COLUMNS - 1 } : { row: MAX_SHEET_ROWS - 1, column: to });
      }
      if (node.kind === "error") return node.value === "#REF!" ? error("#REF!") : unresolved("calculation-required");
      if (node.kind === "structured") {
        const sheet = this.sheetsById.get(context.sheetId)!;
        if (!node.table) { operations += sheet.tables?.length ?? 0; checkLimit(operations, 100_000); }
        const containing = node.table ? undefined : (sheet.tables ?? []).find(table => inside(context, table.range));
        if (!node.table && !containing) return error("#VALUE!");
        const table = this.tables.get((node.table ?? containing!.name).toUpperCase()); if (!table) return error("#NAME?");
        operations += table.columns.length; checkLimit(operations, 100_000);
        const bound = bindStructuredReference(node, { sheet: sheet.name, row: context.row, column: context.column }, [table], containing?.name);
        return "error" in bound ? error(bound.error) : bound.range ? area(this.sheetsByName.get(bound.sheet.toUpperCase()), bound.range.from, bound.range.to) : ranges([]);
      }
      if (node.kind === "name") {
        const bound = this.bindName(node, context, nameContext);
        if (bound?.kind === "error") return error(bound.code);
        if (!bound) {
          if (node.sheet === undefined && this.tables.has(node.name.toUpperCase())) return visit({ kind: "structured", table: node.name, selectors: [], columns: [] }, context, depth + 1, nameContext);
          return error("#NAME?");
        }
        const name = bound.definition;
        const identity = JSON.stringify([name.nameId, bound.location.sheetId, bound.context.unboundSheet]);
        if (path.has(identity)) return unresolved("circular-name");
        let parsed = this.parsedNames.get(name.nameId);
        if (!parsed) { parsed = parseFormula(name.expression, name.origin); this.parsedNames.set(name.nameId, parsed); }
        path.add(identity);
        const result = visit(parsed, bound.location, depth + 1, bound.context);
        path.delete(identity); return result;
      }
      if (node.kind === "binary" && [",", " ", ":"].includes(node.operator)) {
        const left = visit(node.left, context, depth + 1, nameContext); if (left.kind !== "ranges") return left;
        const right = visit(node.right, context, depth + 1, nameContext); if (right.kind !== "ranges") return right;
        if (node.operator === ",") return ranges([...left.areas, ...right.areas]);
        if (node.operator === ":") {
          const combined = [...left.areas, ...right.areas], first = combined[0];
          if (!left.areas.length || !right.areas.length || !first) return error("#REF!");
          if (combined.some(value => value.sheetId !== first.sheetId)) return error("#VALUE!");
          return area(this.sheetsById.get(first.sheetId), { row: Math.min(...combined.map(value => value.from.row)), column: Math.min(...combined.map(value => value.from.column)) }, { row: Math.max(...combined.map(value => value.to.row)), column: Math.max(...combined.map(value => value.to.column)) });
        }
        const result: WorkbookReferenceArea[] = [];
        for (const a of left.areas) for (const b of right.areas) {
          checkLimit(++operations, 100_000); if (a.sheetId !== b.sheetId) continue;
          const from = { row: Math.max(a.from.row, b.from.row), column: Math.max(a.from.column, b.from.column) }, to = { row: Math.min(a.to.row, b.to.row), column: Math.min(a.to.column, b.to.column) };
          if (from.row <= to.row && from.column <= to.column) { checkLimit(result.length + 1, 4096); result.push({ sheetId: a.sheetId, from, to }); }
        }
        return result.length ? ranges(result) : error("#NULL!");
      }
      return unresolved("calculation-required");
    };
    return freeze(visit(ast, origin, 0, context));
  }
}

const references = new WeakMap<DeepReadonly<WorkbookMetadata>, WorkbookReferences>();
/** Immutable revisions share one bounded-source index, never values or caches
 * across documents. Weak ownership does not retain discarded history. */
export function workbookReferences(workbook: DeepReadonly<WorkbookMetadata>): WorkbookReferences {
  let result = references.get(workbook);
  if (!result) { result = new WorkbookReferences(workbook); references.set(workbook, result); }
  return result;
}
