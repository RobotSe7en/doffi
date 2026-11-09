import type { Annotation } from "../contracts/annotations.js";
import type { WorkbookCoordinate, WorkbookDefinedName, WorkbookMetadata } from "../contracts/workbook.js";
import type { WorkbookChartReference } from "../contracts/workbook-charts.js";
import type { WorkbookOperation } from "../contracts/workbook-operations.js";
import type { WorkbookModel } from "../runtime.js";
import { DocumentError } from "../errors.js";
import { parseFormula, type FormulaExpression } from "../workbook/formula.js";
import { rewriteFormula, type FormulaRewriteBindings } from "../workbook/formula-rewrite.js";
import { WorkbookCellEdit } from "./workbook-cells.js";
import { WorkbookContentEdit } from "./workbook-content.js";
import { cloneWorkbookDrawing } from "./workbook-drawings.js";
import type { SourceStyleEdit } from "./source-styles.js";
import { createWorkbookCommentCloner } from "./workbook-comment-copy.js";

const copy = <T>(value: T): T => JSON.parse(JSON.stringify(value)) as T;
const upper = (value: string) => value.toUpperCase();
function invalid(message: string): never { throw new DocumentError("InvalidOperation", message); }
/** A worksheet copy owns fresh source identities inside the same candidate.
 * Outside consumers, source chunks, names and native provenance stay intact. */
export function* copyWorkbookSheet(workbook: WorkbookMetadata, cells: WorkbookCellEdit, content: WorkbookContentEdit, annotations: Annotation[], base: WorkbookModel,
  operation: Extract<WorkbookOperation, { op: "copyWorkbookSheet" }>, allocate: () => string, styles: SourceStyleEdit): Generator<{ completed: number; total: number }, void> {
  const source = workbook.sheets.find(sheet => sheet.sheetId === operation.sheetId) ?? invalid("An earlier operation removed the copied worksheet.");
  if (operation.name.startsWith("'") || operation.name.endsWith("'") || workbook.sheets.some(sheet => upper(sheet.name) === upper(operation.name))) invalid("The copied worksheet requires a valid unused name.");
  const before = operation.beforeSheetId === undefined ? undefined : workbook.sheets.find(sheet => sheet.sheetId === operation.beforeSheetId) ?? invalid("The worksheet insertion target no longer exists.");
  let work = 0, count = 0;
  for (const _cell of cells.entries(source.sheetId)) { count++; if (++work % 256 === 0) yield { completed: work, total: 0 }; }
  cells.checkAdditionalCells(count);
  const definitions = [...workbook.names ?? []], scopes = new Map<string | undefined, Map<string, WorkbookDefinedName>>(), sheets = new Map(workbook.sheets.map(sheet => [upper(sheet.name), sheet.sheetId]));
  for (const definition of definitions) { const scope = scopes.get(definition.sheetId) ?? new Map<string, WorkbookDefinedName>(); scope.set(upper(definition.name), definition); scopes.set(definition.sheetId, scope); if (++work % 256 === 0) yield { completed: work, total: 0 }; }
  const globals = scopes.get(undefined), locals = scopes.get(source.sheetId), tableSheets = new Map(workbook.sheets.flatMap(sheet => (sheet.tables ?? []).map(table => [upper(table.name), sheet.sheetId] as const)));
  const tableObjectTokens = new Set(workbook.sheets.flatMap(sheet => (sheet.tables ?? []).map(table => upper(table.objectName ?? table.name))));
  const resolve = (name: string, owner: string, explicit: boolean, book: 0 | undefined, named: boolean, localOnly = false) => {
    if (book === 0) return globals?.get(upper(name));
    const sheetId = sheets.get(upper(owner)), local = sheetId === undefined ? undefined : scopes.get(sheetId)?.get(upper(name));
    return named && !explicit ? globals?.get(upper(name)) : local ?? (localOnly ? undefined : globals?.get(upper(name)));
  };
  const selected = new Set([...locals?.values() ?? []].map(name => name.nameId)), reverse = new Map<string, Set<string>>();
  // A global definition is copied only when its references reach an identity
  // that is being copied. Work through aliases and cycles without evaluation.
  for (const definition of globals?.values() ?? []) {
    const pending: FormulaExpression[] = [parseFormula(definition.expression, definition.origin)]; let direct = false;
    while (pending.length) {
      const node = pending.pop()!; if (++work % 256 === 0) yield { completed: work, total: 0 };
      if ((node.kind === "cell" || node.kind === "axisRange") && node.sheet !== undefined && upper(node.sheet) === upper(source.name)) direct = true;
      else if (node.kind === "range") pending.push(node.from, node.to);
      else if (node.kind === "structured" && node.table !== undefined && tableSheets.get(upper(node.table)) === source.sheetId) direct = true;
      else if (node.kind === "name") {
        const dependency = resolve(node.name, node.sheet || source.name, node.sheet !== undefined, node.workbook, true, node.sheet === "");
        if (dependency && (dependency.sheetId === undefined || dependency.sheetId === source.sheetId)) {
          const consumers = reverse.get(dependency.nameId) ?? new Set<string>(); consumers.add(definition.nameId); reverse.set(dependency.nameId, consumers);
        } else if (!dependency && node.sheet === undefined && tableSheets.get(upper(node.name)) === source.sheetId) direct = true;
      } else if (node.kind === "binary") pending.push(node.left, node.right);
      else if (node.kind === "unary") pending.push(node.operand);
      else if (node.kind === "call") for (const argument of node.arguments) if (argument) pending.push(argument);
    }
    if (direct) selected.add(definition.nameId);
  }
  const queue = [...selected]; for (let at = 0; at < queue.length; at++) for (const consumer of reverse.get(queue[at]!) ?? []) {
    if (!selected.has(consumer)) { selected.add(consumer); queue.push(consumer); } if (++work % 256 === 0) yield { completed: work, total: 0 };
  }
  const copiedNames = new Map<string, string>(), choices = new Map<string, string>(), allTokens = new Set([...definitions.map(name => upper(name.name)), ...tableSheets.keys()]);
  for (const choice of operation.nameConflicts ?? []) {
    if (choices.has(choice.nameId)) invalid("A copied-name conflict cannot have repeated choices."); choices.set(choice.nameId, choice.name);
  }
  const usedLocal = new Set<string>();
  for (const definition of definitions.filter(name => name.sheetId === source.sheetId)) { copiedNames.set(definition.nameId, definition.name); usedLocal.add(upper(definition.name)); }
  for (const definition of globals?.values() ?? []) if (selected.has(definition.nameId)) {
    let name = definition.name;
    if (usedLocal.has(upper(name))) {
      const resolved = choices.get(definition.nameId);
      if (resolved === undefined) throw new DocumentError("InvalidOperation", "The copied global name conflicts with a local name; choose an unused name explicitly.", [{ code: "copy-name-conflict", path: [definition.nameId], message: definition.name }]);
      const parsed = parseFormula("=" + resolved, definition.origin);
      if (parsed.kind !== "name" || parsed.sheet !== undefined || parsed.workbook !== undefined || parsed.name !== resolved || allTokens.has(upper(resolved))) invalid("A copied name conflict requires an unused unqualified name.");
      name = resolved; choices.delete(definition.nameId); allTokens.add(upper(name));
    }
    if (usedLocal.has(upper(name))) invalid("Copied local names must remain unique.");
    copiedNames.set(definition.nameId, name); usedLocal.add(upper(name));
  }
  if (choices.size) invalid("Name choices must refer only to actual copied global/local conflicts.");
  const added = copy(source), copiedTables = new Map<string, string>(); added.sheetId = allocate(); added.name = operation.name;
  for (const table of added.tables ?? []) {
    const original = table.name; let serial = 2, name: string;
    do { const suffix = "_" + serial++; name = original.slice(0, 255 - suffix.length) + suffix; } while (allTokens.has(upper(name)) || tableObjectTokens.has(upper(name)));
    allTokens.add(upper(name)); tableObjectTokens.add(upper(name)); copiedTables.set(upper(original), name); table.name = name; if (table.objectName !== undefined) table.objectName = name;
    table.tableId = allocate(); for (const column of table.columns) column.columnId = allocate();
    if (++work % 256 === 0) yield { completed: work, total: 0 };
  }
  for (const rule of [...added.conditionalRules ?? [], ...added.dataValidations ?? []]) { rule.ruleId = allocate(); if (++work % 256 === 0) yield { completed: work, total: 0 }; }
  const rewrite = (expression: string, origin: WorkbookCoordinate, named = false) => {
    const bindings: FormulaRewriteBindings = { copiedNames, copiedTables, definedName: named, resolveName: (name, owner, explicit, book, localOnly) => resolve(name, owner, explicit, book, named, localOnly)?.nameId,
      isTableName: (name, owner, book) => !globals?.has(upper(name)) && (book === 0 || named || !scopes.get(sheets.get(upper(owner)))?.has(upper(name))) && tableSheets.has(upper(name)) };
    return rewriteFormula(expression, { ...origin, sheet: source.name }, { ...origin, sheet: added.name }, { kind: "copySheet", from: source.name, to: added.name }, bindings);
  };
  for (const definition of definitions) if (copiedNames.has(definition.nameId)) {
    const value = copy(definition); value.nameId = allocate(); value.name = copiedNames.get(definition.nameId)!; value.sheetId = added.sheetId;
    value.expression = rewrite(definition.expression, definition.origin, true); (workbook.names ??= []).push(value);
    if (++work % 256 === 0) yield { completed: work, total: 0 };
  }
  for (const rule of [...added.conditionalRules ?? [], ...added.dataValidations ?? []]) {
    if ("expression" in rule) rule.expression = rewrite(rule.expression, rule.origin);
    if ("formulas" in rule) rule.formulas = rule.formulas.map(formula => rewrite(formula, rule.origin));
    if (rule.kind === "colorScale") for (const stop of rule.stops) if (stop.kind === "formula") stop.expression = rewrite(stop.expression, rule.origin);
    if (++work % 256 === 0) yield { completed: work, total: 0 };
  }
  for (const table of added.tables ?? []) for (const [index, column] of table.columns.entries()) {
    if (column.calculatedFormula) column.calculatedFormula = rewrite(column.calculatedFormula, { row: table.range.from.row + Number(table.header), column: table.range.from.column + index });
    if (column.totalFormula) column.totalFormula = rewrite(column.totalFormula, { row: table.range.to.row, column: table.range.from.column + index });
    if (++work % 256 === 0) yield { completed: work, total: 0 };
  }
  // Hyperlink destinations are stored strings, not formula dependencies.
  // Qualified locations retain their original target in a worksheet copy;
  // an unqualified location naturally resolves in its new owning sheet.
  workbook.sheets.splice(before ? workbook.sheets.indexOf(before) : workbook.sheets.length, 0, added);
  for (const cell of cells.entries(source.sheetId)) {
    cells.relocate(cell, cell.row, cell.column, added.sheetId);
    if (cell.formula) cells.rewriteFormula(added.sheetId, cell.row, cell.column, rewrite(cell.formula.expression, cell), false);
    if (++work % 256 === 0) yield { completed: work, total: 0 };
  }
  const identities = new Map<string, string>();
  const reference = (value: WorkbookChartReference | undefined): void => {
    if (!value) return; value.expression = rewrite(value.expression, { row: 0, column: 0 });
    // Unqualified references acquire a new owning worksheet even if their
    // spelling is unchanged. Their historical vectors are never current data.
    if (value.cache) value.cache.state = "stale";
  };
  for (const original of [...workbook.objects ?? []].filter(object => object.sheetId === source.sheetId)) {
    const cloned = cloneWorkbookDrawing(workbook, content, original, allocate); cloned.object.sheetId = added.sheetId;
    for (const [from, to] of cloned.identities) identities.set(from, to);
    if (cloned.object.kind === "chart") {
      const chart = workbook.charts!.find(chart => chart.nodeId === cloned.object.nodeId)!;
      if (chart.categories && "expression" in chart.categories) reference(chart.categories);
      reference(chart.titleReference); reference(chart.categoryAxisTitle?.reference); reference(chart.valueAxisTitle?.reference);
      for (const series of chart.series) {
        if ("expression" in series.values) reference(series.values); if (series.name?.kind === "reference") reference(series.name.reference);
        if (++work % 256 === 0) yield { completed: work, total: 0 };
      }
    }
    yield { completed: ++work, total: 0 };
  }
  const cloneComment = createWorkbookCommentCloner(workbook, content, allocate);
  for (const original of [...annotations]) if (original.anchor.kind === "cells" && original.anchor.sheetId === source.sheetId) {
    const cloned = cloneComment(original, { ...original.anchor, sheetId: added.sheetId }, base.snapshot().manifest.revision);
    for (const [from, to] of cloned.identities) identities.set(from, to);
    annotations.push(cloned.annotation); yield { completed: ++work, total: 0 };
  }
  styles.aliasNodeIds(identities);
  if (added.visibility === "visible") workbook.activeSheetId = added.sheetId;
  yield { completed: work, total: work };
}
