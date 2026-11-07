import type { DeepReadonly, WorkbookModel } from "../runtime.js";
import type { FormulaExpression } from "./formula.js";
import { parseFormula } from "./formula.js";
import { printFormula } from "./formula-print.js";
import type { WorkbookReferenceArea, WorkbookReferenceLocation } from "./reference.js";

/** Shared, bounded source-reference discovery for calculation and validation.
 * Reference rectangles stay compact; function evaluation remains separate. */
export function workbookExpressionDependencies(model: WorkbookModel, expression: FormulaExpression, origin: WorkbookReferenceLocation,
  step: () => void, add: (area: DeepReadonly<WorkbookReferenceArea>) => void, call?: (name: string) => void,
  resolve: WorkbookModel["resolveReference"] = (expression, at) => model.resolveReference(expression, at)): void {
  const workbook = model.snapshot().metadata.workbook, definitions = workbook.names ?? [], sheets = workbook.sheets;
  const seen = new Set<string>();
  const emit = (area: DeepReadonly<WorkbookReferenceArea>): void => {
    const identity = JSON.stringify([area.sheetId, area.from.row, area.from.column, area.to.row, area.to.column]);
    if (!seen.has(identity)) { seen.add(identity); add(area); }
  };
  const pending: { ast: FormulaExpression; at: WorkbookReferenceLocation; names: readonly string[] }[] = [{ ast: expression, at: origin, names: [] }];
  while (pending.length) {
    step(); const { ast, at, names } = pending.pop()!;
    if (ast.kind === "call") {
      const name = ast.name.toUpperCase().replace(/^_XLFN\./, ""); call?.(name);
      if ((name === "SUMIF" || name === "AVERAGEIF") && ast.arguments.length === 3 && ast.arguments[0] && ast.arguments[2]) {
        const criteria = resolve(printFormula(ast.arguments[0], at), at), sum = resolve(printFormula(ast.arguments[2], at), at);
        if (criteria.kind === "ranges" && criteria.areas.length === 1 && sum.kind === "ranges" && sum.areas.length === 1) {
          const a = criteria.areas[0]!, b = sum.areas[0]!, to = { row: b.from.row + a.to.row - a.from.row, column: b.from.column + a.to.column - a.from.column };
          if (to.row < 1_048_576 && to.column < 16_384) emit({ sheetId: b.sheetId, from: b.from, to });
        }
      }
      if (name === "SUBTOTAL") for (const sheet of sheets) for (const filter of [...sheet.autoFilter ? [sheet.autoFilter] : [], ...sheet.tables?.flatMap(table => table.autoFilter ? [table.autoFilter] : []) ?? []]) {
        for (const column of filter.columns) if (column.criterion && filter.range.from.row < filter.range.to.row) {
          step(); emit({ sheetId: sheet.sheetId, from: { row: filter.range.from.row + 1, column: column.column }, to: { row: filter.range.to.row, column: column.column } });
        }
      }
      for (const argument of ast.arguments) if (argument) pending.push({ ast: argument, at, names });
    } else if (ast.kind === "unary") pending.push({ ast: ast.operand, at, names });
    else if (ast.kind === "binary" && ![",", " ", ":"].includes(ast.operator)) pending.push({ ast: ast.left, at, names }, { ast: ast.right, at, names });
    else if (["cell", "range", "axisRange", "structured", "name", "binary"].includes(ast.kind)) {
      const bound = resolve(printFormula(ast, at), at);
      if (bound.kind === "ranges") { for (const area of bound.areas) { step(); emit(area); } }
      // A reference-valued name can still contain calculation/visibility
      // dependencies in INDEX/IF parameters. Successful binding does not make
      // its defining expression irrelevant to dependency or volatility flags.
      if (ast.kind === "name") {
        const scope = ast.sheet === undefined ? at.sheetId : sheets.find(sheet => sheet.name.toUpperCase() === ast.sheet!.toUpperCase())?.sheetId;
        const definition = definitions.find(name => name.name.toUpperCase() === ast.name.toUpperCase() && name.sheetId === scope)
          ?? (ast.sheet === undefined ? definitions.find(name => name.name.toUpperCase() === ast.name.toUpperCase() && name.sheetId === undefined) : undefined);
        if (definition && !names.includes(definition.nameId)) pending.push({ ast: parseFormula(definition.expression, definition.origin), at: { ...at, sheetId: definition.sheetId ?? at.sheetId }, names: [...names, definition.nameId] });
      }
    }
  }
}
