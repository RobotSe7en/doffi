import type { DeepReadonly, WorkbookModel } from "../runtime.js";
import type { FormulaExpression } from "./formula.js";
import { parseFormula } from "./formula.js";
import { workbookReferences, type WorkbookNameContext, type WorkbookReferenceArea, type WorkbookReferenceLocation, type WorkbookReferenceResult } from "./reference.js";
import type { WorkbookDefinedName } from "../contracts/workbook.js";

/** Shared, bounded source-reference discovery for calculation and validation.
 * Reference rectangles stay compact; function evaluation remains separate. */
export function workbookExpressionDependencies(model: WorkbookModel, expression: FormulaExpression, origin: WorkbookReferenceLocation,
  step: () => void, add: (area: DeepReadonly<WorkbookReferenceArea>) => void, call?: (name: string) => void,
  resolve: (expression: FormulaExpression, at: WorkbookReferenceLocation, context: WorkbookNameContext) => WorkbookReferenceResult = (expression, at, context) => workbookReferences(model.snapshot().metadata.workbook).resolveExpression(expression, at, context),
  options: { readonly onName?: (definition: DeepReadonly<WorkbookDefinedName>, location: WorkbookReferenceLocation, context: WorkbookNameContext) => void; readonly includeFilterRanges?: boolean; readonly context?: WorkbookNameContext } = {}): void {
  const workbook = model.snapshot().metadata.workbook, sheets = workbook.sheets, references = workbookReferences(workbook);
  const seen = new Set<string>();
  const emit = (area: DeepReadonly<WorkbookReferenceArea>): void => {
    const identity = JSON.stringify([area.sheetId, area.from.row, area.from.column, area.to.row, area.to.column]);
    if (!seen.has(identity)) { seen.add(identity); add(area); }
  };
  const pending: { ast: FormulaExpression; at: WorkbookReferenceLocation; names: readonly string[]; context: WorkbookNameContext }[] = [{ ast: expression, at: origin, names: [], context: options.context ?? {} }];
  while (pending.length) {
    step(); const { ast, at, names, context } = pending.pop()!;
    if (ast.kind === "call") {
      const name = ast.name.toUpperCase().replace(/^_XLFN\./, ""); call?.(name);
      if ((name === "SUMIF" || name === "AVERAGEIF") && ast.arguments.length === 3 && ast.arguments[0] && ast.arguments[2]) {
        const criteria = resolve(ast.arguments[0], at, context), sum = resolve(ast.arguments[2], at, context);
        if (criteria.kind === "ranges" && criteria.areas.length === 1 && sum.kind === "ranges" && sum.areas.length === 1) {
          const a = criteria.areas[0]!, b = sum.areas[0]!, to = { row: b.from.row + a.to.row - a.from.row, column: b.from.column + a.to.column - a.from.column };
          if (to.row < 1_048_576 && to.column < 16_384) emit({ sheetId: b.sheetId, from: b.from, to });
        }
      }
      if (name === "SUBTOTAL" && options.includeFilterRanges !== false) for (const sheet of sheets) for (const filter of [...sheet.autoFilter ? [sheet.autoFilter] : [], ...sheet.tables?.flatMap(table => table.autoFilter ? [table.autoFilter] : []) ?? []]) {
        for (const column of filter.columns) if (column.criterion && filter.range.from.row < filter.range.to.row) {
          step(); emit({ sheetId: sheet.sheetId, from: { row: filter.range.from.row + 1, column: column.column }, to: { row: filter.range.to.row, column: column.column } });
        }
      }
      for (const argument of ast.arguments) if (argument) pending.push({ ast: argument, at, names, context });
    } else if (ast.kind === "unary") pending.push({ ast: ast.operand, at, names, context });
    else if (ast.kind === "binary" && ![",", " ", ":"].includes(ast.operator)) pending.push({ ast: ast.left, at, names, context }, { ast: ast.right, at, names, context });
    else if (["cell", "range", "axisRange", "structured", "name", "binary"].includes(ast.kind)) {
      const bound = resolve(ast, at, context);
      if (bound.kind === "ranges") { for (const area of bound.areas) { step(); emit(area); } }
      // A reference-valued name can still contain calculation/visibility
      // dependencies in INDEX/IF parameters. Successful binding does not make
      // its defining expression irrelevant to dependency or volatility flags.
      if (ast.kind === "binary") pending.push({ ast: ast.left, at, names, context }, { ast: ast.right, at, names, context });
      if (ast.kind === "name") {
        const boundName = references.bindName(ast, at, context);
        if (boundName?.kind === "name") {
          const definition = boundName.definition;
          options.onName?.(definition, boundName.location, boundName.context);
          const identity = JSON.stringify([definition.nameId, boundName.location.sheetId, boundName.context.unboundSheet === true]);
          if (!names.includes(identity)) pending.push({ ast: parseFormula(definition.expression, definition.origin), at: boundName.location, names: [...names, identity], context: boundName.context });
        }
      }
    }
  }
}
