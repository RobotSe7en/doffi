import type { DeepReadonly } from "../runtime.js";
import type { FormulaExpression } from "./formula.js";
import { parseFormula } from "./formula.js";
import type { WorkbookReferenceArea } from "./reference.js";
import type { ScalarEvaluation, ScalarFunctionContext } from "./scalar.js";
import type { ScalarValue } from "./scalar-values.js";
import { scalarBlank, scalarError, scalarNumber, scalarToNumber } from "./scalar-values.js";
import { matchesWorkbookCriterion } from "./scalar-criteria.js";
type Area = DeepReadonly<WorkbookReferenceArea>;
function singleArea(value: ScalarEvaluation): Area | undefined { return !("type" in value) && value.kind === "references" && value.areas.length === 1 ? value.areas[0] : undefined; }
export function evaluateConditionalAggregate(name: string, args: readonly (FormulaExpression | null)[], context: ScalarFunctionContext): ScalarEvaluation | undefined {
  if (!["SUMIF", "SUMIFS", "COUNTIF", "COUNTIFS", "AVERAGEIF", "AVERAGEIFS"].includes(name)) return undefined;
  const plural = name.endsWith("IFS"), count = name.startsWith("COUNT"), average = name.startsWith("AVERAGE");
  if (!plural && (args.length < 2 || args.length > (count ? 2 : 3)) || plural && (args.length < (count ? 2 : 3) || args.length % 2 !== (count ? 0 : 1) || args.length > (count ? 254 : 255))) return scalarError("#VALUE!");
  const areas: Area[] = [], criteria: ScalarValue[] = [];
  let sum: Area | undefined;
  if (plural && !count) { const value = context.visit(args[0]!); if (!("type" in value) && value.kind === "unavailable") return value; sum = singleArea(value); if (!sum) return scalarError("#VALUE!"); }
  const first = plural && !count ? 1 : 0, end = plural ? args.length : 2;
  for (let index = first; index < end; index += 2) {
    const value = context.visit(args[index]!), area = singleArea(value); if (!("type" in value) && value.kind === "unavailable") return value; if (!area) return scalarError("#VALUE!");
    const criterion = context.scalar(context.visit(args[index + 1]!)); if (!("type" in criterion) || criterion.type === "error") return criterion;
    areas.push(area); criteria.push(criterion);
  }
  const rows = areas[0]!.to.row - areas[0]!.from.row + 1, columns = areas[0]!.to.column - areas[0]!.from.column + 1, total = rows * columns;
  if (!plural && !count) {
    if (args.length === 3) { const value = context.visit(args[2]!); if (!("type" in value) && value.kind === "unavailable") return value; sum = singleArea(value); if (!sum) return scalarError("#VALUE!"); }
    else sum = areas[0]!;
    sum = { ...sum, to: { row: sum.from.row + rows - 1, column: sum.from.column + columns - 1 } };
    if (sum.to.row >= 1_048_576 || sum.to.column >= 16_384) return scalarError("#REF!");
  }
  if ([...areas, ...(sum ? [sum] : [])].some(area => area.to.row - area.from.row + 1 !== rows || area.to.column - area.from.column + 1 !== columns)) return scalarError("#VALUE!");
  const offsets = new Set<number>();
  for (const area of [...areas, ...(sum ? [sum] : [])]) for (const location of context.stored(area)) { context.step(); offsets.add((location.row - area.from.row) * columns + location.column - area.from.column); }
  let matched = count && criteria.every(criterion => matchesWorkbookCriterion(scalarBlank, criterion, context.step, name !== "COUNTIF")) ? total - offsets.size : 0, accumulated = 0;
  for (const offset of offsets) {
    const row = Math.floor(offset / columns), column = offset % columns; let match = true;
    for (let index = 0; index < areas.length; index++) {
      const area = areas[index]!, value = context.read({ sheetId: area.sheetId, row: area.from.row + row, column: area.from.column + column });
      if (!("type" in value)) return value;
      if (name === "AVERAGEIF" && (value.type === "boolean" || value.type === "blank")) { match = false; break; }
      const criterionValue = name === "AVERAGEIFS" && value.type === "boolean" ? scalarNumber(Number(value.value)) : value;
      if (!matchesWorkbookCriterion(criterionValue, criteria[index]!, context.step, name !== "COUNTIF")) { match = false; break; }
    }
    if (!match) continue;
    if (count) { matched++; continue; }
    const value = context.read({ sheetId: sum!.sheetId, row: sum!.from.row + row, column: sum!.from.column + column });
    if (!("type" in value) || value.type === "error") return value;
    if (value.type === "number") { accumulated += value.value; matched++; }
  }
  return count ? scalarNumber(matched) : average ? matched ? scalarNumber(accumulated / matched) : scalarError("#DIV/0!") : scalarNumber(accumulated);
}
function containsSubtotal(expression: string, origin: { row: number; column: number }, context: ScalarFunctionContext): boolean {
  const pending = [parseFormula(expression, origin)];
  while (pending.length) {
    context.step(); const node = pending.pop()!;
    if (node.kind === "call") { if (node.name.toUpperCase().replace(/^_XLFN\./, "") === "SUBTOTAL") return true; for (const argument of node.arguments) if (argument) pending.push(argument); }
    else if (node.kind === "unary") pending.push(node.operand); else if (node.kind === "binary") pending.push(node.left, node.right);
  }
  return false;
}
export function evaluateSubtotal(args: readonly (FormulaExpression | null)[], context: ScalarFunctionContext): ScalarEvaluation {
  if (args.length < 2 || args.length > 255) return scalarError("#VALUE!");
  const source = context.scalar(context.visit(args[0]!)); if (!("type" in source)) return source;
  const number = scalarToNumber(source); if (typeof number !== "number") return number;
  const code = Math.trunc(number), fn = code > 100 ? code - 100 : code; if (fn < 1 || fn > 11 || code > 11 && code < 101 || code > 111) return scalarError("#VALUE!");
  let count = 0, present = 0, total = 0, product = 1, minimum = Infinity, maximum = -Infinity, mean = 0, m2 = 0;
  for (const argument of args.slice(1)) {
    const value = context.visit(argument); if ("type" in value) return value.type === "error" ? value : scalarError("#VALUE!"); if (value.kind === "unavailable") return value;
    for (const area of value.areas) for (const location of context.stored(area)) {
      context.step(); const filtered = context.filtered(location); if (typeof filtered !== "boolean") return filtered;
      if (filtered || code > 100 && context.hidden(location)) continue;
      const expression = context.formula(location); if (expression && containsSubtotal(expression, location, context)) continue;
      if (fn === 3) { if (context.hasContent(location)) present++; continue; }
      const item = context.read(location); if (!("type" in item)) return item;
      if (item.type !== "blank" || expression) present++;
      if (fn === 2 || fn === 3) { if (item.type === "number") count++; continue; }
      if (item.type === "error") return item; if (item.type !== "number") continue;
      count++; total += item.value; product *= item.value; minimum = Math.min(minimum, item.value); maximum = Math.max(maximum, item.value);
      const delta = item.value - mean; mean += delta / count; m2 += delta * (item.value - mean);
    }
  }
  if (fn === 1) return count ? scalarNumber(total / count) : scalarError("#DIV/0!");
  if (fn === 2 || fn === 3) return scalarNumber(fn === 2 ? count : present);
  if (fn === 4 || fn === 5) return scalarNumber(count ? fn === 4 ? maximum : minimum : 0);
  if (fn === 6) return scalarNumber(count ? product : 0);
  if (fn === 9) return scalarNumber(total);
  const denominator = count - (fn === 7 || fn === 10 ? 1 : 0);
  return denominator <= 0 ? scalarError("#DIV/0!") : scalarNumber(fn === 7 || fn === 8 ? Math.sqrt(Math.max(0, m2 / denominator)) : m2 / denominator);
}
