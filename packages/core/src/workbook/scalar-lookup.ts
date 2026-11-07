import type { DeepReadonly } from "../runtime.js";
import type { FormulaExpression } from "./formula.js";
import type { WorkbookReferenceArea, WorkbookReferenceLocation } from "./reference.js";
import type { ScalarEvaluation, ScalarFunctionContext, ScalarResolved } from "./scalar.js";
import type { ScalarValue } from "./scalar-values.js";
import { scalarCompare, scalarError, scalarNumber, scalarToBoolean, scalarToNumber } from "./scalar-values.js";
import { wildcardMatch } from "./scalar-text.js";
type Area = DeepReadonly<WorkbookReferenceArea>;
function areas(value: ScalarEvaluation): readonly Area[] | undefined { return !("type" in value) && value.kind === "references" ? value.areas : undefined; }
function match(needle: ScalarValue, area: Area, vertical: boolean, mode: number, context: ScalarFunctionContext): number | ScalarResolved {
  const count = vertical ? area.to.row - area.from.row + 1 : area.to.column - area.from.column + 1;
  const axis = vertical ? "row" : "column", stored = new Map<number, WorkbookReferenceLocation>();
  for (const location of context.stored(area)) { context.step(); stored.set(location[axis] - area.from[axis], location); }
  const positions = [...stored.keys()].sort((a, b) => a - b);
  // One representative at each end of a blank gap handles exact/approximate
  // zero lookups without expanding whole rows or columns.
  const candidates = new Set(positions); let previous = -1;
  for (const position of [...positions, count]) { if (previous + 1 < position) { candidates.add(previous + 1); candidates.add(position - 1); } previous = position; }
  let found = -1;
  for (const position of [...candidates].sort((a, b) => a - b)) {
    context.step(); const location = stored.get(position) ?? { sheetId: area.sheetId, row: area.from.row + (vertical ? position : 0), column: area.from.column + (vertical ? 0 : position) };
    const result = context.read(location); if (!("type" in result)) return result; if (result.type === "error") continue;
    const value = result.type === "blank" ? scalarNumber(0) : result, target = needle.type === "blank" ? scalarNumber(0) : needle;
    if (value.type !== target.type) continue;
    if (mode === 0 && target.type === "text" && value.type === "text") {
      if (wildcardMatch(target.value, value.value, context.step) >= 0) return position; continue;
    }
    const order = scalarCompare(value, target); if (typeof order !== "number") continue;
    if (mode === 0 && order === 0) return position;
    if (mode === 1 && order <= 0 || mode === -1 && order >= 0) found = position;
  }
  return found < 0 ? scalarError("#N/A") : found;
}
export function evaluateLookupFunction(name: string, args: readonly (FormulaExpression | null)[], context: ScalarFunctionContext): ScalarEvaluation | undefined {
  if (!["INDEX", "MATCH", "VLOOKUP", "HLOOKUP"].includes(name)) return undefined;
  const scalar = (index: number): ScalarResolved => context.scalar(context.visit(args[index] ?? null));
  if (name === "INDEX") {
    if (args.length < 2 || args.length > 4) return scalarError("#VALUE!");
    const value = context.visit(args[0]!); if (!("type" in value) && value.kind === "unavailable") return value;
    const source = areas(value); if (!source) return scalarError("#VALUE!");
    const values: number[] = [];
    for (let index = 1; index < args.length; index++) { const item = scalar(index); if (!("type" in item)) return item; const number = scalarToNumber(item); if (typeof number !== "number") return number; values.push(Math.trunc(number)); }
    const area = source[(values[2] ?? 1) - 1]; if (!area) return scalarError("#REF!");
    const rows = area.to.row - area.from.row + 1, columns = area.to.column - area.from.column + 1;
    let row = values[0]!, column = values[1] ?? (columns === 1 ? 1 : 0);
    if (args.length === 2 && rows === 1) { column = row; row = 1; }
    if (row < 0 || column < 0) return scalarError("#VALUE!"); if (row > rows || column > columns) return scalarError("#REF!");
    return { kind: "references", areas: [{ sheetId: area.sheetId,
      from: { row: row ? area.from.row + row - 1 : area.from.row, column: column ? area.from.column + column - 1 : area.from.column },
      to: { row: row ? area.from.row + row - 1 : area.to.row, column: column ? area.from.column + column - 1 : area.to.column },
    }] };
  }
  if (args.length < (name === "MATCH" ? 2 : 3) || args.length > (name === "MATCH" ? 3 : 4)) return scalarError("#VALUE!");
  const needle = scalar(0); if (!("type" in needle) || needle.type === "error") return needle;
  const value = context.visit(args[1]!); if (!("type" in value) && value.kind === "unavailable") return value;
  const source = areas(value); if (source?.length !== 1) return scalarError("#VALUE!");
  const area = source[0]!, vertical = name !== "HLOOKUP" && (name !== "MATCH" || area.from.column === area.to.column);
  if (name === "MATCH" && area.from.row !== area.to.row && area.from.column !== area.to.column) return scalarError("#N/A");
  let mode = 1, index = 1;
  if (name === "MATCH" && args.length === 3) { const item = scalar(2); if (!("type" in item)) return item; const number = scalarToNumber(item); if (typeof number !== "number") return number; mode = number > 0 ? 1 : number < 0 ? -1 : 0; }
  if (name !== "MATCH") {
    const item = scalar(2); if (!("type" in item)) return item; const number = scalarToNumber(item); if (typeof number !== "number") return number; index = Math.trunc(number);
    if (index < 1) return scalarError("#VALUE!"); if (index > (vertical ? area.to.column - area.from.column + 1 : area.to.row - area.from.row + 1)) return scalarError("#REF!");
    if (args.length === 4) { const value = scalar(3); if (!("type" in value)) return value; const approximate = scalarToBoolean(value); if (typeof approximate !== "boolean") return approximate; mode = approximate ? 1 : 0; }
  }
  const search: Area = { ...area, to: { row: vertical ? area.to.row : area.from.row, column: vertical ? area.from.column : area.to.column } };
  const found = match(needle, search, vertical, mode, context); if (typeof found !== "number") return found;
  return name === "MATCH" ? scalarNumber(found + 1) : context.read({ sheetId: area.sheetId, row: area.from.row + (vertical ? found : index - 1), column: area.from.column + (vertical ? index - 1 : found) });
}
