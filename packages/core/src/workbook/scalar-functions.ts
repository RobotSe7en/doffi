import type { FormulaExpression } from "./formula.js";
import type { ScalarEvaluation, ScalarFunctionContext, ScalarResolved } from "./scalar.js";
import { decimalComponents } from "./format-decimal.js";
import { scalarBinary, scalarBoolean, scalarError, scalarNumber, scalarToBoolean, scalarToNumber, scalarToText } from "./scalar-values.js";
import { evaluateTextFunction } from "./scalar-text.js";
import { evaluateDateFunction } from "./scalar-dates.js";
import { evaluateConditionalAggregate, evaluateSubtotal } from "./scalar-statistics.js";
import { evaluateLookupFunction } from "./scalar-lookup.js";

const aggregateNames = new Set(["SUM", "MIN", "MAX", "AVERAGE", "COUNT", "COUNTA", "COUNTBLANK"]);
function aggregate(name: string, args: readonly (FormulaExpression | null)[], context: ScalarFunctionContext): ScalarResolved {
  if (name === "COUNTBLANK" && args.length !== 1) return scalarError("#VALUE!");
  let count = 0, sum = 0, minimum = Infinity, maximum = -Infinity;
  for (const argument of args) {
    context.step(); const value = context.visit(argument);
    if (name === "COUNTBLANK" && ("type" in value || value.kind !== "references")) return "type" in value ? scalarError("#VALUE!") : value.kind === "unavailable" ? value : scalarError("#VALUE!");
    for (const entry of context.entries(value, name === "COUNTA")) {
      context.step(); const item = entry.value;
      if (name === "COUNTA" && entry.reference) { if (entry.hasContent) count += entry.count; continue; }
      if (!("type" in item)) return item;
      if (name === "COUNTA") { if (item.type !== "blank") count += entry.count; continue; }
      if (name === "COUNTBLANK") { if (item.type === "blank" || item.type === "text" && item.value === "") count += entry.count; continue; }
      if (name === "COUNT") {
        if (item.type === "number" || !entry.reference && (item.type === "boolean" || item.type === "text" && typeof scalarToNumber(item) === "number")) count += entry.count;
        continue;
      }
      if (entry.reference && ["blank", "boolean", "text"].includes(item.type)) continue;
      const number = scalarToNumber(item); if (typeof number !== "number") return number;
      count += entry.count; sum += number * entry.count; minimum = Math.min(minimum, number); maximum = Math.max(maximum, number);
    }
  }
  if (["COUNT", "COUNTA", "COUNTBLANK"].includes(name)) return scalarNumber(count);
  return name === "AVERAGE" ? count ? scalarNumber(sum / count) : scalarError("#DIV/0!")
    : name === "MIN" ? scalarNumber(count ? minimum : 0) : name === "MAX" ? scalarNumber(count ? maximum : 0) : scalarNumber(sum);
}

/** Formula rounding uses its own argument conversion and direction rules.
 * Only the general-purpose finite decimal decomposition is shared with display;
 * no scaling multiplication can overflow an otherwise unchanged finite input. */
function rounded(value: number, places: number, direction: "nearest" | "down" | "up"): ScalarResolved {
  if (value === 0 || places > 400) return scalarNumber(value);
  if (places < -400) return direction === "up" ? scalarError("#NUM!") : scalarNumber(0);
  const digits = Math.trunc(places), source = decimalComponents(value), exponent = source.exponent + digits;
  if (exponent >= 0) return scalarNumber(value);
  const coefficient = BigInt(source.digits), divisor = 10n ** BigInt(-exponent);
  const units = direction === "nearest" ? (coefficient + divisor / 2n) / divisor : direction === "up" ? (coefficient + divisor - 1n) / divisor : coefficient / divisor;
  return scalarNumber(Number(`${value < 0 ? "-" : ""}${units}e${-digits}`));
}

export function evaluateScalarFunction(node: Extract<FormulaExpression, { kind: "call" }>, context: ScalarFunctionContext): ScalarEvaluation {
  const name = node.name.toUpperCase().replace(/^_XLFN\./, ""), args = node.arguments;
  const scalar = (index: number): ScalarResolved => context.scalar(context.visit(args[index] ?? null));
  if (name === "IFS") {
    if (args.length < 2 || args.length % 2 || args.length > 254) return scalarError("#VALUE!");
    for (let index = 0; index < args.length; index += 2) {
      const value = scalar(index); if (!("type" in value)) return value;
      const test = scalarToBoolean(value); if (typeof test !== "boolean") return test;
      if (test) return context.visit(args[index + 1]!);
    }
    return scalarError("#N/A");
  }
  if (name === "IF") {
    if (args.length < 1 || args.length > 3) return scalarError("#VALUE!");
    const value = scalar(0); if (!("type" in value)) return value;
    const test = scalarToBoolean(value); if (typeof test !== "boolean") return test;
    return test ? args.length >= 2 ? context.visit(args[1]!) : scalarNumber(0) : args.length >= 3 ? context.visit(args[2]!) : scalarBoolean(false);
  }
  if (name === "IFERROR" || name === "IFNA") {
    if (args.length !== 2) return scalarError("#VALUE!");
    const value = scalar(0);
    return "type" in value && value.type === "error" && (name === "IFERROR" || value.value === "#N/A") ? context.visit(args[1]!) : value;
  }
  if (name === "TRUE" || name === "FALSE") return args.length ? scalarError("#VALUE!") : scalarBoolean(name === "TRUE");
  if (name === "AND" || name === "OR") {
    let found = false, result = name === "AND";
    for (const argument of args) for (const entry of context.entries(context.visit(argument))) {
      context.step(); const value = entry.value; if (!("type" in value)) return value;
      if (entry.reference && (value.type === "text" || value.type === "blank")) continue;
      const logical = scalarToBoolean(value); if (typeof logical !== "boolean") return logical;
      found = true; result = name === "AND" ? result && logical : result || logical;
    }
    return found ? scalarBoolean(result) : scalarError("#VALUE!");
  }
  if (name === "NOT") {
    if (args.length !== 1) return scalarError("#VALUE!");
    const value = scalar(0); if (!("type" in value)) return value;
    const test = scalarToBoolean(value); return typeof test === "boolean" ? scalarBoolean(!test) : test;
  }
  if (["ISBLANK", "ISNUMBER", "ISTEXT", "ISERROR", "ISNA"].includes(name)) {
    if (args.length !== 1) return scalarError("#VALUE!");
    if (name === "ISBLANK") {
      const source = context.visit(args[0]!);
      if (!("type" in source) && source.kind === "references" && source.areas.length === 1) {
        const area = source.areas[0]!, row = area.from.row === area.to.row ? area.from.row : context.origin.row, column = area.from.column === area.to.column ? area.from.column : context.origin.column;
        if (!(area.from.row !== area.to.row && area.from.column !== area.to.column) && row >= area.from.row && row <= area.to.row && column >= area.from.column && column <= area.to.column)
          return scalarBoolean(!context.hasContent({ sheetId: area.sheetId, row, column }));
      }
      const value = context.scalar(source); return "type" in value ? scalarBoolean(value.type === "blank") : value;
    }
    const value = scalar(0); if (!("type" in value)) return value;
    return scalarBoolean(name === "ISNUMBER" ? value.type === "number"
      : name === "ISTEXT" ? value.type === "text" : name === "ISERROR" ? value.type === "error" : value.type === "error" && value.value === "#N/A");
  }
  if (aggregateNames.has(name)) return aggregate(name, args, context);
  if (["ROW", "COLUMN", "ROWS", "COLUMNS"].includes(name)) {
    if (!args.length && (name === "ROW" || name === "COLUMN")) return scalarNumber(context.origin[name === "ROW" ? "row" : "column"] + 1);
    if (args.length !== 1) return scalarError("#VALUE!");
    const value = context.visit(args[0]!);
    if ("type" in value) return value.type === "error" ? value : name === "ROWS" || name === "COLUMNS" ? scalarNumber(1) : scalarError("#VALUE!");
    if (value.kind === "unavailable") return value;
    if (value.areas.length !== 1) return scalarError("#VALUE!");
    const area = value.areas[0]!, axis = name === "ROW" || name === "ROWS" ? "row" : "column";
    return scalarNumber(name === "ROW" || name === "COLUMN" ? area.from[axis] + 1 : area.to[axis] - area.from[axis] + 1);
  }
  if (name === "LEN") {
    if (args.length !== 1) return scalarError("#VALUE!");
    const value = scalar(0); if (!("type" in value)) return value;
    const text = scalarToText(value); return typeof text === "string" ? scalarNumber(text.length) : text;
  }
  const textResult = evaluateTextFunction(name, args, context); if (textResult !== undefined) return textResult;
  const conditionalResult = evaluateConditionalAggregate(name, args, context); if (conditionalResult !== undefined) return conditionalResult;
  if (name === "SUBTOTAL") return evaluateSubtotal(args, context);
  const lookupResult = evaluateLookupFunction(name, args, context); if (lookupResult !== undefined) return lookupResult;
  if (["DATE", "YEAR", "MONTH", "DAY", "TIME", "HOUR", "MINUTE", "SECOND", "DAYS", "EDATE", "EOMONTH", "WEEKDAY", "TODAY", "NOW"].includes(name)) {
    const values: import("./scalar-values.js").ScalarValue[] = [];
    for (let index = 0; index < args.length; index++) { const value = scalar(index); if (!("type" in value)) return value; values.push(value); }
    return evaluateDateFunction(name, values, context)!;
  }
  const counts: Readonly<Record<string, readonly [number, number]>> = {
    ABS: [1, 1], INT: [1, 1], TRUNC: [1, 2], MOD: [2, 2], POWER: [2, 2], SQRT: [1, 1], ROUND: [2, 2], ROUNDUP: [2, 2], ROUNDDOWN: [2, 2],
  };
  const count = counts[name];
  if (!count) return { kind: "unavailable", reason: "unsupported-function", functionName: node.name };
  if (args.length < count[0] || args.length > count[1]) return scalarError("#VALUE!");
  const numbers: number[] = [];
  for (let index = 0; index < args.length; index++) {
    const value = scalar(index); if (!("type" in value)) return value;
    const number = scalarToNumber(value); if (typeof number !== "number") return number; numbers.push(number);
  }
  const a = numbers[0]!, b = numbers[1] ?? 0;
  if (name === "ABS") return scalarNumber(Math.abs(a));
  if (name === "INT") return scalarNumber(Math.floor(a));
  if (name === "TRUNC" || name === "ROUNDDOWN") return rounded(a, b, "down");
  if (name === "ROUND" || name === "ROUNDUP") return rounded(a, b, name === "ROUND" ? "nearest" : "up");
  if (name === "MOD") {
    if (b === 0) return scalarError("#DIV/0!");
    const remainder = a % b;
    return scalarNumber(remainder !== 0 && Math.sign(remainder) !== Math.sign(b) ? remainder + b : remainder);
  }
  if (name === "POWER") return scalarBinary("^", scalarNumber(a), scalarNumber(b));
  return a < 0 ? scalarError("#NUM!") : scalarNumber(Math.sqrt(a));
}
