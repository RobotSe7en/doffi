import type { FormulaExpression } from "./formula.js";
import type { ScalarEvaluation, ScalarFunctionContext, ScalarResolved } from "./scalar.js";
import { scalarBoolean, scalarError, scalarNumber, scalarText, scalarToBoolean, scalarToNumber, scalarToText } from "./scalar-values.js";
import { createWorkbookNumberFormatter } from "./number-formatter.js";
import { parseWorkbookValueText } from "./scalar-parse-value.js";

/** Bounded wildcard matching avoids constructing executable patterns or
 * backtracking regular expressions from workbook criteria. */
export function wildcardMatch(pattern: string, text: string, step: () => void, whole = true): number {
  const tokens: { kind: "literal" | "one" | "many"; text?: string }[] = [];
  for (let index = 0; index < pattern.length; index++) {
    step(); const character = pattern[index]!;
    if (character === "~" && index + 1 < pattern.length && ["~", "*", "?"].includes(pattern[index + 1]!)) tokens.push({ kind: "literal", text: pattern[++index]!.toUpperCase() });
    else tokens.push(character === "*" ? { kind: "many" } : character === "?" ? { kind: "one" } : { kind: "literal", text: character.toUpperCase() });
  }
  for (let start = 0; start <= (whole ? 0 : text.length); start++) {
    let token = 0, cursor = start, star = -1, retry = start;
    while (cursor < text.length) {
      step(); const current = tokens[token];
      if (current?.kind === "one" || current?.kind === "literal" && current.text === text[cursor]!.toUpperCase()) { token++; cursor++; }
      else if (current?.kind === "many") { star = token++; retry = cursor; }
      else if (!whole && token === tokens.length) return start;
      else if (star >= 0) { token = star + 1; cursor = ++retry; }
      else break;
    }
    while (tokens[token]?.kind === "many") { token++; step(); }
    if (token === tokens.length && (!whole || cursor === text.length)) return start;
  }
  return -1;
}
export function evaluateTextFunction(name: string, args: readonly (FormulaExpression | null)[], context: ScalarFunctionContext): ScalarEvaluation | undefined {
  const counts: Readonly<Record<string, readonly [number, number]>> = { LEFT: [1, 2], RIGHT: [1, 2], MID: [3, 3], TRIM: [1, 1], UPPER: [1, 1], LOWER: [1, 1], CONCAT: [1, 253], CONCATENATE: [1, 255], TEXTJOIN: [3, 254], SUBSTITUTE: [3, 4], REPLACE: [4, 4], FIND: [2, 3], SEARCH: [2, 3], EXACT: [2, 2], TEXT: [2, 2], VALUE: [1, 1] };
  const count = counts[name]; if (!count) return undefined;
  if (args.length < count[0] || args.length > count[1]) return scalarError("#VALUE!");
  const values: ScalarResolved[] = [];
  if (!["CONCAT", "CONCATENATE", "TEXTJOIN"].includes(name)) for (const argument of args) {
    const value = context.scalar(context.visit(argument)); if (!("type" in value) || value.type === "error") return value; values.push(value);
  }
  const text = (index: number): string => scalarToText(values[index] as import("./scalar-values.js").ScalarValue) as string;
  const number = (index: number, fallback = 1): number | ScalarResolved => values[index] === undefined ? fallback : scalarToNumber(values[index] as import("./scalar-values.js").ScalarValue);
  if (["CONCAT", "CONCATENATE", "TEXTJOIN"].includes(name)) {
    let delimiter = "", ignore = false, start = 0, output = "", found = false;
    if (name === "TEXTJOIN") {
      const a = context.scalar(context.visit(args[0]!)), b = context.scalar(context.visit(args[1]!));
      if (!("type" in a)) return a; if (!("type" in b)) return b;
      const separator = scalarToText(a), skip = scalarToBoolean(b); if (typeof separator !== "string") return separator; if (typeof skip !== "boolean") return skip;
      delimiter = separator; ignore = skip; start = 2;
    }
    for (const argument of args.slice(start)) {
      const source = context.visit(argument), entries = name === "CONCATENATE" ? [{ value: context.scalar(source), count: 1 }] : context.orderedEntries(source);
      for (const entry of entries) {
        context.step(); if (!("type" in entry.value)) return entry.value;
        const value = scalarToText(entry.value); if (typeof value !== "string") return value; if (ignore && value === "") continue;
        if (!value && !delimiter) continue;
        const length = output.length + entry.count * value.length + (found ? entry.count : entry.count - 1) * delimiter.length;
        if (length > 32767) return scalarError("#VALUE!");
        for (let index = 0; index < entry.count; index++) { context.step(); output += (found ? delimiter : "") + value; found = true; }
      }
    }
    return scalarText(output);
  }
  if (name === "TRIM") return scalarText(text(0).replace(/^ +| +$/g, "").replace(/ +/g, " "));
  if (name === "UPPER" || name === "LOWER") return scalarText(name === "UPPER" ? text(0).toUpperCase() : text(0).toLowerCase());
  if (name === "EXACT") return scalarBoolean(text(0) === text(1));
  if (name === "LEFT" || name === "RIGHT") {
    const amount = number(1); if (typeof amount !== "number") return amount; if (amount < 0) return scalarError("#VALUE!");
    const count = Math.trunc(amount), value = text(0); return scalarText(name === "LEFT" ? value.slice(0, count) : count === 0 ? "" : value.slice(Math.max(0, value.length - count)));
  }
  if (name === "MID" || name === "REPLACE") {
    const from = number(1), length = number(2); if (typeof from !== "number") return from; if (typeof length !== "number") return length;
    if (from < 1 || length < 0) return scalarError("#VALUE!");
    const value = text(0), start = Math.trunc(from) - 1, count = Math.trunc(length);
    return scalarText(name === "MID" ? value.slice(start, start + count) : value.slice(0, start) + text(3) + value.slice(start + count));
  }
  if (name === "SUBSTITUTE") {
    const source = text(0), old = text(1), replacement = text(2), instance = number(3, 0);
    if (typeof instance !== "number") return instance; if (args.length === 4 && instance < 1) return scalarError("#VALUE!");
    if (!old) return scalarText(source);
    let result = "", start = 0, count = 0;
    while (start < source.length) {
      context.step(); const index = source.indexOf(old, start); if (index < 0) break; count++;
      const value = !instance || count === Math.trunc(instance) ? replacement : old;
      if (result.length + index - start + value.length > 32767) return scalarError("#VALUE!");
      result += source.slice(start, index) + value; start = index + old.length;
    }
    return scalarText(result + source.slice(start));
  }
  if (name === "FIND" || name === "SEARCH") {
    const from = number(2); if (typeof from !== "number") return from; const start = Math.trunc(from) - 1, source = text(1);
    if (start < 0 || start >= source.length && !(start === source.length && text(0) === "")) return scalarError("#VALUE!");
    const found = name === "FIND" ? source.indexOf(text(0), start) : wildcardMatch(text(0), source.slice(start), context.step, false);
    return found < 0 ? scalarError("#VALUE!") : scalarNumber(found + 1 + (name === "SEARCH" ? start : 0));
  }
  if (name === "TEXT") {
    const value = values[0]!; if (!("type" in value)) return value;
    const parsed = value.type === "text" ? parseWorkbookValueText(value.value, context.locale ?? "zh-CN", context.dateSystem) : scalarToNumber(value);
    const numeric = typeof parsed === "number" ? parsed : value.type === "text" ? value.value : parsed; if (typeof numeric !== "number" && typeof numeric !== "string") return numeric;
    const formatter = createWorkbookNumberFormatter({ dateSystem: context.dateSystem, ...(context.locale ? { hostLocale: context.locale } : {}) });
    try { const result = formatter.format(numeric, { code: text(1) }); return result.kind === "parts" ? scalarText(result.parts.map(part => part.kind === "text" ? part.text : part.kind === "space" ? " " : "").join("")) : scalarError("#VALUE!"); }
    finally { formatter.dispose(); }
  }
  const value = values[0]!; if (!("type" in value)) return value;
  if (value.type !== "text") return value.type === "number" ? value : scalarError("#VALUE!");
  const result = parseWorkbookValueText(value.value, context.locale ?? "zh-CN", context.dateSystem); return typeof result === "number" ? scalarNumber(result) : result;
}
