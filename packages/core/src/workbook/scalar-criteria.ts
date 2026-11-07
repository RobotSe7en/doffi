import type { ScalarValue } from "./scalar-values.js";
import { scalarBlank, scalarCompare, scalarToNumber } from "./scalar-values.js";
import { wildcardMatch } from "./scalar-text.js";

export function matchesWorkbookCriterion(value: ScalarValue, criterion: ScalarValue, step: () => void, blankAsZero = true): boolean {
  step();
  let operator = "=", target = criterion.type === "blank" ? { type: "number" as const, value: 0 } : criterion;
  if (target.type === "text") {
    const match = /^(<=|>=|<>|=|<|>)([\s\S]*)$/.exec(target.value);
    if (match) { operator = match[1]!; target = { type: "text", value: match[2]! }; }
    if (target.value === "") {
      const empty = value.type === "blank" || value.type === "text" && value.value === "";
      return operator === "=" ? empty : operator === "<>" ? !empty : false;
    }
    if (/^#(?:NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|GETTING_DATA)$/i.test(target.value)) target = { type: "error", value: target.value.toUpperCase() as Extract<ScalarValue, { type: "error" }>["value"] };
    const numeric = scalarToNumber(target);
    if (typeof numeric === "number") target = { type: "number", value: numeric };
    else if (target.type === "text" && (operator === "=" || operator === "<>") && /[?*~]/.test(target.value)) {
      const match = value.type === "text" && wildcardMatch(target.value, value.value, step) >= 0; return operator === "=" ? match : !match;
    }
  }
  if (value.type === "error" || target.type === "error") {
    const equal = (value.type === "error" || value.type === "text") && (target.type === "error" || target.type === "text") && value.value.toUpperCase() === target.value.toUpperCase();
    return operator === "=" ? equal : operator === "<>" ? !equal : false;
  }
  if (value.type === "blank" && !blankAsZero) return false;
  if (target.type === "number" && value.type === "text") {
    if (operator !== "=" && operator !== "<>") return false;
    const numeric = scalarToNumber(value), equal = typeof numeric === "number" && numeric === target.value;
    return operator === "=" ? equal : !equal;
  }
  const order = scalarCompare(value.type === "blank" ? scalarBlank : value, target); if (typeof order !== "number") return false;
  return operator === "=" ? order === 0 : operator === "<>" ? order !== 0 : operator === "<" ? order < 0 : operator === ">" ? order > 0 : operator === "<=" ? order <= 0 : order >= 0;
}
