import type { WordMetadata } from "../contracts/word.js";
interface Token {
  from: number;
  to: number;
  value: string;
  quoted: boolean;
}
/** Read literal operands without interpreting field code or matching switches
 * inside quoted tooltips, style lists or other field arguments. */
function tokens(instruction: string): Token[] {
  const result: Token[] = [];
  for (let at = 0; at < instruction.length; ) {
    if (/\s/.test(instruction[at]!)) {
      at++;
      continue;
    }
    const from = at,
      quoted = instruction[at] === '"';
    if (quoted) {
      at++;
      let value = "";
      for (; at < instruction.length && instruction[at] !== '"'; at++) {
        if (instruction[at] === "\\" && instruction[at + 1] === '"') {
          value += '"';
          at++;
        } else value += instruction[at];
      }
      if (at === instruction.length) return [];
      at++;
      result.push({ from, to: at, value, quoted });
    } else {
      while (at < instruction.length && !/\s/.test(instruction[at]!)) at++;
      result.push({ from, to: at, value: instruction.slice(from, at), quoted });
    }
  }
  return result;
}
/** Rewrite literal bookmark operands while retaining switches, quoting and
 * exact offsets of nested instruction fields after the changed operand. */
export function remapWordFieldBookmarks(
  field: WordMetadata["fields"][number],
  resolve: (name: string) => string,
): void {
  const parts = tokens(field.instruction),
    kind = parts[0]?.value.toUpperCase();
  let target: Token | undefined;
  if (kind && ["REF", "PAGEREF", "NOTEREF"].includes(kind)) target = parts[1];
  else if (kind === "TOC") {
    const index = parts.findIndex((part) => !part.quoted && part.value.toLowerCase() === "\\b");
    if (index >= 0) target = parts[index + 1];
  } else if (kind === "HYPERLINK") {
    for (let index = 1; index < parts.length; index++) {
      const part = parts[index]!;
      if (!part.quoted && /^\\[a-z*]$/i.test(part.value)) {
        const flag = part.value.toLowerCase();
        if (["\\l", "\\o", "\\t", "\\s", "\\*"].includes(flag)) {
          if (flag === "\\l") target = parts[index + 1];
          index++;
        } else if (!["\\m", "\\n", "\\h"].includes(flag)) return;
      } else if (part.value) return;
    }
  }
  if (
    !target ||
    !target.value ||
    target.value.startsWith("\\") ||
    field.instructionFields?.some(
      (child) => child.offset >= target!.from && child.offset < target!.to,
    )
  )
    return;
  const next = resolve(target.value);
  if (next === target.value) return;
  const replacement = target.quoted ? '"' + next + '"' : next,
    delta = replacement.length - target.to + target.from;
  field.instruction =
    field.instruction.slice(0, target.from) + replacement + field.instruction.slice(target.to);
  for (const nested of field.instructionFields ?? [])
    if (nested.offset >= target.to) nested.offset += delta;
}
