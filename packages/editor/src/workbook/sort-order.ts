import type { DeepReadonly, WorkbookValue } from "@doffi/core";

// Microsoft Support, "Sort data in a range or table in Excel": text digits
// precede punctuation, apostrophes/hyphens are ignored, errors compare equal,
// and blank cells remain last in both sort directions. Locale governs letters.
const punctuation = ' !"#$%&()*,./:;?@[\\]^_`{|}~+<=>';
type Value = DeepReadonly<WorkbookValue>;
function parts(text: string): { group: number; value: string }[] {
  const result: { group: number; value: string }[] = [];
  for (const character of text.normalize("NFC").replace(/['-]/g, "")) {
    result.push({ group: /[0-9]/.test(character) ? 0 : punctuation.includes(character) ? 1 : 2, value: character });
  }
  return result;
}
export function workbookSortComparison(locale: string, caseSensitive: boolean): (a: Value, b: Value, descending: boolean) => number {
  const collator = new Intl.Collator(locale, { usage: "sort", sensitivity: caseSensitive ? "variant" : "accent", caseFirst: "lower", numeric: false });
  const parsed = new Map<string, ReturnType<typeof parts>>(), rank = { number: 0, text: 1, boolean: 2, error: 3, blank: 4 };
  function text(a: string, b: string): number {
    let left = parsed.get(a); if (!left) { left = parts(a); parsed.set(a, left); }
    let right = parsed.get(b); if (!right) { right = parts(b); parsed.set(b, right); }
    for (let i = 0; i < Math.min(left.length, right.length); i++) {
      const x = left[i]!, y = right[i]!;
      const order = x.group - y.group || (x.group === 1 ? punctuation.indexOf(x.value) - punctuation.indexOf(y.value) : x.group === 0 ? x.value.charCodeAt(0) - y.value.charCodeAt(0) : collator.compare(x.value, y.value));
      if (order) return order;
    }
    return left.length - right.length || Number(a.includes("-")) - Number(b.includes("-"));
  }
  return (a, b, descending) => {
    if (a.type === "blank" || b.type === "blank") return Number(a.type === "blank") - Number(b.type === "blank");
    const order = a.type !== b.type ? rank[a.type] - rank[b.type] : a.type === "number" && b.type === "number" ? a.value < b.value ? -1 : a.value > b.value ? 1 : 0
      : a.type === "text" && b.type === "text" ? text(a.value, b.value) : a.type === "boolean" && b.type === "boolean" ? Number(a.value) - Number(b.value) : 0;
    return descending ? -order : order;
  };
}
