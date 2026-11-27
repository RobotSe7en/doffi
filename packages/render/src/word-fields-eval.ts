import {
  formatWordBibliography,
  formatWordCitation,
  type DeepReadonly,
  type WordBibliographyStyle,
  type WordMetadata,
} from "@doffi/core";
import { formatWordDatePicture } from "./word-table-formulas.js";
import { pageLabel } from "./pagination.js";

const ONES = [
  "zero",
  "one",
  "two",
  "three",
  "four",
  "five",
  "six",
  "seven",
  "eight",
  "nine",
  "ten",
  "eleven",
  "twelve",
  "thirteen",
  "fourteen",
  "fifteen",
  "sixteen",
  "seventeen",
  "eighteen",
  "nineteen",
];
const TENS = ["", "", "twenty", "thirty", "forty", "fifty", "sixty", "seventy", "eighty", "ninety"];
const ORD_ONES = [
  "zeroth",
  "first",
  "second",
  "third",
  "fourth",
  "fifth",
  "sixth",
  "seventh",
  "eighth",
  "ninth",
  "tenth",
  "eleventh",
  "twelfth",
  "thirteenth",
  "fourteenth",
  "fifteenth",
  "sixteenth",
  "seventeenth",
  "eighteenth",
  "nineteenth",
];
const ORD_TENS = [
  "",
  "",
  "twentieth",
  "thirtieth",
  "fortieth",
  "fiftieth",
  "sixtieth",
  "seventieth",
  "eightieth",
  "ninetieth",
];

function englishCardinal(n: number): string {
  if (!Number.isFinite(n) || !Number.isInteger(n)) return String(n);
  if (n < 0) return "minus " + englishCardinal(-n);
  if (n < 20) return ONES[n]!;
  if (n < 100) {
    const t = Math.floor(n / 10);
    const r = n % 10;
    return r ? `${TENS[t]}-${ONES[r]}` : TENS[t]!;
  }
  if (n < 1000) {
    const h = Math.floor(n / 100);
    const r = n % 100;
    return r ? `${ONES[h]} hundred ${englishCardinal(r)}` : `${ONES[h]} hundred`;
  }
  if (n < 1_000_000) {
    const k = Math.floor(n / 1000);
    const r = n % 1000;
    return r ? `${englishCardinal(k)} thousand ${englishCardinal(r)}` : `${englishCardinal(k)} thousand`;
  }
  return String(n);
}

function englishOrdinal(n: number): string {
  if (!Number.isFinite(n) || !Number.isInteger(n) || n < 0) return String(n);
  if (n < 20) return ORD_ONES[n]!;
  if (n < 100) {
    const t = Math.floor(n / 10);
    const r = n % 10;
    return r ? `${TENS[t]}-${ORD_ONES[r]}` : ORD_TENS[t]!;
  }
  if (n < 1000) {
    const h = Math.floor(n / 100);
    const r = n % 100;
    return r ? `${ONES[h]} hundred ${englishOrdinal(r)}` : `${ONES[h]} hundredth`;
  }
  return englishCardinal(n) + "th";
}

function chineseNumber(n: number, financial: boolean): string {
  if (!Number.isFinite(n) || !Number.isInteger(n) || n < 0 || n > 99999999) return String(n);
  const digits = financial
    ? ["零", "壹", "贰", "叁", "肆", "伍", "陆", "柒", "捌", "玖"]
    : ["〇", "一", "二", "三", "四", "五", "六", "七", "八", "九"];
  const units = financial ? ["", "拾", "佰", "仟"] : ["", "十", "百", "千"];
  if (n === 0) return digits[0]!;
  const section = (v: number): string => {
    let out = "";
    let zero = false;
    const s = String(v);
    for (let i = 0; i < s.length; i++) {
      const d = Number(s[i]);
      const pos = s.length - 1 - i;
      if (d === 0) {
        zero = out.length > 0;
      } else {
        if (zero) {
          out += digits[0];
          zero = false;
        }
        if (!(d === 1 && pos === 1 && out === "" && !financial)) out += digits[d];
        out += units[pos];
      }
    }
    return out;
  };
  const high = Math.floor(n / 10000);
  const low = n % 10000;
  if (!high) return section(low);
  const highStr = section(high) + (financial ? "万" : "万");
  if (!low) return highStr;
  return highStr + (low < 1000 ? digits[0] : "") + section(low);
}

export function formatNumericPicture(rawNum: number, picture: string): string {
  if (!Number.isFinite(rawNum)) return String(rawNum);
  const sections = picture.split(";");
  const chosen =
    rawNum < 0 && sections[1]
      ? sections[1]
      : rawNum === 0 && sections[2]
        ? sections[2]
        : sections[0]!;
  if (!/[0#]/.test(chosen)) {
    return chosen.replaceAll("'", "").replaceAll('"', "");
  }
  const isPercent = chosen.includes("%");
  const value = (rawNum < 0 && sections[1] ? Math.abs(rawNum) : rawNum) * (isPercent ? 100 : 1);
  const hasComma = chosen.includes(",");
  const dotIdx = chosen.indexOf(".");
  let decimals = 0;
  if (dotIdx >= 0) {
    const fracPart = chosen.slice(dotIdx + 1);
    decimals = (fracPart.match(/[0#]/g) ?? []).length;
  }
  const fixed = value.toFixed(decimals);
  const [intPart = "0", fracPart = ""] = fixed.split(".");
  const groupedInt = hasComma ? intPart.replace(/\B(?=(\d{3})+(?!\d))/g, ",") : intPart;
  const formatted = decimals > 0 ? `${groupedInt}.${fracPart}` : groupedInt;
  const prefixMatch = /^[^0#.,]+/.exec(chosen);
  const suffixMatch = /[^0#.,]+$/.exec(chosen);
  const prefix = prefixMatch ? prefixMatch[0].replaceAll("'", "").replaceAll('"', "") : "";
  const suffix = suffixMatch ? suffixMatch[0].replaceAll("'", "").replaceAll('"', "") : "";
  return `${prefix}${formatted}${suffix}`;
}

export function applyWordFieldSwitches(rawValue: string, instruction: string): string {
  let result = rawValue;
  const numSwitch = /\\#\s+(?:"([^"]+)"|([^\s\\]+))/.exec(instruction);
  if (numSwitch) {
    const pic = numSwitch[1] ?? numSwitch[2] ?? "";
    const parsed = Number(result.replace(/,/g, "").trim());
    if (pic && Number.isFinite(parsed) && result.trim() !== "") {
      result = formatNumericPicture(parsed, pic);
    }
  }
  const dateSwitch = /\\@\s+(?:"([^"]+)"|([^\s\\]+))/.exec(instruction);
  if (dateSwitch && result.trim()) {
    const parsedDate = Date.parse(result);
    if (Number.isFinite(parsedDate)) {
      result = formatWordDatePicture(new Date(parsedDate), instruction, false);
    }
  }
  for (const match of instruction.matchAll(/\\\*\s+([A-Za-z0-9]+)/g)) {
    const fmt = match[1]!;
    const lower = fmt.toLowerCase();
    const numeric = Number(result.trim());
    if (lower === "upper") result = result.toUpperCase();
    else if (lower === "lower") result = result.toLowerCase();
    else if (lower === "firstcap")
      result = result.replace(/^(\s*)(\S)/, (_, ws: string, ch: string) => ws + ch.toUpperCase());
    else if (lower === "caps")
      result = result.replace(/\b([a-z])/g, (m) => m.toUpperCase());
    else if (fmt === "ROMAN" && Number.isInteger(numeric) && numeric > 0)
      result = pageLabel(numeric, "upperRoman");
    else if (fmt === "roman" && Number.isInteger(numeric) && numeric > 0)
      result = pageLabel(numeric, "lowerRoman");
    else if (fmt === "ALPHABETIC" && Number.isInteger(numeric) && numeric > 0)
      result = pageLabel(numeric, "upperLetter");
    else if (fmt === "alphabetic" && Number.isInteger(numeric) && numeric > 0)
      result = pageLabel(numeric, "lowerLetter");
    else if (lower === "arabic" && Number.isFinite(numeric))
      result = String(Math.trunc(numeric));
    else if (lower === "chinesenum3" && Number.isInteger(numeric))
      result = chineseNumber(numeric, false);
    else if (lower === "chinesenum2" && Number.isInteger(numeric))
      result = chineseNumber(numeric, true);
    else if (lower === "cardtext" && Number.isInteger(numeric))
      result = englishCardinal(numeric);
    else if (lower === "ordtext" && Number.isInteger(numeric))
      result = englishOrdinal(numeric);
    else if (lower === "ordinal" && Number.isInteger(numeric)) {
      const mod100 = Math.abs(numeric) % 100;
      const mod10 = Math.abs(numeric) % 10;
      const suffix =
        mod100 >= 11 && mod100 <= 13
          ? "th"
          : mod10 === 1
            ? "st"
            : mod10 === 2
              ? "nd"
              : mod10 === 3
                ? "rd"
                : "th";
      result = `${numeric}${suffix}`;
    }
  }
  return result;
}

export function tokenizeWordFieldArgs(input: string): string[] {
  const tokens: string[] = [];
  let i = 0;
  const s = input.trim();
  while (i < s.length) {
    while (i < s.length && /\s/.test(s[i]!)) i++;
    if (i >= s.length) break;
    if (s[i] === '"') {
      i++;
      let buf = "";
      while (i < s.length && s[i] !== '"') {
        if (s[i] === "\\" && i + 1 < s.length && (s[i + 1] === '"' || s[i + 1] === "\\")) {
          buf += s[i + 1];
          i += 2;
        } else {
          buf += s[i++];
        }
      }
      if (i < s.length && s[i] === '"') i++;
      tokens.push(buf);
    } else if (s[i] === "\\" && i + 1 < s.length) {
      tokens.push(s.slice(i, i + 2));
      i += 2;
    } else {
      let buf = "";
      while (i < s.length && !/\s/.test(s[i]!)) buf += s[i++];
      tokens.push(buf);
    }
  }
  return tokens;
}

function wildcardMatch(text: string, pattern: string): boolean {
  const escaped = pattern
    .replace(/[.+^${}()|[\]\\]/g, "\\$&")
    .replaceAll("*", ".*")
    .replaceAll("?", ".");
  return new RegExp(`^${escaped}$`, "i").test(text);
}

export function evaluateWordCondition(leftRaw: string, op: string, rightRaw: string): boolean {
  const leftNum = Number(leftRaw.replace(/,/g, "").trim());
  const rightNum = Number(rightRaw.replace(/,/g, "").trim());
  const bothNum =
    leftRaw.trim() !== "" &&
    rightRaw.trim() !== "" &&
    Number.isFinite(leftNum) &&
    Number.isFinite(rightNum);
  const cmp = bothNum
    ? leftNum - rightNum
    : leftRaw.localeCompare(rightRaw, undefined, { sensitivity: "accent" });
  switch (op) {
    case "=":
    case "==":
      return bothNum ? leftNum === rightNum : wildcardMatch(leftRaw, rightRaw);
    case "<>":
    case "!=":
      return bothNum ? leftNum !== rightNum : !wildcardMatch(leftRaw, rightRaw);
    case "<":
      return cmp < 0;
    case "<=":
      return cmp <= 0;
    case ">":
      return cmp > 0;
    case ">=":
      return cmp >= 0;
    default:
      return false;
  }
}

export function evaluateIfOrCompare(instruction: string, isCompare: boolean): string {
  const body = instruction.replace(/^\s*(?:IF|COMPARE)\s+/i, "");
  const tokens = tokenizeWordFieldArgs(body).filter((t) => !/^\\[*#@]$/.test(t));
  // Strip switch pairs (\* Upper, etc.) from trailing tokens
  const positional: string[] = [];
  for (let i = 0; i < tokens.length; i++) {
    const tok = tokens[i]!;
    if (/^\\[a-z*#@]$/i.test(tok)) {
      i++;
      continue;
    }
    positional.push(tok);
  }
  if (positional.length < 3) return isCompare ? "0" : "";
  const matched = evaluateWordCondition(positional[0]!, positional[1]!, positional[2]!);
  if (isCompare) return matched ? "1" : "0";
  return matched ? (positional[3] ?? "") : (positional[4] ?? "");
}

export function evaluateSymbolField(instruction: string): string {
  const match = /^\s*SYMBOL\s+(0x[0-9a-f]+|\d+)/i.exec(instruction);
  if (!match) return "";
  const raw = match[1]!;
  const code = raw.toLowerCase().startsWith("0x") ? parseInt(raw.slice(2), 16) : Number(raw);
  if (!Number.isInteger(code) || code <= 0 || code > 0x10ffff) return "";
  try {
    return String.fromCodePoint(code);
  } catch {
    return "";
  }
}

export function evaluateCitationField(
  instruction: string,
  word: DeepReadonly<WordMetadata> | undefined,
  citationOrder: ReadonlyMap<string, number>,
): string {
  const match = /^\s*CITATION\s+(?:"([^"]+)"|([^\s\\]+))/i.exec(instruction);
  const tag = match?.[1] ?? match?.[2];
  if (!tag) return "(未知引文)";
  const bib = word?.bibliography;
  const source = bib?.sources.find((item) => item.tag === tag);
  if (!source) return `(${tag})`;
  const style: WordBibliographyStyle = bib?.style ?? "APA";
  const pagesMatch = /\\p\s+(?:"([^"]+)"|([^\s\\]+))/i.exec(instruction);
  const pages = pagesMatch?.[1] ?? pagesMatch?.[2];
  const suppressAuthor = /\\n\b/i.test(instruction);
  const suppressYear = /\\y\b/i.test(instruction);
  const suppressTitle = /\\t\b/i.test(instruction);
  const ordinal = citationOrder.get(tag);
  return formatWordCitation(source, style, {
    ...(pages ? { pages } : {}),
    suppressAuthor,
    suppressYear,
    suppressTitle,
    ...(ordinal !== undefined ? { ordinal } : {}),
  });
}

export function evaluateBibliographyField(
  word: DeepReadonly<WordMetadata> | undefined,
  citedTags: readonly string[],
): string {
  const bib = word?.bibliography;
  if (!bib?.sources.length) return "此文档中没有引文源。";
  const style: WordBibliographyStyle = bib.style ?? "APA";
  return formatWordBibliography(bib.sources, style, citedTags)
    .map((entry) => entry.text)
    .join("\n");
}

export interface ParsedXeEntry {
  readonly main: string;
  readonly sub?: string;
  readonly pageNumber: number;
  readonly pageLabel: string;
  readonly bold: boolean;
  readonly italic: boolean;
  readonly seeAlso?: string;
}

export function parseXeInstruction(instruction: string): Omit<ParsedXeEntry, "pageNumber" | "pageLabel"> | undefined {
  const body = instruction.replace(/^\s*XE\s+/i, "");
  const tokens = tokenizeWordFieldArgs(body);
  if (!tokens.length) return undefined;
  let rawEntry = "";
  let bold = false;
  let italic = false;
  let seeAlso: string | undefined;
  for (let i = 0; i < tokens.length; i++) {
    const tok = tokens[i]!;
    if (tok === "\\b") bold = true;
    else if (tok === "\\i") italic = true;
    else if (tok === "\\t" && i + 1 < tokens.length) seeAlso = tokens[++i];
    else if (/^\\[a-z]$/i.test(tok)) {
      if (i + 1 < tokens.length && !tokens[i + 1]!.startsWith("\\")) i++;
    } else if (!rawEntry) {
      rawEntry = tok;
    }
  }
  if (!rawEntry.trim()) return undefined;
  const parts = rawEntry.split(":");
  const main = parts[0]!.trim();
  const sub = parts.slice(1).join(":").trim() || undefined;
  return {
    main,
    ...(sub ? { sub } : {}),
    bold,
    italic,
    ...(seeAlso ? { seeAlso } : {}),
  };
}

function compressPages(entries: readonly { num: number; label: string }[], sep: string, rangeSep: string): string {
  const unique = new Map<number, string>();
  for (const item of entries) unique.set(item.num, item.label);
  const sorted = [...unique.entries()].sort((a, b) => a[0] - b[0]);
  const out: string[] = [];
  for (let i = 0; i < sorted.length; ) {
    let j = i;
    while (j + 1 < sorted.length && sorted[j + 1]![0] === sorted[j]![0] + 1) j++;
    if (j >= i + 2) out.push(`${sorted[i]![1]}${rangeSep}${sorted[j]![1]}`);
    else {
      for (let k = i; k <= j; k++) out.push(sorted[k]![1]);
    }
    i = j + 1;
  }
  return out.join(sep);
}

export function formatIndexField(entries: readonly ParsedXeEntry[], instruction: string): string {
  if (!entries.length) return "未找到索引项。";
  const sepMatch = /\\e\s+"([^"]*)"/i.exec(instruction);
  const rangeMatch = /\\g\s+"([^"]*)"/i.exec(instruction);
  const headMatch = /\\h\s+"([^"]*)"/i.exec(instruction);
  const sep = sepMatch?.[1] ?? ", ";
  const rangeSep = rangeMatch?.[1] ?? "-";
  const showAlphaHeader = Boolean(headMatch && headMatch[1]!.length > 0);

  const grouped = new Map<
    string,
    {
      pages: { num: number; label: string }[];
      seeAlso: Set<string>;
      subs: Map<string, { pages: { num: number; label: string }[]; seeAlso: Set<string> }>;
    }
  >();
  for (const item of entries) {
    let group = grouped.get(item.main);
    if (!group) {
      group = { pages: [], seeAlso: new Set(), subs: new Map() };
      grouped.set(item.main, group);
    }
    if (item.sub) {
      let sub = group.subs.get(item.sub);
      if (!sub) {
        sub = { pages: [], seeAlso: new Set() };
        group.subs.set(item.sub, sub);
      }
      if (item.seeAlso) sub.seeAlso.add(item.seeAlso);
      else sub.pages.push({ num: item.pageNumber, label: item.pageLabel });
    } else {
      if (item.seeAlso) group.seeAlso.add(item.seeAlso);
      else group.pages.push({ num: item.pageNumber, label: item.pageLabel });
    }
  }
  const mains = [...grouped.keys()].sort((a, b) =>
    a.localeCompare(b, undefined, { sensitivity: "base" }),
  );
  const lines: string[] = [];
  let lastLetter = "";
  for (const main of mains) {
    if (showAlphaHeader) {
      const firstChar = (main[0] ?? "#").toUpperCase();
      const letter = /[A-Z]/.test(firstChar) ? firstChar : firstChar;
      if (letter !== lastLetter) {
        if (lines.length) lines.push("");
        lines.push(letter);
        lastLetter = letter;
      }
    }
    const group = grouped.get(main)!;
    const refs: string[] = [];
    if (group.pages.length) refs.push(compressPages(group.pages, sep, rangeSep));
    if (group.seeAlso.size) refs.push([...group.seeAlso].join("; "));
    lines.push(refs.length ? `${main}, ${refs.join("; ")}` : main);
    const subKeys = [...group.subs.keys()].sort((a, b) =>
      a.localeCompare(b, undefined, { sensitivity: "base" }),
    );
    for (const sub of subKeys) {
      const subData = group.subs.get(sub)!;
      const subRefs: string[] = [];
      if (subData.pages.length) subRefs.push(compressPages(subData.pages, sep, rangeSep));
      if (subData.seeAlso.size) subRefs.push([...subData.seeAlso].join("; "));
      lines.push(`    ${sub}${subRefs.length ? `, ${subRefs.join("; ")}` : ""}`);
    }
  }
  return lines.join("\n");
}

export const TOA_CATEGORIES: readonly string[] = [
  "全部",
  "案例",
  "法规",
  "其他引文",
  "规则",
  "条约",
  "条例",
  "宪法条文",
  "类别 8",
  "类别 9",
  "类别 10",
  "类别 11",
  "类别 12",
  "类别 13",
  "类别 14",
  "类别 15",
  "类别 16",
];

export interface ParsedTaEntry {
  readonly longCitation: string;
  readonly shortCitation: string;
  readonly category: number;
  readonly pageNumber: number;
  readonly pageLabel: string;
}

export function parseTaInstruction(instruction: string): Omit<ParsedTaEntry, "pageNumber" | "pageLabel"> | undefined {
  const body = instruction.replace(/^\s*TA\s+/i, "");
  const tokens = tokenizeWordFieldArgs(body);
  let longCitation = "";
  let shortCitation = "";
  let category = 1;
  for (let i = 0; i < tokens.length; i++) {
    const tok = tokens[i]!;
    if (tok === "\\l" && i + 1 < tokens.length) longCitation = tokens[++i]!;
    else if (tok === "\\s" && i + 1 < tokens.length) shortCitation = tokens[++i]!;
    else if (tok === "\\c" && i + 1 < tokens.length) {
      const num = Number(tokens[++i]);
      if (Number.isInteger(num) && num >= 1 && num <= 16) category = num;
    }
  }
  if (!longCitation && !shortCitation) return undefined;
  return {
    longCitation: longCitation || shortCitation,
    shortCitation: shortCitation || longCitation,
    category,
  };
}

export function formatToaField(entries: readonly ParsedTaEntry[], instruction: string): string {
  const catMatch = /\\c\s+(\d+)/i.exec(instruction);
  const targetCategory = catMatch ? Number(catMatch[1]) : undefined;
  const usePassim = /\\p\b/i.test(instruction);
  const showHeader = /\\h\b/i.test(instruction);
  const filtered = entries.filter(
    (entry) => targetCategory === undefined || entry.category === targetCategory,
  );
  if (!filtered.length) return "未找到引文目录项。";

  const byCategory = new Map<number, Map<string, { long: string; pages: { num: number; label: string }[] }>>();
  for (const item of filtered) {
    let catMap = byCategory.get(item.category);
    if (!catMap) {
      catMap = new Map();
      byCategory.set(item.category, catMap);
    }
    const key = item.shortCitation.toLowerCase();
    let rec = catMap.get(key);
    if (!rec) {
      rec = { long: item.longCitation, pages: [] };
      catMap.set(key, rec);
    }
    rec.pages.push({ num: item.pageNumber, label: item.pageLabel });
  }
  const lines: string[] = [];
  for (const cat of [...byCategory.keys()].sort((a, b) => a - b)) {
    if (showHeader) lines.push(TOA_CATEGORIES[cat] ?? `类别 ${cat}`);
    const catMap = byCategory.get(cat)!;
    const sorted = [...catMap.values()].sort((a, b) =>
      a.long.localeCompare(b.long, undefined, { sensitivity: "base" }),
    );
    for (const item of sorted) {
      const distinctPages = new Set(item.pages.map((p) => p.num));
      const pageStr =
        usePassim && distinctPages.size >= 5
          ? "passim"
          : compressPages(item.pages, ", ", "-");
      lines.push(`${item.long}\t${pageStr}`);
    }
  }
  return lines.join("\n");
}
