import type { ScalarFunctionContext, ScalarResolved } from "./scalar.js";
import type { ScalarValue } from "./scalar-values.js";
import { scalarError, scalarNumber, scalarToNumber } from "./scalar-values.js";
import { workbookSerialDate } from "./calendar.js";
import { scalarDateNumber } from "./scalar-parse-value.js";

const dayMs = 86_400_000;
export function dateSerial(year: number, month: number, day: number, system: "1900" | "1904"): number {
  const normalized = new Date(0); normalized.setUTCFullYear(year, month - 1, 1);
  const adjusted = normalized.getTime();
  const epoch = system === "1900" ? Date.UTC(1899, 11, 31) : Date.UTC(1904, 0, 1);
  const serial = (adjusted - epoch) / dayMs + day - 1;
  return serial + (system === "1900" && adjusted >= Date.UTC(1900, 2, 1) ? 1 : 0);
}
export function evaluateDateFunction(name: string, values: readonly ScalarValue[], context: ScalarFunctionContext): ScalarResolved | undefined {
  const counts: Readonly<Record<string, readonly [number, number]>> = { DATE: [3, 3], YEAR: [1, 1], MONTH: [1, 1], DAY: [1, 1], TIME: [3, 3], HOUR: [1, 1], MINUTE: [1, 1], SECOND: [1, 1], DAYS: [2, 2], EDATE: [2, 2], EOMONTH: [2, 2], WEEKDAY: [1, 2], TODAY: [0, 0], NOW: [0, 0] };
  const count = counts[name]; if (!count) return undefined;
  if (values.length < count[0] || values.length > count[1]) return scalarError("#VALUE!");
  const numbers: number[] = [];
  for (const [index, value] of values.entries()) {
    const dateArgument = name === "DAYS" || index === 0 && !["DATE", "TIME"].includes(name);
    const number = dateArgument ? scalarDateNumber(value, context.locale ?? "zh-CN", context.dateSystem) : scalarToNumber(value);
    if (typeof number !== "number") return number; numbers.push(number);
  }
  const a = numbers[0] ?? 0, b = numbers[1] ?? 0, maximum = context.dateSystem === "1900" ? 2_958_465 : 2_957_003;
  const bounded = (serial: number): ScalarResolved => serial >= 0 && serial < maximum + 1 ? scalarNumber(serial) : scalarError("#NUM!");
  if (name === "NOW" || name === "TODAY") {
    if (!context.now) return { kind: "unavailable", reason: "calculation-required" };
    const time = context.now.unixMilliseconds + context.now.utcOffsetMinutes * 60_000, date = new Date(time);
    const serial = dateSerial(date.getUTCFullYear(), date.getUTCMonth() + 1, date.getUTCDate(), context.dateSystem);
    return bounded(name === "TODAY" ? serial : serial + ((time % dayMs + dayMs) % dayMs) / dayMs);
  }
  if (name === "DATE") {
    let year = Math.trunc(a); if (year >= 0 && year < 1900) year += 1900;
    if (year < 0 || year > 9999 || Math.abs(b) > 120_000 || Math.abs(numbers[2]!) > 4_000_000) return scalarError("#NUM!");
    return bounded(dateSerial(year, Math.trunc(b), Math.trunc(numbers[2]!), context.dateSystem));
  }
  if (name === "TIME") {
    if (numbers.some(value => value < 0 || value > 32767)) return scalarError("#NUM!");
    return scalarNumber((Math.trunc(a) * 3600 + Math.trunc(b) * 60 + Math.trunc(numbers[2]!)) % 86400 / 86400);
  }
  if (name === "DAYS") return a < 0 || a >= maximum + 1 || b < 0 || b >= maximum + 1 ? scalarError("#NUM!") : scalarNumber(Math.trunc(a) - Math.trunc(b));
  const date = a >= 0 && a <= maximum + 1 ? workbookSerialDate(a, context.dateSystem) : undefined;
  if (!date) return scalarError("#NUM!");
  if (name === "YEAR" || name === "MONTH" || name === "DAY") return scalarNumber(date[name === "YEAR" ? "year" : name === "MONTH" ? "month" : "day"]);
  if (name === "HOUR" || name === "MINUTE" || name === "SECOND") {
    const seconds = Math.floor(date.dayFraction * 86400 + 1e-7) % 86400;
    return scalarNumber(name === "HOUR" ? Math.floor(seconds / 3600) : name === "MINUTE" ? Math.floor(seconds / 60) % 60 : seconds % 60);
  }
  if (name === "WEEKDAY") {
    const type = numbers.length > 1 ? Math.trunc(b) : 1, starts: Record<number, number> = { 1: 0, 2: 1, 3: 1, 11: 1, 12: 2, 13: 3, 14: 4, 15: 5, 16: 6, 17: 0 };
    if (!Object.hasOwn(starts, type)) return scalarError("#NUM!");
    return scalarNumber((date.weekday - starts[type]! + 7) % 7 + (type === 3 ? 0 : 1));
  }
  if (Math.abs(b) > 120_000) return scalarError("#NUM!");
  const index = date.year * 12 + date.month - 1 + Math.trunc(b), year = Math.floor(index / 12), month = (index % 12 + 12) % 12 + 1;
  if (year < 1900 || year > 9999) return scalarError("#NUM!");
  const last = context.dateSystem === "1900" && year === 1900 && month === 2 ? 29 : new Date(Date.UTC(year, month, 0)).getUTCDate();
  return bounded(dateSerial(year, month, name === "EOMONTH" ? last : Math.min(date.day, last), context.dateSystem));
}
