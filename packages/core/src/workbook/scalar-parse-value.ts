import type { ScalarError, ScalarValue } from "./scalar-values.js";
import { scalarError, scalarToNumber } from "./scalar-values.js";

/** Explicit, locale-bound parsing. Date.parse and the machine's timezone are
 * intentionally absent; ISO/native date fields and time fields stay typed. */
export function parseWorkbookValueText(source: string, locale: string, system: "1900" | "1904"): number | ScalarError {
  const value = source.trim();
  const parts = new Intl.NumberFormat(locale).formatToParts(1234.5);
  const group = parts.find(part => part.type === "group")?.value ?? ",", decimal = parts.find(part => part.type === "decimal")?.value ?? ".";
  const numeric = value.replace(/^[¥￥$€£]\s*|\s*[¥￥$€£]$/g, "").split(group).join("").split(decimal).join(".");
  const number = scalarToNumber({ type: "text", value: numeric }); if (typeof number === "number") return number;
  let dateText = value, fraction = 0;
  const time = /(?:^|[ T])([0-9]{1,2}):([0-9]{1,2})(?::([0-9]{1,2}(?:\.[0-9]+)?))?\s*(AM|PM)?$/i.exec(value);
  if (time) {
    let hour = Number(time[1]); const minute = Number(time[2]), second = Number(time[3] ?? 0), period = time[4]?.toUpperCase();
    if (minute >= 60 || second >= 60 || period && (hour < 1 || hour > 12) || !period && hour >= 24) return scalarError("#VALUE!");
    if (period) hour = hour % 12 + (period === "PM" ? 12 : 0);
    fraction = (hour * 3600 + minute * 60 + second) / 86400; dateText = value.slice(0, time.index).trim();
    if (!dateText) return fraction;
  }
  let year: number, month: number, day: number;
  const iso = /^([0-9]{4})[-\/]([0-9]{1,2})[-\/]([0-9]{1,2})$/.exec(dateText), chinese = /^([0-9]{4})年([0-9]{1,2})月([0-9]{1,2})日?$/.exec(dateText);
  if (iso || chinese) { const match = iso ?? chinese!; year = Number(match[1]); month = Number(match[2]); day = Number(match[3]); }
  else {
    const fields = /^([0-9]{1,2})[\/.-]([0-9]{1,2})[\/.-]([0-9]{2}|[0-9]{4})$/.exec(dateText); if (!fields) return scalarError("#VALUE!");
    const order = new Intl.DateTimeFormat(locale, { year: "numeric", month: "numeric", day: "numeric", timeZone: "UTC" }).formatToParts(new Date(Date.UTC(2001, 10, 22))).filter(part => ["year", "month", "day"].includes(part.type)).map(part => part.type);
    if (order[0] === "year") return scalarError("#VALUE!");
    year = Number(fields[3]); if (fields[3]!.length === 2) year += year < 30 ? 2000 : 1900;
    month = Number(fields[order[0] === "month" ? 1 : 2]); day = Number(fields[order[0] === "day" ? 1 : 2]);
  }
  if (year < (system === "1900" ? 1900 : 1904) || year > 9999 || month < 1 || month > 12 || day < 1) return scalarError("#VALUE!");
  const last = system === "1900" && year === 1900 && month === 2 ? 29 : new Date(Date.UTC(year, month, 0)).getUTCDate(); if (day > last) return scalarError("#VALUE!");
  if (system === "1900" && year === 1900 && month === 2 && day === 29) return 60 + fraction;
  const timestamp = Date.UTC(year, month - 1, day), epoch = system === "1900" ? Date.UTC(1899, 11, 31) : Date.UTC(1904, 0, 1);
  return (timestamp - epoch) / 86_400_000 + (system === "1900" && timestamp >= Date.UTC(1900, 2, 1) ? 1 : 0) + fraction;
}
export function scalarDateNumber(value: ScalarValue, locale: string, system: "1900" | "1904"): number | ScalarError {
  return value.type === "text" ? parseWorkbookValueText(value.value, locale, system) : scalarToNumber(value);
}
