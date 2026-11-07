import type { DeepReadonly } from "../runtime.js";
import type { WorkbookAutoFilter, WorkbookFilterCriterion, WorkbookFilterDate, WorkbookSheet } from "../contracts/workbook.js";
import type { WorkbookReferenceArea, WorkbookReferenceLocation } from "./reference.js";
import type { ScalarResolved, WorkbookScalarContext, WorkbookScalarUnavailable } from "./scalar.js";
import { workbookSerialDate } from "./calendar.js";
import { dateSerial } from "./scalar-dates.js";
import { matchesWorkbookCriterion } from "./scalar-criteria.js";
import { scalarToText } from "./scalar-values.js";

export interface WorkbookFilterContext extends WorkbookScalarContext {
  readonly dateSystem: "1900" | "1904";
  read(location: WorkbookReferenceLocation): ScalarResolved;
  stored(area: DeepReadonly<WorkbookReferenceArea>): Generator<WorkbookReferenceLocation>;
  step(): void;
}
type Filter = DeepReadonly<WorkbookAutoFilter>;
type Criterion = DeepReadonly<WorkbookFilterCriterion>;
const operators = { equal: "=", notEqual: "<>", lessThan: "<", lessThanOrEqual: "<=", greaterThan: ">", greaterThanOrEqual: ">=" };
const fields = ["year", "month", "day", "hour", "minute", "second"] as const;
function dateMatches(value: number, group: DeepReadonly<WorkbookFilterDate>, system: "1900" | "1904"): boolean {
  const date = workbookSerialDate(value, system); if (!date) return false;
  const seconds = Math.floor(date.dayFraction * 86400 + 1e-7) % 86400;
  const parts = { ...date, hour: Math.floor(seconds / 3600), minute: Math.floor(seconds / 60) % 60, second: seconds % 60 };
  return fields.slice(0, fields.indexOf(group.grouping) + 1).every(field => parts[field] === group[field]);
}
/** One immutable query owns filter statistics. No native cached visibility or
 * default clock is mistaken for the result of a current-source filter. */
export function createWorkbookRowFilter(sheet: DeepReadonly<WorkbookSheet>, context: WorkbookFilterContext): (row: number) => boolean | WorkbookScalarUnavailable {
  const filters: Filter[] = [...sheet.autoFilter ? [sheet.autoFilter] : [], ...sheet.tables?.flatMap(table => table.autoFilter ? [table.autoFilter] : []) ?? []];
  const statistics = new Map<Criterion, number | WorkbookScalarUnavailable>();
  function threshold(filter: Filter, column: number, criterion: Criterion): number | WorkbookScalarUnavailable {
    const known = statistics.get(criterion); if (known !== undefined) return known;
    const values: number[] = [], area = { sheetId: sheet.sheetId, from: { row: filter.range.from.row + 1, column }, to: { row: filter.range.to.row, column } };
    if (area.from.row <= area.to.row) for (const location of context.stored(area)) {
      context.step(); const value = context.read(location); if (!("type" in value)) return value;
      if (value.type === "number") values.push(value.value);
    }
    let result = NaN;
    if (values.length && criterion.kind === "top") {
      values.sort((a, b) => { context.step(); return a - b; });
      const count = Math.min(values.length, Math.max(1, criterion.percent ? Math.ceil(values.length * criterion.count / 100) : Math.floor(criterion.count)));
      result = values[criterion.top ? values.length - count : count - 1]!;
    } else if (values.length) {
      let scale = 0; for (const value of values) { context.step(); scale = Math.max(scale, Math.abs(value)); }
      result = scale ? values.reduce((sum, value) => { context.step(); return sum + value / scale / values.length; }, 0) * scale : 0;
    }
    statistics.set(criterion, result); return result;
  }
  function dynamic(value: number, type: Extract<Criterion, { kind: "dynamic" }>["type"]): boolean | WorkbookScalarUnavailable {
    const date = workbookSerialDate(value, context.dateSystem); if (!date) return false;
    if (/^M\d+$/.test(type)) return date.month === Number(type.slice(1));
    if (/^Q[1-4]$/.test(type)) return Math.floor((date.month - 1) / 3) + 1 === Number(type.slice(1));
    if (!context.now) return { kind: "unavailable", reason: "calculation-required" };
    const clock = new Date(context.now.unixMilliseconds + context.now.utcOffsetMinutes * 60000), year = clock.getUTCFullYear(), month = clock.getUTCMonth() + 1;
    const today = dateSerial(year, month, clock.getUTCDate(), context.dateSystem), day = Math.floor(value);
    if (["today", "tomorrow", "yesterday"].includes(type)) return day === today + (type === "tomorrow" ? 1 : type === "yesterday" ? -1 : 0);
    if (type === "yearToDate") return day >= dateSerial(year, 1, 1, context.dateSystem) && day <= today;
    const offset = type.startsWith("next") ? 1 : type.startsWith("last") ? -1 : 0;
    if (type.endsWith("Week")) {
      // Office's dynamic calendar weeks begin on Sunday. Use the real host
      // calendar weekday rather than the 1900 compatibility weekday.
      const first = today - clock.getUTCDay() + offset * 7; return day >= first && day < first + 7;
    }
    const months = type.endsWith("Month") ? 1 : type.endsWith("Quarter") ? 3 : 12;
    const firstMonth = months === 1 ? month : months === 3 ? Math.floor((month - 1) / 3) * 3 + 1 : 1;
    const first = dateSerial(year, firstMonth + offset * months, 1, context.dateSystem), last = dateSerial(year, firstMonth + (offset + 1) * months, 1, context.dateSystem);
    return day >= first && day < last;
  }
  function matches(filter: Filter, row: number, column: number, criterion: Criterion): boolean | WorkbookScalarUnavailable {
    const value = context.read({ sheetId: sheet.sheetId, row, column }); if (!("type" in value)) return value;
    if (criterion.kind === "values") {
      if (value.type === "blank" || value.type === "text" && value.value === "") return criterion.blank;
      const text = value.type === "error" ? value.value : scalarToText(value);
      if (typeof text === "string" && criterion.values.some(item => { context.step(); return item.toUpperCase() === text.toUpperCase(); })) return true;
      return value.type === "number" && criterion.dates.some(group => { context.step(); return dateMatches(value.value, group, context.dateSystem); });
    }
    if (criterion.kind === "custom") {
      const tests = criterion.tests.map(test => matchesWorkbookCriterion(value, { type: "text", value: operators[test.operator] + test.value }, context.step));
      return criterion.and ? tests.every(Boolean) : tests.some(Boolean);
    }
    if (value.type !== "number") return false;
    if (criterion.kind === "top" || ["aboveAverage", "belowAverage"].includes(criterion.type)) {
      const bound = threshold(filter, column, criterion); if (typeof bound !== "number") return bound;
      return criterion.kind === "top" ? criterion.top ? value.value >= bound : value.value <= bound : criterion.type === "aboveAverage" ? value.value > bound : value.value < bound;
    }
    return dynamic(value.value, criterion.type);
  }
  return row => {
    for (const filter of filters) {
      context.step(); if (row <= filter.range.from.row || row > filter.range.to.row) continue;
      for (const column of filter.columns) if (column.criterion) {
        context.step(); const match = matches(filter, row, column.column, column.criterion); if (typeof match !== "boolean") return match; if (!match) return true;
      }
    }
    return false;
  };
}
