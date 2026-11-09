import { DocumentError } from "../errors.js";
import { checkLimit, freeze, invalid } from "../document/validation.js";
import type { WorkbookDateSystem } from "./calendar.js";
import { workbookBuiltinNumberFormat } from "./builtin-number-format.js";
import { workbookDateNames } from "./date-names.js";
import type { WorkbookDateNames } from "./date-names.js";
import { resolveWorkbookDisplayLocale, workbookLocaleSeparators } from "./display-locale.js";
import type { WorkbookDisplayLocale } from "./display-locale.js";
import { formatDateSection } from "./format-date.js";
import { formatGeneralSection } from "./format-general-section.js";
import type { WorkbookNumberMeasure } from "./format-general-section.js";
import type { NumberDisplayPart, NumberSeparators } from "./format-numeric.js";
import { formatLiteralSection, selectNumberFormat } from "./format-select.js";
import { formatWorkbookValue } from "./format-value.js";
import { workbookLocaleDatePattern } from "./locale-date-pattern.js";
import { compileNumberFormat } from "./number-format.js";
import type { NumberFormatColor, NumberFormatSection } from "./number-format.js";

export interface WorkbookNumberFormatterOptions {
  readonly dateSystem: WorkbookDateSystem;
  readonly hostLocale?: string;
  readonly documentLocale?: string;
}
export type WorkbookDisplayFormat = ({ readonly code: string; readonly builtinId?: never } | { readonly builtinId: number; readonly code?: never }) & { readonly formatId?: string };
interface DisplayMetadata { readonly locale: string; readonly localeSource: WorkbookDisplayLocale["source"]; readonly color?: NumberFormatColor }
export type WorkbookNumberDisplay = DisplayMetadata & (
  | { readonly kind: "parts"; readonly parts: readonly NumberDisplayPart[] }
  | { readonly kind: "overflow"; readonly reason: "width" | "invalidDate" | "unmatchedCondition" }
);
export interface WorkbookNumberFormatter {
  /** Resolve the source code with this formatter's explicit display locale. */
  code(definition?: WorkbookDisplayFormat): string;
  /** With no measurement, General uses its full display envelope. Other formats
   * produce parts for measured spacing/fill and width handling by the renderer. */
  format(value: number | string, definition?: WorkbookDisplayFormat, measurement?: WorkbookNumberMeasure): WorkbookNumberDisplay;
  dispose(): void;
}
interface LocaleContext { display: WorkbookDisplayLocale; separators: NumberSeparators; names?: WorkbookDateNames }

/** Instance-owned bounded caches contain format source and locale vocabulary,
 * never cell values. No document's source or locale enters a global cache. */
export function createWorkbookNumberFormatter(input: WorkbookNumberFormatterOptions): WorkbookNumberFormatter {
  if (!input || !["1900", "1904"].includes(input.dateSystem) || Object.keys(input).some(key => !["dateSystem", "hostLocale", "documentLocale"].includes(key))) invalid("Invalid workbook number formatter options.");
  for (const locale of [input.hostLocale, input.documentLocale]) if (locale !== undefined) {
    if (typeof locale !== "string" || !locale) invalid("A display locale must be a nonempty language tag."); checkLimit(locale.length, 255);
  }
  const options = { ...input }, formats = new Map<string, readonly NumberFormatSection[]>(), locales = new Map<number, LocaleContext>(); let disposed = false;
  const context = (native = 0): LocaleContext => {
    let known = locales.get(native);
    if (!known) {
      const display = resolveWorkbookDisplayLocale({ ...options, formatLocale: native });
      known = { display, separators: workbookLocaleSeparators(display) };
      if (locales.size >= 64) locales.delete(locales.keys().next().value!);
      locales.set(native, known);
    }
    return known;
  };
  const compile = (code: string): readonly NumberFormatSection[] => {
    let known = formats.get(code);
    if (!known) {
      known = compileNumberFormat(code);
      if (formats.size >= 256) formats.delete(formats.keys().next().value!);
      formats.set(code, known);
    }
    return known;
  };
  const code = (definition: WorkbookDisplayFormat = { builtinId: 0 }): string => {
    if (disposed) throw new DocumentError("InvalidOperation", "This workbook number formatter has been disposed.");
    if (!definition || typeof definition !== "object" || Object.keys(definition).some(key => !["code", "builtinId", "formatId"].includes(key))
      || (definition.code !== undefined) === (definition.builtinId !== undefined) || definition.code !== undefined && typeof definition.code !== "string"
      || definition.formatId !== undefined && (typeof definition.formatId !== "string" || !definition.formatId)) invalid("A display format requires exactly one source code or built-in ID.");
    const result = definition.code ?? workbookBuiltinNumberFormat(definition.builtinId!, context().display); compile(result); return result;
  };
  return Object.freeze({
    code,
    format(value: number | string, definition: WorkbookDisplayFormat = { builtinId: 0 }, measurement?: WorkbookNumberMeasure): WorkbookNumberDisplay {
      if (disposed) throw new DocumentError("InvalidOperation", "This workbook number formatter has been disposed.");
      if (typeof value !== "string" && (typeof value !== "number" || !Number.isFinite(value))) invalid("A number format requires a finite numeric or text scalar.");
      if (typeof value === "string") checkLimit(value.length, 32767);
      if (measurement && (!Number.isFinite(measurement.availableWidth) || measurement.availableWidth < 0 || typeof measurement.measure !== "function")) invalid("Number display requires a finite width and text measurement function.");
      const sections = compile(code(definition));
      const selection = selectNumberFormat(value, sections), section = "section" in selection ? selection.section : undefined;
      const selected = context(section?.locale), metadata: DisplayMetadata = { locale: selected.display.locale, localeSource: selected.display.source, ...(section?.color ? { color: section.color } : {}) };
      if (selection.kind === "unmatchedCondition") return freeze({ ...metadata, kind: "overflow", reason: "unmatchedCondition" });
      if (typeof value === "string") return freeze({ ...metadata, kind: "parts", parts: selection.kind === "text"
        ? formatLiteralSection(value, selection.section, true) : [{ kind: "text", text: value, source: { from: 0, to: value.length } }] });
      let result = formatWorkbookValue(value, sections, selected.separators);
      if (typeof value === "number" && selected.display.systemFormat) {
        const pattern = compile(workbookLocaleDatePattern(selected.display, selected.display.systemFormat))[0]!;
        result = { kind: "date", value, section: pattern, automaticNegative: false };
      }
      if (result.kind === "parts") return freeze({ ...metadata, kind: "parts", parts: result.parts });
      if (result.kind === "date") {
        selected.names ??= workbookDateNames(selected.display);
        const date = formatDateSection(result.value, result.section, options.dateSystem, selected.names, selected.separators.decimal);
        return freeze(date.kind === "parts" ? { ...metadata, ...date } : { ...metadata, kind: "overflow", reason: "invalidDate" });
      }
      if (result.kind === "general") {
        const parts = formatGeneralSection(result.value, result.section, result.automaticNegative, selected.separators.decimal, measurement);
        return freeze(parts ? { ...metadata, kind: "parts", parts } : { ...metadata, kind: "overflow", reason: "width" });
      }
      return freeze({ ...metadata, kind: "overflow", reason: "unmatchedCondition" });
    },
    dispose(): void { disposed = true; formats.clear(); locales.clear(); },
  });
}
