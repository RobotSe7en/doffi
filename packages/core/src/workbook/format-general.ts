import { invalid } from "../document/validation.js";
import { decimalComponents, decimalMagnitude } from "./format-decimal.js";

export type GeneralNumberDisplay = { readonly kind: "text"; readonly text: string; readonly notation: "decimal" | "scientific" } | { readonly kind: "overflow" };
interface Candidate { readonly text: string; readonly notation: "decimal" | "scientific"; readonly precision: number }

/** General uses the native eleven-position display envelope (excluding the
 * negative sign). Candidate precision is bounded, independent of cell width.
 * Fonts/measurement belong to the caller; no source scalar is rounded in place. */
export function formatGeneralNumber(value: number, decimal: string, availableWidth: number, measure: (text: string) => number, powerAdjustment = 0): GeneralNumberDisplay {
  if (!Number.isFinite(value) || !Number.isFinite(availableWidth) || availableWidth < 0 || !decimal || decimal.length > 8 || /[\u0000-\u001f]/.test(decimal)
    || !Number.isSafeInteger(powerAdjustment) || Math.abs(powerAdjustment) > 8192) invalid("Invalid General number-format scalar, width, decimal symbol or scale.");
  for (const character of decimal) { const point = character.codePointAt(0)!; if (point >= 0xd800 && point <= 0xdfff) invalid("A General decimal symbol must be well-formed Unicode."); }
  const original = decimalComponents(value), source = { ...original, exponent: original.exponent + powerAdjustment }, leading = value === 0 ? 0 : source.digits.length + source.exponent - 1;
  const significant = Math.max(1, source.digits.replace(/0+$/, "").length), sign = value < 0 ? "-" : "";
  const candidates: Candidate[] = [], seen = new Set<string>();
  const add = (integer: string, fraction: string, suffix: string, notation: Candidate["notation"], precision: number): void => {
    const trimmed = fraction.replace(/0+$/, ""), text = sign + integer + (trimmed ? decimal + trimmed : "") + suffix;
    if (integer.length + (trimmed ? 1 + trimmed.length : 0) + suffix.length > 11 || seen.has(text)) return;
    seen.add(text); candidates.push({ text, notation, precision: Math.min(significant, precision) });
  };
  // Only decimal places may be dropped in fixed notation. Narrowing a column
  // cannot turn an exact integer such as 12345 into an unrelated shorter value.
  if (leading < 11) {
    const integerPlaces = Math.max(1, leading + 1), maximum = Math.min(Math.max(0, -source.exponent), Math.max(0, 10 - integerPlaces));
    for (let places = maximum; places >= 0; places--) {
      const rounded = decimalMagnitude(value, places, powerAdjustment);
      if (value !== 0 && rounded.integer === "0" && !/[1-9]/.test(rounded.fraction)) continue;
      add(rounded.integer, rounded.fraction, "", "decimal", Math.max(1, leading + 1 + places));
    }
  }
  // A full-precision decimal outranks every scientific candidate (ties
  // prefer decimal). Measure it first, without constructing unused rounds.
  const exact = candidates.find(candidate => candidate.precision === significant);
  if (exact) {
    const width = measure(exact.text);
    if (!Number.isFinite(width) || width < 0) invalid("General number-format measurement must be finite and nonnegative.");
    if (width <= availableWidth) return { kind: "text", text: exact.text, notation: exact.notation };
  }
  if (value !== 0) for (let precision = 15; precision >= 1; precision--) {
    let exponent = leading, rounded = decimalMagnitude(value, precision - 1, powerAdjustment - exponent);
    if (rounded.integer.length > 1) { exponent++; rounded = decimalMagnitude(value, precision - 1, powerAdjustment - exponent); }
    const suffix = "E" + (exponent < 0 ? "-" : "+") + Math.abs(exponent).toString().padStart(2, "0");
    add(rounded.integer, rounded.fraction, suffix, "scientific", precision);
  }
  // For very small decimals scientific notation may retain more digits in the
  // same width. Equal precision prefers familiar fixed notation.
  candidates.sort((a, b) => b.precision - a.precision || (a.notation === b.notation ? 0 : a.notation === "decimal" ? -1 : 1));
  for (const candidate of candidates) {
    if (candidate === exact) continue; // Its one measurement already failed.
    const width = measure(candidate.text);
    if (!Number.isFinite(width) || width < 0) invalid("General number-format measurement must be finite and nonnegative.");
    if (width <= availableWidth) return { kind: "text", text: candidate.text, notation: candidate.notation };
  }
  return { kind: "overflow" };
}
