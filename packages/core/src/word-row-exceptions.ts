import { expandSourceCssProperty } from "./document/source-css.js";

const prefix = "--word-row-exception-";
/** Row exceptions have table semantics, but never become table/cell direct formatting. */
export function wordRowExceptionKey(key: string): string {
  return prefix + (key.startsWith("--word-table-") ? key.slice(13) : key);
}
export function wordRowExceptionProperties(
  css: Readonly<Record<string, string>>,
): Record<string, string> {
  return Object.fromEntries(
    Object.entries(css).flatMap(([key, value]) => {
      if (!key.startsWith(prefix) || value === "initial") return [];
      const name = key.slice(prefix.length);
      return [
        [
          name.startsWith("cell-spacing") || name === "width-type" ? "--word-table-" + name : name,
          value,
        ],
      ];
    }),
  );
}
export function wordRowExceptionCss(css: Readonly<Record<string, string>>): Record<string, string> {
  return Object.fromEntries(
    Object.entries(css).flatMap(([key, value]) =>
      (key.startsWith("--") ? [[key, value]] : expandSourceCssProperty(key, value)).map(
        ([name, value]) => [wordRowExceptionKey(name!), value!],
      ),
    ),
  );
}
/** Word treats width, indentation and fixed layout on the first physical row
 * as table properties. Since Word 2013, alignment has the same scope; legacy
 * alignment remains row-local. Keep the native declarations on their rows. */
export function wordEffectiveTableProperties(
  table: Readonly<Record<string, string>>,
  firstRow: Readonly<Record<string, string>>,
  compatibilityMode = 12,
): Record<string, string> {
  const exception = wordRowExceptionProperties(firstRow), result = { ...table };
  if (exception.width !== undefined) {
    result.width = /^0(?:\.0+)?(?:pt|%)?$/.test(exception.width) ? "auto" : exception.width;
    if (exception["--word-table-width-type"] !== undefined)
      result["--word-table-width-type"] = exception["--word-table-width-type"]!;
  }
  if (exception["table-layout"] === "fixed") result["table-layout"] = "fixed";
  if (exception.indent !== undefined) {
    result["margin-left"] = exception.indent;
    result["--word-table-indent-origin"] = "native";
  }
  const alignment = firstRow["--word-row-alignment"] ?? exception.alignment;
  if (compatibilityMode >= 15 && alignment) {
    result["margin-left"] = alignment === "left" || alignment === "start"
      ? exception.indent ?? "0pt" : "auto";
    result["margin-right"] = alignment === "center" ? "auto" : "0pt";
  }
  return result;
}
/** Native cell properties (including style cells) override table exceptions. */
export function wordRowExceptionCellProperties(
  css: Readonly<Record<string, string>>,
  column: number,
  span: number,
  columns: number,
): Record<string, string> {
  const out = Object.fromEntries(
    Object.entries(css).filter(([key]) => key.startsWith("padding-") || key === "background-color"),
  );
  for (const side of ["top", "bottom", "left", "right"]) {
    const inner = (side === "left" && column > 0) || (side === "right" && column + span < columns);
    const border = inner ? css["--word-border-inside-v"] : undefined;
    if (border)
      Object.assign(out, Object.fromEntries(expandSourceCssProperty("border-" + side, border)));
    else
      for (const part of ["width", "style", "color"]) {
        const key = "border-" + side + "-" + part;
        if (css[key] !== undefined) out[key] = css[key]!;
      }
  }
  return out;
}
