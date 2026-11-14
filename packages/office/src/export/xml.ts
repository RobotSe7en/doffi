import {wordSettingsOrder}from'../ooxml/word-settings-order.js';
import { DocumentError } from "@doffi/core";
import type { DocumentReadLimits } from "@doffi/core";

/** XML text and attributes are always data. Reject characters XML 1.0 cannot
 * encode rather than changing the user's UTF-16 source. */
export function escapeXml(value: string): string {
  if (/[\u0000-\u0008\u000b\u000c\u000e-\u001f\ud800-\udfff\ufffe\uffff]/u.test(value))
    throw new DocumentError("UnsupportedFeature", "The document contains a character that XML 1.0 cannot encode.");
  return value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/\r/g, "&#13;");
}
export function attributes(values: Readonly<Record<string, string | number | undefined>>): string {
  return Object.entries(values).filter((entry) => entry[1] !== undefined).map(([name, value]) => " " + name + '="' + escapeXml(String(value)).replace(/\t/g, "&#9;").replace(/\n/g, "&#10;") + '"').join("");
}
export const empty = (name: string, values: Readonly<Record<string, string | number | undefined>> = {}): string => "<" + name + attributes(values) + "/>";
export const value = (name: string, val: string | number | undefined): string => val === undefined ? "" : empty("w:" + name, { "w:val": val });
export const declaration = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
export const namespaces = ' xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';
const propertyOrder: Readonly<Record<string, readonly string[]>> = {
  settings:wordSettingsOrder.filter(name=>!name.includes(":")),
  rPr: ["rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike", "snapToGrid", "color", "spacing", "kern", "sz", "szCs", "highlight", "u", "shd", "vertAlign", "lang", "rPrChange"],
  pPr: ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "widowControl", "numPr", "pBdr", "shd", "tabs", "suppressAutoHyphens", "overflowPunct", "snapToGrid", "spacing", "ind", "jc", "outlineLvl", "rPr", "sectPr", "pPrChange"],
  tblPr: ["tblStyle", "tblStyleRowBandSize", "tblStyleColBandSize", "tblW", "jc", "tblCellSpacing", "tblInd", "tblBorders", "shd", "tblLayout", "tblCellMar", "tblLook", "tblPrChange"],
  tcPr: ["tcW", "gridSpan", "vMerge", "tcBorders", "shd", "noWrap", "tcMar", "textDirection", "vAlign", "tcPrChange"],
};
/** Put generated property children in the schema sequence. Values have already
 * been escaped; this tokenizer sees only our generated markup, never raw HTML. */
export function orderedProperties(kind: string, source: string): string {
  const order = propertyOrder[kind]; if (!order || !source) return source;
  const fragments: { xml: string; name: string }[] = []; let depth = 0, start = 0, name = "";
  for (const match of source.matchAll(/<\/?[^>]+>/g)) {
    const token = match[0];
    if (depth === 0) { start = match.index; name = /^<w:([^\s/>]+)/.exec(token)?.[1] ?? ""; }
    if (token.startsWith("</")) depth--; else if (!token.endsWith("/>")) depth++;
    if (depth === 0) fragments.push({ xml: source.slice(start, match.index + token.length), name });
  }
  if (depth !== 0 || fragments.some((fragment) => !order.includes(fragment.name))) throw new DocumentError("ExportFailed", "Generated native properties violate their supported sequence.");
  return fragments.sort((left, right) => order.indexOf(left.name) - order.indexOf(right.name)).map((fragment) => fragment.xml).join("");
}
export interface XmlPartBudget { bytes: number; nodes: number; files: number }

export class XmlParts {
  readonly files: Map<string, Uint8Array> = new Map<string, Uint8Array>();
  private readonly encoder = new (globalThis as unknown as { TextEncoder: new () => { encode(value: string): Uint8Array } }).TextEncoder();
  constructor(readonly limits: DocumentReadLimits, private readonly budget:XmlPartBudget = {bytes:0,nodes:0,files:0}) {}
  buffer(path: string): { add(value: string): void; finish(): void } {
    const chunks: string[] = []; let length = 0;
    return { add: (value) => {
      const bytes = this.encoder.encode(value).length;
      length += bytes; this.budget.bytes += bytes; this.budget.nodes += (value.match(/<(?![!?/])/g) ?? []).length;
      if (length > this.limits.maxFileBytes || this.budget.bytes > this.limits.maxPackageBytes || this.budget.nodes > this.limits.maxNodes)
        throw new DocumentError("ResourceLimitExceeded", "Native Office XML exceeds its output budget.");
      chunks.push(value);
    }, finish: () => {
      if (this.files.has(path) || this.budget.files >= this.limits.maxFiles) throw new DocumentError("ResourceLimitExceeded", "Native Office part count exceeds its budget.");
      this.budget.files++;
      this.files.set(path, this.encoder.encode(chunks.join("")));
    } };
  }
  put(path: string, data: string | Uint8Array): void {
    if (this.files.has(path)) throw new DocumentError("ExportFailed", "An OPC output part was repeated.");
    const bytes = typeof data === "string" ? this.encoder.encode(data) : new Uint8Array(data);
    if (typeof data === "string") this.budget.nodes += (data.match(/<(?![!?/])/g) ?? []).length;
    this.budget.bytes += bytes.length;
    if (bytes.length > this.limits.maxFileBytes || this.budget.bytes > this.limits.maxPackageBytes || this.budget.files + 1 > this.limits.maxFiles || this.budget.nodes > this.limits.maxNodes)
      throw new DocumentError("ResourceLimitExceeded", "Native Office output exceeds its package budget.");
    this.budget.files++;
    this.files.set(path, bytes);
  }
}
