import type { WorkbookCoordinate, WorkbookRange } from "@doffi/core";
import { attributes, escapeXml } from "../export/xml.js";
import { R, S } from "../ooxml/xml.js";

export const sheetNamespaces: string = ' xmlns="' + S + '" xmlns:r="' + R + '"';
export const workbookPart = "xl/workbook.xml";
export const sheetContentType = (kind: string): string =>
  "application/vnd.openxmlformats-officedocument.spreadsheetml." + kind + "+xml";
export const bool = (value: boolean | undefined): number | undefined =>
  value === undefined ? undefined : Number(value);

/** ST_Xstring escapes must be shielded before escaping the surrounding XML. */
export const xstring = (value: string): string => value.replace(/_(?=x[\da-fA-F]{4}_)/g, "_x005F_");
export const element = (tag: string, content: string, attrs: Parameters<typeof attributes>[0] = {}): string =>
  "<" + tag + attributes(attrs) + ">" + content + "</" + tag + ">";
export const textElement = (tag: string, content: string, attrs?: Parameters<typeof attributes>[0]): string =>
  element(tag, escapeXml(content), attrs);
export const stringText = (content: string): string =>
  element("t", escapeXml(xstring(content)), { "xml:space": "preserve" });
export const formulaText = (expression: string): string => escapeXml(expression.slice(1));

export function columnName(column: number): string {
  let label = "";
  for (let index = column + 1; index > 0; index = Math.floor((index - 1) / 26))
    label = String.fromCharCode(65 + (index - 1) % 26) + label;
  return label;
}
export function address(cell: WorkbookCoordinate, absolute = false): string {
  const marker = absolute ? "$" : "";
  return marker + columnName(cell.column) + marker + (cell.row + 1);
}
export function rangeRef(range: WorkbookRange, absolute = false): string {
  const from = address(range.from, absolute), to = address(range.to, absolute);
  return from === to ? from : from + ":" + to;
}
export const sheetReference = (name: string): string => "'" + name.replaceAll("'", "''") + "'!";
