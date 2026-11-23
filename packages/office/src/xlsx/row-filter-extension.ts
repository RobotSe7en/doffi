import { check, fail, S } from "../ooxml/xml.js";
import type { XmlNode } from "../ooxml/xml.js";
import { attributes, boolean, one, value } from "./values.js";

export const FILTER_HIDDEN_NAMESPACE = "urn:doffi:xlsx:row-visibility:1";

/** SpreadsheetML uses the same hidden bit for manual hiding and filtering. */
export function isFilterHiddenRow(row: XmlNode): boolean {
  if (!row.content.length) return false;
  check(row, ["extLst"], S);
  const list = one(row, "extLst");
  if (!list) return false;
  attributes(list, []);
  check(list, ["ext"], S);
  const extension = one(list, "ext");
  if (!extension || value(extension, "uri") !== FILTER_HIDDEN_NAMESPACE) fail(list, "Unknown worksheet row extension.");
  attributes(extension, ["uri"]);
  check(extension, ["filterHidden"], FILTER_HIDDEN_NAMESPACE);
  const marker = one(extension, "filterHidden", FILTER_HIDDEN_NAMESPACE);
  if (!marker || !boolean(row, "hidden")) fail(extension, "A filter-hidden row must have its native hidden flag set.");
  attributes(marker, []);
  check(marker, [], FILTER_HIDDEN_NAMESPACE);
  return true;
}
