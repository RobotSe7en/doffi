import { attr, check, elements, fail, MC, S, unsupported } from "../ooxml/xml.js";
import type { XmlNode } from "../ooxml/xml.js";
import { COMMENT_REVISION } from "./comment-attributes.js";
import { attributes, integer, one, value } from "./values.js";

const spreadsheet = "http://schemas.microsoft.com/office/spreadsheetml/";
const x14 = spreadsheet + "2009/9/main";
const x15 = spreadsheet + "2010/11/main";
const absolutePath = spreadsheet + "2010/11/ac";
const revision6 = spreadsheet + "2016/revision6";
const revision10 = spreadsheet + "2016/revision10";
const calc = spreadsheet + "2018/calcfeatures";
const uuid = /^\{[\da-f]{8}(?:-[\da-f]{4}){3}-[\da-f]{12}\}$/i;

/** Validate producer declarations without using them to discard unknown content. */
export function excelRootAttributes(root: XmlNode, revisionUid = false): void {
  const ignorable = attr(root, "Ignorable", MC);
  if (ignorable !== undefined && ignorable.trim().split(/[\t\r\n ]+/).some(prefix => !prefix || !Object.hasOwn(root.namespaces, prefix))) {
    fail(root, "A SpreadsheetML ignorable prefix is undeclared.");
  }
  const uid = attr(root, "uid", COMMENT_REVISION);
  if (revisionUid && uid !== undefined && !uuid.test(uid)) fail(root, "A worksheet revision UID is invalid.");
  const ignored = new Set([MC + " Ignorable", ...(revisionUid ? [COMMENT_REVISION + " uid"] : [])]);
  attributes({ ...root, attrs: new Map([...root.attrs].filter(([key]) => !ignored.has(key))) }, []);
}

/** Excel's local save path and empty revision pointer are producer metadata. */
export function excelWorkbookContent(root: XmlNode): XmlNode {
  const alternate = one(root, "AlternateContent", MC), revision = one(root, "revisionPtr", COMMENT_REVISION);
  if (alternate) {
    attributes(alternate, []); check(alternate, ["Choice"], MC);
    const choice = one(alternate, "Choice", MC);
    if (!choice) fail(alternate, "An Excel save-path wrapper has no choice.");
    const prefix = value(choice, "Requires");
    if (!prefix || !Object.hasOwn(choice.namespaces, prefix) || choice.namespaces[prefix] !== x15) unsupported(choice);
    attributes(choice, ["Requires"]); check(choice, ["absPath"], absolutePath);
    const path = one(choice, "absPath", absolutePath);
    if (!path || !value(path, "url")) fail(choice, "An Excel save-path wrapper has no path.");
    attributes(path, ["url"]); check(path, [], absolutePath);
  }
  if (revision) {
    check(revision, [], COMMENT_REVISION);
    const qualified = new Set([revision6 + " coauthVersionLast", revision6 + " coauthVersionMax", revision10 + " uidLastSave"]);
    attributes({ ...revision, attrs: new Map([...revision.attrs].filter(([key]) => !qualified.has(key))) }, ["revIDLastSave", "documentId"]);
    if (integer(revision, "revIDLastSave") !== 0) unsupported(revision, "Workbook revision history is not supported.");
    for (const name of ["coauthVersionLast", "coauthVersionMax"]) {
      const version = attr(revision, name, revision6);
      if (version !== undefined && !/^\d{1,10}$/.test(version)) fail(revision, "A workbook coauthoring version is invalid.");
    }
    const lastSave = attr(revision, "uidLastSave", revision10);
    if (lastSave !== undefined && !uuid.test(lastSave)) fail(revision, "A workbook save UID is invalid.");
  }
  return alternate || revision ? { ...root, content: root.content.filter(node => node !== alternate && node !== revision) } : root;
}

export function readExcelCalcFeatures(extension: XmlNode): boolean {
  if (value(extension, "uri")?.toUpperCase() !== "{B58B0392-4F1F-4190-BB64-5DF3571DCE5F}") return false;
  check(extension, ["calcFeatures"], calc);
  const list = one(extension, "calcFeatures", calc);
  if (!list) fail(extension, "Excel calculation-version metadata is missing.");
  attributes(list, []); check(list, ["feature"], calc);
  const features = elements(list, "feature", calc);
  if (!features.length) fail(list, "Excel calculation-version metadata is empty.");
  for (const feature of features) {
    attributes(feature, ["name"]); check(feature, [], calc);
    if (!value(feature, "name")) fail(feature, "An Excel calculation-version feature has no name.");
  }
  return true;
}

export function readExcelStyleDefaults(list: XmlNode | undefined): void {
  if (!list) return;
  attributes(list, []); check(list, ["ext"], S);
  const definitions = new Map<string, readonly [string, string, string]>([
    ["{EB79DEF2-80B8-43E5-95BD-54CBDDF9020C}", [x14, "slicerStyles", "defaultSlicerStyle"]],
    ["{9260A510-F301-46A8-8635-F512D64BE5F5}", [x15, "timelineStyles", "defaultTimelineStyle"]],
  ]);
  for (const extension of elements(list, "ext", S)) {
    attributes(extension, ["uri"]);
    const uri = value(extension, "uri")?.toUpperCase() ?? "", definition = definitions.get(uri);
    if (!definition) unsupported(extension);
    definitions.delete(uri);
    const [ns, name, attribute] = definition;
    check(extension, [name], ns);
    const defaults = one(extension, name, ns);
    if (!defaults) fail(extension, "An Excel style-default extension is empty.");
    attributes(defaults, [attribute]); check(defaults, [], ns);
  }
}
