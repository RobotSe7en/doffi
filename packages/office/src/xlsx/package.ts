import { createAssetDescriptor, DocumentError } from "@doffi/core";
import type { OpcPackage } from "../ooxml/package.js";
import { attr, check, elements, fail, PKG, R, S } from "../ooxml/xml.js";
import type { XlsxRecords } from "./read.js";
import type { WorkbookDrawings } from "./drawing.js";
import { attributes, one } from "./values.js";

const prefix = "application/vnd.openxmlformats-officedocument.";
const roles = new Map<string, string>([
  [prefix + "spreadsheetml.sheet.main+xml", "workbook"], [prefix + "spreadsheetml.worksheet+xml", "worksheet"],
  [prefix + "spreadsheetml.styles+xml", "styles"], [prefix + "spreadsheetml.sharedStrings+xml", "sharedStrings"],
  [prefix + "spreadsheetml.table+xml", "table"], [prefix + "spreadsheetml.comments+xml", "comments"],
  [prefix + "spreadsheetml.calcChain+xml", "calcChain"], [prefix + "drawing+xml", "drawing"],
  [prefix + "drawingml.chart+xml", "chart"], [prefix + "theme+xml", "theme"], [prefix + "vmlDrawing", "vmlDrawing"],
  [prefix + "extended-properties+xml", "extended-properties"], ["application/vnd.openxmlformats-package.core-properties+xml", "core-properties"],
  [prefix + "custom-properties+xml", "custom-properties"],
  ["application/vnd.ms-excel.person+xml", "person"], ["application/vnd.ms-excel.threadedcomments+xml", "threadedComment"],
  ["image/png", "image"], ["image/jpeg", "image"], ["image/webp", "image"],
]);
const allowed: Readonly<Record<string, readonly string[]>> = {
  root: [R + "/officeDocument", R + "/extended-properties", R + "/custom-properties", PKG + "/metadata/core-properties", PKG + "/metadata/thumbnail"],
  workbook: ["worksheet", "styles", "sharedStrings", "theme", "calcChain"].map(name => R + "/" + name).concat("http://schemas.microsoft.com/office/2017/10/relationships/person"),
  worksheet: ["drawing", "vmlDrawing", "comments", "table", "hyperlink"].map(name => R + "/" + name).concat("http://schemas.microsoft.com/office/2017/10/relationships/threadedComment"),
  drawing: [R + "/image", R + "/chart", R + "/hyperlink"],
  chart: [R + "/hyperlink"],
};
function relationshipSource(path: string): string | undefined {
  if (path === "_rels/.rels") return "";
  const match = /^(.*\/)?_rels\/([^/]+)\.rels$/.exec(path);
  return match ? (match[1] ?? "") + match[2]! : undefined;
}

/** Audit every container entry and relationship, including otherwise unused
 * metadata. Feature adapters still have to consume each allowed semantic part. */
export function auditXlsxPackage(opc: OpcPackage): void {
  const CT = "http://schemas.openxmlformats.org/package/2006/content-types", types = opc.read("[Content_Types].xml");
  attributes(types, []); check(types, ["Default", "Override"], CT);
  for (const node of elements(types, undefined, CT)) { attributes(node, node.name === "Default" ? ["Extension", "ContentType"] : ["PartName", "ContentType"]); check(node, [], CT); }
  const fileRoles = new Map<string, string>(), relationshipFiles = new Set<string>();
  for (const path of opc.files.keys()) {
    if (path === "[Content_Types].xml") continue;
    const source = relationshipSource(path);
    if (source !== undefined) {
      if (source && !opc.files.has(source)) fail({ part: path, name: "Relationships" }, "A relationship part has no owning source part.");
      if (opc.contentType(path) !== "application/vnd.openxmlformats-package.relationships+xml") fail({ part: path, name: "Relationships" }, "A relationship part has an invalid content type.");
      relationshipFiles.add(path); continue;
    }
    const role = roles.get(opc.contentType(path) ?? "");
    if (!role) throw new DocumentError("UnsupportedFeature", "The workbook contains a native part outside the declared reading scope.");
    if (role === "custom-properties") {
      // Calc may add an empty container on a second save. Validate its exact
      // emptiness; populated custom properties still require source mapping.
      const root = opc.read(path), uri = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties";
      if (root.name !== "Properties" || root.uri !== uri) fail(root, "A custom-property container has an invalid root.");
      attributes(root, []); check(root, [], uri);
    }
    fileRoles.set(path, role);
  }
  const graph = new Map<string, string[]>();
  for (const path of relationshipFiles) {
    const source = relationshipSource(path)!, root = opc.read(path);
    attributes(root, []); check(root, ["Relationship"], PKG);
    for (const node of elements(root, undefined, PKG)) { attributes(node, ["Id", "Type", "Target", "TargetMode"]); check(node, [], PKG); }
    const role = source ? fileRoles.get(source) : "root", permitted = allowed[role ?? ""] ?? [], edges: string[] = [];
    for (const relation of opc.relationships(source).values()) {
      if (!permitted.includes(relation.type)) throw new DocumentError("UnsupportedFeature", "A workbook relationship is not supported for its owning part.");
      if (relation.external) {
        if (relation.type !== R + "/hyperlink") throw new DocumentError("UnsupportedFeature", "External workbook resources and data are not supported.");
        continue;
      }
      if (relation.type === R + "/hyperlink") fail(root, "A hyperlink relationship must contain an external URI target.");
      const kind = relation.type.split("/").at(-1), expected = kind === "officeDocument" ? "workbook" : kind === "thumbnail" ? "image" : kind;
      if (fileRoles.get(relation.target) !== expected) fail(root, "A workbook relationship targets an incompatible native part.");
      edges.push(relation.target);
    }
    graph.set(source, edges);
  }
  const reached = new Set<string>(), pending = [""];
  while (pending.length) for (const part of graph.get(pending.pop()!) ?? []) if (!reached.has(part)) { reached.add(part); pending.push(part); }
  for (const path of fileRoles.keys()) if (!reached.has(path)) fail({ part: path, name: "part" }, "A native workbook part is disconnected from the package root.");
}

/** A relationship's presence is not evidence that its feature was mapped.
 * Invoke only after the record, comment and drawing adapters all succeeded. */
export function auditXlsxPartCoverage(records: XlsxRecords, drawings: WorkbookDrawings): void {
  const { opc, native } = records, consumed = new Set([native.part, ...native.sheets.map(sheet => sheet.part), ...drawings.parts, ...drawings.charts.parts]);
  const related = (owner: string, type: string): void => { const part = opc.related(owner, type); if (part) consumed.add(part); };
  for (const type of ["styles", "sharedStrings", "theme", "calcChain", "person"]) related(native.part, type);
  for (const type of ["metadata/core-properties", "extended-properties", "custom-properties"]) related("", type);
  for (const sheet of records.sheets) {
    for (const type of ["comments", "threadedComment"]) related(sheet.source.part, type);
    const references = [one(sheet.root, "legacyDrawing"), ...elements(one(sheet.root, "tableParts"), "tablePart", S)];
    for (const reference of references) if (reference) {
      const relation = opc.relationships(sheet.source.part).get(attr(reference, "id", R) ?? "");
      if (!relation || relation.external) fail(reference, "A mapped native part has no valid relationship."); consumed.add(relation.target);
    }
  }
  // A package thumbnail is a derived preview, not an independently editable
  // image. Validate its actual container before discarding the old preview.
  const thumbnail = opc.related("", "metadata/thumbnail");
  if (thumbnail) {
    const bytes = opc.files.get(thumbnail)!;
    if (bytes.byteLength > records.limits.source.maxFileBytes) throw new DocumentError("ResourceLimitExceeded", "The native thumbnail exceeds its resource budget.");
    createAssetDescriptor(thumbnail, opc.contentType(thumbnail)!, bytes); consumed.add(thumbnail);
  }
  for (const part of opc.files.keys()) if (part !== "[Content_Types].xml" && relationshipSource(part) === undefined && !consumed.has(part)) {
    fail({ part, name: "part" }, "A native workbook part has no consumed source feature.");
  }
}
