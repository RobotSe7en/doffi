import { parseWorkbookFormula } from "@doffi/core";
import type { WorkbookDefinedName, WorkbookFormulaExpression, WorkbookMetadata, WorkbookRange, WorkbookSheet } from "@doffi/core";
import { attr, check, elements, fail, R, S, text, unsupported } from "../ooxml/xml.js";
import type { XmlNode } from "../ooxml/xml.js";
import type { OpcPackage } from "../ooxml/package.js";
import { attributes, boolean, integer, one, spreadsheetText, value } from "./values.js";
import type { PrintNames } from "./print.js";
import { excelRootAttributes, excelWorkbookContent, readExcelCalcFeatures } from "./excel-metadata.js";

export interface NativeSheet {
  readonly part: string;
  readonly nativeId: number;
  readonly sheetId: string;
  readonly name: string;
  readonly visibility: WorkbookSheet["visibility"];
  readonly printNames: PrintNames;
  filterName?: WorkbookRange;
}
export interface NativeWorkbook {
  readonly part: string;
  readonly dateSystem: WorkbookMetadata["dateSystem"];
  readonly calculationMode: WorkbookMetadata["calculationMode"];
  readonly activeSheetId: string;
  readonly sheets: readonly NativeSheet[];
  readonly names: WorkbookDefinedName[];
}
function extension(root: XmlNode): void {
  const list = one(root, "extLst"); if (!list) return;
  attributes(list, []); check(list, ["ext"], S);
  const extensions = elements(list, "ext", S);
  if (extensions.length > 1) unsupported(list);
  for (const item of extensions) {
    attributes(item, ["uri"]);
    if (readExcelCalcFeatures(item)) continue;
    if (value(item, "uri") !== "{7626C862-2A13-11E5-B345-FEFF819CDC9F}") unsupported(item);
    check(item, ["extCalcPr"], "http://schemas.libreoffice.org/");
    const settings = one(item, "extCalcPr", "http://schemas.libreoffice.org/");
    if (!settings) fail(item, "A Calc extension has no calculation settings.");
    attributes(settings, ["stringRefSyntax"]); check(settings, [], S);
    if (value(settings, "stringRefSyntax") !== "CalcA1ExcelA1") unsupported(settings, "This Calc reference dialect is not supported.");
  }
}
function protection(node: XmlNode | undefined): void {
  if (!node) return;
  attributes(node, ["lockStructure", "lockWindows", "lockRevision"]); check(node, [], S);
  if (["lockStructure", "lockWindows", "lockRevision"].some(key => boolean(node, key))) unsupported(node, "Workbook protection is not supported.");
}
function boundRange(ast: WorkbookFormulaExpression, sheet: NativeSheet, node: XmlNode): WorkbookRange {
  const scope = (name: string | undefined): void => { if (!name || name.toUpperCase() !== sheet.name.toUpperCase()) fail(node, "A built-in worksheet name targets another or unspecified worksheet."); };
  if (ast.kind === "cell") {
    scope(ast.sheet); if (!ast.absoluteRow || !ast.absoluteColumn) fail(node, "A built-in worksheet reference must be absolute.");
    return { from: { row: ast.row, column: ast.column }, to: { row: ast.row, column: ast.column } };
  }
  if (ast.kind === "range") {
    const from = boundRange(ast.from, sheet, node).from, to = boundRange(ast.to, sheet, node).to;
    if (from.row > to.row || from.column > to.column) fail(node, "A built-in worksheet range is reversed.");
    return { from, to };
  }
  if (ast.kind === "axisRange") {
    scope(ast.sheet);
    if (!ast.from.absolute || !ast.to.absolute || ast.from.index > ast.to.index) fail(node, "A built-in worksheet axis range must be ordered and absolute.");
    return ast.axis === "row" ? { from: { row: ast.from.index, column: 0 }, to: { row: ast.to.index, column: 16383 } }
      : { from: { row: 0, column: ast.from.index }, to: { row: 1_048_575, column: ast.to.index } };
  }
  return unsupported(node, "This built-in worksheet name is not a fixed range.");
}
function terms(ast: WorkbookFormulaExpression): WorkbookFormulaExpression[] {
  return ast.kind === "binary" && ast.operator === "," ? [...terms(ast.left), ...terms(ast.right)] : [ast];
}

/** Workbook-level XML only; worksheets and shared strings remain streaming parts. */
export function readWorkbook(opc: OpcPackage): NativeWorkbook {
  const part = opc.related("", "officeDocument");
  if (!part) fail({ part: "_rels/.rels", name: "Relationships" }, "The package has no workbook relationship.");
  if (opc.contentType(part) !== "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml") unsupported({ part, name: "workbook", uri: S, attrs: new Map(), namespaces: {}, content: [] }, "The native input is not an ordinary XLSX workbook.");
  const source = opc.read(part);
  if (source.uri !== S || source.name !== "workbook") fail(source, "The workbook root is invalid.");
  excelRootAttributes(source);
  const root = excelWorkbookContent(source);
  check(root, ["fileVersion", "workbookPr", "workbookProtection", "bookViews", "sheets", "definedNames", "calcPr", "extLst"], S);
  extension(root); protection(one(root, "workbookProtection"));
  const props = one(root, "workbookPr"), calculation = one(root, "calcPr"), views = one(root, "bookViews"), sheetsRoot = one(root, "sheets");
  if (props) {
    check(props, [], S); attributes(props, ["date1904", "backupFile", "showObjects", "filterPrivacy", "saveExternalLinkValues", "updateLinks", "defaultThemeVersion", "dateCompatibility", "checkCompatibility", "hidePivotFieldList", "showPivotChartFilter", "allowRefreshQuery", "refreshAllConnections"]);
    for (const key of ["backupFile", "filterPrivacy", "saveExternalLinkValues", "dateCompatibility", "checkCompatibility", "hidePivotFieldList", "showPivotChartFilter", "allowRefreshQuery", "refreshAllConnections"]) if (value(props, key) !== undefined) boolean(props, key);
    if (value(props, "showObjects") !== undefined && value(props, "showObjects") !== "all") unsupported(props, "Hidden or placeholder workbook objects are not supported.");
    if (value(props, "defaultThemeVersion") !== undefined) integer(props, "defaultThemeVersion");
    if (value(props, "updateLinks") !== undefined && !["userSet", "never", "always"].includes(value(props, "updateLinks")!)) fail(props, "Invalid link update setting.");
  }
  let calculationMode: WorkbookMetadata["calculationMode"] = "auto";
  if (calculation) {
    check(calculation, [], S); attributes(calculation, ["calcId", "calcMode", "fullCalcOnLoad", "refMode", "iterate", "iterateCount", "iterateDelta", "fullPrecision", "calcCompleted", "calcOnSave", "concurrentCalc", "concurrentManualCount", "forceFullCalc"]);
    if (boolean(calculation, "iterate") || !boolean(calculation, "fullPrecision", true)) unsupported(calculation, "Iterative or displayed-precision calculation is not supported.");
    const mode = value(calculation, "calcMode") ?? "auto";
    if (!["auto", "autoNoTable", "manual"].includes(mode)) fail(calculation, "Invalid workbook calculation mode.");
    calculationMode = mode === "manual" ? "manual" : "auto";
    if (value(calculation, "refMode") !== undefined && !["A1", "R1C1"].includes(value(calculation, "refMode")!)) fail(calculation, "Invalid formula reference display mode.");
    for (const key of ["fullCalcOnLoad", "calcCompleted", "calcOnSave", "concurrentCalc", "forceFullCalc"]) if (value(calculation, key) !== undefined) boolean(calculation, key);
    for (const key of ["calcId", "iterateCount", "concurrentManualCount"]) if (value(calculation, key) !== undefined) integer(calculation, key);
  }
  if (!sheetsRoot) fail(root, "A workbook requires worksheets.");
  check(sheetsRoot, ["sheet"], S); attributes(sheetsRoot, []);
  const sheets: NativeSheet[] = [], nativeIds = new Set<number>(), paths = new Set<string>(), sheetNames = new Set<string>();
  for (const sheet of elements(sheetsRoot, "sheet", S)) {
    check(sheet, [], S);
    // Relationship id is the only permitted non-namespace qualified attribute.
    for (const [key] of sheet.attrs) if (!key.startsWith("http://www.w3.org/2000/xmlns/ ") && ![" name", " sheetId", " state", R + " id"].includes(key)) unsupported(sheet);
    const nativeId = integer(sheet, "sheetId", 1), name = spreadsheetText(value(sheet, "name") ?? "", sheet, 31), visibility = value(sheet, "state") ?? "visible", relation = opc.relationships(part).get(attr(sheet, "id", R) ?? "");
    if (!name || nativeIds.has(nativeId) || sheetNames.has(name.toUpperCase()) || !["visible", "hidden", "veryHidden"].includes(visibility)) fail(sheet, "Worksheet identity, name or visibility is invalid.");
    if (!relation || relation.external || relation.type !== R + "/worksheet" || paths.has(relation.target)) fail(sheet, "A worksheet relationship is invalid or repeated.");
    if (opc.contentType(relation.target) !== "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml") fail(sheet, "A worksheet part has an invalid content type.");
    nativeIds.add(nativeId); paths.add(relation.target); sheetNames.add(name.toUpperCase());
    sheets.push({ part: relation.target, nativeId, sheetId: `sheet-${nativeId}`, name, visibility: visibility as NativeSheet["visibility"], printNames: { areas: [] } });
  }
  if (!sheets.length) fail(root, "A workbook requires worksheets.");
  if (views) { attributes(views, []); check(views, ["workbookView"], S); }
  const view = elements(views, "workbookView", S)[0], activeIndex = view ? integer(view, "activeTab", 0, sheets.length - 1, 0) : 0;
  if (sheets[activeIndex]!.visibility !== "visible") fail(root, "The active worksheet must be visible.");
  const names: WorkbookDefinedName[] = [], definitions = one(root, "definedNames"), seen = new Map<string, string>();
  if (definitions) { attributes(definitions, []); check(definitions, ["definedName"], S); }
  for (const node of elements(definitions, "definedName", S)) {
    attributes(node, ["name", "localSheetId", "hidden", "comment", "function", "vbProcedure", "xlm"]);
    if (node.content.some(item => typeof item !== "string")) fail(node, "Defined name content must be text.");
    if (boolean(node, "function") || boolean(node, "vbProcedure") || boolean(node, "xlm")) unsupported(node, "Macro names are not supported.");
    // Native defined-name offsets use A1, not the selected cell or active tab;
    // see the independent pinned-Calc name-origin probes in scripts/.
    const name = spreadsheetText(value(node, "name") ?? "", node, 255), expression = "=" + text(node), origin = { row: 0, column: 0 }, ast = parseWorkbookFormula(expression, origin);
    const index = value(node, "localSheetId") === undefined ? undefined : integer(node, "localSheetId", 0, sheets.length - 1), scope = index === undefined ? undefined : sheets[index]!;
    const key = JSON.stringify([index ?? null, name.toUpperCase()]);
    if (name.startsWith("_xlnm.")) {
      if (!scope) fail(node, "A built-in worksheet name requires a local scope.");
      let canonical: string;
      if (name === "_xlnm.Print_Area") {
        const areas = terms(ast).map(term => boundRange(term, scope, node)); canonical = JSON.stringify(areas);
        if (!seen.has(key)) scope.printNames.areas.push(...areas);
      } else if (name === "_xlnm.Print_Titles") {
        const result: Omit<PrintNames, "areas"> = {};
        for (const term of terms(ast)) {
          boundRange(term, scope, node);
          if (term.kind !== "axisRange") unsupported(node, "Print titles must reference whole rows or columns.");
          const property = term.axis === "row" ? "repeatRows" : "repeatColumns";
          if (result[property]) fail(node, "Print title axes are repeated.");
          result[property] = { from: term.from.index, to: term.to.index };
        }
        canonical = JSON.stringify([result.repeatRows ?? null, result.repeatColumns ?? null]);
        if (!seen.has(key)) Object.assign(scope.printNames, result);
      } else if (name === "_xlnm._FilterDatabase") {
        const range = boundRange(ast, scope, node); canonical = JSON.stringify(range);
        if (!seen.has(key)) scope.filterName = range;
      } else unsupported(node, "This built-in defined name is not supported.");
      if (seen.has(key) && (name !== "_xlnm.Print_Titles" || seen.get(key) !== canonical)) fail(node, "A built-in worksheet name is repeated or conflicting.");
      seen.set(key, canonical);
    } else {
      if (seen.has(key)) fail(node, "A defined name is repeated in its scope.");
      seen.set(key, expression);
      names.push({ nameId: `name-${names.length + 1}`, name, expression, origin, ...(scope ? { sheetId: scope.sheetId } : {}),
        hidden: boolean(node, "hidden"), ...(value(node, "comment") === undefined ? {} : { description: spreadsheetText(value(node, "comment")!, node) }) });
    }
  }
  return { part, dateSystem: boolean(props, "date1904") ? "1904" : "1900", calculationMode, activeSheetId: sheets[activeIndex]!.sheetId, sheets, names };
}
