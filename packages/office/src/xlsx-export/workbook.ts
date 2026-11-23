import { parseWorkbookFormula, printWorkbookFormula, resolveSourceCssColor } from "@doffi/core";
import type { DeepReadonly, WorkbookDefinedName, WorkbookFormulaExpression, WorkbookSnapshot } from "@doffi/core";
import type { OutputOpc } from "../export/opc.js";
import { fontName } from "../export/values.js";
import { declaration, empty } from "../export/xml.js";
import { A } from "../ooxml/xml.js";
import { themeSlots } from "../xlsx/colors.js";
import { builtInNames } from "./print.js";
import type { WorkbookStylesWriter } from "./styles.js";
import { bool, element, formulaText, sheetContentType, sheetNamespaces, textElement, workbookPart, xstring } from "./xml.js";

/** Native names use an A1 origin with wrapping for relative references. */
function nativeNameFormula(name: DeepReadonly<WorkbookDefinedName>): string {
  if (!name.origin.row && !name.origin.column) return name.expression;
  const normalize = (value: number, absolute: boolean, limit: number): number =>
    absolute ? value : (value % limit + limit) % limit;
  const cell = (node: Extract<WorkbookFormulaExpression, { kind: "cell" }>): typeof node => ({
    ...node, row: normalize(node.row, node.absoluteRow, 1_048_576), column: normalize(node.column, node.absoluteColumn, 16_384),
  });
  const visit = (node: WorkbookFormulaExpression): WorkbookFormulaExpression => {
    switch (node.kind) {
      case "cell": return cell(node);
      case "range": return { ...node, from: cell(node.from), to: cell(node.to) };
      case "axisRange": {
        const limit = node.axis === "row" ? 1_048_576 : 16_384;
        const shift = (at: typeof node.from): typeof node.from => ({ ...at, index: normalize(at.index, at.absolute, limit) });
        return { ...node, from: shift(node.from), to: shift(node.to) };
      }
      case "binary": return { ...node, left: visit(node.left), right: visit(node.right) };
      case "unary": return { ...node, operand: visit(node.operand) };
      case "call": return { ...node, arguments: node.arguments.map(item => item && visit(item)) };
      default: return node;
    }
  };
  return printWorkbookFormula(visit(parseWorkbookFormula(name.expression, name.origin)), { row: 0, column: 0 });
}

export function writeWorkbook(opc: OutputOpc, snapshot: WorkbookSnapshot, needsRecalculation: boolean): void {
  const { workbook } = snapshot.metadata;
  const sheets = new Map(workbook.sheets.map((sheet, index) => [sheet.sheetId, index]));
  const part = opc.part(workbookPart, sheetContentType("sheet.main"));
  part.add(declaration + "<workbook" + sheetNamespaces + ">");
  part.add(empty("workbookPr", { date1904: bool(workbook.dateSystem === "1904") }));
  part.add(element("bookViews", empty("workbookView", { activeTab: sheets.get(workbook.activeSheetId) })));
  part.add("<sheets>");
  for (const [index, sheet] of workbook.sheets.entries()) part.add(empty("sheet", {
    name: xstring(sheet.name), state: sheet.visibility, sheetId: index + 1,
    "r:id": opc.relate(workbookPart, "worksheet", "worksheets/sheet" + (index + 1) + ".xml"),
  }));
  part.add("</sheets>");
  const names = (workbook.names ?? []).map(name => element("definedName",
    formulaText(nativeNameFormula(name)), {
      name: xstring(name.name), localSheetId: name.sheetId === undefined ? undefined : sheets.get(name.sheetId),
      hidden: bool(name.hidden), comment: name.description === undefined ? undefined : xstring(name.description),
    })).join("") + workbook.sheets.map(builtInNames).join("");
  if (names) part.add(element("definedNames", names));
  part.add(empty("calcPr", { calcId: 0, calcMode: workbook.calculationMode, fullCalcOnLoad: bool(needsRecalculation), calcCompleted: bool(!needsRecalculation) }));
  part.add("</workbook>");
  part.finish();
  opc.relate("", "officeDocument", workbookPart);
}

export function writeWorkbookTheme(opc: OutputOpc, styles: WorkbookStylesWriter): void {
  const { theme, themeRaw } = styles;
  if (!theme["--office-theme-lt1"]) return;
  const slots = ["dk1", "lt1", "dk2", "lt2", ...themeSlots.slice(4)];
  const palette = slots.map(slot => element("a:" + slot, empty("a:srgbClr", {
    val: resolveSourceCssColor(theme["--office-theme-" + slot] ?? "#000000"),
  }))).join("");
  const fonts = ["major", "minor"].map(kind => {
    const face = (role: string): string => {
      const source = themeRaw["--office-" + kind + "-" + role];
      if (role !== "ascii" && (!source || /^var\(/.test(source))) return "";
      return fontName(theme["--office-" + kind + "-" + role] ?? '"Arial"');
    };
    const scripts = Object.entries(theme).flatMap(([name, value]) => {
      const script = new RegExp("^--office-" + kind + "-script-([A-Za-z]{4})$").exec(name)?.[1];
      return script ? [empty("a:font", { script, typeface: fontName(value) })] : [];
    });
    return element("a:" + kind + "Font", empty("a:latin", { typeface: face("ascii") }) +
      empty("a:ea", { typeface: face("east-asia") }) + empty("a:cs", { typeface: face("complex-script") }) + scripts.join(""));
  }).join("");
  const solid = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>';
  const format = element("a:fmtScheme", element("a:fillStyleLst", solid.repeat(3)) +
    element("a:lnStyleLst", element("a:ln", solid, { w: 12700 }).repeat(3)) +
    element("a:effectStyleLst", "<a:effectStyle><a:effectLst/></a:effectStyle>".repeat(3)) +
    element("a:bgFillStyleLst", solid.repeat(3)), { name: "Office" });
  const themeXml = element("a:theme", element("a:themeElements", element("a:clrScheme", palette, { name: "Office" }) +
    element("a:fontScheme", fonts, { name: "Office" }) + format), { "xmlns:a": A, name: "Office" });
  const part = opc.part("xl/theme/theme1.xml", "application/vnd.openxmlformats-officedocument.theme+xml");
  part.add(declaration + themeXml);
  part.finish();
  opc.relate(workbookPart, "theme", "theme/theme1.xml");
}

export function writeWorkbookProperties(opc: OutputOpc, metadata: WorkbookSnapshot["metadata"]): void {
  const { properties } = metadata;
  const cp = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
  const coreFields = { "dc:title": metadata.title, "dc:language": metadata.locale, "dc:creator": properties?.creator, "dc:subject": properties?.subject,
    "dc:description": properties?.description, "dc:identifier": properties?.identifier, "cp:lastModifiedBy": properties?.lastModifiedBy,
    "cp:revision": properties?.revision, "cp:category": properties?.category, "cp:contentStatus": properties?.contentStatus,
    "cp:contentType": properties?.contentType, "cp:version": properties?.version, "cp:keywords": properties?.keywords,
    "cp:lastPrinted": properties?.lastPrinted, "dcterms:created": properties?.created, "dcterms:modified": properties?.modified };
  const core = Object.entries(coreFields).map(([key, value]) => value === undefined ? "" : textElement(key, value,
    key.startsWith("dcterms:") ? { "xsi:type": "dcterms:W3CDTF" } : undefined)).join("");
  if (core) {
    const part = opc.part("docProps/core.xml", "application/vnd.openxmlformats-package.core-properties+xml");
    part.add(declaration + element("cp:coreProperties", core, { "xmlns:cp": cp, "xmlns:dc": "http://purl.org/dc/elements/1.1/",
      "xmlns:dcterms": "http://purl.org/dc/terms/", "xmlns:xsi": "http://www.w3.org/2001/XMLSchema-instance" }));
    part.finish();
    opc.relate("", "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties", "docProps/core.xml");
  }
  const appFields = { Template: properties?.template, Manager: properties?.manager, Company: properties?.company,
    Application: properties?.application, AppVersion: properties?.applicationVersion, TotalTime: properties?.totalEditingMinutes };
  const app = Object.entries(appFields).map(([key, value]) => value === undefined ? "" : textElement(key, String(value))).join("");
  if (app) {
    const part = opc.part("docProps/app.xml", "application/vnd.openxmlformats-officedocument.extended-properties+xml");
    part.add(declaration + element("Properties", app, { xmlns: "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" }));
    part.finish();
    opc.relate("", "extended-properties", "docProps/app.xml");
  }
}
