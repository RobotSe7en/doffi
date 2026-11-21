import type { DeepReadonly, WorkbookPrint, WorkbookPrintBand, WorkbookSheet } from "@doffi/core";
import { nativePaperSize } from "../xlsx/print.js";
import { empty } from "../export/xml.js";
import { bool, columnName, element, rangeRef, sheetReference, textElement, xstring } from "./xml.js";

const fields = { page: "P", pages: "N", sheet: "A", file: "F", path: "Z", date: "D", time: "T" };
function band(value: DeepReadonly<WorkbookPrintBand>): string {
  return (["left", "center", "right"] as const).map(section => !value[section].length ? "" :
    "&" + section[0]!.toUpperCase() + value[section].map(run => run.kind === "field" ?
      "&" + fields[run.field] : run.text.replaceAll("&", "&&")).join("")).join("");
}
function breaks(tag: string, items: DeepReadonly<WorkbookPrint["rowBreaks"]>): string {
  return items.length ? element(tag, items.map(item => empty("brk", { id: item.before, min: item.from, max: item.to, man: 1 })).join(""),
    { count: items.length, manualBreakCount: items.length }) : "";
}
export function printSettings(print: DeepReadonly<WorkbookPrint>): string {
  const { marginsPt, headerFooter: hf } = print;
  const options = empty("printOptions", { horizontalCentered: bool(print.horizontalCentered), verticalCentered: bool(print.verticalCentered),
    headings: bool(print.headings), gridLines: bool(print.gridLines) });
  const margins = marginsPt ? empty("pageMargins", Object.fromEntries(Object.entries(marginsPt).map(([name, value]) => [name, value / 72]))) : "";
  const paperSize = print.paper && nativePaperSize(print.paper);
  const setup = empty("pageSetup", {
    orientation: print.orientation, pageOrder: print.pageOrder, blackAndWhite: bool(print.blackAndWhite),
    paperSize,
    paperWidth: print.paper && paperSize === undefined ? print.paper.widthPt.toFixed(12).replace(/0+$/, "").replace(/\.$/, "") + "pt" : undefined,
    paperHeight: print.paper && paperSize === undefined ? print.paper.heightPt.toFixed(12).replace(/0+$/, "").replace(/\.$/, "") + "pt" : undefined,
    scale: print.scale.kind === "percent" ? print.scale.percent : undefined, fitToWidth: print.scale.kind === "fit" ? print.scale.width : undefined,
    fitToHeight: print.scale.kind === "fit" ? print.scale.height : undefined, firstPageNumber: print.firstPage, useFirstPageNumber: bool(print.firstPage !== undefined),
  });
  const bands = (["oddHeader", "oddFooter", "evenHeader", "evenFooter", "firstHeader", "firstFooter"] as const)
    .map(name => hf[name] ? textElement(name, xstring(band(hf[name]))) : "").join("");
  const headers = element("headerFooter", bands, { differentFirst: bool(hf.differentFirst), differentOddEven: bool(hf.differentOddEven),
    scaleWithDoc: bool(hf.scaleWithDocument), alignWithMargins: bool(hf.alignWithMargins) });
  return options + margins + setup + headers + breaks("rowBreaks", print.rowBreaks) + breaks("colBreaks", print.columnBreaks);
}
export function builtInNames(sheet: DeepReadonly<WorkbookSheet>, index: number): string {
  const prefix = sheetReference(sheet.name);
  const local = (name: string, values: readonly string[], hidden?: number): string =>
    values.length ? textElement("definedName", values.join(","), { name: "_xlnm." + name, localSheetId: index, hidden }) : "";
  const areas = sheet.print?.areas.map(range => prefix + rangeRef(range, true)) ?? [];
  const cols = sheet.print?.repeatColumns, rows = sheet.print?.repeatRows;
  const titles = [
    ...(rows ? [prefix + "$" + (rows.from + 1) + ":$" + (rows.to + 1)] : []),
    ...(cols ? [prefix + "$" + columnName(cols.from) + ":$" + columnName(cols.to)] : []),
  ];
  return local("Print_Area", areas) + local("Print_Titles", titles) +
    local("_FilterDatabase", sheet.autoFilter ? [prefix + rangeRef(sheet.autoFilter.range, true)] : [], 1);
}
