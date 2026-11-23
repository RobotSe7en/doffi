import type { WorkbookPrint, WorkbookPrintBand, WorkbookPrintRun } from "@doffi/core";
import { check, elements, fail, S, text, unsupported } from "../ooxml/xml.js";
import type { XmlNode } from "../ooxml/xml.js";
import { attributes, boolean, finite, integer, number, one, spreadsheetText, value } from "./values.js";

export type PrintNames = Pick<WorkbookPrint, "areas" | "repeatRows" | "repeatColumns">;
const fields: Readonly<Record<string, Extract<WorkbookPrintRun, { kind: "field" }>["field"]>> = { P: "page", N: "pages", A: "sheet", F: "file", Z: "path", D: "date", T: "time" };

/** Native compact fields become typed source runs, never executable HTML. */
export function printBand(source: string, node: XmlNode): WorkbookPrintBand {
  const band: WorkbookPrintBand = { left: [], center: [], right: [] };
  let section: keyof WorkbookPrintBand = "center", literal = "";
  const flush = (): void => { if (literal) { band[section].push({ kind: "text", text: literal }); literal = ""; } };
  const decoded = spreadsheetText(source, node);
  for (let index = 0; index < decoded.length; index++) {
    const character = decoded[index]!;
    if (character !== "&") { literal += character; continue; }
    const code = decoded[++index];
    if (code === "&") { literal += "&"; continue; }
    flush();
    if (code === "L" || code === "C" || code === "R") section = code === "L" ? "left" : code === "C" ? "center" : "right";
    else if (code && fields[code]) {
      if (code === "P" && /^[+-]\d/.test(decoded.slice(index + 1))) unsupported(node, "Page-number offsets are not implemented.");
      band[section].push({ kind: "field", field: fields[code]! });
    }
    else unsupported(node, "This print header/footer control is not implemented.");
  }
  flush(); return band;
}

// ISO/IEC 29500 pageSetup paperSize values; dimensions are unrotated points.
const inch = (width: number, height: number): { widthPt: number; heightPt: number } => ({ widthPt: width * 72, heightPt: height * 72 });
const mm = (width: number, height: number): { widthPt: number; heightPt: number } => inch(width / 25.4, height / 25.4);
const papers: Readonly<Record<number, { widthPt: number; heightPt: number }>> = {
  1: inch(8.5, 11), 2: inch(8.5, 11), 3: inch(11, 17), 4: inch(17, 11), 5: inch(8.5, 14), 6: inch(5.5, 8.5), 7: inch(7.25, 10.5),
  8: mm(297, 420), 9: mm(210, 297), 10: mm(210, 297), 11: mm(148, 210), 12: mm(250, 353), 13: mm(176, 250),
};
export function nativePaperSize(size: NonNullable<WorkbookPrint["paper"]>): number | undefined {
  const match = Object.entries(papers).find(([, value]) => Math.abs(value.widthPt - size.widthPt) < 1e-8 && Math.abs(value.heightPt - size.heightPt) < 1e-8);
  return match ? Number(match[0]) : undefined;
}

function paper(setup: XmlNode | undefined): WorkbookPrint["paper"] {
  if (!setup) return undefined;
  const width = value(setup, "paperWidth"), height = value(setup, "paperHeight");
  if (width !== undefined || height !== undefined) {
    const measure = (source: string | undefined): number => {
      const match = source && /^([+]?(?:\d+(?:\.\d*)?|\.\d+))(mm|cm|in|pt|pc|pi)$/.exec(source);
      if (!match) return fail(setup, "A custom paper dimension is invalid.");
      return finite(match[1]!, setup) * ({ mm: 72 / 25.4, cm: 72 / 2.54, in: 72, pt: 1, pc: 12, pi: 12 }[match[2]!]!);
    };
    return { widthPt: measure(width), heightPt: measure(height) };
  }
  const code = integer(setup, "paperSize", 0, 65535, 0);
  if (code === 0) return undefined; // Printer-dependent paper remains unspecified.
  if (!papers[code]) unsupported(setup, "This native paper size is not implemented.");
  return { ...papers[code]! };
}
function pageBreaks(root: XmlNode, name: "rowBreaks" | "colBreaks"): WorkbookPrint["rowBreaks"] {
  const parent = one(root, name), breaks = elements(parent, "brk", S), result: WorkbookPrint["rowBreaks"] = [];
  if (!parent) return result;
  check(parent, ["brk"], S); attributes(parent, ["count", "manualBreakCount"]);
  if (value(parent, "count") !== undefined && integer(parent, "count") !== breaks.length) fail(parent, "Page break count is inconsistent.");
  const row = name === "rowBreaks"; let manual = 0;
  for (const item of breaks) {
    check(item, [], S); attributes(item, ["id", "min", "max", "man", "pt"]);
    if (boolean(item, "pt")) unsupported(item, "Pivot table page breaks are not supported.");
    const before = integer(item, "id", 1, row ? 1_048_575 : 16383), from = integer(item, "min", 0, row ? 16383 : 1_048_575, 0), to = integer(item, "max", 0, row ? 16383 : 1_048_575, 0);
    if (from > to) fail(item, "A page break span is reversed.");
    if (boolean(item, "man")) { result.push({ before, from, to }); manual++; }
  }
  if (value(parent, "manualBreakCount") !== undefined && integer(parent, "manualBreakCount") !== manual) fail(parent, "Manual page break count is inconsistent.");
  result.sort((a, b) => a.before - b.before);
  return result;
}

export function readPrint(root: XmlNode, names: PrintNames): WorkbookPrint {
  const setup = one(root, "pageSetup"), margins = one(root, "pageMargins"), options = one(root, "printOptions"), hf = one(root, "headerFooter"), properties = one(one(root, "sheetPr"), "pageSetUpPr");
  if (setup) {
    check(setup, [], S);
    attributes(setup, ["paperSize", "paperWidth", "paperHeight", "orientation", "scale", "fitToWidth", "fitToHeight", "pageOrder", "firstPageNumber", "useFirstPageNumber", "blackAndWhite", "draft", "cellComments", "errors", "horizontalDpi", "verticalDpi", "copies", "usePrinterDefaults"]);
    if (boolean(setup, "draft") || value(setup, "cellComments") !== undefined && value(setup, "cellComments") !== "none"
      || value(setup, "errors") !== undefined && value(setup, "errors") !== "displayed") unsupported(setup, "This print rendering mode is not implemented.");
    for (const key of ["horizontalDpi", "verticalDpi", "copies"]) if (value(setup, key) !== undefined) integer(setup, key, 1);
    if (value(setup, "usePrinterDefaults") !== undefined) boolean(setup, "usePrinterDefaults");
  }
  if (properties) { attributes(properties, ["autoPageBreaks", "fitToPage"]); check(properties, [], S); if (value(properties, "autoPageBreaks") !== undefined) boolean(properties, "autoPageBreaks"); }
  const orientation = value(setup, "orientation") ?? "default", pageOrder = value(setup, "pageOrder") ?? "downThenOver";
  if (!["default", "portrait", "landscape"].includes(orientation) || !["downThenOver", "overThenDown"].includes(pageOrder)) fail(root, "A page layout enumeration is invalid.");
  let marginsPt: WorkbookPrint["marginsPt"];
  if (margins) {
    attributes(margins, ["left", "right", "top", "bottom", "header", "footer"]); check(margins, [], S);
    marginsPt = { left: number(margins, "left") * 72, right: number(margins, "right") * 72, top: number(margins, "top") * 72, bottom: number(margins, "bottom") * 72, header: number(margins, "header") * 72, footer: number(margins, "footer") * 72 };
  }
  if (options) { attributes(options, ["horizontalCentered", "verticalCentered", "headings", "gridLines", "gridLinesSet"]); check(options, [], S); if (value(options, "gridLinesSet") !== undefined) boolean(options, "gridLinesSet"); }
  if (hf) { attributes(hf, ["differentFirst", "differentOddEven", "scaleWithDoc", "alignWithMargins"]); check(hf, ["oddHeader", "oddFooter", "evenHeader", "evenFooter", "firstHeader", "firstFooter"], S); }
  const band = (key: string): WorkbookPrintBand => { const node = one(hf, key); if (!node) return { left: [], center: [], right: [] }; attributes(node, []); if (node.content.some(item => typeof item !== "string")) fail(node, "Header/footer content must be text."); return printBand(text(node), node); };
  const headerFooter: WorkbookPrint["headerFooter"] = {
    differentFirst: boolean(hf, "differentFirst"), differentOddEven: boolean(hf, "differentOddEven"),
    scaleWithDocument: boolean(hf, "scaleWithDoc", true), alignWithMargins: boolean(hf, "alignWithMargins", true), oddHeader: band("oddHeader"), oddFooter: band("oddFooter"),
  };
  for (const key of ["evenHeader", "evenFooter", "firstHeader", "firstFooter"] as const) if (one(hf, key)) headerFooter[key] = band(key);
  const dimension = paper(setup);
  return {
    ...names, ...(dimension ? { paper: dimension } : {}), ...(marginsPt ? { marginsPt } : {}),
    orientation: orientation as WorkbookPrint["orientation"], pageOrder: pageOrder as WorkbookPrint["pageOrder"],
    scale: boolean(properties, "fitToPage") ? { kind: "fit", width: setup ? integer(setup, "fitToWidth", 0, 32767, 1) : 1, height: setup ? integer(setup, "fitToHeight", 0, 32767, 1) : 1 } : { kind: "percent", percent: setup ? integer(setup, "scale", 10, 400, 100) : 100 },
    horizontalCentered: boolean(options, "horizontalCentered"), verticalCentered: boolean(options, "verticalCentered"), gridLines: boolean(options, "gridLines"), headings: boolean(options, "headings"), blackAndWhite: boolean(setup, "blackAndWhite"),
    ...(boolean(setup, "useFirstPageNumber") ? { firstPage: integer(setup!, "firstPageNumber", 1, 32767, 1) } : {}),
    rowBreaks: pageBreaks(root, "rowBreaks"), columnBreaks: pageBreaks(root, "colBreaks"), headerFooter,
  };
}
