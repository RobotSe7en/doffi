import { A, check, child, elements, fail, text, unsupported } from "../ooxml/xml.js";
import type { XmlNode } from "../ooxml/xml.js";
import type { Css } from "../ooxml/styles.js";
import { bool, Context, integer, points, value } from "./context.js";
import type { Html } from "./context.js";
import { fill, outline } from "./styles.js";
import { importText } from "./text.js";

const basicAccentStyle = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}";
const styledEdges = [["lnL", "left"], ["lnR", "right"], ["lnT", "top"], ["lnB", "bottom"]] as const;

export function importTable(context: Context, source: XmlNode, frame: Html, part: string): number[] {
  check(source, ["tblPr", "tblGrid", "tr"], A);
  const properties = child(source, "tblPr", A);
  let basicAccent = false;
  if (properties) {
    check(properties, ["noFill", "solidFill", "effectLst", "tableStyleId", "extLst"], A);
    const style = child(properties, "tableStyleId", A);
    if (style) {
      if (style.content.some(item => typeof item !== "string")) fail(style, "A table style ID must be text.");
      const id = text(style).trim().toUpperCase();
      if (id && !/^\{?00000000-0000-0000-0000-000000000000\}?$/.test(id)) {
        if (!context.compatiblePreview || id !== basicAccentStyle ||
            ["bandRow", "bandCol", "firstCol", "lastCol", "lastRow", "rtl"].some(name => bool(properties, name)))
          unsupported(style, "This presentation table style is not supported.");
        const main = context.related("", "officeDocument")!;
        const stylesPart = context.related(main, "tableStyles", false);
        if (stylesPart) {
          const styles = context.part(stylesPart, "presentationml.tableStyles", "tblStyleLst", A);
          check(styles, ["tblStyle"], A);
          if (elements(styles, "tblStyle", A).some(entry => value(entry, "styleId")?.toUpperCase() === id))
            unsupported(style, "A custom definition overrides this preview table style.");
        }
        basicAccent = true;
        context.retainPreview("table-style", frame.attrs["data-node-id"]!, properties, "Built-in table styles require compatible preview.");
      }
    }
  }
  const columns = elements(child(source, "tblGrid", A), "gridCol", A);
  const grid = columns.map(column => points(column, "w"));
  if (!grid.length || grid.some(width => width <= 0)) fail(source, "A presentation table grid is invalid.");
  const rows = elements(source, "tr", A);
  if (!rows.length) fail(source, "A presentation table requires rows.");
  context.count(rows.length * grid.length);
  const occupied = Array.from({ length: rows.length }, () => Array<number>(grid.length).fill(0));
  const origins = new Map<number, { row: number; column: number }>();
  const table = context.make("table");
  context.style(table, { "border-collapse": "collapse", "table-layout": "fixed", width: "100%", ...fill(properties, "--office-table-fill") });
  const tbody = context.make("tbody");
  let identity = 0;
  rows.forEach((sourceRow, rowIndex) => {
    check(sourceRow, ["tc", "extLst"], A);
    const height = points(sourceRow, "h");
    if (height <= 0) fail(sourceRow, "A table row height must be positive.");
    const row = context.make("tr"); context.style(row, { height: height + "pt" });
    const cells = elements(sourceRow, "tc", A);
    if (cells.length !== grid.length) fail(sourceRow, "Each DrawingML row must have one cell record per grid column.");
    cells.forEach((sourceCell, columnIndex) => {
      check(sourceCell, ["txBody", "tcPr", "extLst"], A);
      const continuation = bool(sourceCell, "hMerge") || bool(sourceCell, "vMerge");
      const rowSpan = integer(sourceCell, "rowSpan", 1), colSpan = integer(sourceCell, "gridSpan", 1);
      if (rowSpan < 1 || colSpan < 1 || rowIndex + rowSpan > rows.length || columnIndex + colSpan > grid.length) fail(sourceCell, "A cell merge exceeds its table.");
      if (continuation) {
        const origin = origins.get(occupied[rowIndex]![columnIndex]!);
        if (!origin || bool(sourceCell, "hMerge") !== (origin.column < columnIndex) || bool(sourceCell, "vMerge") !== (origin.row < rowIndex)) fail(sourceCell, "A merge continuation direction disagrees with its origin.");
        const body = child(sourceCell, "txBody", A);
        if (body && elements(body, "p", A).some(p => elements(p, undefined, A).some(run => run.name === "br" || ["r", "fld"].includes(run.name) && text(child(run, "t", A)) !== ""))) unsupported(sourceCell, "Merged continuation cells cannot discard their text.");
        return;
      }
      const cellIdentity = ++identity;
      origins.set(cellIdentity, { row: rowIndex, column: columnIndex });
      for (let y = rowIndex; y < rowIndex + rowSpan; y++) for (let x = columnIndex; x < columnIndex + colSpan; x++) {
        context.count();
        if (occupied[y]![x]) fail(sourceCell, "Presentation table merges overlap.");
        occupied[y]![x] = cellIdentity;
      }
      const cell = context.make("td", { ...(rowSpan > 1 ? { rowspan: String(rowSpan) } : {}), ...(colSpan > 1 ? { colspan: String(colSpan) } : {}) });
      const p = child(sourceCell, "tcPr", A);
      const css: Css = { ...fill(p, "--office-cell-fill") };
      if (basicAccent) {
        if (!css["--office-cell-fill-kind"]) unsupported(sourceCell, "Compatible table styles require explicit cell fills.");
        const header = rowIndex === 0 && bool(properties!, "firstRow");
        css.color = header ? "var(--office-theme-lt1,#ffffff)" : "var(--office-theme-dk1,#000000)";
        if (header) css["font-weight"] = "bold";
        for (const [, side] of styledEdges) css["border-" + side] = "1pt solid var(--office-theme-lt1,#ffffff)";
      }
      if (p) {
        check(p, ["lnL", "lnR", "lnT", "lnB", "lnTlToBr", "lnBlToTr", "cell3D", "noFill", "solidFill", "gradFill", "headers", "extLst"], A);
        if (child(p, "lnTlToBr", A) || child(p, "lnBlToTr", A) || child(p, "cell3D", A)) unsupported(p, "Diagonal cell borders and 3D cells are not supported.");
        for (const [tag, side] of styledEdges) {
          const line = child(p, tag, A), declarations = outline(line, "--office-cell-border-" + side);
          Object.assign(css, declarations);
          if (line) {
            const width = points(line, "w", 9525), color = declarations["--office-cell-border-" + side + "-color"] ?? "transparent";
            css["border-" + side] = width + "pt " + (declarations["--office-cell-border-"+side+"-kind"]==="none"?"none":declarations["--office-cell-border-" + side + "-dash"]?.includes("dot") ? "dotted" : declarations["--office-cell-border-" + side + "-dash"]?.includes("ash") ? "dashed" : "solid") + " " + color;
          }
        }
        for (const [name, side] of [["marL", "left"], ["marR", "right"], ["marT", "top"], ["marB", "bottom"]]) {
          const amount = points(p, name!, side === "left" || side === "right" ? 91440 : 45720);
          if (amount < 0) fail(p, "Cell margins cannot be negative.");
          css["padding-" + side] = amount + "pt";
        }
        const anchor = value(p, "anchor") ?? "t";
        if (!["t", "ctr", "b"].includes(anchor)) unsupported(p, "This cell alignment is not supported.");
        css["vertical-align"] = anchor === "t" ? "top" : anchor === "b" ? "bottom" : "middle";
        if (value(p, "vert") && value(p, "vert") !== "horz") unsupported(p, "Vertical table text is not supported.");
      }
      const textBody = child(sourceCell, "txBody", A);
      if (!textBody) fail(sourceCell, "A table cell text body is missing.");
      const body = context.make("div"); const textProperties=importText(context, textBody, body, part);
      if(textProperties?.wrap===false){context.bytes(120);body.attrs.style=(body.attrs.style?body.attrs.style+";":"")+"white-space:nowrap";}
      cell.children.push(body); context.style(cell, css); row.children.push(cell);
    });
    tbody.children.push(row);
  });
  if (occupied.some(row => row.some(v => !v))) fail(source, "A presentation table grid is not fully covered.");
  table.children.push(tbody); frame.children.push(table);
  return grid;
}
