import type { DocumentReadLimits, WorkbookCoordinate } from "@doffi/core";
import { check, fail, S } from "../ooxml/xml.js";
import type { XmlBudget, XmlNode } from "../ooxml/xml.js";
import { streamXml } from "../ooxml/xml-stream.js";
import type { ImportControl } from "../import-control.js";
import { address, attributes, integer, value } from "./values.js";
import { isFilterHiddenRow } from "./row-filter-extension.js";

export interface WorksheetScanOptions {
  readonly limits: DocumentReadLimits;
  readonly budget: XmlBudget;
  readonly maxRetainedNodes: number;
  readonly control: ImportControl;
  readonly sheetId: string;
  readonly pass: 0 | 1;
  readonly cell: (node: XmlNode, coordinate: WorkbookCoordinate) => void;
  readonly row?: (node: XmlNode, row: number, filterHidden: boolean) => void;
}

/** One cell record at a time, including very wide sparse rows. Missing native
 * coordinates are inferred in sequence; duplicate or reversed positions fail. */
export async function scanWorksheet(bytes: Uint8Array, part: string, options: WorksheetScanOptions): Promise<XmlNode> {
  let root: XmlNode | undefined, data: XmlNode | undefined, rowNode: XmlNode | undefined, row = -1, column = -1;
  const cells = new WeakMap<XmlNode, WorkbookCoordinate>();
  return streamXml(bytes, part, {
    limits: options.limits, budget: options.budget, maxRetainedNodes: options.maxRetainedNodes,
    maxTextCodeUnits: 32767 * 7, control: options.control,
    progress: (completed, total) => ({ phase: "worksheet", completed: completed + options.pass * total, total: total * 2, sheetId: options.sheetId }),
    ignoreWhitespace: node => node.uri === S && !["t", "v", "f", "formula", "formula1", "formula2", "oddHeader", "oddFooter", "evenHeader", "evenFooter", "firstHeader", "firstFooter"].includes(node.name),
    open(node, parent) {
      if (!parent) {
        if (node.uri !== S || node.name !== "worksheet") fail(node, "The worksheet root is invalid.");
        root = node;
      } else if (parent === root && node.uri === S && node.name === "sheetData") {
        if (data) fail(node, "A worksheet has duplicate cell tables.");
        attributes(node, []); data = node;
      } else if (parent === data && node.uri === S && node.name === "row") {
        const next = integer(node, "r", 1, 1_048_576, row + 2) - 1;
        if (next <= row) fail(node, "Worksheet rows must have unique increasing coordinates.");
        row = next; column = -1; rowNode = node;
      } else if (parent === rowNode && node.uri === S && node.name === "c") {
        const source = value(node, "r");
        if (source !== undefined && !/^[A-Za-z]{1,3}[1-9]\d*$/.test(source)) fail(node, "A native cell coordinate is invalid.");
        const coordinate = source === undefined ? { row, column: column + 1 } : address(source, node);
        if (coordinate.row !== row || coordinate.column <= column || coordinate.column >= 16384) fail(node, "Worksheet cells must have unique increasing coordinates in their row.");
        column = coordinate.column; cells.set(node, coordinate);
      } else if (node.uri === S && ["row", "c", "sheetData"].includes(node.name)) {
        fail(node, "A worksheet cell-table element is misplaced.");
      }
    },
    complete(node, parent) {
      const coordinate = cells.get(node);
      if (coordinate) { options.cell(node, coordinate); return false; }
      if (node === rowNode && parent === data) {
        const filterHidden = isFilterHiddenRow(node);
        options.row?.(node, row, filterHidden);
        rowNode = undefined;
        return false;
      }
      if (node === data) check(node, [], S);
      if (node === root && !data) fail(node, "A worksheet has no cell table.");
      return true;
    },
  });
}
