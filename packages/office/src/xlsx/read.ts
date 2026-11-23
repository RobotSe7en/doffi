import { createSharedWorkbookFormulas, DocumentError } from "@doffi/core";
import type { WorkbookSheet } from "@doffi/core";
import { readDocumentZipAsync } from "../document-zip.js";
import { ImportControl } from "../import-control.js";
import { openOpcFiles } from "../ooxml/package.js";
import type { OpcPackage } from "../ooxml/package.js";
import { fail } from "../ooxml/xml.js";
import type { XmlNode } from "../ooxml/xml.js";
import { cellFormula, readCell } from "./cell.js";
import { WorkbookChunks } from "./chunks.js";
import { xlsxLimits } from "./options.js";
import type { XlsxReadOptions } from "./options.js";
import { scanWorksheet } from "./scan.js";
import { readSheet, rowDimension } from "./sheet.js";
import { readSharedStrings } from "./strings.js";
import { XlsxStyles } from "./styles.js";
import { readWorkbook } from "./workbook.js";
import type { NativeSheet, NativeWorkbook } from "./workbook.js";
import { one } from "./values.js";
import { readConditionalRules, readDataValidations } from "./rules.js";
import { auditXlsxPackage } from "./package.js";
import { readTables, readTableStyles } from "./tables.js";
import type { NativeTableStyles } from "./tables.js";
import { readNoteBoxes } from "./note-boxes.js";
import type { NativeNoteBoxSources } from "./note-boxes.js";
import { readClassicComments } from "./classic-comments.js";
import type { NativeClassicNotes } from "./classic-comments.js";
import { readThreadedComments } from "./threaded-comments.js";
import type { NativeThreads } from "./threaded-comments.js";
import { readAutoFilter } from "./filter.js";
import { NativeCalculationChain } from "./calc-chain.js";
import { readWorkbookProperties } from "./properties.js";

type RowDimension = WorkbookSheet["rowDimensions"][number];
/** Join identical adjacent explicit runs without expanding implicit rows. */
function appendRow(rows: RowDimension[], row: RowDimension): void {
  const previous = rows.at(-1), keys = ["heightPt", "hidden", "cssClass", "numberFormatId", "autoHeight"] as const;
  if (previous && previous.to + 1 === row.from && keys.every(key => previous[key] === row[key])) previous.to = row.to;
  else rows.push(row);
}
export interface NativeSheetRecords {
  readonly source: NativeSheet;
  readonly root: XmlNode;
  readonly metadata: WorkbookSheet;
}
export interface XlsxRecords {
  readonly opc: OpcPackage;
  readonly control: ImportControl;
  readonly limits: ReturnType<typeof xlsxLimits>;
  readonly native: NativeWorkbook;
  readonly properties: ReturnType<typeof readWorkbookProperties>;
  readonly styles: XlsxStyles;
  readonly tableStyles: NativeTableStyles;
  readonly chunks: WorkbookChunks;
  readonly sheets: readonly NativeSheetRecords[];
  readonly threads: NativeThreads;
  readonly notes: NativeClassicNotes;
  readonly noteBoxes: ReadonlyMap<string, NativeNoteBoxSources>;
  readonly counts: { readonly cells: number; readonly formulas: number };
}

/** Internal first stage. It deliberately returns remaining feature headers and
 * native style associations for complete R3-07 mapping, not a public document
 * package that could silently omit tables, drawings, rules or comments. */
export async function readXlsxRecords(bytes: Uint8Array, options: XlsxReadOptions = {}): Promise<XlsxRecords> {
  const limits = xlsxLimits(options), control = new ImportControl(options), budget = { nodes: 0 };
  const files = await readDocumentZipAsync(bytes, { limits: limits.xml, maxZipBytes: limits.native.maxZipBytes }, control);
  await control.checkpoint({ phase: "metadata", completed: 0, total: 1 });
  const opc = openOpcFiles(files, { limits: { ...limits.xml, maxFileBytes: limits.native.maxMetadataBytes } }, budget, limits.native.maxRetainedNodes);
  const main = opc.related("", "officeDocument");
  if (main && opc.contentType(main) !== "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml") throw new DocumentError("UnsupportedFormat", "The input is not an ordinary XLSX workbook.");
  auditXlsxPackage(opc);
  const properties = readWorkbookProperties(opc);
  const native = readWorkbook(opc), styles = new XlsxStyles(opc, native.part, limits.native.maxCssCodeUnits), chunks = new WorkbookChunks(limits.source);
  const calculationChain = new NativeCalculationChain(opc, native, limits.workbook.maxFormulas);
  const tableStyles = readTableStyles(styles);
  const stringPart = opc.related(native.part, "sharedStrings");
  if (stringPart && opc.contentType(stringPart) !== "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml") fail({ part: stringPart, name: "sst" }, "The shared string content type is invalid.");
  const strings = await readSharedStrings(stringPart ? files.get(stringPart) : undefined, stringPart ?? native.part, styles, {
    limits: limits.xml, budget, control, maxRetainedNodes: limits.native.maxRetainedNodes,
    maxStrings: limits.native.maxSharedStrings, maxCodeUnits: limits.native.maxSharedStringCodeUnits,
  });
  const sheets: NativeSheetRecords[] = []; let cells = 0, formulas = 0;
  for (const source of native.sheets) {
    const rows: RowDimension[] = [], shared = createSharedWorkbookFormulas(limits.workbook.maxFormulas, limits.workbook.maxFormulaKeyCodeUnits);
    const scan = { limits: limits.xml, budget, maxRetainedNodes: limits.native.maxRetainedNodes, control, sheetId: source.sheetId };
    const sheetBytes = files.get(source.part)!;
    await scanWorksheet(sheetBytes, source.part, { ...scan, pass: 0,
      cell(node, coordinate) {
        if (++cells > limits.workbook.maxCells || one(node, "f") && ++formulas > limits.workbook.maxFormulas) throw new DocumentError("ResourceLimitExceeded", "Native cells or formulas exceed their cumulative workbook budget.");
        cellFormula(node, coordinate, shared, 0);
        if (one(node, "f")) calculationChain.formula(source.nativeId, coordinate);
      },
      row(node, row, filterHidden) {
        const dimension = rowDimension(node, row, styles);
        if (dimension) {
          if (filterHidden) dimension.hidden = false;
          appendRow(rows, dimension);
        }
      },
    });
    chunks.startSheet(source.sheetId);
    const root = await scanWorksheet(sheetBytes, source.part, { ...scan, pass: 1,
      cell(node, coordinate) { chunks.append(coordinate, readCell(node, coordinate, styles, strings, shared).html); },
    });
    const metadata = readSheet(root, source, rows, opc, styles);
    metadata.conditionalRules = readConditionalRules(root, source.sheetId, styles);
    Object.assign(metadata, readDataValidations(root, source.sheetId));
    metadata.tables = readTables(root, opc, styles, tableStyles);
    const autoFilter = readAutoFilter(one(root, "autoFilter"));
    if (autoFilter) metadata.autoFilter = autoFilter;
    if (source.filterName && (!autoFilter || JSON.stringify(source.filterName) !== JSON.stringify(autoFilter.range))) fail(root, "The built-in filter name does not match its worksheet auto-filter.");
    // Calc often writes hidden="false" on every row. Keep overrides only when
    // they differ from the sheet default; formatting/height remains explicit.
    metadata.rowDimensions = [];
    for (const row of rows) {
      if (row.hidden === metadata.defaultRowsHidden) delete row.hidden;
      if (row.autoHeight === true && !metadata.defaultRowHeightCustom) delete row.autoHeight;
      if (Object.keys(row).length > 2) appendRow(metadata.rowDimensions, row);
    }
    sheets.push({ source, root, metadata });
    await control.checkpoint({ phase: "worksheet", completed: sheetBytes.length * 2, total: sheetBytes.length * 2, sheetId: source.sheetId });
  }
  chunks.finish(); strings.validateReferences(); calculationChain.finish();
  const notes = readClassicComments(opc, native.sheets, styles, limits.source.maxNodes, limits.source.maxFileBytes);
  const noteBoxes = readNoteBoxes(opc, sheets.map(sheet => ({ sheetId: sheet.source.sheetId, root: sheet.root })), notes, styles, limits.source.maxNodes);
  const threads = readThreadedComments(opc, native.part, native.sheets, limits.source.maxNodes);
  return { opc, control, limits, native, properties, styles, tableStyles, chunks, sheets, threads, notes, noteBoxes, counts: { cells, formulas } };
}
