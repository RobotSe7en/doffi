import { captureWorkbookModel, collectWorkbookFilteredRows, copyWorkbookScalarContext, createWorkbookCalculatedEvaluator, createWorkbookCalculator, DocumentError, openDocument } from "@doffi/core";
import type { WorkbookCalculationResult, WorkbookModel } from "@doffi/core";
import { writeOpcZipAsync } from "./document-zip.js";
import { OutputOpc } from "./export/opc.js";
import { XmlParts } from "./export/xml.js";
import { xlsxLimits } from "./xlsx/options.js";
import type { XlsxExportOptions, XlsxExportProgress, XlsxExportResult } from "./xlsx-export/options.js";
import { WorkbookCommentsWriter } from "./xlsx-export/comments.js";
import { WorkbookDrawingsWriter } from "./xlsx-export/drawings.js";
import { cellKey, writeSheet } from "./xlsx-export/sheet.js";
import { WorkbookStylesWriter } from "./xlsx-export/styles.js";
import { writeWorkbook, writeWorkbookProperties, writeWorkbookTheme } from "./xlsx-export/workbook.js";

export type { XlsxExportOptions, XlsxExportProgress, XlsxExportReport, XlsxExportResult, XlsxExportLoss } from "./xlsx-export/options.js";

const optionNames = new Set(["limits", "workbookLimits", "nativeLimits", "signal", "onProgress", "calculation", "calculationContext"]);
const scheduler = globalThis as unknown as { setTimeout(callback: () => void, delay: number): unknown };
const yieldWork = (): Promise<void> => new Promise(resolve => scheduler.setTimeout(resolve, 0));

/** Compile one immutable source revision into editable SpreadsheetML. Office
 * parts are derived from source and same-revision calculation, without mutating either. */
export async function exportXlsx(document: WorkbookModel, options: XlsxExportOptions = {}): Promise<XlsxExportResult> {
  if (!document || document.kind !== "workbook")
    throw new DocumentError("UnsupportedFormat", "XLSX export requires a workbook package.");
  if (!options || typeof options !== "object" || Array.isArray(options) || Object.keys(options).some(key => !optionNames.has(key)) ||
    (options.onProgress !== undefined && typeof options.onProgress !== "function") ||
    (options.signal !== undefined && (!options.signal || typeof options.signal !== "object" || typeof options.signal.aborted !== "boolean")) ||
    (options.calculation !== undefined && options.calculationContext !== undefined))
    throw new DocumentError("InvalidOperation", "Invalid XLSX export options.");
  const limits = xlsxLimits(options);
  const captured = captureWorkbookModel(document), { onProgress, signal, calculation: supplied } = options;
  const calculationContext = copyWorkbookScalarContext(options.calculationContext);
  const guard = (): void => {
    if (signal?.aborted) throw new DocumentError("OperationCancelled", "XLSX export was cancelled.");
    if (captureWorkbookModel(document) !== captured)
      throw new DocumentError("RevisionConflict", "The workbook changed during XLSX export.");
  };
  const checkpoint = async (progress?: XlsxExportProgress): Promise<void> => {
    guard();
    try { if (progress) onProgress?.(Object.freeze(progress)); }
    catch { throw new DocumentError("HostSubmissionFailed", "The host export progress callback failed."); }
    await yieldWork();
    guard();
  };
  guard();
  if (supplied) createWorkbookCalculatedEvaluator(captured, supplied).dispose();
  const model = openDocument(captured.serialize(), { kind: "workbook", limits: limits.source, workbookLimits: limits.workbook });
  const snapshot = model.snapshot(), { workbook } = snapshot.metadata;
  await checkpoint({ phase: "calculate", completed: 0, total: 1 });
  let calculation: WorkbookCalculationResult | undefined = supplied;
  if (!calculation && snapshot.manifest.chunks.length) {
    const calculator = createWorkbookCalculator(model, { maxFormulas: limits.workbook.maxFormulas });
    try {
      calculation = await calculator.calculate({ ...(signal ? { signal } : {}), context: calculationContext,
        yield: async () => { await yieldWork(); guard(); } });
    } finally { calculator.dispose(); }
  }
  await checkpoint({ phase: "calculate", completed: 1, total: 1 });
  const cells = new Map(calculation?.cells.map(cell => [cellKey(cell.sheetId, cell.row, cell.column), cell]));
  const opc = new OutputOpc(new XmlParts(limits.xml)), styles = new WorkbookStylesWriter(snapshot, limits.source);
  const caches = { written: 0, omitted: 0 }, total = workbook.sheets.length;
  const exportedAnnotationIds: string[] = [], unexportedAnnotations: XlsxExportResult["report"]["unexportedAnnotations"][number][] = [];
  const losses: XlsxExportResult["report"]["losses"][number][] = [];
  const comments = new WorkbookCommentsWriter(model, opc, styles, { exportedAnnotationIds, unexportedAnnotations, losses });
  const drawings = new WorkbookDrawingsWriter(model, opc, styles, limits.source, cells, losses, limits.workbook.maxCells);
  let firstTable = 1;
  for (const [index, sheet] of workbook.sheets.entries()) {
    await checkpoint({ phase: "worksheet", completed: index, total, sheetId: sheet.sheetId });
    const partName = "xl/worksheets/sheet" + (index + 1) + ".xml";
    const filter = await collectWorkbookFilteredRows(supplied ? captured : model, sheet.sheetId,
      { ...(calculation ? { calculation } : { context: calculationContext }), yield: checkpoint, ...(signal ? { signal } : {}) });
    if (filter.unresolved)
      losses.push({ kind: "metadata", sheetId: sheet.sheetId, message: "Some native filter visibility could not be computed. The filter criteria are preserved." });
    await writeSheet({ model, sheet, ...(filter.rows ? { filteredRows: filter.rows } : {}), active: sheet.sheetId === workbook.activeSheetId, partName,
      opc, styles, calculated: cells, caches, firstTable, checkpoint, tail: drawings.sheet(sheet, partName, index + 1) + comments.sheet(sheet.sheetId, partName, index + 1) });
    firstTable += sheet.tables?.length ?? 0;
  }
  await checkpoint({ phase: "worksheet", completed: total, total });
  comments.finish();
  writeWorkbook(opc, snapshot, caches.omitted > 0);
  styles.write(opc);
  writeWorkbookTheme(opc, styles);
  writeWorkbookProperties(opc, snapshot.metadata);
  await checkpoint({ phase: "package", completed: 0, total: 1 });
  opc.finish();
  for (const [path, bytes] of opc.parts.files) if ((path.endsWith(".xml") || path.endsWith(".rels") || path.endsWith(".vml")) &&
    !/^xl\/worksheets\/sheet\d+\.xml$/.test(path) && bytes.length > limits.native.maxMetadataBytes)
    throw new DocumentError("ResourceLimitExceeded", "Native XLSX metadata exceeds its part budget.");
  const bytes = await writeOpcZipAsync(opc.parts.files, { limits: limits.xml, maxZipBytes: limits.native.maxZipBytes }, checkpoint);
  await checkpoint({ phase: "package", completed: 1, total: 1 });
  return Object.freeze({ bytes, report: Object.freeze({ documentId: snapshot.manifest.documentId,
    revision: snapshot.manifest.revision, formulaCaches: Object.freeze(caches),
    exportedAnnotationIds: Object.freeze(exportedAnnotationIds), unexportedAnnotations: Object.freeze(unexportedAnnotations.map(loss => Object.freeze(loss))),
    losses: Object.freeze(losses.map(loss => Object.freeze(loss))),
  }) });
}
