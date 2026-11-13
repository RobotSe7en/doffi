export { DEFAULT_MAX_ZIP_BYTES, readDocumentZip, writeDocumentZip } from "./document-zip.js";
export type { DocumentZipOptions } from "./document-zip.js";
export { importDocx, importDocxWithReport } from "./docx-import.js";
export type { DocxImportOptions, DocxImportReport, DocxImportResult } from "./docx-import.js";
export { importPptx } from "./pptx-import.js";
export type { PptxImportOptions } from "./pptx-import.js";
export { importXlsx } from "./xlsx-import.js";
export type { XlsxImportOptions } from "./xlsx-import.js";
export { DEFAULT_XLSX_NATIVE_LIMITS } from "./xlsx/options.js";
export type { XlsxNativeLimits } from "./xlsx/options.js";
export type { OfficeImportControl, OfficeImportProgress } from "./import-control.js";
export { exportDocx } from "./docx-export.js";
export type { DocxExportOptions, DocxExportReport, DocxExportResult } from "./docx-export.js";
export type { UnexportedAnnotation } from "./export/annotations.js";

export { exportPptx } from "./pptx-export.js";
export type { PptxExportOptions, PptxExportReport, PptxExportResult, PptxAnnotationLoss, UnexportedPptxAnnotation } from "./pptx-export.js";
