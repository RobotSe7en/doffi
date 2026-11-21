import type { WorkbookCalculationResult, WorkbookScalarContext } from "@doffi/core";
import type { XlsxReadOptions } from "../xlsx/options.js";

export interface XlsxExportProgress {
  readonly phase: "calculate" | "worksheet" | "package";
  readonly completed: number;
  readonly total: number;
  readonly sheetId?: string;
}
export interface XlsxExportOptions extends Pick<XlsxReadOptions, "limits" | "workbookLimits" | "nativeLimits" | "signal"> {
  /** An already completed first-party batch for the exact exported revision. */
  readonly calculation?: WorkbookCalculationResult;
  /** Host-provided time/locale for a fresh batch. */
  readonly calculationContext?: WorkbookScalarContext;
  readonly onProgress?: (progress: XlsxExportProgress) => void;
}
export interface XlsxExportLoss {
  readonly kind: "anchor" | "thread" | "status" | "format" | "metadata";
  readonly message: string;
  readonly annotationId?: string;
  readonly sheetId?: string;
  readonly nodeId?: string;
}
export interface XlsxExportReport {
  readonly documentId: string;
  readonly revision: number;
  readonly formulaCaches: { readonly written: number; readonly omitted: number };
  readonly exportedAnnotationIds: readonly string[];
  readonly unexportedAnnotations: readonly {
    readonly annotationId: string;
    readonly reason: "missing_coordinate" | "representative_collision" | "unsupported_anchor";
  }[];
  readonly losses: readonly XlsxExportLoss[];
}
export interface XlsxExportResult {
  readonly bytes: Uint8Array;
  readonly report: XlsxExportReport;
}
