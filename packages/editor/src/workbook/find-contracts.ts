import type { DeepReadonly, WorkbookCalculationResult, WorkbookRangeChange } from "@doffi/core";

export interface WorkbookFindQuery {
  readonly text: string;
  /** Source searches formulas as expressions and constants as their source
   * text. Display searches formatted current values, excluding layout fills. */
  readonly field?: "source" | "formula" | "display";
  readonly matchCase?: boolean;
  readonly entireCell?: boolean;
  /** Omission includes all sheets, including hidden sheets. */
  readonly scope?: readonly WorkbookRangeChange[];
}
export interface WorkbookFindOptions {
  readonly calculation?: WorkbookCalculationResult;
  readonly locale?: string;
  readonly maxCells?: number;
  readonly maxMatches?: number;
  readonly maxCodeUnits?: number;
  readonly maxWindows?: number;
  readonly maxSteps?: number;
  readonly maxValues?: number;
  readonly windowCells?: number;
  readonly signal?: { readonly aborted: boolean };
  readonly yield?: () => Promise<void>;
  readonly onProgress?: (progress: WorkbookFindProgress) => void;
}
export interface WorkbookFindProgress { readonly cells: number; readonly windows: number; readonly matches: number }
export interface WorkbookFindMatch {
  readonly sheetId: string; readonly row: number; readonly column: number;
  /** Exact UTF-16 offsets in the selected source/display field. */
  readonly from: number; readonly to: number; readonly text: string;
}
export interface WorkbookFindDiagnostic {
  readonly sheetId: string; readonly row: number; readonly column: number;
  readonly reason: "calculation-required" | "unavailable" | "format-overflow";
}
export interface WorkbookFindResult {
  readonly documentId: string; readonly baseRevision: number;
  readonly matches: readonly WorkbookFindMatch[];
  readonly diagnostics: readonly WorkbookFindDiagnostic[];
  readonly complete: boolean;
  readonly cells: number; readonly windows: number;
}
export interface WorkbookReplaceOptions extends WorkbookFindOptions {
  /** Default preserves constant types and rich text; explicit auto allows the
   * same conversions as typed input. Formula expressions remain formulas. */
  readonly valueMode?: "preserve" | "auto" | "text";
  /** Replace one previously found occurrence, with revision and source proof.
   * The query is rerun at this coordinate before preparing any operation. */
  readonly only?: { readonly documentId: string; readonly baseRevision: number; readonly match: DeepReadonly<WorkbookFindMatch> };
}
