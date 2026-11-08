import type { DeepReadonly, DocumentOperation, DocumentSelection, TransactionPrepareOptions, TransactionReceipt, WorkbookRange } from "@doffi/core";
import type { WorkbookPreviewOptions, WorkbookRenderReport } from "@doffi/render";
import type { WorkbookInputState } from "./input-controller.js";

export interface WorkbookValidationWarning {
  readonly sheetId: string;
  readonly row: number;
  readonly column: number;
  readonly ruleId: string;
  readonly message: string;
}
export interface WorkbookValidationConfirmation {
  readonly documentId: string;
  readonly baseRevision: number;
  readonly warnings: readonly WorkbookValidationWarning[];
}
export interface WorkbookEditorOptions {
  readonly title?: string;
  readonly widthPt?: number;
  readonly heightPt?: number;
  readonly locale?: string;
  readonly limits?: WorkbookPreviewOptions["limits"];
  /** The host supplies authorized assets, explicit time/locale and any warning
   * acknowledgment. Stop validation always remains in core. */
  readonly prepareOptions?: () => TransactionPrepareOptions;
  /** Explicit synchronous host decision for warning/information violations.
   * Only true accepts these rules for this command. Stop/unavailable rules
   * never reach this callback. Declining retains the original draft. */
  readonly confirmValidationWarnings?: (confirmation: WorkbookValidationConfirmation) => boolean;
  readonly onError?: (error: Error) => void;
  readonly onSelectionChange?: (selection: DocumentSelection | undefined) => void;
  readonly onPendingChange?: (pending: boolean) => void;
}
export interface WorkbookEditorView {
  readonly element: HTMLDivElement;
  readonly ready: Promise<void>;
  settled(): Promise<WorkbookRenderReport>;
  selection(): DocumentSelection | undefined;
  inputState(): WorkbookInputState;
  hasPendingInput(): boolean;
  selectRange(sheetId: string, range: DeepReadonly<WorkbookRange>): Promise<DocumentSelection>;
  focusInput(selectAll?: boolean): void;
  applyInput(): TransactionReceipt | undefined;
  cancelInput(): void;
  perform(operations: readonly DocumentOperation[], options?: TransactionPrepareOptions): TransactionReceipt;
  recalculate(): Promise<WorkbookRenderReport>;
  dispose(): void;
}
