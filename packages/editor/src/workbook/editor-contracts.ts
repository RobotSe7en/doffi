import type { DeepReadonly, DocumentOperation, DocumentSelection, TransactionPrepareOptions, TransactionReceipt, WorkbookRange } from "@doffi/core";
import type { WorkbookPreviewOptions, WorkbookRenderReport } from "@doffi/render";
import type { WorkbookInputState } from "./input-controller.js";

export interface WorkbookEditorOptions {
  readonly title?: string;
  readonly widthPt?: number;
  readonly heightPt?: number;
  readonly locale?: string;
  readonly limits?: WorkbookPreviewOptions["limits"];
  /** The host supplies authorized assets, explicit time/locale and any warning
   * acknowledgment. Stop validation always remains in core. */
  readonly prepareOptions?: () => TransactionPrepareOptions;
  readonly onError?: (error: Error) => void;
  readonly onSelectionChange?: (selection: DocumentSelection | undefined) => void;
}
export interface WorkbookEditorView {
  readonly element: HTMLDivElement;
  readonly ready: Promise<void>;
  settled(): Promise<WorkbookRenderReport>;
  selection(): DocumentSelection | undefined;
  inputState(): WorkbookInputState;
  selectRange(sheetId: string, range: DeepReadonly<WorkbookRange>): Promise<DocumentSelection>;
  focusInput(selectAll?: boolean): void;
  applyInput(): TransactionReceipt | undefined;
  cancelInput(): void;
  perform(operations: readonly DocumentOperation[], options?: TransactionPrepareOptions): TransactionReceipt;
  recalculate(): Promise<WorkbookRenderReport>;
  dispose(): void;
}
