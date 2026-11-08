import type { DeepReadonly, DocumentOperation, TransactionPrepareOptions, WorkbookRange, WorkbookSnapshot } from "@doffi/core";

export type WorkbookControlsIdentity = Pick<WorkbookSnapshot["manifest"], "documentId" | "revision">;
/** Auxiliary forms read a bounded view and submit source commands to its owner. */
export interface WorkbookControlsHost {
  snapshot(): WorkbookSnapshot | undefined;
  selection(): { sheetId: string; range: DeepReadonly<WorkbookRange> } | undefined;
  idle(): boolean;
  prepareOptions?(): TransactionPrepareOptions;
  readonly locale?: string;
  perform(operations: readonly DocumentOperation[], options?: TransactionPrepareOptions, expected?: WorkbookControlsIdentity): Promise<unknown>;
  run(action: () => Promise<void>): void;
  draftChanged?(): void;
  selectSheet?(sheetId: string): Promise<unknown>;
  readonly signal: AbortSignal;
}
