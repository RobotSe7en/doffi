import type { DeepReadonly, DocumentSelection, TransactionPrepareOptions, TransactionReceipt, WorkbookCandidateChange, WorkbookEditProposal, WorkbookInteractionRequest, WorkbookInteractionRequestInput, WorkbookInteractionScope, WorkbookProposalReview } from "@doffi/core";
import type { WorkbookRenderReport } from "@doffi/render";
import type { WorkbookEditorOptions, WorkbookEditorView } from "./editor-contracts.js";
import type { WorkbookRemoteEditorView } from "./editor-remote.js";
import type { WorkbookSessionCandidate, WorkbookSessionReceipt } from "./session-transport.js";

export type WorkbookWorkbenchCandidate = WorkbookCandidateChange | WorkbookSessionCandidate;
export type WorkbookWorkbenchEvent =
  | { readonly type: "selectionChanged"; readonly selection: DocumentSelection | undefined }
  | { readonly type: "annotationsSubmitted"; readonly request: WorkbookInteractionRequest }
  | { readonly type: "proposalReady"; readonly proposal: WorkbookEditProposal; readonly candidate: WorkbookWorkbenchCandidate; readonly review: DeepReadonly<WorkbookProposalReview> };
export interface WorkbookWorkbenchOptions extends WorkbookEditorOptions {
  readonly scope?: DeepReadonly<WorkbookInteractionScope>;
  readonly allowGlobalScope?: boolean;
  readonly readRanges?: WorkbookInteractionRequestInput["readRanges"];
  readonly supportedOperations?: WorkbookInteractionRequestInput["supportedOperations"];
  readonly contextLimits?: WorkbookInteractionRequestInput["limits"];
  readonly onSubmitAnnotations?: (request: DeepReadonly<WorkbookInteractionRequest>, context: { readonly signal: AbortSignal }) => void | WorkbookEditProposal | Promise<void | WorkbookEditProposal>;
  readonly onEvent?: (event: WorkbookWorkbenchEvent) => void;
}
export interface WorkbookWorkbenchView {
  readonly element: HTMLDivElement;
  readonly editor: WorkbookEditorView | WorkbookRemoteEditorView;
  readonly ready: Promise<void>;
  settled(): Promise<WorkbookRenderReport>;
  request(): WorkbookInteractionRequest | undefined;
  buildRequest(annotationIds?: readonly string[], scope?: DeepReadonly<WorkbookInteractionScope>): Promise<WorkbookInteractionRequest>;
  submitAnnotations(annotationIds?: readonly string[]): Promise<void>;
  loadProposal(request: WorkbookInteractionRequest, proposal: WorkbookEditProposal, resources?: TransactionPrepareOptions): Promise<WorkbookWorkbenchCandidate>;
  applyProposal(): Promise<TransactionReceipt | WorkbookSessionReceipt>;
  rejectProposal(): Promise<void>;
  cancel(): void;
  dispose(): void;
}
