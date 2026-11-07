import type { Annotation, Anchor } from "./contracts/annotations.js";
import type { AssetDescriptor, ContentDocumentManifest, ContentDocumentMetadata, WorkbookManifest } from "./contracts/document.js";
import type { WorkbookDocumentMetadata, WorkbookValue, WorkbookFormulaSource, WorkbookCoordinate } from "./contracts/workbook.js";
import type { DocumentEvent } from "./contracts/events.js";
import type { ContextStyle, DocumentInteractionRequest, EditProposal } from "./contracts/interaction.js";
import type { ContentOperationName, DocumentTransaction } from "./contracts/operations.js";
import type { PresentationInteractionRequest,PresentationEditProposal,PresentationOperationName } from "./contracts/presentation-interaction.js";
import type { WorkbookReferenceLocation, WorkbookReferenceResult } from "./workbook/reference.js";

export type DeepReadonly<Value> = Value extends readonly (infer Item)[]
  ? readonly DeepReadonly<Item>[]
  : Value extends object ? { readonly [Key in keyof Value]: DeepReadonly<Value[Key]> } : Value;

/** Virtual package files; manifest.json is contained here, never duplicated as authority. */
export type DocumentPackage = ReadonlyMap<string, Uint8Array>;

/** Bind each provider to the host's principal/session; package paths never authorize reads. */
export interface AssetProvider {
  readAsset(asset: DeepReadonly<AssetDescriptor>): Promise<Uint8Array | undefined>;
}

/** A derived, read-only view, not a second persistence format. */
export interface ContentDocumentSnapshot {
  readonly manifest: DeepReadonly<ContentDocumentManifest>;
  readonly html: string;
  readonly styles: readonly DeepReadonly<ContextStyle>[];
  readonly metadata: DeepReadonly<ContentDocumentMetadata>;
  readonly annotations: readonly DeepReadonly<Annotation>[];
}
export interface WorkbookSnapshot {
  readonly manifest: DeepReadonly<WorkbookManifest>;
  readonly styles: readonly DeepReadonly<ContextStyle>[];
  readonly metadata: DeepReadonly<WorkbookDocumentMetadata>;
  readonly annotations: readonly DeepReadonly<Annotation>[];
  readonly content?: { readonly path: string; readonly html: string };
}
export type DocumentSnapshot = ContentDocumentSnapshot | WorkbookSnapshot;
export interface WorkbookCell extends WorkbookCoordinate {
  readonly sheetId: string;
  readonly value: DeepReadonly<WorkbookValue>;
  readonly formula?: DeepReadonly<WorkbookFormulaSource>;
  readonly html: string;
  readonly text: string;
  readonly cssClass?: string;
  readonly formatId?: string;
}
export interface WorkbookWindow extends WorkbookCoordinate {
  readonly sheetId: string; readonly rows: number; readonly columns: number;
}
export interface WorkbookReadLimits {
  readonly maxCells: number; readonly maxFormulas: number; readonly maxWindowCells: number;
  readonly maxFormulaTemplates: number; readonly maxFormulaKeyCodeUnits: number;
}

export interface DocumentReadLimits {
  readonly maxFiles: number;
  readonly maxPackageBytes: number;
  readonly maxFileBytes: number;
  readonly maxNodes: number;
  readonly maxDepth: number;
  readonly maxCssNodes: number;
}

export interface OpenDocumentOptions {
  readonly limits?: Partial<DocumentReadLimits>;
  readonly workbookLimits?: Partial<WorkbookReadLimits>;
  /** Optional expected format; a mismatch fails without converting content. */
  readonly kind?: "document" | "presentation" | "workbook";
}

/** A derived node view. Text offsets use UTF-16; br is \n and img is U+FFFC. */
export interface DocumentNode {
  readonly nodeId: string;
  readonly tagName: string;
  readonly attributes: Readonly<Record<string, string>>;
  readonly parentNodeId: string | null;
  readonly childNodeIds: readonly string[];
  readonly text: string;
  readonly html: string;
}

/** Read-only document loading and serialization; editing belongs to DocumentSession. */
export interface ContentDocumentModel {
  readonly kind: "document" | "presentation";
  snapshot(): ContentDocumentSnapshot;
  getNode(nodeId: string): DocumentNode | undefined;
  serialize(): DocumentPackage;
}
export interface WorkbookModel {
  readonly kind: "workbook";
  snapshot(): WorkbookSnapshot;
  getNode(nodeId: string): DocumentNode | undefined;
  getCell(sheetId: string, row: number, column: number): WorkbookCell | undefined;
  readWindow(window: WorkbookWindow): readonly WorkbookCell[];
  /** Copy one manifest-declared asset without serializing cell chunks. The
   * optional byte budget is checked before allocating the returned copy. */
  getAsset(path: string, maxBytes?: number): Uint8Array | undefined;
  /** Resolve source ranges without evaluating formulas or expanding cells. */
  resolveReference(expression: string, origin: WorkbookReferenceLocation): WorkbookReferenceResult;
  serialize(): DocumentPackage;
}
export type DocumentModel = ContentDocumentModel | WorkbookModel;

export interface DocumentSelection {
  readonly documentId: string;
  readonly baseRevision: number;
  readonly anchor: DeepReadonly<Anchor>;
}

export interface InteractionRequestInput {
  readonly annotationIds: readonly string[];
  readonly supportedOperations?: readonly ContentOperationName[];
  readonly limits?: {
    readonly maxBytes: number;
    readonly maxNodes: number;
    readonly maxAssets: number;
  };
}
/** Explicit PPT dispatch preserves the existing V1 Word request contract. */
export interface PresentationInteractionRequestInput {
  readonly kind:"presentation";
  readonly annotationIds:readonly string[];
  readonly scope?:"slides"|"global";
  readonly supportedOperations?:readonly PresentationOperationName[];
  readonly assets?:readonly DeepReadonly<AssetDescriptor>[];
  readonly limits?:InteractionRequestInput["limits"];
}
export type PresentationAnnotationSubmissionHandler=(request:DeepReadonly<PresentationInteractionRequest>)=>void|PresentationEditProposal|Promise<void|PresentationEditProposal>;

// Type-only opacity is supplemented by private candidate identity inside the session.
declare const candidateBrand: unique symbol;
export interface CandidateChange<Model extends DocumentModel = ContentDocumentModel> {
  readonly [candidateBrand]: true;
  readonly documentId: string;
  readonly baseRevision: number;
  readonly preview: ReturnType<Model["snapshot"]>;
  /** Immutable candidate source, including staged resources; no editing session. */
  readonly model: Model;
  readonly affectedNodeIds: readonly string[];
  /** Stable slide identities whose rendered content or dependencies changed. PPT only. */
  readonly affectedSlideIds?: readonly string[];
  readonly affectedSheetIds?: readonly string[];
  readonly changedRanges?: readonly WorkbookRangeChange[];
  readonly changes: readonly DocumentNodeChange[];
  readonly changedAnnotationIds: readonly string[];
  readonly mapping: ChangeMapping;
}

export interface DocumentNodeChange {
  readonly nodeId: string;
  readonly kind: "added" | "removed" | "modified";
}

export interface DocumentSessionLimits {
  readonly maxOperations: number;
  readonly maxTransactionBytes: number;
  readonly maxCandidates: number;
  readonly maxCandidateBytes: number;
  readonly maxHistorySteps: number;
  readonly maxHistoryBytes: number;
}

export interface DocumentSessionOptions extends OpenDocumentOptions {
  readonly sessionLimits?: Partial<DocumentSessionLimits>;
  readonly calculationLimits?: Partial<import("./workbook/calculation.js").WorkbookCalculationLimits>;
}

export interface DocumentHistoryState {
  readonly undoSteps: number;
  readonly redoSteps: number;
  readonly retainedBytes: number;
}

export interface TextOffsetMapping {
  readonly nodeId: string;
  readonly from: number;
  readonly to: number;
  readonly insertedLength: number;
}

export interface ChangeMapping {
  readonly nodes: ReadonlyMap<string, string | null>;
  readonly textEdits: readonly TextOffsetMapping[];
  /** Ordered structural steps; each step maps the preceding source state. */
  readonly workbook?: readonly WorkbookStructureMapping[];
}

export interface WorkbookRangeChange { readonly sheetId: string; readonly range: DeepReadonly<import("./contracts/workbook.js").WorkbookRange> }
export type WorkbookStructureMapping =
  | { readonly kind: "axis"; readonly sheetId: string; readonly axis: "row" | "column"; readonly action: "insert" | "delete"; readonly index: number; readonly count: number }
  | { readonly kind: "move"; readonly source: WorkbookRangeChange; readonly target: WorkbookRangeChange }
  | { readonly kind: "sort"; readonly sheetId: string; readonly range: DeepReadonly<import("./contracts/workbook.js").WorkbookRange>; readonly rows: readonly number[] }
  | { readonly kind: "deleteSheet"; readonly sheetId: string }
  | { readonly kind: "restoreSheet"; readonly sheetId: string };

export interface TransactionReceipt {
  readonly documentId: string;
  readonly revision: number;
  readonly changed: boolean;
  readonly affectedNodeIds: readonly string[];
  readonly affectedSlideIds?: readonly string[];
  readonly affectedSheetIds?: readonly string[];
  readonly changedRanges?: readonly WorkbookRangeChange[];
  readonly mapping: ChangeMapping;
}

export type DocumentEventListener = (event: DeepReadonly<DocumentEvent>) => void;
export type AnnotationSubmissionHandler = (request: DeepReadonly<DocumentInteractionRequest>) => void | Promise<void>;

/** Host-authorized immutable resources staged with the same isolated candidate. */
export interface TransactionPrepareOptions {
  readonly assets?: readonly { readonly descriptor: AssetDescriptor; readonly bytes: Uint8Array }[];
  readonly calculationContext?: import("./workbook/scalar.js").WorkbookScalarContext;
  /** Warning/information rules require an explicit host acknowledgment. Stop
   * rules cannot be bypassed by this list. */
  readonly validation?: { readonly acceptedRuleIds: readonly string[] };
}
export interface TransactionPreparationProgress {
  readonly phase: "operations" | "cells" | "chunks" | "validation" | "calculation";
  readonly completed: number;
  /** Zero when this phase cannot know its final work count in advance. */
  readonly total: number;
}
/** Hosts provide a macrotask checkpoint for message-based cancellation. Core
 * owns neither the Worker nor its event loop. */
export interface TransactionPreparationExecution {
  readonly signal?: { readonly aborted: boolean };
  readonly yield?: () => Promise<void>;
  readonly onProgress?: (progress: TransactionPreparationProgress) => void;
}

/** Shared transaction and history boundary for editing views. */
export interface SessionLifecycle<Candidate extends CandidateChange<DocumentModel>> {
  getAnnotations(): readonly DeepReadonly<Annotation>[];
  prepare(transaction: DocumentTransaction, options?: TransactionPrepareOptions): Candidate;
  commit(candidate: Candidate): TransactionReceipt;
  discard(candidate: Candidate): void;
  undo(): TransactionReceipt;
  redo(): TransactionReceipt;
  history(): DocumentHistoryState;
  clearHistory(): void;
  subscribe(listener: DocumentEventListener): () => void;
  dispose(): void;
}
export interface DocumentEditingSession extends ContentDocumentModel, SessionLifecycle<CandidateChange> {}
export type WorkbookCandidateChange = CandidateChange<WorkbookModel>;
export interface WorkbookSession extends WorkbookModel, SessionLifecycle<WorkbookCandidateChange> {
  prepareAsync(transaction: DocumentTransaction, options?: TransactionPrepareOptions, execution?: TransactionPreparationExecution): Promise<WorkbookCandidateChange>;
  calculate(options?: import("./workbook/calculation.js").WorkbookCalculationOptions): Promise<import("./workbook/calculation.js").WorkbookCalculationResult>;
  calculateSync(options?: Omit<import("./workbook/calculation.js").WorkbookCalculationOptions, "yield">): import("./workbook/calculation.js").WorkbookCalculationResult;
  getCalculatedCell(sheetId: string, row: number, column: number): import("./workbook/calculation.js").WorkbookCalculatedCell;
}

/** Revision-bound collaboration on the same transaction boundary. */
export interface DocumentSession extends DocumentEditingSession {
  buildInteractionRequest(input:PresentationInteractionRequestInput):PresentationInteractionRequest;
  buildInteractionRequest(input: InteractionRequestInput): DocumentInteractionRequest;
  prepareProposal(request:PresentationInteractionRequest,proposal:PresentationEditProposal,options?:TransactionPrepareOptions):CandidateChange;
  prepareProposal(request: DocumentInteractionRequest, proposal: EditProposal): CandidateChange;
}
