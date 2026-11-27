import { DocumentError } from "@doffi/core";
import type { DeepReadonly, DocumentErrorCode, DocumentErrorData, PageSetup, Rectangle, ValidationIssue, WorkbookRange, WorkbookScalarContext } from "@doffi/core";

/** Font names stay in the source. This report describes only the current preview. */
export interface FontResolution {
  readonly requestedFamily: string;
  readonly selectedFamily: string;
  readonly substituted: boolean;
  readonly missingCharacters: readonly string[];
  readonly glyphAuditComplete: boolean;
}

export interface RenderLimits {
  readonly maxNodes: number;
  readonly maxPages: number;
  readonly maxLayoutPasses: number;
  readonly maxPrintNodes: number;
}
export const DEFAULT_RENDER_LIMITS: Readonly<RenderLimits> = Object.freeze({
  maxNodes: 50_000,
  maxPages: 256,
  maxLayoutPasses: 8,
  maxPrintNodes: 250_000,
});
export function renderLimits(input: Partial<RenderLimits> = {}): Readonly<RenderLimits> {
  const limits = { ...DEFAULT_RENDER_LIMITS, ...input };
  for (const value of Object.values(limits))
    if (!Number.isSafeInteger(value) || value <= 0)
      throw new DocumentError("InvalidOperation", "Render limits must be positive safe integers.");
  return Object.freeze(limits);
}
export interface RenderedPage {
  readonly pageId: string;
  readonly index: number;
  readonly sectionId: string;
  readonly number: number;
  readonly label: string;
  readonly rect: Readonly<Rectangle>;
  readonly setup: Readonly<PageSetup>;
  readonly blank: boolean;
}
export interface LayoutDiagnostic {
  readonly code:
    | "UnsupportedField"
    | "UnresolvedReference"
    | "SectionGeometryChange"
    | "ContentOverflow"
    | "NativeFormatting";
  readonly nodeId?: string;
  readonly message: string;
}
export interface LayoutFieldResult {
  readonly fieldNodeId: string;
  readonly resultNodeIds: readonly string[];
  readonly value: string;
  readonly kind:
    | "PAGE"
    | "NUMPAGES"
    | "TOC"
    | "TOF"
    | "INDEX"
    | "TOA"
    | "CITATION"
    | "BIBLIOGRAPHY"
    | "REF"
    | "PAGEREF"
    | "NOTEREF"
    | "STYLEREF"
    | "SEQ"
    | "IF"
    | "COMPARE"
    | "SET"
    | "QUOTE"
    | "SYMBOL"
    | "DATE"
    | "TIME"
    | "CREATEDATE"
    | "SAVEDATE"
    | "PRINTDATE"
    | "EDITTIME"
    | "AUTHOR"
    | "TITLE"
    | "SUBJECT"
    | "KEYWORDS"
    | "COMMENTS"
    | "LASTSAVEDBY"
    | "REVNUM"
    | "DOCPROPERTY"
    | "FILENAME"
    | "TEMPLATE"
    | "NUMWORDS"
    | "NUMCHARS"
    | "SECTION"
    | "SECTIONPAGES"
    | "FORMULA";
}
/** Derived geometry is bound to one immutable document revision. All values use points. */
interface RenderReport {
  readonly documentId: string;
  readonly revision: number;
  readonly layout: "pages" | "continuous";
  readonly pages: readonly RenderedPage[];
  readonly nodePages: Readonly<Record<string, readonly string[]>>;
  readonly fields: readonly LayoutFieldResult[];
  readonly diagnostics: readonly LayoutDiagnostic[];
  readonly fonts: readonly FontResolution[];
}

export interface WordRenderReport extends RenderReport {
  readonly kind?: "document";
}
export interface RenderedSlide {
  readonly slideId: string;
  readonly rootNodeId: string;
  readonly index: number;
  readonly hidden: boolean;
  readonly rect: Readonly<Rectangle>;
}
export interface PresentationRenderReport extends RenderReport {
  readonly kind: "presentation";
  readonly slides: readonly RenderedSlide[];
}
export type DocumentRenderReport = WordRenderReport | PresentationRenderReport;

export interface WorkbookPrintSelection {
  readonly sheetId: string;
  /** Overrides this sheet's saved print areas when supplied. */
  readonly areas?: readonly DeepReadonly<WorkbookRange>[];
}
export interface WorkbookPrintSettings {
  /** Defaults to all visible worksheets in workbook order. */
  readonly sheets?: readonly WorkbookPrintSelection[];
  readonly formulaValues?: "current" | "cached";
  readonly context?: WorkbookScalarContext;
  readonly fileName?: string;
  readonly filePath?: string;
  /** Cumulative paper area at 96 DPI before generating pages. */
  readonly maxPaperPixels?: number;
}
export interface WorkbookPrintedPage extends RenderedPage {
  readonly sheetId: string;
  readonly areaIndex: number;
  readonly range: DeepReadonly<WorkbookRange>;
  readonly scale: number;
}
export interface WorkbookPrintReport extends RenderReport {
  readonly kind: "workbook";
  readonly layout: "pages";
  readonly pages: readonly WorkbookPrintedPage[];
  readonly formulaValues: "current" | "cached";
  readonly formulaCells: Readonly<{ valid: number; imported: number; stale: number; unavailable: number }>;
}
export type PortableRenderReport = DocumentRenderReport | WorkbookPrintReport;

/** Byte conversion contracts shared by the host and its managed converter. */
export type ConversionErrorCode = DocumentErrorCode | "Cancelled" | "Timeout" | "Busy" | "InvalidRequest";
export interface ConversionRequestV1 {
  readonly protocolVersion: 1;
  readonly taskId: string;
  readonly format: "pdf" | "png";
  readonly input: {
    readonly format: "document-package-zip" | "docx" | "pptx" | "xlsx";
    readonly bytesBase64: string;
    readonly documentId?: string;
  };
  readonly options?: {
    readonly pageIndex?: number;
    readonly scale?: number;
    readonly timeoutMs?: number;
    readonly limits?: Partial<RenderLimits>;
    readonly workbook?: WorkbookPrintSettings;
  };
}
export type ConversionResponseV1 = {
  readonly protocolVersion: 1; readonly taskId: string; readonly ok: true;
  readonly output: { readonly mimeType: "application/pdf" | "image/png"; readonly bytesBase64: string;
    readonly report: PortableRenderReport; readonly engine: { readonly name: "chromium"; readonly version: string } };
} | {
  readonly protocolVersion: 1; readonly taskId: string | null; readonly ok: false;
  readonly error: { readonly code: ConversionErrorCode; readonly message: string; readonly details?: readonly ValidationIssue[] };
};

export interface BridgeWebFont {
  readonly family: string;
  readonly aliases: readonly string[];
  readonly url: string;
  readonly weight?: string;
  readonly style?: string;
  readonly equivalentAliases?: boolean;
  readonly fallbackOnly?: boolean;
}

export interface RenderOptions {
  readonly title?: string;
  readonly layout?: "pages" | "continuous";
  readonly limits?: Partial<RenderLimits>;
  readonly webFonts?: readonly BridgeWebFont[];
  readonly workbook?: WorkbookPrintSettings;
}
export type BrowserRenderResult =
  | { readonly ok: true; readonly report: PortableRenderReport }
  | { readonly ok: false; readonly error: DocumentErrorData };
export interface BrowserRenderBridge {
  mount(
    files: readonly (readonly [string, string])[],
    options?: RenderOptions,
  ): Promise<BrowserRenderResult>;
  preparePrint(): Promise<BrowserRenderResult>;
  pageClip(index: number): {
    readonly x: number;
    readonly y: number;
    readonly width: number;
    readonly height: number;
  };
  dispose(): void;
}
