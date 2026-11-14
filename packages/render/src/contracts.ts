import { DocumentError } from "@doffi/core";
import type { DocumentErrorData, PageSetup, Rectangle } from "@doffi/core";


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
  maxNodes: 50_000, maxPages: 256, maxLayoutPasses: 8, maxPrintNodes: 250_000,
});
export function renderLimits(input: Partial<RenderLimits> = {}): Readonly<RenderLimits> {
  const limits = { ...DEFAULT_RENDER_LIMITS, ...input };
  for (const value of Object.values(limits)) if (!Number.isSafeInteger(value) || value <= 0)
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
  readonly code: "UnsupportedField" | "SectionGeometryChange" | "ContentOverflow" | "NativeFormatting";
  readonly nodeId?: string;
  readonly message: string;
}
export interface LayoutFieldResult {
  readonly fieldNodeId: string;
  readonly resultNodeIds: readonly string[];
  readonly value: string;
  readonly kind: "PAGE" | "NUMPAGES" | "TOC" | 'REF' | 'PAGEREF' | 'NOTEREF' | 'SEQ';
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

export interface WordRenderReport extends RenderReport { readonly kind?: "document" }
export interface RenderedSlide {
  readonly slideId:string;
  readonly rootNodeId:string;
  readonly index:number;
  readonly hidden:boolean;
  readonly rect:Readonly<Rectangle>;
}
export interface PresentationRenderReport extends RenderReport {
  readonly kind:"presentation";
  readonly slides:readonly RenderedSlide[];
}
export type DocumentRenderReport=WordRenderReport|PresentationRenderReport;

export interface RenderOptions { readonly title?: string; readonly layout?: "pages" | "continuous"; readonly limits?: Partial<RenderLimits> }
export type BrowserRenderResult = { readonly ok: true; readonly report: DocumentRenderReport } | { readonly ok: false; readonly error: DocumentErrorData };
export interface BrowserRenderBridge {
  mount(files: readonly (readonly [string, string])[], options?: RenderOptions): Promise<BrowserRenderResult>;
  preparePrint(): Promise<BrowserRenderResult>;
  pageClip(index: number): { readonly x: number; readonly y: number; readonly width: number; readonly height: number };
  dispose(): void;
}
