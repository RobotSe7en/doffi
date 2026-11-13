import { z } from "zod";
import { FORMAT_VERSION, IdentifierSchema, PackagePathSchema, RevisionSchema } from "./primitives.js";
import { WordMetadataSchema } from "./word.js";
import { PresentationDocumentMetadataSchema } from "./presentation.js";
import type { PresentationDocumentMetadata } from "./presentation.js";
import { WorkbookChunkDescriptorSchema, WorkbookDocumentMetadataSchema } from "./workbook.js";
import type { WorkbookDocumentMetadata } from "./workbook.js";

export const AssetDescriptorSchema = z.strictObject({
  path: PackagePathSchema,
  mimeType: z.string().min(1),
  sha256: z.string().length(64).regex(/^[a-f0-9]{64}$/),
});
export type AssetDescriptor = z.infer<typeof AssetDescriptorSchema>;

const manifestFields = {
  format: z.literal("doffi"),
  formatVersion: z.literal(FORMAT_VERSION),
  documentId: IdentifierSchema,
  revision: RevisionSchema,
  styles: z.array(PackagePathSchema),
  metadata: PackagePathSchema,
  annotations: PackagePathSchema,
  assets: z.array(AssetDescriptorSchema),
};
export const ContentDocumentManifestSchema = z.strictObject({ ...manifestFields, kind: z.enum(["document", "presentation"]), entrypoint: PackagePathSchema });
export type ContentDocumentManifest = z.infer<typeof ContentDocumentManifestSchema>;
export const WorkbookManifestSchema = z.strictObject({ ...manifestFields, kind: z.literal("workbook"), chunks: z.array(WorkbookChunkDescriptorSchema), content: PackagePathSchema.optional() });
export type WorkbookManifest = z.infer<typeof WorkbookManifestSchema>;
export const DocumentManifestSchema = z.union([ContentDocumentManifestSchema, WorkbookManifestSchema]);
export type DocumentManifest = z.infer<typeof DocumentManifestSchema>;

export const PageSetupSchema = z.strictObject({
  unit: z.literal("pt"),
  width: z.number().positive(),
  height: z.number().positive(),
  margins: z.strictObject({
    top: z.number().min(0),
    right: z.number().min(0),
    bottom: z.number().min(0),
    left: z.number().min(0),
  }),
});
export type PageSetup = z.infer<typeof PageSetupSchema>;

export const DocumentSectionSchema = z.strictObject({
  sectionId: IdentifierSchema,
  rootNodeId: IdentifierSchema,
  pageSetup: PageSetupSchema.optional(),
  breakType: z.enum(["continuous", "nextPage", "evenPage", "oddPage"]).optional(),
  differentFirstPage: z.boolean().optional(),
  columns: z.strictObject({ count: z.number().int().positive(), gap: z.number().nonnegative(), separator: z.boolean().optional() }).optional(),
  grid: z.strictObject({ type: z.enum(["default", "lines", "linesAndChars", "snapToChars"]), linePitch: z.number().positive().optional(), characterSpacing: z.number().int().optional() }).optional(),
  gutter: z.number().nonnegative().optional(),
  gutterPosition: z.enum(["left", "right", "top"]).optional(),
  pageNumberStart: z.number().int().min(0).optional(),
  pageNumberFormat: z.enum(["decimal", "upperRoman", "lowerRoman", "upperLetter", "lowerLetter"]).optional(),
  headerDistance: z.number().min(0).optional(),
  footerDistance: z.number().min(0).optional(),
  headers: z.strictObject({ default: IdentifierSchema.optional(), first: IdentifierSchema.optional(), even: IdentifierSchema.optional() }).optional(),
  footers: z.strictObject({ default: IdentifierSchema.optional(), first: IdentifierSchema.optional(), even: IdentifierSchema.optional() }).optional(),
});
export type DocumentSection = z.infer<typeof DocumentSectionSchema>;

export const WordDocumentMetadataSchema = z.strictObject({
  title: z.string().optional(),
  locale: IdentifierSchema.optional(),
  pageSetup: PageSetupSchema,
  sections: z.array(DocumentSectionSchema).min(1),
  word: WordMetadataSchema.optional(),
});
export type WordDocumentMetadata = z.infer<typeof WordDocumentMetadataSchema> & { readonly presentation?: never; readonly workbook?: never };
export type ContentDocumentMetadata = WordDocumentMetadata | PresentationDocumentMetadata;
export const ContentDocumentMetadataSchema = z.union([WordDocumentMetadataSchema, PresentationDocumentMetadataSchema]);
export const DocumentMetadataSchema = z.union([WordDocumentMetadataSchema, PresentationDocumentMetadataSchema, WorkbookDocumentMetadataSchema]);
export type DocumentMetadata = ContentDocumentMetadata | WorkbookDocumentMetadata;
