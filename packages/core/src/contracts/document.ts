import { z } from "zod";
import { FORMAT_VERSION, IdentifierSchema, PackagePathSchema, RevisionSchema } from "./primitives.js";
import { WordMetadataSchema } from "./word.js";
import {PageSetupSchema,DocumentSectionSchema} from './word-page.js';
export {PageSetupSchema,DocumentSectionSchema} from './word-page.js';
export type {PageSetup,DocumentSection} from './word-page.js';
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
