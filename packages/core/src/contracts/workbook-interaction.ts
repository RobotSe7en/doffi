import { z } from "zod";
import { AnnotationSchema, AnnotationsSchema } from "./annotations.js";
import { AssetDescriptorSchema } from "./document.js";
import { ContextStyleSchema } from "./interaction.js";
import { FORMAT_VERSION, IdentifierSchema, PROTOCOL_VERSION, RevisionSchema } from "./primitives.js";
import { WorkbookDocumentMetadataSchema, WorkbookCoordinateSchema, WorkbookFormulaSourceSchema, WorkbookRangeSchema, WorkbookValueSchema } from "./workbook.js";
import { WorkbookOperationSchema } from "./workbook-operations.js";

export const WorkbookInteractionRangeSchema = z.strictObject({ sheetId: IdentifierSchema, range: WorkbookRangeSchema });
export type WorkbookInteractionRange = z.infer<typeof WorkbookInteractionRangeSchema>;
export const WorkbookInteractionScopeSchema = z.discriminatedUnion("kind", [
  z.strictObject({ kind: z.literal("ranges"), ranges: z.array(WorkbookInteractionRangeSchema).min(1), objectIds: z.array(IdentifierSchema).optional(), newObjectSheetIds: z.array(IdentifierSchema).optional() }),
  z.strictObject({ kind: z.literal("tables"), tableIds: z.array(IdentifierSchema).min(1), objectIds: z.array(IdentifierSchema).optional(), newObjectSheetIds: z.array(IdentifierSchema).optional() }),
  z.strictObject({ kind: z.literal("global") }),
]);
export type WorkbookInteractionScope = z.infer<typeof WorkbookInteractionScopeSchema>;
export const WorkbookInteractionLimitsSchema = z.strictObject({
  maxBytes: z.number().int().positive(), maxCells: z.number().int().positive(), maxNodes: z.number().int().positive(),
  maxDependencies: z.number().int().positive(), maxSteps: z.number().int().positive(), maxAssets: z.number().int().positive(),
});
export type WorkbookInteractionLimits = z.infer<typeof WorkbookInteractionLimitsSchema>;
export const WorkbookInteractionCalculationContextSchema = z.strictObject({
  now: z.strictObject({ unixMilliseconds: z.number().min(-8.64e15).max(8.64e15), utcOffsetMinutes: z.number().int().min(-1440).max(1440) }).optional(),
  locale: z.string().min(1).max(255).optional(),
});
export const WorkbookInteractionValueSchema = z.discriminatedUnion("kind", [
  z.strictObject({ kind: z.literal("value"), value: WorkbookValueSchema }),
  z.strictObject({ kind: z.literal("unavailable"), reason: z.enum(["unsupported-function", "circular-reference", "circular-name", "ambiguous-name-scope", "calculation-required"]), functionName: z.string().optional() }),
]);
export type WorkbookInteractionValue = z.infer<typeof WorkbookInteractionValueSchema>;
export const WorkbookInteractionCellSchema = WorkbookCoordinateSchema.extend({
  sheetId: IdentifierSchema, access: z.enum(["editable", "read-only"]),
  value: WorkbookValueSchema, formula: WorkbookFormulaSourceSchema.optional(), html: z.string(), text: z.string(),
  cssClass: z.string().optional(), formatId: IdentifierSchema.optional(), current: WorkbookInteractionValueSchema,
});
export type WorkbookInteractionCell = z.infer<typeof WorkbookInteractionCellSchema>;
export const WorkbookInteractionContextSchema = z.strictObject({
  formatVersion: z.literal(FORMAT_VERSION), kind: z.literal("workbook"), scope: WorkbookInteractionScopeSchema,
  writableRanges: z.array(WorkbookInteractionRangeSchema).min(1), writableObjectIds: z.array(IdentifierSchema), writableNewObjectSheetIds: z.array(IdentifierSchema), writableAnnotationIds: z.array(IdentifierSchema),
  /** Rectangles describe complete sparse coverage, including implicit blanks.
   * The cells array holds every stored cell in that coverage, never a sample. */
  requestedReadRanges: z.array(WorkbookInteractionRangeSchema), dependencies: z.array(WorkbookInteractionRangeSchema), cells: z.array(WorkbookInteractionCellSchema), annotations: AnnotationsSchema,
  blocks: z.array(z.strictObject({ nodeId: IdentifierSchema, html: z.string().min(1), access: z.enum(["editable", "read-only"]) })),
  styles: z.array(ContextStyleSchema), metadata: WorkbookDocumentMetadataSchema,
  assets: z.array(AssetDescriptorSchema), providedAssets: z.array(AssetDescriptorSchema),
  calculationContext: WorkbookInteractionCalculationContextSchema, limits: WorkbookInteractionLimitsSchema,
});
export type WorkbookInteractionContext = z.infer<typeof WorkbookInteractionContextSchema>;
export const WorkbookProposalOperationSchema = WorkbookOperationSchema;
export const WorkbookProposalOperationNameSchema = z.enum(WorkbookProposalOperationSchema.options.map(schema => schema.shape.op.value));
export type WorkbookProposalOperationName = z.infer<typeof WorkbookProposalOperationNameSchema>;
export const WorkbookInteractionRequestSchema = z.strictObject({
  protocolVersion: z.literal(PROTOCOL_VERSION), requestId: IdentifierSchema, documentId: IdentifierSchema, baseRevision: RevisionSchema,
  annotations: z.array(z.union([AnnotationSchema.options[0].extend({ status: z.literal("open") }), AnnotationSchema.options[1].extend({ status: z.literal("open") })])).min(1),
  context: WorkbookInteractionContextSchema, supportedOperations: z.array(WorkbookProposalOperationNameSchema),
});
export type WorkbookInteractionRequest = z.infer<typeof WorkbookInteractionRequestSchema>;
export const WorkbookEditProposalSchema = z.strictObject({
  protocolVersion: z.literal(PROTOCOL_VERSION), requestId: IdentifierSchema, proposalId: IdentifierSchema, documentId: IdentifierSchema, baseRevision: RevisionSchema,
  annotationIds: z.array(IdentifierSchema), summary: z.string(), operations: z.array(WorkbookProposalOperationSchema),
});
export type WorkbookEditProposal = z.infer<typeof WorkbookEditProposalSchema>;
/** Current-source calculation is separate from source changes. Recalculated
 * consumers can be read-only even when their displayed result changes. */
export const WorkbookProposalReviewSchema = z.strictObject({
  requestId: IdentifierSchema, proposalId: IdentifierSchema, summary: z.string(),
  calculatedChanges: z.array(WorkbookCoordinateSchema.extend({ sheetId: IdentifierSchema, access: z.enum(["editable", "read-only"]), before: WorkbookInteractionValueSchema, after: WorkbookInteractionValueSchema })),
});
export type WorkbookProposalReview = z.infer<typeof WorkbookProposalReviewSchema>;
