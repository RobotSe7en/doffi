import { z } from "zod";
import { IdentifierSchema } from "./primitives.js";
import { WorkbookDrawingOperationSchema } from "./workbook-drawing-operations.js";
import { WorkbookCommentOperationSchema } from "./workbook-comment-operations.js";
export * from "./workbook-drawing-operations.js";
export * from "./workbook-comment-operations.js";
import { WorkbookAutoFilterSchema, WorkbookCellStylesSchema, WorkbookConditionalRuleSchema, WorkbookCoordinateSchema, WorkbookDataValidationSchema, WorkbookDefinedNameSchema, WorkbookHyperlinkSchema, WorkbookNumberFormatSchema, WorkbookRangeSchema, WorkbookSheetSchema, WorkbookTableSchema, WorkbookTableStyleReferenceSchema, WorkbookTableStyleSchema, WorkbookValueSchema } from "./workbook.js";

export const WorkbookCellSourceSchema = z.discriminatedUnion("kind", [
  z.strictObject({ kind: z.literal("value"), value: WorkbookValueSchema }),
  z.strictObject({ kind: z.literal("formula"), expression: z.string().min(2).max(8192).startsWith("=") }),
  z.strictObject({ kind: z.literal("richText"), html: z.string().max(2 * 1024 * 1024) }),
]);
export type WorkbookCellSource = z.infer<typeof WorkbookCellSourceSchema>;
export const WorkbookCellInputSchema = WorkbookCoordinateSchema.extend({
  source: WorkbookCellSourceSchema,
  cssClass: z.string().min(1).max(4096).nullable().optional(),
  formatId: IdentifierSchema.nullable().optional(),
});
export type WorkbookCellInput = z.infer<typeof WorkbookCellInputSchema>;
export const WorkbookCellFormatInputSchema = WorkbookCoordinateSchema.extend({
  cssClass: z.string().min(1).max(4096).nullable().optional(), formatId: IdentifierSchema.nullable().optional(),
});
export type WorkbookCellFormatInput = z.infer<typeof WorkbookCellFormatInputSchema>;
export const WorkbookOperationSchema = z.discriminatedUnion("op", [
  ...WorkbookDrawingOperationSchema.options,
  ...WorkbookCommentOperationSchema.options,
  z.strictObject({ op: z.literal("setWorkbookCells"), sheetId: IdentifierSchema, cells: z.array(WorkbookCellInputSchema).min(1) }),
  z.strictObject({ op: z.literal("setWorkbookCellFormats"), sheetId: IdentifierSchema, cells: z.array(WorkbookCellFormatInputSchema).min(1) }),
  z.strictObject({ op: z.literal("setWorkbookNumberFormat"), format: WorkbookNumberFormatSchema }),
  z.strictObject({ op: z.literal("setWorkbookStyleClass"), cssClass: IdentifierSchema, properties: z.record(z.string(), z.string().min(1)) }),
  z.strictObject({ op: z.literal("formatWorkbookRanges"), sheetId: IdentifierSchema, ranges: z.array(WorkbookRangeSchema).min(1), properties: z.record(z.string(), z.string().min(1)), numberFormatId: IdentifierSchema.optional() }),
  z.strictObject({ op: z.literal("formatWorkbookConditionalRule"), sheetId: IdentifierSchema, ruleId: IdentifierSchema, stopIndex: z.number().int().min(0).max(2).optional(), properties: z.record(z.string(), z.string().min(1)), numberFormatId: IdentifierSchema.nullable().optional() }),
  z.strictObject({ op: z.literal("formatWorkbookTableStyleElement"), styleId: IdentifierSchema, region: WorkbookTableStyleSchema.shape.elements.element.shape.region, properties: z.record(z.string(), z.string().min(1)), numberFormatId: IdentifierSchema.nullable().optional() }),
  z.strictObject({ op: z.literal("formatWorkbookText"), sheetId: IdentifierSchema, at: WorkbookCoordinateSchema, from: z.number().int().min(0).max(32767), to: z.number().int().min(0).max(32767), expectedText: z.string().max(32767), properties: z.record(z.string(), z.string().min(1)) }),
  z.strictObject({ op: z.literal("setWorkbookCellStyleDefinition"), definition: WorkbookCellStylesSchema.shape.definitions.element.omit({ cssClass: true }), properties: z.record(z.string(), z.string().min(1)) }),
  z.strictObject({ op: z.literal("setWorkbookCellStyleName"), name: WorkbookCellStylesSchema.shape.names.element }),
  z.strictObject({ op: z.literal("removeWorkbookCellStyleName"), name: z.string().min(1).max(255) }),
  z.strictObject({ op: z.literal("setWorkbookHyperlink"), sheetId: IdentifierSchema, range: WorkbookRangeSchema, target: WorkbookHyperlinkSchema.omit({ range: true }).nullable() }),
  z.strictObject({ op: z.literal("replaceWorkbookHyperlinks"), sheetId: IdentifierSchema, range: WorkbookRangeSchema, links: z.array(WorkbookHyperlinkSchema) }),
  z.strictObject({ op: z.literal("clearWorkbookRanges"), sheetId: IdentifierSchema, ranges: z.array(WorkbookRangeSchema).min(1), mode: z.enum(["contents", "all"]) }),
  z.strictObject({ op: z.literal("mergeWorkbookCells"), sheetId: IdentifierSchema, range: WorkbookRangeSchema, coveredContent: z.enum(["reject", "clear"]).optional() }),
  z.strictObject({ op: z.literal("unmergeWorkbookCells"), sheetId: IdentifierSchema, range: WorkbookRangeSchema }),
  z.strictObject({ op: z.literal("setWorkbookAutoFilter"), sheetId: IdentifierSchema, filter: WorkbookAutoFilterSchema.nullable() }),
  z.strictObject({ op: z.literal("setWorkbookDefinedName"), name: WorkbookDefinedNameSchema }),
  z.strictObject({ op: z.literal("removeWorkbookDefinedName"), nameId: IdentifierSchema }),
  z.strictObject({ op: z.literal("setWorkbookTable"), sheetId: IdentifierSchema, table: WorkbookTableSchema }),
  z.strictObject({ op: z.literal("removeWorkbookTable"), sheetId: IdentifierSchema, tableId: IdentifierSchema }),
  z.strictObject({ op: z.literal("setWorkbookTableStyle"), style: WorkbookTableStyleSchema }),
  z.strictObject({ op: z.literal("removeWorkbookTableStyle"), styleId: IdentifierSchema }),
  z.strictObject({ op: z.literal("setWorkbookConditionalRule"), sheetId: IdentifierSchema, rule: WorkbookConditionalRuleSchema }),
  z.strictObject({ op: z.literal("removeWorkbookConditionalRule"), sheetId: IdentifierSchema, ruleId: IdentifierSchema }),
  z.strictObject({ op: z.literal("setWorkbookDataValidation"), sheetId: IdentifierSchema, rule: WorkbookDataValidationSchema }),
  z.strictObject({ op: z.literal("removeWorkbookDataValidation"), sheetId: IdentifierSchema, ruleId: IdentifierSchema }),
  z.strictObject({ op: z.literal("editWorkbookAxis"), sheetId: IdentifierSchema, axis: z.enum(["row", "column"]), action: z.enum(["insert", "delete"]),
    index: z.number().int().min(0).max(1_048_575), count: z.number().int().min(1).max(1_048_576), formatSource: z.enum(["none", "before", "after"]).optional() }),
  z.strictObject({ op: z.literal("reorderWorkbookRows"), sheetId: IdentifierSchema, range: WorkbookRangeSchema,
    rows: z.array(z.number().int().min(0).max(1_048_575)).min(1).max(1_048_576) }),
  z.strictObject({ op: z.literal("moveWorkbookRange"), sheetId: IdentifierSchema, range: WorkbookRangeSchema, targetSheetId: IdentifierSchema,
    target: WorkbookCoordinateSchema, overwrittenContent: z.enum(["reject", "clear"]).optional() }),
  z.strictObject({ op: z.literal("setWorkbookSettings"), changes: z.strictObject({
    calculationMode: z.enum(["auto", "manual"]).optional(), activeSheetId: IdentifierSchema.optional(), defaultTableStyle: WorkbookTableStyleReferenceSchema.nullable().optional(),
  }) }),
  z.strictObject({ op: z.literal("insertWorkbookSheet"), name: WorkbookSheetSchema.shape.name, beforeSheetId: IdentifierSchema.optional() }),
  z.strictObject({ op: z.literal("copyWorkbookSheet"), sheetId: IdentifierSchema, name: WorkbookSheetSchema.shape.name, beforeSheetId: IdentifierSchema.optional(),
    nameConflicts: z.array(z.strictObject({ nameId: IdentifierSchema, name: z.string().min(1).max(255) })).optional() }),
  z.strictObject({ op: z.literal("updateWorkbookSheet"), sheetId: IdentifierSchema, changes: WorkbookSheetSchema.pick({
    name: true, visibility: true, defaultRowHeightPt: true, defaultRowHeightCustom: true, defaultColumnWidth: true,
    defaultColumnWidthUnit: true, defaultRowsHidden: true, rowDimensions: true, columnDimensions: true,
    frozen: true, showGridLines: true, rightToLeft: true, tabColor: true,
  }).partial() }),
  z.strictObject({ op: z.literal("removeWorkbookSheet"), sheetId: IdentifierSchema }),
  z.strictObject({ op: z.literal("moveWorkbookSheet"), sheetId: IdentifierSchema, beforeSheetId: IdentifierSchema.optional() }),
]);
export type WorkbookOperation = z.infer<typeof WorkbookOperationSchema>;
export const WorkbookOperationNameSchema = z.enum(WorkbookOperationSchema.options.map(schema => schema.shape.op.value));
export type WorkbookOperationName = z.infer<typeof WorkbookOperationNameSchema>;
