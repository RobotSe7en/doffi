import { z } from "zod";
import { DocumentPropertiesSchema } from "./properties.js";
import { IdentifierSchema, PackagePathSchema } from "./primitives.js";
import { WorkbookColumnSchema, WorkbookCoordinateSchema, WorkbookRangeSchema, WorkbookRowSchema } from "./workbook-coordinates.js";
import { WorkbookConditionalRuleSchema, WorkbookDataValidationSchema } from "./workbook-rules.js";
import { WorkbookAutoFilterSchema, WorkbookTableSchema, WorkbookTableStyleReferenceSchema, WorkbookTableStyleSchema } from "./workbook-tables.js";
import { WorkbookCommentAuthorSchema, WorkbookCommentSchema } from "./workbook-comments.js";
import { WorkbookObjectSchema } from "./workbook-drawing.js";
import { WorkbookChartSchema } from "./workbook-charts.js";
import { WorkbookCellStylesSchema } from "./workbook-cell-styles.js";
export * from "./workbook-cell-styles.js";
export * from "./workbook-charts.js";
export * from "./workbook-comments.js";
export * from "./workbook-drawing.js";
export * from "./workbook-coordinates.js";
export * from "./workbook-rules.js";
export * from "./workbook-tables.js";

export const WorkbookChunkDescriptorSchema = z.strictObject({
  sheetId: IdentifierSchema,
  row: WorkbookRowSchema.multipleOf(256),
  column: WorkbookColumnSchema.multipleOf(32),
  path: PackagePathSchema,
  sha256: z.string().length(64).regex(/^[a-f0-9]{64}$/),
});
export type WorkbookChunkDescriptor = z.infer<typeof WorkbookChunkDescriptorSchema>;

export const WorkbookDefinedNameSchema = z.strictObject({
  nameId: IdentifierSchema, name: z.string().min(1).max(255), sheetId: IdentifierSchema.optional(),
  expression: z.string().min(2).max(8192).startsWith("="), origin: WorkbookCoordinateSchema,
  hidden: z.boolean().optional(), description: z.string().max(32767).optional(),
});
export type WorkbookDefinedName = z.infer<typeof WorkbookDefinedNameSchema>;
export const WorkbookHyperlinkSchema = z.strictObject({
  range: WorkbookRangeSchema, url: z.string().min(1).max(8192).optional(), location: z.string().min(1).max(8192).optional(),
  tooltip: z.string().max(32767).optional(), display: z.string().max(32767).optional(),
});
export type WorkbookHyperlink = z.infer<typeof WorkbookHyperlinkSchema>;
export const WorkbookObjectHyperlinkSchema = WorkbookHyperlinkSchema.omit({ range: true, display: true }).extend({ nodeId: IdentifierSchema });
export type WorkbookObjectHyperlink = z.infer<typeof WorkbookObjectHyperlinkSchema>;
export const WorkbookPrintRunSchema = z.discriminatedUnion("kind", [
  z.strictObject({ kind: z.literal("text"), text: z.string().max(32767), cssClass: IdentifierSchema.optional() }),
  z.strictObject({ kind: z.literal("field"), field: z.enum(["page", "pages", "sheet", "file", "path", "date", "time"]), cssClass: IdentifierSchema.optional() }),
]);
export const WorkbookPrintBandSchema = z.strictObject({
  left: z.array(WorkbookPrintRunSchema), center: z.array(WorkbookPrintRunSchema), right: z.array(WorkbookPrintRunSchema),
});
export const WorkbookPrintSchema = z.strictObject({
  areas: z.array(WorkbookRangeSchema),
  repeatRows: z.strictObject({ from: WorkbookRowSchema, to: WorkbookRowSchema }).optional(),
  repeatColumns: z.strictObject({ from: WorkbookColumnSchema, to: WorkbookColumnSchema }).optional(),
  paper: z.strictObject({ widthPt: z.number().positive().max(14400), heightPt: z.number().positive().max(14400) }).optional(),
  orientation: z.enum(["default", "portrait", "landscape"]), pageOrder: z.enum(["downThenOver", "overThenDown"]),
  marginsPt: z.strictObject({ left: z.number().nonnegative(), right: z.number().nonnegative(), top: z.number().nonnegative(), bottom: z.number().nonnegative(), header: z.number().nonnegative(), footer: z.number().nonnegative() }).optional(),
  scale: z.discriminatedUnion("kind", [
    z.strictObject({ kind: z.literal("percent"), percent: z.number().int().min(10).max(400) }),
    z.strictObject({ kind: z.literal("fit"), width: z.number().int().min(0).max(32767), height: z.number().int().min(0).max(32767) }),
  ]),
  horizontalCentered: z.boolean(), verticalCentered: z.boolean(), gridLines: z.boolean(), headings: z.boolean(),
  blackAndWhite: z.boolean(), firstPage: z.number().int().min(1).max(32767).optional(),
  rowBreaks: z.array(z.strictObject({ before: WorkbookRowSchema.min(1), from: WorkbookColumnSchema, to: WorkbookColumnSchema })),
  columnBreaks: z.array(z.strictObject({ before: WorkbookColumnSchema.min(1), from: WorkbookRowSchema, to: WorkbookRowSchema })),
  headerFooter: z.strictObject({
    differentFirst: z.boolean(), differentOddEven: z.boolean(), scaleWithDocument: z.boolean(), alignWithMargins: z.boolean(),
    oddHeader: WorkbookPrintBandSchema, oddFooter: WorkbookPrintBandSchema,
    evenHeader: WorkbookPrintBandSchema.optional(), evenFooter: WorkbookPrintBandSchema.optional(),
    firstHeader: WorkbookPrintBandSchema.optional(), firstFooter: WorkbookPrintBandSchema.optional(),
  }),
});
export type WorkbookPrint = z.infer<typeof WorkbookPrintSchema>;
export type WorkbookPrintBand = z.infer<typeof WorkbookPrintBandSchema>;
export type WorkbookPrintRun = z.infer<typeof WorkbookPrintRunSchema>;

/** Built-in identities can be locale dependent; do not freeze them into an
 * arbitrary locale's format string during import. Custom formats have one code. */
export const WorkbookNumberFormatSchema = z.union([
  z.strictObject({ formatId: IdentifierSchema, code: z.string().min(1).max(2048) }),
  z.strictObject({ formatId: IdentifierSchema, builtinId: z.number().int().min(0).max(163) }),
]);
export type WorkbookNumberFormat = z.infer<typeof WorkbookNumberFormatSchema>;
const workbookOutline = { outlineLevel: z.number().int().min(1).max(7).optional(), collapsed: z.boolean().optional() };
export const WorkbookRowDimensionSchema = z.strictObject({
  from: WorkbookRowSchema, to: WorkbookRowSchema,
  ...workbookOutline,
  /** Explicit automatic sizing can override a manually fixed sheet default. */
  autoHeight: z.boolean().optional(),
  heightPt: z.number().min(0).max(409).optional(), hidden: z.boolean().optional(), cssClass: IdentifierSchema.optional(), numberFormatId: IdentifierSchema.optional(),
});
export const WorkbookColumnDimensionSchema = z.strictObject({
  from: WorkbookColumnSchema, to: WorkbookColumnSchema,
  ...workbookOutline,
  width: z.number().min(0).max(255).optional(), unit: z.literal("characters"), hidden: z.boolean().optional(), cssClass: IdentifierSchema.optional(), numberFormatId: IdentifierSchema.optional(),
});
export const WorkbookSheetSchema = z.strictObject({
  sheetId: IdentifierSchema,
  name: z.string().min(1).max(31).regex(/^[^\\/\[\]:?*\u0000-\u001f]+(?![\s\S])/),
  visibility: z.enum(["visible", "hidden", "veryHidden"]),
  defaultRowHeightPt: z.number().positive().max(409),
  defaultRowHeightCustom: z.boolean().optional(),
  defaultColumnWidth: z.number().positive().max(255),
  defaultColumnWidthUnit: z.enum(["characters", "baseCharacters"]).optional(),
  defaultRowsHidden: z.boolean().optional(),
  defaultCssClass: IdentifierSchema,
  defaultNumberFormatId: IdentifierSchema.optional(),
  rowDimensions: z.array(WorkbookRowDimensionSchema),
  columnDimensions: z.array(WorkbookColumnDimensionSchema),
  merges: z.array(WorkbookRangeSchema),
  frozen: z.strictObject({ rows: z.number().int().min(0).max(1_048_575), columns: z.number().int().min(0).max(16_383) }),
  showGridLines: z.boolean(),
  rightToLeft: z.boolean(),
  tabColor: z.string().length(7).regex(/^#[a-fA-F0-9]{6}$/).optional(),
  tabCssClass: IdentifierSchema.optional(),
  hyperlinks: z.array(WorkbookHyperlinkSchema).optional(),
  print: WorkbookPrintSchema.optional(),
  conditionalRules: z.array(WorkbookConditionalRuleSchema).optional(),
  dataValidations: z.array(WorkbookDataValidationSchema).optional(),
  disableValidationPrompts: z.boolean().optional(),
  autoFilter: WorkbookAutoFilterSchema.optional(),
  tables: z.array(WorkbookTableSchema).optional(),
});
export type WorkbookSheet = z.infer<typeof WorkbookSheetSchema>;
export const WorkbookMetadataSchema = z.strictObject({
  dateSystem: z.enum(["1900", "1904"]),
  calculationMode: z.enum(["auto", "manual"]),
  activeSheetId: IdentifierSchema,
  sheets: z.array(WorkbookSheetSchema).min(1),
  numberFormats: z.array(WorkbookNumberFormatSchema),
  cellStyles: WorkbookCellStylesSchema.optional(),
  /** CSS font used to interpret character-based column widths (native Normal).
   * Older packages without this reference use the worksheet default font. */
  columnWidthFontCssClass: IdentifierSchema.optional(),
  themeCssClass: IdentifierSchema.optional(),
  names: z.array(WorkbookDefinedNameSchema).optional(),
  tableStyles: z.array(WorkbookTableStyleSchema).optional(),
  defaultTableStyle: WorkbookTableStyleReferenceSchema.optional(),
  commentAuthors: z.array(WorkbookCommentAuthorSchema).optional(),
  comments: z.array(WorkbookCommentSchema).optional(),
  objects: z.array(WorkbookObjectSchema).optional(),
  objectHyperlinks: z.array(WorkbookObjectHyperlinkSchema).optional(),
  charts: z.array(WorkbookChartSchema).optional(),
  stories: z.array(z.strictObject({ storyId: IdentifierSchema, rootNodeId: IdentifierSchema, kind: z.literal("comment"), sheetId: IdentifierSchema })).optional(),
});
export type WorkbookMetadata = z.infer<typeof WorkbookMetadataSchema>;
export const WorkbookDocumentMetadataSchema = z.strictObject({
  title: z.string().optional(), locale: IdentifierSchema.optional(), properties: DocumentPropertiesSchema.optional(), workbook: WorkbookMetadataSchema,
});
export type WorkbookDocumentMetadata = z.infer<typeof WorkbookDocumentMetadataSchema> & { readonly word?: never; readonly presentation?: never };

export const WorkbookErrorValueSchema = z.enum(["#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#GETTING_DATA"]);
export const WorkbookValueSchema = z.discriminatedUnion("type", [
  z.strictObject({ type: z.literal("blank") }),
  z.strictObject({ type: z.literal("text"), value: z.string().max(32_767) }),
  z.strictObject({ type: z.literal("number"), value: z.number() }),
  z.strictObject({ type: z.literal("boolean"), value: z.boolean() }),
  z.strictObject({ type: z.literal("error"), value: WorkbookErrorValueSchema }),
]);
export type WorkbookValue = z.infer<typeof WorkbookValueSchema>;
export const WorkbookFormulaSourceSchema = z.strictObject({
  expression: z.string().min(2).max(8192).startsWith("="),
  cacheState: z.enum(["imported", "valid", "stale", "unavailable"]),
});
export type WorkbookFormulaSource = z.infer<typeof WorkbookFormulaSourceSchema>;
