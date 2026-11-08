import { z } from "zod";
import { IdentifierSchema, PackagePathSchema } from "./primitives.js";
import { DrawingPresetSchema, DrawingTextPropertiesSchema } from "./drawing.js";
import { WorkbookDrawingAnchorSchema, WorkbookDrawingLocksSchema } from "./workbook-drawing.js";
import { WorkbookChartSchema, WorkbookChartVectorSchema, WorkbookChartLabelSchema } from "./workbook-charts.js";

const html = z.string().max(2 * 1024 * 1024), properties = z.record(z.string(), z.string().min(1));
export const WorkbookObjectInputSchema = z.discriminatedUnion("kind", [
  z.strictObject({ kind: z.literal("text"), html, text: DrawingTextPropertiesSchema.optional() }),
  z.strictObject({ kind: z.literal("geometry"), preset: DrawingPresetSchema, html: html.optional(), text: DrawingTextPropertiesSchema.optional() }),
  z.strictObject({ kind: z.literal("image"), assetPath: PackagePathSchema, alt: z.string().max(32767) }),
]);
export type WorkbookObjectInput = z.infer<typeof WorkbookObjectInputSchema>;
/** Chart input owns either references or native literals, never a second data set.
 * Retained series identities must belong to the original chart. */
export const WorkbookChartDataInputSchema = z.strictObject({
  categories: WorkbookChartVectorSchema.optional(),
  series: z.array(z.strictObject({ seriesId: IdentifierSchema.optional(), name: WorkbookChartLabelSchema.optional(), values: WorkbookChartVectorSchema })).min(1),
});
export type WorkbookChartDataInput = z.infer<typeof WorkbookChartDataInputSchema>;
export const WorkbookDrawingOperationSchema = z.discriminatedUnion("op", [
  z.strictObject({ op: z.literal("insertWorkbookObject"), sheetId: IdentifierSchema, anchor: WorkbookDrawingAnchorSchema, object: WorkbookObjectInputSchema, name: z.string().max(32767).optional(), properties: properties.optional() }),
  z.strictObject({ op: z.literal("updateWorkbookObject"), nodeId: IdentifierSchema, changes: z.strictObject({
    anchor: WorkbookDrawingAnchorSchema.optional(), name: z.string().max(32767).nullable().optional(), locked: z.boolean().optional(), printable: z.boolean().optional(), locks: WorkbookDrawingLocksSchema.nullable().optional(),
    rotation: z.number().optional(), flipH: z.boolean().optional(), flipV: z.boolean().optional(), hidden: z.boolean().optional(), preset: DrawingPresetSchema.optional(),
  }) }),
  z.strictObject({ op: z.literal("setWorkbookObjectText"), nodeId: IdentifierSchema, html, text: DrawingTextPropertiesSchema.nullable().optional() }),
  z.strictObject({ op: z.literal("setWorkbookObjectStyle"), nodeId: IdentifierSchema, properties }),
  z.strictObject({ op: z.literal("replaceWorkbookImage"), nodeId: IdentifierSchema, assetPath: PackagePathSchema, alt: z.string().max(32767) }),
  z.strictObject({ op: z.literal("setWorkbookImageCrop"), nodeId: IdentifierSchema, crop: z.strictObject({ left: z.number().min(0).lt(1), top: z.number().min(0).lt(1), right: z.number().min(0).lt(1), bottom: z.number().min(0).lt(1) }) }),
  z.strictObject({ op: z.literal("removeWorkbookObject"), nodeId: IdentifierSchema }),
  z.strictObject({ op: z.literal("duplicateWorkbookObjects"), nodeIds: z.array(IdentifierSchema).min(1), offset: z.strictObject({ xPt: z.number(), yPt: z.number() }) }),
  z.strictObject({ op: z.literal("reorderWorkbookObjects"), nodeIds: z.array(IdentifierSchema).min(1), position: z.enum(["front", "back", "forward", "backward"]) }),
  z.strictObject({ op: z.literal("insertWorkbookChart"), sheetId: IdentifierSchema, anchor: WorkbookDrawingAnchorSchema,
    kind: z.enum(["column", "line", "pie"]), data: WorkbookChartDataInputSchema, title: html.optional(), name: z.string().max(32767).optional() }),
  z.strictObject({ op: z.literal("setWorkbookChart"), nodeId: IdentifierSchema,
    changes: WorkbookChartSchema.omit({ chartId: true, nodeId: true, series: true, categories: true, titleNodeId: true, titleReference: true, categoryAxisTitle: true, valueAxisTitle: true }).partial(),
    clear: z.array(z.enum(["minimum", "maximum", "majorUnit", "numberFormat", "labelContents", "seriesStyles", "pointStyles", "categories"])).optional() }),
  z.strictObject({ op: z.literal("setWorkbookChartData"), nodeId: IdentifierSchema, data: WorkbookChartDataInputSchema }),
  z.strictObject({ op: z.literal("setWorkbookChartTitle"), nodeId: IdentifierSchema, target: z.enum(["title", "categoryAxis", "valueAxis"]),
    source: z.union([z.null(), z.strictObject({ html, text: DrawingTextPropertiesSchema.optional(), cssClass: IdentifierSchema.optional() }),
      z.strictObject({ reference: WorkbookChartSchema.shape.titleReference.unwrap(), cssClass: IdentifierSchema.optional() })]) }),
]);
export type WorkbookDrawingOperation = z.infer<typeof WorkbookDrawingOperationSchema>;
