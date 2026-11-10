import { z } from "zod";
import { IdentifierSchema } from "./primitives.js";
import { ChartAppearanceSchema } from "./chart.js";
import { DrawingTextPropertiesSchema } from "./drawing.js";

const index = z.number().int().min(0).max(0xffff_ffff);
export const WorkbookChartCacheSchema = z.strictObject({
  state: z.enum(["imported", "valid", "stale"]), count: index,
  points: z.array(z.strictObject({ index, value: z.union([z.number(), z.string().max(32767), z.null()]), numberFormat: z.string().max(2048).optional() })),
  numberFormat: z.string().max(2048).optional(),
});
export type WorkbookChartCache = z.infer<typeof WorkbookChartCacheSchema>;
/** A reference is the source. Native caches are explicitly derived values and
 * never become an independently writable chart data table. Origin is A1 on the
 * owning worksheet unless explicitly bound to another current worksheet. This
 * binding preserves caller-dependent names when the drawing changes owner. */
export const WorkbookChartReferenceSchema = z.strictObject({
  expression: z.string().min(2).max(8192).startsWith("="), valueType: z.enum(["text", "number"]), cache: WorkbookChartCacheSchema.optional(),
  originSheetId: IdentifierSchema.optional(),
});
export type WorkbookChartReference = z.infer<typeof WorkbookChartReferenceSchema>;
/** Native literal vectors own their values. They have neither a formula nor a
 * calculation/cache state, and never fabricate worksheet cells. */
export const WorkbookChartLiteralSchema = WorkbookChartCacheSchema.omit({ state: true }).extend({
  kind: z.literal("literal"), valueType: z.enum(["text", "number"]),
});
export type WorkbookChartLiteral = z.infer<typeof WorkbookChartLiteralSchema>;
export const WorkbookChartVectorSchema = z.union([WorkbookChartReferenceSchema, WorkbookChartLiteralSchema]);
export type WorkbookChartVector = z.infer<typeof WorkbookChartVectorSchema>;
export const WorkbookChartLabelSchema = z.discriminatedUnion("kind", [
  z.strictObject({ kind: z.literal("text"), text: z.string().max(32767) }),
  z.strictObject({ kind: z.literal("reference"), reference: WorkbookChartReferenceSchema }),
]);
export type WorkbookChartLabel = z.infer<typeof WorkbookChartLabelSchema>;
const axisTitle = z.strictObject({
  nodeId: IdentifierSchema.optional(), reference: WorkbookChartReferenceSchema.optional(),
  cssClass: IdentifierSchema.optional(), text: DrawingTextPropertiesSchema.optional(),
});
export const WorkbookChartSchema = ChartAppearanceSchema.extend({
  chartId: IdentifierSchema, nodeId: IdentifierSchema,
  categories: WorkbookChartVectorSchema.optional(),
  series: z.array(z.strictObject({ seriesId: IdentifierSchema, nativeIndex: index.optional(), name: WorkbookChartLabelSchema.optional(), values: WorkbookChartVectorSchema })).min(1),
  titleReference: WorkbookChartReferenceSchema.optional(),
  titleCssClass: IdentifierSchema.optional(), titleText: DrawingTextPropertiesSchema.optional(),
  categoryAxisTitle: axisTitle.optional(), valueAxisTitle: axisTitle.optional(),
  visibleCellsOnly: z.boolean().optional(), varyColors: z.boolean().optional(), showLabelsOverMaximum: z.boolean().optional(),
  language: z.string().max(255).optional(), dateSystem: z.enum(["1900", "1904"]).optional(),
});
export type WorkbookChart = z.infer<typeof WorkbookChartSchema>;
