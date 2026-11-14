import { WordTextStyleSchema } from './word-text-style.js';
import { z } from 'zod';
import { IdentifierSchema, PackagePathSchema } from './primitives.js';
export const WordChartKindSchema = z.enum(['column', 'line', 'pie', 'bar', 'area', 'doughnut', 'scatter', 'combo']);
const seriesKind = z.enum(['column', 'line', 'bar', 'area', 'scatter']);
export const WordChartTextStyleSchema = WordTextStyleSchema.extend({ fonts:z.strictObject({ascii:z.string().max(255).optional(),hAnsi:z.string().max(255).optional(),eastAsia:z.string().max(255).optional(),cs:z.string().max(255).optional()}).optional(),fontSize: z.number().min(1).max(4000).optional() });
export const WordChartAxisSchema = z.strictObject({ visible: z.boolean().optional(), reverse: z.boolean().optional(), crosses: z.enum(['autoZero', 'min', 'max']).optional(), crossesAt: z.number().optional(), labelPosition: z.enum(['nextTo', 'low', 'high', 'none']).optional(), crossBetween: z.enum(['between', 'midCat']).optional(), position: z.enum(['l', 'r', 't', 'b']).optional(), majorGridlines: z.boolean().optional(), numberFormat: z.string().optional(), title: z.string().optional(), textStyle: WordChartTextStyleSchema.optional() });
export const WordChartPropertiesSchema = z.strictObject({
    kind: WordChartKindSchema, textStyle: WordChartTextStyleSchema.optional(), titleStyle: WordChartTextStyleSchema.optional(), legendStyle: WordChartTextStyleSchema.optional(), labelStyle: WordChartTextStyleSchema.optional(), axes: z.strictObject({ category: WordChartAxisSchema.optional(), value: WordChartAxisSchema.optional(), secondaryCategory: WordChartAxisSchema.optional(), secondaryValue: WordChartAxisSchema.optional() }).optional(), title: z.string().optional(), legend: z.enum(['none', 'left', 'right', 'top', 'bottom']), showLabels: z.boolean(), labelContents: z.array(z.enum(['value', 'category', 'series', 'percent'])).optional(),
    categories: z.array(z.string()), series: z.array(z.strictObject({ seriesId: IdentifierSchema, name: z.string(), values: z.array(z.union([z.number(), z.null()])), xValues: z.array(z.union([z.number(), z.null()])).optional(), kind: seriesKind.optional(), axis: z.enum(['primary', 'secondary']).optional(), color: z.string().regex(/^#[0-9a-fA-F]{6}$/).optional(), marker: z.enum(['none', 'circle', 'square', 'diamond', 'triangle', 'x', 'plus', 'star', 'dot', 'dash', 'auto']).optional(), markerSize: z.number().min(2).max(72).optional(), smooth: z.boolean().optional(), lineVisible: z.boolean().optional(), pointColors: z.array(z.string().regex(/^#[0-9a-fA-F]{6}$/).nullable()).optional() })).min(1),
    scatterStyle: z.enum(['marker', 'line', 'lineMarker', 'smooth', 'smoothMarker']).optional(), numberFormat: z.string().optional(), minimum: z.number().optional(), maximum: z.number().optional(), majorUnit: z.number().positive().optional(), xMinimum: z.number().optional(), xMaximum: z.number().optional(), xMajorUnit: z.number().positive().optional(), secondaryXMinimum: z.number().optional(), secondaryXMaximum: z.number().optional(), secondaryXMajorUnit: z.number().positive().optional(), secondaryMinimum: z.number().optional(), secondaryMaximum: z.number().optional(), secondaryMajorUnit: z.number().positive().optional(),
    grouping: z.enum(['standard', 'clustered', 'stacked', 'percentStacked']).optional(), gapWidth: z.number().min(0).max(500).optional(), overlap: z.number().min(-100).max(100).optional(), holeSize: z.number().int().min(10).max(90).optional(), firstSliceAngle: z.number().min(0).max(360).optional(), blankAs: z.enum(['gap', 'zero', 'span']).optional(),
});
/** Original chart parts are immutable templates for unmodeled formatting. The
 * typed vectors above are the sole current data; output rewrites their caches
 * and embedded worksheet cells before exposing the native graph. */
export const WordChartNativeSchema = z.strictObject({
    baseline: WordChartPropertiesSchema.omit({ categories: true, series: true }).extend({ series: z.array(WordChartPropertiesSchema.shape.series.element.omit({ values: true, xValues: true })) }).optional(),
    primaryAxisIds: z.array(z.string()).optional(), secondaryAxisIds: z.array(z.string()).optional(),
    previewWarnings: z.array(z.string()).optional(),
    chartPart: PackagePathSchema, workbookPart: PackagePathSchema.optional(),
    parts: z.array(z.strictObject({ sourcePart: PackagePathSchema, assetPath: PackagePathSchema, contentType: z.string().min(1) })),
    relationships: z.array(z.strictObject({ sourcePart: PackagePathSchema, id: IdentifierSchema, type: z.string().min(1), target: z.string().min(1), external: z.boolean() })),
    bindings: z.array(z.strictObject({ seriesId: IdentifierSchema, sourceIndex: z.number().int().nonnegative().optional(), color: z.string().optional(), pointColors: z.array(z.string().nullable()).optional(), nameFormula: z.string().optional(), categoryFormula: z.string().optional(), valueFormula: z.string().optional(), xFormula: z.string().optional() })),
});
export const WordChartSchema = WordChartPropertiesSchema.extend({ native: WordChartNativeSchema.optional() });
export type WordChart = z.infer<typeof WordChartSchema>;
