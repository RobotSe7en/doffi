import { WordChartSchema, WordChartPropertiesSchema } from "./word-chart.js";
import { z } from "zod";
import { IdentifierSchema, PackagePathSchema, RevisionSchema } from "./primitives.js";
import { DrawingGeometrySchema, DrawingPresetSchema } from "./drawing.js";
export const WordDrawingLayoutSchema = z.strictObject({
  mode: z.enum(["inline", "floating"]),
  horizontal: z
    .enum([
      "page",
      "margin",
      "column",
      "character",
      "leftMargin",
      "rightMargin",
      "insideMargin",
      "outsideMargin",
    ])
    .optional(),
  vertical: z
    .enum([
      "page",
      "margin",
      "paragraph",
      "line",
      "topMargin",
      "bottomMargin",
      "insideMargin",
      "outsideMargin",
    ])
    .optional(),
  alignH: z.enum(["left", "center", "right", "inside", "outside"]).optional(),
  alignV: z.enum(["top", "center", "bottom", "inside", "outside"]).optional(),
  wrap: z.enum(["none", "square", "tight", "through", "topAndBottom"]).optional(),
  wrapSide: z.enum(["bothSides", "left", "right", "largest"]).optional(),
  distances: z
    .strictObject({
      top: z.number().nonnegative(),
      right: z.number().nonnegative(),
      bottom: z.number().nonnegative(),
      left: z.number().nonnegative(),
    })
    .optional(),
  behindText: z.boolean().optional(),
  hidden: z.boolean().optional(),
  effectExtent: z.strictObject({ top: z.number().nonnegative(), right: z.number().nonnegative(), bottom: z.number().nonnegative(), left: z.number().nonnegative() }).optional(),
  allowOverlap: z.boolean().optional(),
  layoutInCell: z.boolean().optional(),
  locked: z.boolean().optional(),
  zOrder: z.number().int().nonnegative().optional(),
  relativeWidth: z
    .strictObject({
      relativeFrom: z.enum([
        "page",
        "margin",
        "leftMargin",
        "rightMargin",
        "insideMargin",
        "outsideMargin",
      ]),
      percent: z.number().nonnegative().max(2147483.647),
    })
    .optional(),
  relativeHeight: z
    .strictObject({
      relativeFrom: z.enum([
        "page",
        "margin",
        "topMargin",
        "bottomMargin",
        "insideMargin",
        "outsideMargin",
      ]),
      percent: z.number().nonnegative().max(2147483.647),
    })
    .optional(),
  wrapPolygon: z
    .array(z.strictObject({ x: z.number(), y: z.number() }))
    .min(3)
    .optional(),
  wrapPolygonEdited: z.boolean().optional(),
});
const crop = z.strictObject({
  left: z.number().nonnegative().lt(1),
  top: z.number().nonnegative().lt(1),
  right: z.number().nonnegative().lt(1),
  bottom: z.number().nonnegative().lt(1),
});
const connector = z.strictObject({
  routing: z.enum(["straight", "elbow"]),
  startNodeId: IdentifierSchema.optional(),
  endNodeId: IdentifierSchema.optional(),
  startSite: z.number().int().min(0).max(3).optional(),
  endSite: z.number().int().min(0).max(3).optional(),
});
export const WordDrawingTextSchema = z.strictObject({
  insets: z
    .strictObject({ top: z.number(), right: z.number(), bottom: z.number(), left: z.number() })
    .optional(),
  vertical: z.enum(["top", "middle", "bottom"]).optional(),
  direction: z.enum(["horizontal", "vertical", "vertical270"]).optional(),
  rotation: z.number().optional(),
  upright: z.boolean().optional(),
  wrap: z.boolean().optional(),
  autoFit: z.enum(["none", "normal", "shape"]).optional(),
  fontScale: z.number().positive().max(1).optional(),
  lineSpacingReduction: z.number().min(0).max(1).optional(),
});
const base = {
  nodeId: IdentifierSchema,
  geometry: DrawingGeometrySchema,
  layout: WordDrawingLayoutSchema,
  name: z.string().optional(),
  watermark: z.boolean().optional(),
  locks: z.strictObject({ noChangeArrowheads: z.boolean().optional(), noChangeShapeType: z.boolean().optional() }).optional(),
};
export const WordDrawingSchema = z.discriminatedUnion("kind", [
  z.strictObject({ ...base, kind: z.literal("chart"), chart: WordChartSchema }),
  z.strictObject({ ...base, kind: z.literal("image"), crop: crop.optional() }),
  z.strictObject({
    ...base,
    kind: z.literal("shape"),
    preset: DrawingPresetSchema,
    textStoryId: IdentifierSchema.optional(),
    text: WordDrawingTextSchema.optional(),
    adjustments: z.record(z.string(), z.number().int()).optional(),
  }),
  z.strictObject({
    ...base,
    kind: z.literal("textBox"),
    textStoryId: IdentifierSchema,
    text: WordDrawingTextSchema.optional(),
  }),
  z.strictObject({ ...base, kind: z.literal("connector"), connector }),
  z.strictObject({
    ...base,
    kind: z.literal("group"),
    coordinateSpace: z.strictObject({
      x: z.number(),
      y: z.number(),
      width: z.number().positive(),
      height: z.number().positive(),
    }),
  }),
]);
export type WordDrawing = z.infer<typeof WordDrawingSchema>;
export type WordDrawingLayout = z.infer<typeof WordDrawingLayoutSchema>;
export const WordDrawingInputSchema = z.discriminatedUnion("kind", [
  z.strictObject({ kind: z.literal("chart"), chart: WordChartPropertiesSchema }),
  z.strictObject({
    kind: z.literal("image"),
    assetPath: PackagePathSchema,
    alt: z.string(),
    crop: crop.optional(),
  }),
  z.strictObject({
    kind: z.literal("shape"),
    preset: DrawingPresetSchema,
    text: z.string().optional(),
  }),
  z.strictObject({ kind: z.literal("textBox"), text: z.string() }),
  z.strictObject({ kind: z.literal("connector"), connector }),
]);
export const WordDrawingOperationSchema = z.discriminatedUnion("op", [
  z.strictObject({
    op: z.literal("insertWordDrawing"),
    nodeId: IdentifierSchema,
    offset: RevisionSchema,
    geometry: DrawingGeometrySchema,
    layout: WordDrawingLayoutSchema,
    object: WordDrawingInputSchema,
    properties: z.record(z.string(), z.string()).optional(),
  }),
  z.strictObject({
    op: z.literal("setWordChart"),
    nodeId: IdentifierSchema,
    changes: WordChartPropertiesSchema.partial(),
  }),
  z.strictObject({
    op: z.literal("setWordDrawingGeometry"),
    nodeId: IdentifierSchema,
    geometry: DrawingGeometrySchema,
  }),
  z.strictObject({
    op: z.literal("setWordDrawingLayout"),
    nodeId: IdentifierSchema,
    layout: WordDrawingLayoutSchema,
  }),
  z.strictObject({
    op: z.literal("setWordDrawingAppearance"),
    nodeId: IdentifierSchema,
    properties: z.record(z.string(), z.string()),
  }),
  z.strictObject({ op: z.literal("setWordDrawingCrop"), nodeId: IdentifierSchema, crop }),
  z.strictObject({
    op: z.literal("setWordDrawingPreset"),
    nodeId: IdentifierSchema,
    preset: DrawingPresetSchema,
  }),
  z.strictObject({
    op: z.literal("setWordDrawingText"),
    nodeId: IdentifierSchema,
    text: WordDrawingTextSchema,
  }),
  z.strictObject({ op: z.literal("setWordDrawingConnector"), nodeId: IdentifierSchema, connector }),
  z.strictObject({ op: z.literal("removeWordDrawing"), nodeId: IdentifierSchema }),
  z.strictObject({
    op: z.literal("copyWordDrawing"),
    nodeId: IdentifierSchema,
    offsetX: z.number().optional(),
    offsetY: z.number().optional(),
  }),
  z.strictObject({
    op: z.literal("groupWordDrawings"),
    nodeId: IdentifierSchema,
    nodeIds: z.array(IdentifierSchema).min(2),
  }),
  z.strictObject({ op: z.literal("ungroupWordDrawing"), nodeId: IdentifierSchema }),
]);
export type WordDrawingOperation = z.infer<typeof WordDrawingOperationSchema>;
