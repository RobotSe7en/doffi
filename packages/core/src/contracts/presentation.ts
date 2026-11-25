import { ChartAppearanceSchema } from "./chart.js";
import { z } from "zod";
import { IdentifierSchema, PackagePathSchema } from "./primitives.js";
import {
  DrawingPresetSchema,
  DrawingTextPropertiesSchema,
  DrawingGeometrySchema,
} from "./drawing.js";

export const PresentationSizeSchema = z.strictObject({
  unit: z.literal("pt"),
  width: z.number().positive(),
  height: z.number().positive(),
});
export type PresentationSize = z.infer<typeof PresentationSizeSchema>;
export const PresentationGeometrySchema = DrawingGeometrySchema;
export type PresentationGeometry = z.infer<typeof PresentationGeometrySchema>;
export const PresentationPresetSchema = DrawingPresetSchema;
const placeholder = z.strictObject({
  type: z.enum([
    "title",
    "body",
    "ctrTitle",
    "subTitle",
    "dt",
    "sldNum",
    "ftr",
    "hdr",
    "obj",
    "chart",
    "tbl",
    "clipArt",
    "dgm",
    "media",
    "sldImg",
    "pic",
  ]),
  index: z.number().int().min(0),
});
const objectBase = {
  nodeId: IdentifierSchema,
  name: z.string().optional(),
  nativeId: IdentifierSchema.optional(),
  creationId: IdentifierSchema.optional(),
  placeholder: placeholder.optional(),
};
const textProperties = DrawingTextPropertiesSchema;
const geometryToken = z.string().regex(/^-?\d{1,15}$|^[A-Za-z_][A-Za-z0-9_]{0,63}$/);
const geometryPoint = { x: geometryToken, y: geometryToken };
export const PresentationCustomGeometrySchema = z.strictObject({
  guides: z.array(z.strictObject({ name: geometryToken, operator: z.enum(["val", "*/", "+-"]), args: z.array(geometryToken).min(1).max(3) })).max(128),
  connections: z.array(z.strictObject({ angle: geometryToken, ...geometryPoint })).max(64),
  textRect: z.strictObject({ left: geometryToken, top: geometryToken, right: geometryToken, bottom: geometryToken }).optional(),
  paths: z.array(z.strictObject({
    width: z.number().int().positive().optional(),
    height: z.number().int().positive().optional(),
    fill: z.boolean(),
    stroke: z.boolean(),
    commands: z.array(z.union([z.strictObject({ kind: z.enum(["move", "line"]), ...geometryPoint }), z.strictObject({ kind: z.literal("close") })])).min(1).max(512),
  })).min(1).max(16),
});
export type PresentationCustomGeometry = z.infer<typeof PresentationCustomGeometrySchema>;
export const PresentationObjectSchema = z.discriminatedUnion("kind", [
  z.strictObject({ ...objectBase, kind: z.literal("text"), text: textProperties.optional() }),
  z.strictObject({ ...objectBase, kind: z.literal("image"), customGeometry: PresentationCustomGeometrySchema.optional() }),
  z.strictObject({
    ...objectBase,
    kind: z.literal("geometry"),
    preset: PresentationPresetSchema,
    customGeometry: PresentationCustomGeometrySchema.optional(),
    text: textProperties.optional(),
  }),
  z.strictObject({
    ...objectBase,
    kind: z.literal("connector"),
    routing: z.enum(["straight", "elbow"]),
    startNodeId: IdentifierSchema.optional(),
    endNodeId: IdentifierSchema.optional(),
    startSite: z.number().int().min(0).optional(),
    endSite: z.number().int().min(0).optional(),
  }),
  z.strictObject({
    ...objectBase,
    kind: z.literal("table"),
    grid: z.array(z.number().positive()).min(1),
  }),
  z.strictObject({ ...objectBase, kind: z.literal("chart"), chartId: IdentifierSchema }),
  z.strictObject({ ...objectBase, kind: z.literal("group") }),
]);
export type PresentationObject = z.infer<typeof PresentationObjectSchema>;
export const PresentationDataSetSchema = z.strictObject({
  dataSetId: IdentifierSchema,
  categories: z.array(z.string()),
  series: z
    .array(
      z.strictObject({
        seriesId: IdentifierSchema,
        name: z.string(),
        values: z.array(z.union([z.number(), z.null()])),
      }),
    )
    .min(1),
});
export type PresentationDataSet = z.infer<typeof PresentationDataSetSchema>;
export const PresentationChartSchema = ChartAppearanceSchema.extend({
  chartId: IdentifierSchema,
  nodeId: IdentifierSchema,
  dataSetId: IdentifierSchema,
});
export type PresentationChart = z.infer<typeof PresentationChartSchema>;
export const PresentationMetadataSchema = z.strictObject({
  size: PresentationSizeSchema,
  firstSlideNumber: z.number().int().min(0).max(9999).optional(),
  notesSize: PresentationSizeSchema.optional(),
  commentFormat: z.enum(["classic", "modern"]),
  nativePreview: z.array(z.union([
    z.strictObject({
      feature: z.enum(["theme-defaults", "theme-format", "animation", "transition", "tags", "handout-master", "handout-theme", "table-style"]),
      part: z.string().min(1),
      ownerId: IdentifierSchema,
      xml: z.string().min(1),
    }),
    z.strictObject({
      feature: z.literal("emf-image"),
      part: PackagePathSchema,
      ownerId: IdentifierSchema,
      assetPath: PackagePathSchema,
    }),
  ])).optional(),
  slides: z
    .array(
      z.strictObject({
        slideId: IdentifierSchema,
        rootNodeId: IdentifierSchema,
        layoutId: IdentifierSchema,
        name: z.string().optional(),
        hidden: z.boolean().optional(),
        showMasterShapes: z.boolean().optional(),
      }),
    )
    .min(1),
  themes: z.array(z.strictObject({ themeId: IdentifierSchema, cssClass: IdentifierSchema })).min(1),
  masters: z
    .array(
      z.strictObject({
        masterId: IdentifierSchema,
        rootNodeId: IdentifierSchema,
        themeId: IdentifierSchema,
        titleStyle: IdentifierSchema.optional(),
        bodyStyle: IdentifierSchema.optional(),
        otherStyle: IdentifierSchema.optional(),
      }),
    )
    .min(1),
  layouts: z
    .array(
      z.strictObject({
        layoutId: IdentifierSchema,
        rootNodeId: IdentifierSchema,
        masterId: IdentifierSchema,
        name: z.string().optional(),
        showMasterShapes: z.boolean().optional(),
      }),
    )
    .min(1),
  notesMasters: z.array(
    z.strictObject({
      notesMasterId: IdentifierSchema,
      rootNodeId: IdentifierSchema,
      themeId: IdentifierSchema,
      textStyle: IdentifierSchema.optional(),
    }),
  ),
  stories: z.array(
    z.discriminatedUnion("kind", [
      z.strictObject({
        storyId: IdentifierSchema,
        rootNodeId: IdentifierSchema,
        kind: z.literal("notes"),
        slideId: IdentifierSchema,
        notesMasterId: IdentifierSchema.optional(),
        bodyNodeId: IdentifierSchema.optional(),
      }),
      z.strictObject({
        storyId: IdentifierSchema,
        rootNodeId: IdentifierSchema,
        kind: z.literal("comment"),
        slideId: IdentifierSchema.optional(),
      }),
    ]),
  ),
  objects: z.array(PresentationObjectSchema),
  charts: z.array(PresentationChartSchema),
  dataSets: z.array(PresentationDataSetSchema),
  commentAuthors: z
    .array(
      z.strictObject({
        authorId: IdentifierSchema,
        format: z.enum(["classic", "modern"]),
        nativeId: IdentifierSchema,
        userId: z.string().optional(),
        providerId: z.string().optional(),
      }),
    )
    .optional(),
  comments: z.array(
    z.strictObject({
      annotationId: IdentifierSchema,
      format: z.enum(["classic", "modern"]),
      authorId: IdentifierSchema.optional(),
      recoveryStatus: z.enum(["open", "resolved", "closed"]).optional(),
      replyAuthors: z
        .array(z.strictObject({ replyId: IdentifierSchema, authorId: IdentifierSchema }))
        .optional(),
    }),
  ),
  hyperlinks: z.array(
    z.strictObject({
      nodeId: IdentifierSchema,
      target: z.string().min(1).optional(),
      slideId: IdentifierSchema.optional(),
    }),
  ),
});
export type PresentationMetadata = z.infer<typeof PresentationMetadataSchema>;
export const PresentationDocumentMetadataSchema = z.strictObject({
  title: z.string().optional(),
  locale: IdentifierSchema.optional(),
  presentation: PresentationMetadataSchema,
});
export type PresentationDocumentMetadata = z.infer<typeof PresentationDocumentMetadataSchema> & {
  readonly word?: never;
  readonly workbook?: never;
  readonly pageSetup?: never;
  readonly sections?: never;
};
