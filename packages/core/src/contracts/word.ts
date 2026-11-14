import { z } from "zod";
import { IdentifierSchema, PackagePathSchema } from "./primitives.js";
import { WordDrawingSchema } from './word-drawing.js';
import { WordNoteSchema,WordNoteReferenceSchema,WordNoteOptionsSchema } from './word-reference.js';
import { WordEquationSchema } from './word-math.js';

/** Residual OOXML owns only content that has no editable HTML representation. */
export const WordNativeContentSchema = z.strictObject({
  mainPart: PackagePathSchema,
  fragments: z.array(z.strictObject({
    nativeId: IdentifierSchema,
    nodeId: IdentifierSchema.optional(),
    kind: z.enum(["inline", "block", "property"]),
    runChild: z.boolean().optional(),
    sourcePart: PackagePathSchema,
    assetPath: PackagePathSchema,
    fragmentKey: IdentifierSchema.optional(),
    propertyPath: z.array(z.string().min(1)),
    ownerName: z.string().min(1),
    styleId: IdentifierSchema.optional(),
    label: z.string(),
  })),
  parts: z.array(z.strictObject({ sourcePart: PackagePathSchema, assetPath: PackagePathSchema, contentType: z.string().min(1) })),
  relationships: z.array(z.strictObject({
    sourcePart: z.string(), id: z.string().min(1), type: z.string().min(1), target: z.string().min(1), external: z.boolean(), root: z.boolean().optional(),
  })),
});
export type WordNativeContent = z.infer<typeof WordNativeContentSchema>;

export const WordStorySchema = z.strictObject({
  storyId: IdentifierSchema,
  kind: z.enum(["main", "header", "footer", "comment", "textBox",'footnote','endnote']),
  rootNodeId: IdentifierSchema,
});
export const WordStyleSchema = z.strictObject({
  styleId: IdentifierSchema,
  kind: z.enum(["paragraph", "character", "table", "numbering"]),
  cssClass: IdentifierSchema,
  runCssClass: IdentifierSchema.optional(),
  paragraphProperties: z.array(IdentifierSchema).optional(),
  runProperties: z.array(IdentifierSchema).optional(),
  basedOn: IdentifierSchema.optional(),
  linkedStyleId: IdentifierSchema.optional(),
  nextStyleId: IdentifierSchema.optional(),
  rowBandSize: z.number().int().positive().optional(),
  columnBandSize: z.number().int().positive().optional(),
  tableLookMask: z.number().int().min(0).max(65535).optional(),
  conditions: z
    .array(
      z.strictObject({
        region: z.enum([
          "wholeTable",
          "firstRow",
          "lastRow",
          "firstCol",
          "lastCol",
          "band1Vert",
          "band2Vert",
          "band1Horz",
          "band2Horz",
          "neCell",
          "nwCell",
          "seCell",
          "swCell",
        ]),
        cssClass: IdentifierSchema,
      }),
    )
    .optional(),
});
export const WordNumberLevelSchema = z.strictObject({
  level: z.number().int().min(0).max(8),
  start: z.number().int().min(0),
  format: z.enum([
    "decimal",
    "bullet",
    "lowerLetter",
    "upperLetter",
    "lowerRoman",
    "upperRoman",
    "none",
  ]),
  text: z.string(),
  suffix: z.enum(["tab", "space", "nothing"]),
  restart: z.number().int().min(0).max(8).optional(),
  cssClass: IdentifierSchema,
});
export const WordMetadataSchema = z.strictObject({
  equations:z.array(WordEquationSchema).optional(),
  notes:z.array(WordNoteSchema).optional(),noteReferences:z.array(WordNoteReferenceSchema).optional(),
  footnoteOptions:WordNoteOptionsSchema.optional(),endnoteOptions:WordNoteOptionsSchema.optional(),
  drawings:z.array(WordDrawingSchema).optional(),
  nativeContent: WordNativeContentSchema.optional(),
  stories: z.array(WordStorySchema).min(1),
  evenAndOddHeaders: z.boolean().optional(),
  automaticHyphenation: z.boolean().optional(),
  defaultRunCssClass: IdentifierSchema.optional(),
  styles: z.array(WordStyleSchema),
  numbering: z.array(
    z.strictObject({
      numberingId: IdentifierSchema,
      levels: z.array(WordNumberLevelSchema).min(1).max(9),
    }),
  ),
  hyperlinks: z.array(
    z.strictObject({
      nodeId: IdentifierSchema,
      endNodeId:IdentifierSchema.optional(),
      target: z.string().min(1).optional(),
      bookmark: IdentifierSchema.optional(),
    }),
  ),
  bookmarks: z.array(
    z.strictObject({
      bookmarkId: IdentifierSchema,
      name: IdentifierSchema,
      startNodeId: IdentifierSchema,
      endNodeId: IdentifierSchema,
    }),
  ),
  fields: z.array(
    z.strictObject({
      nodeId: IdentifierSchema,
      instruction: z.string(),
      resultState: z.literal("cached"),
      locked: z.boolean().optional(),
      dirty: z.boolean().optional(),
      endNodeId: IdentifierSchema.optional(),
      resultNodeIds: z.array(IdentifierSchema).optional(),
      instructionFields: z
        .array(
          z.strictObject({
            offset: z.number().int().min(0),
            fieldNodeId: IdentifierSchema,
          }),
        )
        .optional(),
    }),
  ),
  tabs: z.array(
    z.strictObject({
      paragraphNodeId: IdentifierSchema,
      defaultInterval: z.number().positive(),
      stops: z.array(
        z.strictObject({
          position: z.number(),
          alignment: z.enum(["left", "center", "right", "decimal", "clear"]),
          leader: z.enum([
            "none",
            "dot",
            "hyphen",
            "underscore",
            "heavy",
            "middleDot",
          ]),
        }),
      ),
    }),
  ),
  revisions: z.array(
    z.strictObject({
      nodeId: IdentifierSchema,
      sourceId: IdentifierSchema,
      kind: z.enum(["insert", "delete"]),
      author: z.string().optional(),
      date: z.string().optional(),
    }),
  ),
  tables: z.array(
    z.strictObject({
      nodeId: IdentifierSchema,
      grid: z.array(z.number().positive()).min(1),
      lookMask: z.number().int().min(0).max(65535).optional(),
      rowBandSize: z.number().int().positive().optional(),
      columnBandSize: z.number().int().positive().optional(),
    }),
  ),
});
export type WordMetadata = z.infer<typeof WordMetadataSchema>;
export type WordNumberLevel = z.infer<typeof WordNumberLevelSchema>;
