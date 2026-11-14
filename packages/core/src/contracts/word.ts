import {WordConditionalFormatSchema,WordConditionalFormatValueSchema} from './word-conditional.js';
import { z } from "zod";
import { IdentifierSchema, PackagePathSchema } from "./primitives.js";
import { WordDrawingSchema } from './word-drawing.js';
import { WordNoteSchema,WordNoteReferenceSchema,WordNoteOptionsSchema } from './word-reference.js';
import { WordEquationSchema } from './word-math.js';
import { WordControlSchema } from './word-control.js';
import { WordCommentSchema,WordPersonSchema } from './word-comment.js';
import {WordRevisionSchema,WordMoveRangeSchema}from'./word-revision.js';

import { WordTableEmptyParagraphSchema } from './word-table-continuation.js';

export const WordTableRegionSchema=z.enum([
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
        ]);

/** Residual OOXML owns only content that has no editable HTML representation. */
export const WordNativeContentSchema = z.strictObject({
  mainPart: PackagePathSchema,
  identities:z.array(z.strictObject({nodeId:IdentifierSchema,sourcePart:PackagePathSchema,sourceId:z.string().regex(/^\d+$/),kind:z.enum(["frame","object"])})).optional(),
  fragments: z.array(z.strictObject({
    nativeId: IdentifierSchema,
    nodeId: IdentifierSchema.optional(),
    kind: z.enum(["inline", "block", "property"]),
    runChild: z.boolean().optional(),
    preview:z.strictObject({assetPath:PackagePathSchema,width:z.number().positive().optional(),height:z.number().positive().optional()}).optional(),
    sourcePart: PackagePathSchema,
    assetPath: PackagePathSchema,
    fragmentKey: IdentifierSchema.optional(),
    fragmentFormat:z.enum(['xml','json']).optional(),relationshipIds:z.array(z.string().min(1)).optional(),
    propertyPath: z.array(z.string().min(1)),
    propertyIndices:z.array(z.number().int().nonnegative()).optional(),propertyIndex:z.number().int().nonnegative().optional(),
    ownerName: z.string().min(1),
    styleId: IdentifierSchema.optional(),
    styleRegion:z.string().min(1).optional(),
    label: z.string(),
    propertyName:z.string().min(1).optional(),revisionSourceId:IdentifierSchema.optional(),revisionPrevious:z.boolean().optional(),
  })),
  parts: z.array(z.strictObject({ sourcePart: PackagePathSchema, assetPath: PackagePathSchema, contentType: z.string().min(1) })),
  relationships: z.array(z.strictObject({
    sourcePart: z.string(), id: z.string().min(1), type: z.string().min(1), target: z.string().min(1), external: z.boolean(), root: z.boolean().optional(),
  })),
});
export type WordNativeContent = z.infer<typeof WordNativeContentSchema>;
export const WordNativeFragmentCollectionSchema=z.strictObject({format:z.literal('doffi-word-native-fragments'),version:z.literal(1),fragments:z.record(IdentifierSchema,z.string())});

export const WordStorySchema = z.strictObject({
  storyId: IdentifierSchema,
  kind: z.enum(["main", "header", "footer", "comment", "textBox",'footnote','endnote']),
  rootNodeId: IdentifierSchema,
});
const WordStyleComponentsSchema=z.strictObject({p:IdentifierSchema.optional(),r:IdentifierSchema.optional(),table:IdentifierSchema.optional(),cell:IdentifierSchema.optional(),row:IdentifierSchema.optional()});
export const WordStyleSchema = z.strictObject({
  name:z.string().optional(),aliases:z.string().optional(),default:z.boolean().optional(),customStyle:z.boolean().optional(),hidden:z.boolean().optional(),semiHidden:z.boolean().optional(),unhideWhenUsed:z.boolean().optional(),qFormat:z.boolean().optional(),locked:z.boolean().optional(),uiPriority:z.number().int().nonnegative().max(65535).optional(),outlineLevel:z.number().int().min(0).max(9).optional(),
  styleId: IdentifierSchema,
  kind: z.enum(["paragraph", "character", "table", "numbering"]),
  cssClass: IdentifierSchema,
  runCssClass: IdentifierSchema.optional(),
  components:WordStyleComponentsSchema.optional(),
  paragraphProperties: z.array(IdentifierSchema).optional(),
  runProperties: z.array(IdentifierSchema).optional(),
  basedOn: IdentifierSchema.optional(),
  linkedStyleId: IdentifierSchema.optional(),
  nextStyleId: IdentifierSchema.optional(),
  rowBandSize: z.number().int().nonnegative().optional(),
  columnBandSize: z.number().int().nonnegative().optional(),
  tableLookMask: z.number().int().min(0).max(65535).optional(),
  conditions: z
    .array(
      z.strictObject({
        region: WordTableRegionSchema,
        cssClass: IdentifierSchema,
        components:WordStyleComponentsSchema.optional(),
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
  conditionalFormats:z.array(WordConditionalFormatSchema).optional(),
    bordersDoNotSurroundHeader:z.boolean().optional(),bordersDoNotSurroundFooter:z.boolean().optional(),
  comments:z.array(WordCommentSchema).optional(),people:z.array(WordPersonSchema).optional(),
  trackChanges:z.boolean().optional(),trackingAuthor:z.string().optional(),
  moves:z.array(WordMoveRangeSchema).optional(),
  controls:z.array(WordControlSchema).optional(),
  equations:z.array(WordEquationSchema).optional(),
  notes:z.array(WordNoteSchema).optional(),noteReferences:z.array(WordNoteReferenceSchema).optional(),
  footnoteOptions:WordNoteOptionsSchema.optional(),endnoteOptions:WordNoteOptionsSchema.optional(),
  drawings:z.array(WordDrawingSchema).optional(),
  nativeContent: WordNativeContentSchema.optional(),
  stories: z.array(WordStorySchema).min(1),
  mirrorMargins:z.boolean().optional(),
  evenAndOddHeaders: z.boolean().optional(),
  automaticHyphenation: z.boolean().optional(),
    noColumnBalance:z.boolean().optional(),
    adjustLineHeightInTable:z.boolean().optional(),
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
  revisions: z.array(WordRevisionSchema),
  tables: z.array(
    z.strictObject({
      nodeId: IdentifierSchema,
      grid: z.array(z.number().positive()).min(1),
      continuationFormats:z.array(z.strictObject({row:z.number().int().nonnegative(),column:z.number().int().nonnegative(),format:WordConditionalFormatValueSchema.optional(),paragraphs:z.array(WordConditionalFormatValueSchema.nullable()).min(1).optional(),emptyParagraphs:z.array(WordTableEmptyParagraphSchema).min(1).optional()})).optional(),
      styleContexts:z.array(z.strictObject({nodeId:IdentifierSchema,cssClass:IdentifierSchema,kind:z.enum(['table','row','cell','paragraph','run'])})).optional(),
      lookMask: z.number().int().min(0).max(65535).optional(),
      rowBandSize: z.number().int().nonnegative().optional(),
      columnBandSize: z.number().int().nonnegative().optional(),
    }),
  ),
});
export type WordMetadata = z.infer<typeof WordMetadataSchema>;
export type WordNumberLevel = z.infer<typeof WordNumberLevelSchema>;
