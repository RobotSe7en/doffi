import {
  WordConditionalFormatSchema,
  WordConditionalFormatValueSchema,
} from "./word-conditional.js";
import { z } from "zod";
import { IdentifierSchema, PackagePathSchema } from "./primitives.js";
import { WordDrawingSchema } from "./word-drawing.js";
import {
  WordNoteSchema,
  WordNoteReferenceSchema,
  WordNoteOptionsSchema,
} from "./word-reference.js";
import { WordEquationSchema } from "./word-math.js";
import { WordControlSchema } from "./word-control.js";
import { WordCommentSchema, WordPersonSchema } from "./word-comment.js";
import { WordRevisionSchema, WordMoveRangeSchema } from "./word-revision.js";

import { WordTableEmptyParagraphSchema } from "./word-table-continuation.js";
import { WordBibliographySchema } from "../word-bibliography.js";

export const WordTableRegionSchema = z.enum([
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
  identities: z
    .array(
      z.strictObject({
        nodeId: IdentifierSchema,
        sourcePart: PackagePathSchema,
        sourceId: z.string().regex(/^\d+$/),
        kind: z.enum(["frame", "object"]),
      }),
    )
    .optional(),
  fragments: z.array(
    z.strictObject({
      nativeId: IdentifierSchema,
      nodeId: IdentifierSchema.optional(),
      kind: z.enum(["inline", "block", "property"]),
      runChild: z.boolean().optional(),
      /** Zero-width preserved boundary, rather than a visible native object. */
      rangeBoundary: z.boolean().optional(),
      /** Ordered native range controls, indexed by the Office adapter. */
      rangeBoundaries: z
        .array(
          z.strictObject({
            rangeId: IdentifierSchema,
            family: z.string().min(1),
            sourceId: z.string().min(1),
            side: z.enum(["start", "end"]),
            active: z.boolean().optional(),
          }),
        )
        .optional(),
      preview: z
        .strictObject({
          assetPath: PackagePathSchema,
          width: z.number().positive().optional(),
          height: z.number().positive().optional(),
        })
        .optional(),
      sourcePart: PackagePathSchema,
      assetPath: PackagePathSchema,
      fragmentKey: IdentifierSchema.optional(),
      fragmentFormat: z.enum(["xml", "json"]).optional(),
      relationshipIds: z.array(z.string().min(1)).optional(),
      /** Indirect OPC dependencies, such as a content control's custom XML item. */
      dependencyParts: z.array(PackagePathSchema).optional(),
      styleReferences: z.record(IdentifierSchema, IdentifierSchema).optional(),
      numberingReferences: z.record(IdentifierSchema, IdentifierSchema).optional(),
      numberingKind: z.enum(["num", "abstractNum", "numPicBullet"]).optional(),
      numberingSourceId: z.string().regex(/^\d+$/).optional(),
      propertyPath: z.array(z.string().min(1)),
      propertyIndices: z.array(z.number().int().nonnegative()).optional(),
      propertyIndex: z.number().int().nonnegative().optional(),
      ownerName: z.string().min(1),
      styleId: IdentifierSchema.optional(),
      styleRegion: z.string().min(1).optional(),
      label: z.string(),
      propertyName: z.string().min(1).optional(),
      revisionSourceId: IdentifierSchema.optional(),
      revisionPrevious: z.boolean().optional(),
      currentRunPropertiesOnly: z.boolean().optional(),
      insertionSourceId: z.string().min(1).optional(),
    }),
  ),
  parts: z.array(
    z.strictObject({
      sourcePart: PackagePathSchema,
      assetPath: PackagePathSchema,
      contentType: z.string().min(1),
      customXml: z
        .strictObject({
          storeItemId: z.string().min(1),
          propertiesPart: PackagePathSchema,
          copiedStoreItemId: z.string().min(1).optional(),
        })
        .optional(),
    }),
  ),
  relationships: z.array(
    z.strictObject({
      sourcePart: z.string(),
      id: z.string().min(1),
      type: z.string().min(1),
      target: z.string().min(1),
      external: z.boolean(),
      root: z.boolean().optional(),
    }),
  ),
});
export type WordNativeContent = z.infer<typeof WordNativeContentSchema>;
export const WordNativeFragmentCollectionSchema = z.strictObject({
  format: z.literal("doffi-word-native-fragments"),
  version: z.literal(1),
  fragments: z.record(IdentifierSchema, z.string()),
});

export const WordStorySchema = z.strictObject({
  storyId: IdentifierSchema,
  kind: z.enum(["main", "header", "footer", "comment", "textBox", "footnote", "endnote"]),
  rootNodeId: IdentifierSchema,
});
const WordStyleComponentsSchema = z.strictObject({
  p: IdentifierSchema.optional(),
  r: IdentifierSchema.optional(),
  table: IdentifierSchema.optional(),
  cell: IdentifierSchema.optional(),
  row: IdentifierSchema.optional(),
});
export const WordStyleSchema = z.strictObject({
  name: z.string().optional(),
  aliases: z.string().optional(),
  default: z.boolean().optional(),
  customStyle: z.boolean().optional(),
  hidden: z.boolean().optional(),
  semiHidden: z.boolean().optional(),
  unhideWhenUsed: z.boolean().optional(),
  qFormat: z.boolean().optional(),
  locked: z.boolean().optional(),
  uiPriority: z.number().int().nonnegative().max(65535).optional(),
  outlineLevel: z.number().int().min(0).max(9).optional(),
  styleId: IdentifierSchema,
  kind: z.enum(["paragraph", "character", "table", "numbering"]),
  cssClass: IdentifierSchema,
  runCssClass: IdentifierSchema.optional(),
  components: WordStyleComponentsSchema.optional(),
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
        components: WordStyleComponentsSchema.optional(),
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
  pictureBulletAssetPath: PackagePathSchema.optional(),
  pictureBulletWidthPt: z.number().positive().max(144).optional(),
  pictureBulletHeightPt: z.number().positive().max(144).optional(),
  pStyle: IdentifierSchema.optional(),
  isLgl: z.boolean().optional(),
});
const WordKinsokuRuleSchema = z.strictObject({
  language: z
    .string()
    .min(1)
    .max(128)
    .regex(/^[A-Za-z0-9]+(?:-[A-Za-z0-9]+)*$/),
  characters: z
    .string()
    .regex(/^[^\u0000-\u0008\u000b\u000c\u000e-\u001f\ud800-\udfff\ufffe\uffff]*$/u),
});
export const WordTypographySchema = z.strictObject({
  characterSpacingControl: z
    .enum(["doNotCompress", "compressPunctuation", "compressPunctuationAndJapaneseKana"])
    .optional(),
  strictFirstAndLastChars: z.boolean().optional(),
  noLineBreaksAfter: WordKinsokuRuleSchema.optional(),
  noLineBreaksBefore: WordKinsokuRuleSchema.optional(),
});
export type WordTypography = z.infer<typeof WordTypographySchema>;
export const WordPageBackgroundSchema = z.strictObject({
  color: z.string().regex(/^[a-fA-F0-9]{6}$/),
});
export type WordPageBackground = z.infer<typeof WordPageBackgroundSchema>;
export const WordProtectionSchema = z.strictObject({
  edit: z.enum(["none", "readOnly", "comments", "trackedChanges", "forms"]),
  enforcement: z.boolean(),
});
export type WordProtection = z.infer<typeof WordProtectionSchema>;
export const WordThemeSchema = z.strictObject({
  name: z.string().min(1).max(128).optional(),
  colorSchemeName: z.string().min(1).max(128).optional(),
  colors: z
    .strictObject({
      dk1: z.string().regex(/^[a-fA-F0-9]{6}$/).optional(),
      lt1: z.string().regex(/^[a-fA-F0-9]{6}$/).optional(),
      dk2: z.string().regex(/^[a-fA-F0-9]{6}$/).optional(),
      lt2: z.string().regex(/^[a-fA-F0-9]{6}$/).optional(),
      accent1: z.string().regex(/^[a-fA-F0-9]{6}$/).optional(),
      accent2: z.string().regex(/^[a-fA-F0-9]{6}$/).optional(),
      accent3: z.string().regex(/^[a-fA-F0-9]{6}$/).optional(),
      accent4: z.string().regex(/^[a-fA-F0-9]{6}$/).optional(),
      accent5: z.string().regex(/^[a-fA-F0-9]{6}$/).optional(),
      accent6: z.string().regex(/^[a-fA-F0-9]{6}$/).optional(),
      hlink: z.string().regex(/^[a-fA-F0-9]{6}$/).optional(),
      folHlink: z.string().regex(/^[a-fA-F0-9]{6}$/).optional(),
    })
    .optional(),
  fontSchemeName: z.string().min(1).max(128).optional(),
  majorLatin: z.string().min(1).max(128).optional(),
  minorLatin: z.string().min(1).max(128).optional(),
  majorEastAsia: z.string().min(1).max(128).optional(),
  minorEastAsia: z.string().min(1).max(128).optional(),
});
export type WordTheme = z.infer<typeof WordThemeSchema>;
export const WordDocumentPropertiesSchema = z.strictObject({
  creator: z.string().max(512).optional(),
  subject: z.string().max(512).optional(),
  keywords: z.string().max(512).optional(),
  description: z.string().max(2048).optional(),
  category: z.string().max(256).optional(),
});
export type WordDocumentProperties = z.infer<typeof WordDocumentPropertiesSchema>;
export const WordTemplateSchema = z.strictObject({
  name: z.string().min(1).max(256).optional(),
  target: z.string().min(1).max(1024).optional(),
  linkStyles: z.boolean().optional(),
});
export type WordTemplate = z.infer<typeof WordTemplateSchema>;
export const WordBuildingBlockSchema = z.strictObject({
  id: IdentifierSchema,
  name: z.string().min(1).max(128),
  gallery: z.enum(["autoText", "quickParts", "coverPages", "headers", "footers"]),
  category: z.string().max(128).optional(),
  description: z.string().max(512).optional(),
  html: z.string().min(1).max(65536),
});
export type WordBuildingBlock = z.infer<typeof WordBuildingBlockSchema>;
export const WordMetadataSchema = z.strictObject({
  conditionalFormats: z.array(WordConditionalFormatSchema).optional(),
  compatibilityMode: z.number().int().min(1).max(2147483647).optional(),
  pageBackground: WordPageBackgroundSchema.nullable().optional(),
  protection: WordProtectionSchema.nullable().optional(),
  theme: WordThemeSchema.optional(),
  properties: WordDocumentPropertiesSchema.optional(),
  bibliography: WordBibliographySchema.optional(),
  template: WordTemplateSchema.nullable().optional(),
  buildingBlocks: z.array(WordBuildingBlockSchema).max(256).optional(),
  bordersDoNotSurroundHeader: z.boolean().optional(),
  bordersDoNotSurroundFooter: z.boolean().optional(),
  comments: z.array(WordCommentSchema).optional(),
  people: z.array(WordPersonSchema).optional(),
  trackChanges: z.boolean().optional(),
  trackingAuthor: z.string().optional(),
  moves: z.array(WordMoveRangeSchema).optional(),
  controls: z.array(WordControlSchema).optional(),
  equations: z.array(WordEquationSchema).optional(),
  notes: z.array(WordNoteSchema).optional(),
  noteReferences: z.array(WordNoteReferenceSchema).optional(),
  footnoteOptions: WordNoteOptionsSchema.optional(),
  endnoteOptions: WordNoteOptionsSchema.optional(),
  drawings: z.array(WordDrawingSchema).optional(),
  nativeContent: WordNativeContentSchema.optional(),
  stories: z.array(WordStorySchema).min(1),
  mirrorMargins: z.boolean().optional(),
  evenAndOddHeaders: z.boolean().optional(),
  automaticHyphenation: z.boolean().optional(),
  typography: WordTypographySchema.optional(),
  doNotWrapTextWithPunct: z.boolean().optional(),
  noColumnBalance: z.boolean().optional(),
  adjustLineHeightInTable: z.boolean().optional(),
  defaultRunCssClass: IdentifierSchema.optional(),
  styles: z.array(WordStyleSchema),
  numbering: z.array(
    z.strictObject({
      numberingId: IdentifierSchema,
      name: z.string().max(128).optional(),
      multiLevelType: z.enum(["singleLevel", "multilevel", "hybridMultilevel"]).optional(),
      styleLink: IdentifierSchema.optional(),
      numStyleLink: IdentifierSchema.optional(),
      levels: z.array(WordNumberLevelSchema).min(1).max(9),
    }),
  ),
  hyperlinks: z.array(
    z.strictObject({
      nodeId: IdentifierSchema,
      endNodeId: IdentifierSchema.optional(),
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
          leader: z.enum(["none", "dot", "hyphen", "underscore", "heavy", "middleDot"]),
        }),
      ),
    }),
  ),
  revisions: z.array(WordRevisionSchema),
  tables: z.array(
    z.strictObject({
      nodeId: IdentifierSchema,
      grid: z.array(z.number().positive()).min(1),
      continuationFormats: z
        .array(
          z.strictObject({
            row: z.number().int().nonnegative(),
            column: z.number().int().nonnegative(),
            cellCssClass: IdentifierSchema.optional(),
            textDirection: z.enum(["lrTb", "tbRl", "btLr"]).optional(),
            format: WordConditionalFormatValueSchema.optional(),
            paragraphs: z.array(WordConditionalFormatValueSchema.nullable()).min(1).optional(),
            emptyParagraphs: z.array(WordTableEmptyParagraphSchema).min(1).optional(),
          }),
        )
        .optional(),
      styleContexts: z
        .array(
          z.strictObject({
            nodeId: IdentifierSchema,
            cssClass: IdentifierSchema,
            kind: z.enum(["table", "row", "cell", "paragraph", "run"]),
          }),
        )
        .optional(),
      lookMask: z.number().int().min(0).max(65535).optional(),
      rowBandSize: z.number().int().nonnegative().optional(),
      columnBandSize: z.number().int().nonnegative().optional(),
    }),
  ),
});
export type WordMetadata = z.infer<typeof WordMetadataSchema>;
export type WordNumberLevel = z.infer<typeof WordNumberLevelSchema>;
