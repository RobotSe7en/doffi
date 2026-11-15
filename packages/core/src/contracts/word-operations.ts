import { z } from 'zod';
import { IdentifierSchema,RevisionSchema,PackagePathSchema } from './primitives.js';
import { DocumentSectionSchema } from "./document.js";
import { TableChangeSchema } from "./table-operations.js";
import { WordStyleSchema,WordTableRegionSchema,WordNumberLevelSchema,WordMetadataSchema } from './word.js';
import { WordDrawingOperationSchema } from './word-drawing.js';
import { WordReferenceOperationSchema } from './word-reference.js';
import { WordMathOperationSchema } from './word-math.js';
import { WordControlOperationSchema } from './word-control.js';
import { WordCommentOperationSchema } from './word-comment.js';
import {WordRevisionOperationSchema}from'./word-revision.js';
const properties=z.record(z.string().min(1),z.string().min(1));
/** Format operations always name an existing source target for proposal scoping. */
export const WordOperationSchema=z.discriminatedUnion('op',[
  z.strictObject({op:z.literal('setWordWatermark'),nodeId:IdentifierSchema,watermark:z.discriminatedUnion('kind',[z.strictObject({kind:z.literal('text'),text:z.string().min(1),fontFamily:z.string().min(1),fontSize:z.number().positive(),color:z.string().regex(/^#[a-fA-F0-9]{6}$/),rotation:z.number().optional()}),z.strictObject({kind:z.literal('image'),assetPath:PackagePathSchema,alt:z.string(),width:z.number().positive(),height:z.number().positive(),rotation:z.number().optional()})]).nullable()}),
  ...WordCommentOperationSchema.options,
  ...WordRevisionOperationSchema.options,
  ...WordDrawingOperationSchema.options,
  ...WordReferenceOperationSchema.options,
  ...WordMathOperationSchema.options,
  ...WordControlOperationSchema.options,
  z.strictObject({op:z.literal('resetWordText'),nodeId:IdentifierSchema,from:RevisionSchema,to:RevisionSchema,expectedText:z.string(),resetParagraph:z.boolean().optional()}),
  z.strictObject({op:z.literal('setWordOptions'),nodeId:IdentifierSchema,bordersDoNotSurroundHeader:z.boolean().optional(),bordersDoNotSurroundFooter:z.boolean().optional(),mirrorMargins:z.boolean().optional(),evenAndOddHeaders:z.boolean().optional(),automaticHyphenation:z.boolean().optional(),noColumnBalance:z.boolean().optional(),adjustLineHeightInTable:z.boolean().optional()}),
  z.strictObject({op:z.literal('setWordTabs'),nodeId:IdentifierSchema,defaultInterval:z.number().positive(),stops:WordMetadataSchema.shape.tabs.element.shape.stops}),
  z.strictObject({op:z.literal('insertWordField'),nodeId:IdentifierSchema,instruction:z.string().min(1),cachedText:z.string().optional(),offset:RevisionSchema.optional()}),
  z.strictObject({op:z.literal('setWordSection'),nodeId:IdentifierSchema,changes:DocumentSectionSchema.omit({sectionId:true,rootNodeId:true}).partial()}),
  z.strictObject({op:z.literal('splitWordSection'),nodeId:IdentifierSchema,breakType:z.enum(['continuous','nextPage','nextColumn','evenPage','oddPage'])}),
  z.strictObject({op:z.literal('insertWordStory'),nodeId:IdentifierSchema,kind:z.enum(['header','footer']),variant:z.enum(['default','first','even']),html:z.string()}),
  z.strictObject({op:z.literal('cloneWordStory'),nodeId:IdentifierSchema,storyId:IdentifierSchema,variant:z.enum(['default','first','even'])}),
  z.strictObject({op:z.literal('setWordStoryContent'),nodeId:IdentifierSchema,html:z.string()}),
  z.strictObject({op:z.literal('insertWordTable'),nodeId:IdentifierSchema,position:z.enum(['before','after']),rows:z.number().int().positive().safe(),columns:z.number().int().positive().safe(),width:z.number().positive().optional()}),
  z.strictObject({op:z.literal('deleteWordTable'),nodeId:IdentifierSchema}),
  z.strictObject({op:z.literal('editWordTable'),nodeId:IdentifierSchema,change:TableChangeSchema}),
  z.strictObject({op:z.literal('formatWordTable'),nodeId:IdentifierSchema,properties:properties.optional(),styleId:IdentifierSchema.nullable().optional(),lookMask:z.number().int().min(0).max(65535).nullable().optional(),rowBandSize:z.number().int().nonnegative().nullable().optional(),columnBandSize:z.number().int().nonnegative().nullable().optional(),repeatHeaderRows:z.number().int().nonnegative().optional(),allowRowSplit:z.boolean().optional(),rowHeightRule:z.enum(['auto','atLeast','exact']).optional()}),
  z.strictObject({op:z.literal('formatWordCell'),nodeId:IdentifierSchema,properties,textDirection:z.enum(['lrTb','tbRl','btLr']).optional()}),
  z.strictObject({op:z.literal('applyWordCharacterStyle'),nodeId:IdentifierSchema,from:RevisionSchema,to:RevisionSchema,expectedText:z.string(),styleId:IdentifierSchema}),
  z.strictObject({op:z.literal('formatWordText'),nodeId:IdentifierSchema,from:RevisionSchema,to:RevisionSchema,expectedText:z.string(),properties}),
  z.strictObject({op:z.literal('setWordParagraph'),nodeId:IdentifierSchema,outlineLevel:z.number().int().min(0).max(9).nullable().optional(),properties:properties.optional(),styleId:IdentifierSchema.nullable().optional(),clearDirectFormatting:z.boolean().optional()}),
  z.strictObject({op:z.literal('deleteWordStyle'),nodeId:IdentifierSchema,styleId:IdentifierSchema,replacementStyleId:IdentifierSchema}),
  z.strictObject({op:z.literal('setWordStyle'),nodeId:IdentifierSchema,style:WordStyleSchema,paragraphProperties:properties.optional(),runProperties:properties.optional(),removeParagraphProperties:z.array(z.string().min(1)).optional(),removeRunProperties:z.array(z.string().min(1)).optional(),tableProperties:properties.optional(),cellProperties:properties.optional(),rowProperties:properties.optional(),tableCondition:z.strictObject({region:WordTableRegionSchema,clear:z.boolean().optional(),paragraphProperties:properties.optional(),runProperties:properties.optional(),tableProperties:properties.optional(),cellProperties:properties.optional(),rowProperties:properties.optional(),removeParagraphProperties:z.array(z.string().min(1)).optional(),removeRunProperties:z.array(z.string().min(1)).optional(),removeTableProperties:z.array(z.string().min(1)).optional(),removeCellProperties:z.array(z.string().min(1)).optional(),removeRowProperties:z.array(z.string().min(1)).optional()}).optional(),removeTableProperties:z.array(z.string().min(1)).optional(),removeCellProperties:z.array(z.string().min(1)).optional(),removeRowProperties:z.array(z.string().min(1)).optional()}),
  z.strictObject({op:z.literal('setWordNumbering'),nodeId:IdentifierSchema,numberingId:IdentifierSchema,copyFromNumberingId:IdentifierSchema.optional(),levels:z.array(WordNumberLevelSchema).min(1).max(9),levelProperties:z.array(z.strictObject({level:z.number().int().min(0).max(8),properties})).optional()}),
  z.strictObject({op:z.literal('setWordList'),nodeId:IdentifierSchema,numberingId:IdentifierSchema.nullable(),level:z.number().int().min(0).max(8).optional()}),
]);
export type WordOperation=z.infer<typeof WordOperationSchema>;
