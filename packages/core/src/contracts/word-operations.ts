import { z } from 'zod';
import { IdentifierSchema,RevisionSchema } from './primitives.js';
import { DocumentSectionSchema } from "./document.js";
import { TableChangeSchema } from "./table-operations.js";
import { WordStyleSchema,WordNumberLevelSchema,WordMetadataSchema } from './word.js';
import { WordDrawingOperationSchema } from './word-drawing.js';
import { WordReferenceOperationSchema } from './word-reference.js';
import { WordMathOperationSchema } from './word-math.js';
const properties=z.record(z.string().min(1),z.string().min(1));
/** Format operations always name an existing source target for proposal scoping. */
export const WordOperationSchema=z.discriminatedUnion('op',[
  ...WordDrawingOperationSchema.options,
  ...WordReferenceOperationSchema.options,
  ...WordMathOperationSchema.options,
  z.strictObject({op:z.literal('resetWordText'),nodeId:IdentifierSchema,from:RevisionSchema,to:RevisionSchema,expectedText:z.string(),resetParagraph:z.boolean().optional()}),
  z.strictObject({op:z.literal('setWordOptions'),nodeId:IdentifierSchema,evenAndOddHeaders:z.boolean().optional(),automaticHyphenation:z.boolean().optional()}),
  z.strictObject({op:z.literal('setWordTabs'),nodeId:IdentifierSchema,defaultInterval:z.number().positive(),stops:WordMetadataSchema.shape.tabs.element.shape.stops}),
  z.strictObject({op:z.literal('insertWordField'),nodeId:IdentifierSchema,instruction:z.string().min(1),cachedText:z.string().optional(),offset:RevisionSchema.optional()}),
  z.strictObject({op:z.literal('setWordSection'),nodeId:IdentifierSchema,changes:DocumentSectionSchema.omit({sectionId:true,rootNodeId:true}).partial()}),
  z.strictObject({op:z.literal('splitWordSection'),nodeId:IdentifierSchema,breakType:z.enum(['continuous','nextPage','evenPage','oddPage'])}),
  z.strictObject({op:z.literal('insertWordStory'),nodeId:IdentifierSchema,kind:z.enum(['header','footer']),variant:z.enum(['default','first','even']),html:z.string()}),
  z.strictObject({op:z.literal('cloneWordStory'),nodeId:IdentifierSchema,storyId:IdentifierSchema,variant:z.enum(['default','first','even'])}),
  z.strictObject({op:z.literal('setWordStoryContent'),nodeId:IdentifierSchema,html:z.string()}),
  z.strictObject({op:z.literal('insertWordTable'),nodeId:IdentifierSchema,position:z.enum(['before','after']),rows:z.number().int().positive().safe(),columns:z.number().int().positive().safe(),width:z.number().positive().optional()}),
  z.strictObject({op:z.literal('editWordTable'),nodeId:IdentifierSchema,change:TableChangeSchema}),
  z.strictObject({op:z.literal('formatWordTable'),nodeId:IdentifierSchema,properties:properties.optional(),styleId:IdentifierSchema.nullable().optional(),repeatHeaderRows:z.number().int().nonnegative().optional(),allowRowSplit:z.boolean().optional()}),
  z.strictObject({op:z.literal('formatWordCell'),nodeId:IdentifierSchema,properties,textDirection:z.enum(['lrTb','tbRl','btLr']).optional()}),
  z.strictObject({op:z.literal('formatWordText'),nodeId:IdentifierSchema,from:RevisionSchema,to:RevisionSchema,expectedText:z.string(),properties}),
  z.strictObject({op:z.literal('setWordParagraph'),nodeId:IdentifierSchema,properties:properties.optional(),styleId:IdentifierSchema.nullable().optional(),clearDirectFormatting:z.boolean().optional()}),
  z.strictObject({op:z.literal('setWordStyle'),nodeId:IdentifierSchema,style:WordStyleSchema,paragraphProperties:properties.optional(),runProperties:properties.optional()}),
  z.strictObject({op:z.literal('setWordNumbering'),nodeId:IdentifierSchema,numberingId:IdentifierSchema,copyFromNumberingId:IdentifierSchema.optional(),levels:z.array(WordNumberLevelSchema).min(1).max(9),levelProperties:z.array(z.strictObject({level:z.number().int().min(0).max(8),properties})).optional()}),
  z.strictObject({op:z.literal('setWordList'),nodeId:IdentifierSchema,numberingId:IdentifierSchema.nullable(),level:z.number().int().min(0).max(8).optional()}),
]);
export type WordOperation=z.infer<typeof WordOperationSchema>;
