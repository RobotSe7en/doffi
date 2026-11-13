import { z } from 'zod';
import { IdentifierSchema,RevisionSchema } from './primitives.js';
import { TableChangeSchema } from "./table-operations.js";
import { WordStyleSchema,WordNumberLevelSchema } from './word.js';
const properties=z.record(z.string().min(1),z.string().min(1));
/** Format operations always name an existing source target for proposal scoping. */
export const WordOperationSchema=z.discriminatedUnion('op',[
  z.strictObject({op:z.literal('insertWordTable'),nodeId:IdentifierSchema,position:z.enum(['before','after']),rows:z.number().int().positive().safe(),columns:z.number().int().positive().safe(),width:z.number().positive().optional()}),
  z.strictObject({op:z.literal('editWordTable'),nodeId:IdentifierSchema,change:TableChangeSchema}),
  z.strictObject({op:z.literal('formatWordTable'),nodeId:IdentifierSchema,properties:properties.optional(),styleId:IdentifierSchema.nullable().optional(),repeatHeaderRows:z.number().int().nonnegative().optional(),allowRowSplit:z.boolean().optional()}),
  z.strictObject({op:z.literal('formatWordCell'),nodeId:IdentifierSchema,properties,textDirection:z.enum(['lrTb','tbRl','btLr']).optional()}),
  z.strictObject({op:z.literal('formatWordText'),nodeId:IdentifierSchema,from:RevisionSchema,to:RevisionSchema,expectedText:z.string(),properties}),
  z.strictObject({op:z.literal('setWordParagraph'),nodeId:IdentifierSchema,properties:properties.optional(),styleId:IdentifierSchema.nullable().optional(),clearDirectFormatting:z.boolean().optional()}),
  z.strictObject({op:z.literal('setWordStyle'),nodeId:IdentifierSchema,style:WordStyleSchema,paragraphProperties:properties.optional(),runProperties:properties.optional()}),
  z.strictObject({op:z.literal('setWordNumbering'),nodeId:IdentifierSchema,numberingId:IdentifierSchema,levels:z.array(WordNumberLevelSchema).min(1).max(9)}),
  z.strictObject({op:z.literal('setWordList'),nodeId:IdentifierSchema,numberingId:IdentifierSchema.nullable(),level:z.number().int().min(0).max(8).optional()}),
]);
export type WordOperation=z.infer<typeof WordOperationSchema>;
