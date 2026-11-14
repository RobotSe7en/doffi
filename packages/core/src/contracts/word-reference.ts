import { z } from 'zod';
import { IdentifierSchema,RevisionSchema } from './primitives.js';
export const WordNoteSchema=z.strictObject({noteId:IdentifierSchema,storyId:IdentifierSchema,kind:z.enum(['footnote','endnote']),type:z.enum(['normal','separator','continuationSeparator','continuationNotice']).optional()});
export const WordNoteReferenceSchema=z.strictObject({nodeId:IdentifierSchema,noteId:IdentifierSchema,customMark:z.string().optional()});
export const WordNoteOptionsSchema=z.strictObject({format:z.enum(['decimal','lowerRoman','upperRoman','lowerLetter','upperLetter','chicago']).optional(),start:z.number().int().positive().optional(),restart:z.enum(['continuous','eachSect','eachPage']).optional(),position:z.enum(['pageBottom','beneathText','sectEnd','docEnd']).optional()});
const bookmarkRange={nodeId:IdentifierSchema,from:RevisionSchema,to:RevisionSchema,expectedText:z.string(),name:IdentifierSchema,endNodeId:IdentifierSchema.optional()};
export const WordReferenceOperationSchema=z.discriminatedUnion('op',[
 z.strictObject({op:z.literal('insertWordNote'),nodeId:IdentifierSchema,offset:RevisionSchema,kind:z.enum(['footnote','endnote']),text:z.string(),customMark:z.string().min(1).optional()}),
 z.strictObject({op:z.literal('removeWordNote'),nodeId:IdentifierSchema}),
 z.strictObject({op:z.literal('convertWordNote'),nodeId:IdentifierSchema,kind:z.enum(['footnote','endnote'])}),
 z.strictObject({op:z.literal('setWordNoteOptions'),nodeId:IdentifierSchema,kind:z.enum(['footnote','endnote']),options:WordNoteOptionsSchema}),
 z.strictObject({op:z.literal('setWordNoteSeparator'),nodeId:IdentifierSchema,kind:z.enum(['footnote','endnote']),type:z.enum(['separator','continuationSeparator','continuationNotice']),text:z.string().optional()}),
 z.strictObject({op:z.literal('insertWordBookmark'),...bookmarkRange}),
 z.strictObject({op:z.literal('setWordBookmark'),...bookmarkRange,bookmarkId:IdentifierSchema}),
 z.strictObject({op:z.literal('removeWordBookmark'),nodeId:IdentifierSchema,bookmarkId:IdentifierSchema}),
 z.strictObject({op:z.literal('insertWordLink'),nodeId:IdentifierSchema,from:RevisionSchema,to:RevisionSchema,expectedText:z.string(),target:z.string().optional(),bookmark:IdentifierSchema.optional()}),
 z.strictObject({op:z.literal('setWordLink'),nodeId:IdentifierSchema,target:z.string().optional(),bookmark:IdentifierSchema.optional()}),
 z.strictObject({op:z.literal('removeWordLink'),nodeId:IdentifierSchema}),
]);
export type WordReferenceOperation=z.infer<typeof WordReferenceOperationSchema>;
