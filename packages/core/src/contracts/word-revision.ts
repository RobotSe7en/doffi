import { z } from 'zod';
import { IdentifierSchema, RevisionSchema } from './primitives.js';
const attributes = z.partialRecord(z.enum(['style', 'class', 'lang', 'data-word-style', 'data-word-line-rule', 'data-word-numbering', 'data-word-level', 'data-word-outline-level', 'data-word-text-direction', 'data-word-repeat-header']), z.string().nullable());
export const WordRevisionSchema = z.strictObject({
    nodeId: IdentifierSchema, sourceId: IdentifierSchema, revisionId: IdentifierSchema.optional(), groupId: IdentifierSchema.optional(),
    kind: z.enum(['insert', 'delete', 'moveFrom', 'moveTo', 'format', 'paragraphInsert', 'paragraphDelete', 'rowInsert', 'rowDelete', 'cellInsert', 'cellDelete']),
    author: z.string().optional(), date: z.string().optional(), moveId: IdentifierSchema.optional(),
    property: z.enum(['run', 'paragraph', 'table', 'row', 'cell', 'section', 'tableGrid']).optional(),
    before: z.strictObject({ attributes: attributes.optional(), tagName: z.enum(['span', 'p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'table', 'tr', 'td', 'th']).optional(), properties: z.record(z.string(), z.string()).optional(), grid: z.array(z.number().positive()).optional(), tabs: z.strictObject({ defaultInterval: z.number().positive(), stops: z.array(z.strictObject({ position: z.number(), alignment: z.enum(['left', 'center', 'right', 'decimal', 'clear']), leader: z.enum(['none', 'dot', 'hyphen', 'underscore', 'heavy', 'middleDot']) })) }).optional() }).optional(),
});
export type WordRevision = z.infer<typeof WordRevisionSchema>;
export const WordRevisionOperationSchema = z.discriminatedUnion('op', [
    z.strictObject({ op: z.literal('setWordReview'), nodeId: IdentifierSchema, trackChanges: z.boolean(), author: z.string().optional() }),
    z.strictObject({ op: z.literal('reviewWordRevision'), nodeId: IdentifierSchema, revisionId: IdentifierSchema, decision: z.enum(['accept', 'reject']) }),
    z.strictObject({ op: z.literal('reviewWordRevisions'), nodeId: IdentifierSchema, decision: z.enum(['accept', 'reject']), author: z.string().optional() }),
    z.strictObject({ op: z.literal('trackWordText'), nodeId: IdentifierSchema, from: RevisionSchema, to: RevisionSchema, expectedText: z.string(), text: z.string(), properties: z.record(z.string(), z.string()).optional(), author: z.string().optional(), date: z.string().optional() }),
]);
export type WordRevisionOperation = z.infer<typeof WordRevisionOperationSchema>;
