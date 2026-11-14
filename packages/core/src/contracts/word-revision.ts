import{ReconciledTextEditSchema}from'./text-edit.js';
import { z } from 'zod';
import { IdentifierSchema, RevisionSchema } from './primitives.js';
import{DocumentSectionSchema}from'./word-page.js';
const attributes = z.partialRecord(z.enum(['style', 'class', 'lang', 'data-word-style', 'data-word-line-rule', 'data-word-numbering', 'data-word-level', 'data-word-outline-level', 'data-word-text-direction', 'data-word-repeat-header']), z.string().nullable());
export const WordRevisionSchema = z.strictObject({
    nodeId: IdentifierSchema, sourceId: IdentifierSchema, revisionId: IdentifierSchema.optional(), groupId: IdentifierSchema.optional(),
    kind: z.enum(['insert', 'delete', 'moveFrom', 'moveTo', 'format', 'paragraphInsert', 'paragraphDelete','paragraphMoveFrom','paragraphMoveTo', 'rowInsert', 'rowDelete', 'cellInsert', 'cellDelete']),
    author: z.string().optional(), date: z.string().optional(), moveId: IdentifierSchema.optional(),
    property: z.enum(['run', 'paragraph', 'table', 'row', 'cell', 'section', 'tableGrid']).optional(),
    before: z.strictObject({tableStyleOptions:z.strictObject({lookMask:z.number().int().min(0).max(65535).optional(),rowBandSize:z.number().int().nonnegative().optional(),columnBandSize:z.number().int().nonnegative().optional()}).optional(),section:DocumentSectionSchema.omit({sectionId:true,rootNodeId:true}).optional(),inheritNativeProperties:z.boolean().optional(), attributes: attributes.optional(), tagName: z.enum(['span', 'p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'table', 'tr', 'td', 'th','li']).optional(), properties: z.record(z.string(), z.string()).optional(), grid: z.array(z.number().positive()).optional(), tabs: z.strictObject({ defaultInterval: z.number().positive(), stops: z.array(z.strictObject({ position: z.number(), alignment: z.enum(['left', 'center', 'right', 'decimal', 'clear']), leader: z.enum(['none', 'dot', 'hyphen', 'underscore', 'heavy', 'middleDot']) })) }).optional() }).optional(),
});
export type WordRevision = z.infer<typeof WordRevisionSchema>;
export const WordMoveRangeSchema=z.strictObject({moveId:IdentifierSchema,kind:z.enum(['from','to']),sourceId:IdentifierSchema,startNodeId:IdentifierSchema,endNodeId:IdentifierSchema,author:z.string(),date:z.string(),columnFirst:z.number().int().nonnegative().safe().optional(),columnLast:z.number().int().nonnegative().safe().optional(),startDisplaced:z.enum(['prev','next']).optional(),endDisplaced:z.enum(['prev','next']).optional()});
export type WordMoveRange=z.infer<typeof WordMoveRangeSchema>;
export const WordRevisionOperationSchema = z.discriminatedUnion('op', [
    z.strictObject({op:z.literal('reconcileWordTracked'),nodeId:IdentifierSchema,html:z.string(),textEdits:z.array(ReconciledTextEditSchema),author:z.string().optional(),date:z.string().optional()}),
    z.strictObject({ op: z.literal('setWordReview'), nodeId: IdentifierSchema, trackChanges: z.boolean(), author: z.string().optional() }),
    z.strictObject({ op: z.literal('reviewWordRevision'), nodeId: IdentifierSchema, revisionId: IdentifierSchema, decision: z.enum(['accept', 'reject']) }),
    z.strictObject({ op: z.literal('reviewWordRevisions'), nodeId: IdentifierSchema, decision: z.enum(['accept', 'reject']), author: z.string().optional() }),
    z.strictObject({ op: z.literal('trackWordText'), nodeId: IdentifierSchema, from: RevisionSchema, to: RevisionSchema, expectedText: z.string(), text: z.string(), properties: z.record(z.string(), z.string()).optional(), author: z.string().optional(), date: z.string().optional() }),
]);
export type WordRevisionOperation = z.infer<typeof WordRevisionOperationSchema>;
