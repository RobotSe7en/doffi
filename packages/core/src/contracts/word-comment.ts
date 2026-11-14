import { z } from 'zod';
import { IdentifierSchema } from './primitives.js';
const durableId = z.string().regex(/^[0-7][0-9A-Fa-f]{7}$/).refine(value => parseInt(value, 16) > 0 && parseInt(value, 16) < 0x7fffffff);
/** Native identities are bindings to the shared annotation thread, not bodies. */
export const WordCommentSchema = z.strictObject({
    annotationId: IdentifierSchema, durableId: durableId.optional(), followUp: z.boolean().optional(),
    replies: z.array(z.strictObject({ replyId: IdentifierSchema, durableId: durableId.optional(), parentReplyId: IdentifierSchema.optional() })),
});
export const WordPersonSchema = z.strictObject({ author: z.string(), presence: z.strictObject({ providerId: z.string(), userId: z.string() }).optional() });
const identity = { author: z.strictObject({ name: z.string(), initials: z.string().optional() }).optional(), date: z.string().optional() };
export const WordCommentOperationSchema = z.discriminatedUnion('op', [
    z.strictObject({ op: z.literal('addWordCommentReply'), nodeId: IdentifierSchema, annotationId: IdentifierSchema, html: z.string().min(1), parentReplyId: IdentifierSchema.optional(), ...identity }),
    z.strictObject({ op: z.literal('removeWordCommentReply'), nodeId: IdentifierSchema, annotationId: IdentifierSchema, replyId: IdentifierSchema }),
    z.strictObject({ op: z.literal('setWordCommentBody'), nodeId: IdentifierSchema, annotationId: IdentifierSchema, replyId: IdentifierSchema.optional(), html: z.string().min(1) }),
    z.strictObject({ op: z.literal('setWordCommentIdentity'), nodeId: IdentifierSchema, annotationId: IdentifierSchema, replyId: IdentifierSchema.optional(), ...identity }),
]);
export type WordCommentOperation = z.infer<typeof WordCommentOperationSchema>;
