import { z } from "zod";
import { AnchorSchema, AnnotationStatusSchema } from "./annotations.js";
import { IdentifierSchema } from "./primitives.js";
import { WorkbookCommentMentionSchema, WorkbookNoteBoxSchema } from "./workbook-comments.js";

const html = z.string().max(2 * 1024 * 1024), author = z.strictObject({ name: z.string().max(32767), initials: z.string().max(32767).optional() }), date = z.string().max(32767);
const anchor = AnchorSchema.options[0];
const textEdit = z.strictObject({ from: z.number().int().nonnegative(), to: z.number().int().nonnegative(), expectedText: z.string(), text: z.string() });
export const WorkbookCommentOperationSchema = z.discriminatedUnion("op", [
  z.strictObject({ op: z.literal("addWorkbookComment"), anchor, html, format: z.enum(["classic", "modern"]).optional(), author: author.optional(), date: date.optional(), box: WorkbookNoteBoxSchema.optional() }),
  z.strictObject({ op: z.literal("addWorkbookCommentReply"), annotationId: IdentifierSchema, html, author: author.optional(), date: date.optional(), parentReplyId: IdentifierSchema.optional() }),
  z.strictObject({ op: z.literal("removeWorkbookCommentReply"), annotationId: IdentifierSchema, replyId: IdentifierSchema }),
  z.strictObject({ op: z.literal("setWorkbookCommentBody"), annotationId: IdentifierSchema, replyId: IdentifierSchema.optional(), html,
    textEdits: z.array(textEdit).optional(), mentions: z.array(WorkbookCommentMentionSchema).optional() }),
  z.strictObject({ op: z.literal("replaceWorkbookCommentText"), annotationId: IdentifierSchema, replyId: IdentifierSchema.optional(),
    ...textEdit.shape, mentions: z.array(WorkbookCommentMentionSchema).optional() }),
  z.strictObject({ op: z.literal("updateWorkbookComment"), annotationId: IdentifierSchema, changes: z.strictObject({ anchor: anchor.optional(), status: AnnotationStatusSchema.optional(),
    format: z.enum(["classic", "modern"]).optional(), author: author.nullable().optional(), date: date.nullable().optional() }) }),
  z.strictObject({ op: z.literal("updateWorkbookCommentReply"), annotationId: IdentifierSchema, replyId: IdentifierSchema,
    changes: z.strictObject({ author: author.nullable().optional(), date: date.nullable().optional(), status: z.enum(["open", "resolved", "closed"]).nullable().optional(), parentReplyId: IdentifierSchema.nullable().optional() }) }),
  z.strictObject({ op: z.literal("setWorkbookNoteBox"), annotationId: IdentifierSchema, box: WorkbookNoteBoxSchema.nullable() }),
]);
export type WorkbookCommentOperation = z.infer<typeof WorkbookCommentOperationSchema>;
