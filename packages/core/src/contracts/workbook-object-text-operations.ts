import { z } from "zod";
import { IdentifierSchema, RevisionSchema } from "./primitives.js";

const range = { nodeId: IdentifierSchema, paragraphId: IdentifierSchema, from: RevisionSchema, to: RevisionSchema, expectedText: z.string().max(2 * 1024 * 1024) };
const properties = z.record(z.string(), z.string().min(1));
/** Every paragraph is scoped to one base-snapshot drawing. Source identities
 * remain core-owned when paragraphs or font wrappers are introduced. */
export const WorkbookObjectTextOperationSchema = z.discriminatedUnion("op", [
  z.strictObject({ op: z.literal("replaceWorkbookObjectText"), ...range, text: z.string().max(2 * 1024 * 1024) }),
  z.strictObject({ op: z.literal("formatWorkbookObjectText"), ...range, properties }),
  z.strictObject({ op: z.literal("insertWorkbookObjectParagraph"), nodeId: IdentifierSchema, beforeParagraphId: IdentifierSchema.optional(), html: z.string().max(2 * 1024 * 1024) }),
  z.strictObject({ op: z.literal("removeWorkbookObjectParagraph"), nodeId: IdentifierSchema, paragraphId: IdentifierSchema }),
  z.strictObject({ op: z.literal("setWorkbookObjectParagraphStyle"), nodeId: IdentifierSchema, paragraphId: IdentifierSchema, properties }),
]);
export type WorkbookObjectTextOperation = z.infer<typeof WorkbookObjectTextOperationSchema>;
