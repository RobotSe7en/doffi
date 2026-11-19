import { z } from "zod";
import { IdentifierSchema, RevisionSchema } from "./primitives.js";
export const ReconciledTextEditSchema = z.strictObject({
  nodeId: IdentifierSchema,
  from: RevisionSchema,
  to: RevisionSchema,
  expectedText: z.string(),
  text: z.string(),
});
export type ReconciledTextEdit = z.infer<typeof ReconciledTextEditSchema>;
