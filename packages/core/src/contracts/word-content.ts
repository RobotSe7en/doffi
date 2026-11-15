import { z } from "zod";
import { WordMetadataSchema } from "./word.js";
/** Source identities are local transport references. The receiving transaction
 * allocates every actual node ID before attaching the semantic graph. */
export const WordInlineContentSemanticsSchema = WordMetadataSchema.pick({
  fields: true,
  hyperlinks: true,
  bookmarks: true,
  equations: true,
  controls: true,
  conditionalFormats: true,
}).partial();
export type WordInlineContentSemantics = z.infer<typeof WordInlineContentSemanticsSchema>;
export const WordContentSemanticsSchema = WordInlineContentSemanticsSchema.extend({
  tables: WordMetadataSchema.shape.tables.optional(),
  tabs: WordMetadataSchema.shape.tabs.optional(),
  conditionalFormats: WordMetadataSchema.shape.conditionalFormats,
  extraClasses: z
    .record(z.string().min(1), z.record(z.string().min(1), z.string().min(1)))
    .optional(),
});
export type WordContentSemantics = z.infer<typeof WordContentSemanticsSchema>;
