import { z } from "zod";
import { WordMetadataSchema } from "./word.js";
import { IdentifierSchema } from "./primitives.js";
/** Source identities are local transport references. The receiving transaction
 * allocates every actual node ID before attaching the semantic graph. Owned
 * stories are transport fragments, serialized into the one document source. */
export const WordContentSemanticsSchema = WordMetadataSchema.pick({
  fields: true,
  hyperlinks: true,
  bookmarks: true,
  equations: true,
  controls: true,
  conditionalFormats: true,
  drawings: true,
  notes: true,
  noteReferences: true,
  numbering: true,
  tables: true,
  tabs: true,
})
  .partial()
  .extend({
    stories: z
      .array(
        z.strictObject({
          storyId: IdentifierSchema,
          rootNodeId: IdentifierSchema,
          kind: z.enum(["textBox", "footnote", "endnote"]),
          html: z.string().min(1),
        }),
      )
      .optional(),
    extraClasses: z
      .record(z.string().min(1), z.record(z.string().min(1), z.string().min(1)))
      .optional(),
  });
export type WordContentSemantics = z.infer<typeof WordContentSemanticsSchema>;
// Inline describes the receiving fragment; its owned stories may contain blocks.
export const WordInlineContentSemanticsSchema = z.union([
  WordContentSemanticsSchema.omit({ numbering: true, tables: true, tabs: true }),
  WordContentSemanticsSchema.extend({
    stories: WordContentSemanticsSchema.shape.stories.unwrap().min(1),
  }),
]);
export type WordInlineContentSemantics = z.infer<typeof WordInlineContentSemanticsSchema>;
