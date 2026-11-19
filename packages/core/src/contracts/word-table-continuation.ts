import { z } from "zod";
import { IdentifierSchema } from "./primitives.js";
import { WordConditionalFormatValueSchema } from "./word-conditional.js";
const tabStops = z.array(
  z.strictObject({
    position: z.number(),
    alignment: z.enum(["left", "center", "right", "decimal", "clear"]),
    leader: z.enum(["none", "dot", "hyphen", "underscore", "heavy", "middleDot"]),
  }),
);

/** Empty native continuation paragraphs have no second editable text tree.
 * Their direct formats live in the package stylesheet, like style components. */
export const WordTableEmptyParagraphSchema = z.strictObject({
  styleId: IdentifierSchema.optional(),
  paragraphCssClass: IdentifierSchema.optional(),
  runCssClass: IdentifierSchema.optional(),
  format: WordConditionalFormatValueSchema.optional(),
  outlineLevel: z.number().int().min(0).max(9).optional(),
  numbering: z
    .strictObject({ numberingId: IdentifierSchema, level: z.number().int().min(0).max(8) })
    .optional(),
  numberingDisabled: z.literal(true).optional(),
  tabs: tabStops.optional(),
  resolvedTabs: tabStops.optional(),
  defaultTabInterval: z.number().positive().optional(),
  runs: z.array(
    z.strictObject({ cssClass: IdentifierSchema.optional(), styleId: IdentifierSchema.optional() }),
  ),
});
export type WordTableEmptyParagraph = z.infer<typeof WordTableEmptyParagraphSchema>;
