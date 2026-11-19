import { z } from "zod";
/** Shared Word font roles keep source typefaces independent of preview fallback. */
export const WordTextStyleSchema = z.strictObject({
  fonts: z
    .strictObject({
      ascii: z.string().optional(),
      hAnsi: z.string().optional(),
      eastAsia: z.string().optional(),
      cs: z.string().optional(),
    })
    .optional(),
  fontSize: z.number().positive().optional(),
  color: z.union([z.literal("auto"), z.string().regex(/^#[a-f0-9]{6}$/i)]).optional(),
  bold: z.boolean().optional(),
  italic: z.boolean().optional(),
  language: z.string().optional(),
});
export type WordTextStyle = z.infer<typeof WordTextStyleSchema>;
