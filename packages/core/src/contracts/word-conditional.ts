import { z } from "zod";
import { IdentifierSchema } from "./primitives.js";
/** Native optimization flags preserve authored presence, including explicit OFF. */
export const WordConditionalFormatValueSchema = z.strictObject({
  val: z
    .string()
    .regex(/^[01]{12}$/)
    .optional(),
  firstRow: z.boolean().optional(),
  lastRow: z.boolean().optional(),
  firstColumn: z.boolean().optional(),
  lastColumn: z.boolean().optional(),
  oddVBand: z.boolean().optional(),
  evenVBand: z.boolean().optional(),
  oddHBand: z.boolean().optional(),
  evenHBand: z.boolean().optional(),
  firstRowFirstColumn: z.boolean().optional(),
  firstRowLastColumn: z.boolean().optional(),
  lastRowFirstColumn: z.boolean().optional(),
  lastRowLastColumn: z.boolean().optional(),
});
export type WordConditionalFormatValue = z.infer<typeof WordConditionalFormatValueSchema>;
export const WordConditionalFormatSchema = z.strictObject({
  nodeId: IdentifierSchema,
  format: WordConditionalFormatValueSchema,
});
