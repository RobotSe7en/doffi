import { z } from 'zod';
export const WordPageBorderSchema = z.strictObject({
    style: z.enum(['nil', 'none', 'single', 'double', 'dotted', 'dashed', 'dashSmallGap', 'thick']),
    width: z.number().nonnegative().max(12), space: z.number().int().nonnegative().max(31),
    color: z.string().regex(/^(?:auto|[a-fA-F0-9]{6})$/),
    themeColor: z.enum(['dark1', 'light1', 'dark2', 'light2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hyperlink', 'followedHyperlink', 'background1', 'text1', 'background2', 'text2', 'none']).optional(),
    themeTint: z.string().regex(/^[a-fA-F0-9]{2}$/).optional(), themeShade: z.string().regex(/^[a-fA-F0-9]{2}$/).optional(),
    shadow: z.boolean().optional(), frame: z.boolean().optional(),
});
export const WordPageBordersSchema = z.strictObject({
    offsetFrom: z.enum(['page', 'text']), display: z.enum(['allPages', 'firstPage', 'notFirstPage']), zOrder: z.enum(['front', 'back']),
    top: WordPageBorderSchema.optional(), left: WordPageBorderSchema.optional(), bottom: WordPageBorderSchema.optional(), right: WordPageBorderSchema.optional(),
});
export type WordPageBorder = z.infer<typeof WordPageBorderSchema>;
export type WordPageBorders = z.infer<typeof WordPageBordersSchema>;
