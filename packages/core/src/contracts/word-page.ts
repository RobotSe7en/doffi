import { IdentifierSchema } from './primitives.js';
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
export const PageSetupSchema = z.strictObject({
    unit: z.literal("pt"),
    width: z.number().positive(),
    height: z.number().positive(),
    margins: z.strictObject({
        top: z.number().min(0),
        right: z.number().min(0),
        bottom: z.number().min(0),
        left: z.number().min(0),
    }),
});
export type PageSetup = z.infer<typeof PageSetupSchema>;
export const DocumentSectionSchema = z.strictObject({
    sectionId: IdentifierSchema,
    rootNodeId: IdentifierSchema,
    pageSetup: PageSetupSchema.optional(),
    breakType: z.enum(["continuous", "nextPage", "evenPage", "oddPage"]).optional(),
    differentFirstPage: z.boolean().optional(),
    columns: z.strictObject({ count: z.number().int().positive(), gap: z.number().nonnegative(), separator: z.boolean().optional(),widths:z.array(z.number().positive()).min(1).optional(),spaces:z.array(z.number().nonnegative()).min(1).optional() }).optional(),
    pageBorders: WordPageBordersSchema.optional(),
    grid: z.strictObject({ type: z.enum(["default", "lines", "linesAndChars", "snapToChars"]), linePitch: z.number().positive().optional(), characterSpacing: z.number().int().optional() }).optional(),
    gutter: z.number().nonnegative().optional(),
    gutterPosition: z.enum(["left", "right", "top"]).optional(),
    pageNumberStart: z.number().int().min(0).optional(),
    pageNumberFormat: z.enum(["decimal", "upperRoman", "lowerRoman", "upperLetter", "lowerLetter"]).optional(),
    headerDistance: z.number().min(0).optional(),
    footerDistance: z.number().min(0).optional(),
    headers: z.strictObject({ default: IdentifierSchema.optional(), first: IdentifierSchema.optional(), even: IdentifierSchema.optional() }).optional(),
    footers: z.strictObject({ default: IdentifierSchema.optional(), first: IdentifierSchema.optional(), even: IdentifierSchema.optional() }).optional(),
});
export type DocumentSection = z.infer<typeof DocumentSectionSchema>;
