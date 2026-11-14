import { z } from 'zod';
import { IdentifierSchema, PackagePathSchema, RevisionSchema } from './primitives.js';
import { DrawingGeometrySchema } from './drawing.js';
export const WordControlPropertiesSchema = z.strictObject({
    kind: z.enum(['richText', 'text', 'checkBox', 'dropDown', 'comboBox', 'date', 'picture', 'equation', 'group', 'citation', 'bibliography']), level: z.enum(['inline', 'block']),
    alias: z.string().optional(), tag: z.string().optional(), lock: z.enum(['unlocked', 'sdtLocked', 'contentLocked', 'sdtContentLocked']).optional(), temporary: z.boolean().optional(), placeholder: z.string().optional(), showingPlaceholder: z.boolean().optional(), multiline: z.boolean().optional(),
    checked: z.boolean().optional(), checkedSymbol: z.number().int().min(0).max(65535).optional(), uncheckedSymbol: z.number().int().min(0).max(65535).optional(), checkedFont: z.string().optional(), uncheckedFont: z.string().optional(),
    items: z.array(z.strictObject({ displayText: z.string(), value: z.string() })).optional(), value: z.string().optional(),
    fullDate: z.string().optional(), dateFormat: z.string().optional(), locale: z.string().optional(), calendar: z.string().optional(), dateMapping: z.enum(['text', 'date', 'dateTime']).optional(),
});
export const WordControlSchema = WordControlPropertiesSchema.extend({ nodeId: IdentifierSchema, sourceId: z.string().optional() });
export type WordControl = z.infer<typeof WordControlSchema>;
export const WordControlOperationSchema = z.discriminatedUnion('op', [
    z.strictObject({ op: z.literal('insertWordControl'), nodeId: IdentifierSchema, offset: RevisionSchema.optional(), control: WordControlPropertiesSchema, html: z.string().optional(), text: z.string().optional() }),
    z.strictObject({ op: z.literal('setWordControl'), nodeId: IdentifierSchema, changes: WordControlPropertiesSchema.omit({ level: true }).partial() }),
    z.strictObject({ op: z.literal('setWordControlValue'), nodeId: IdentifierSchema, value: z.union([z.string(), z.boolean()]) }),
    z.strictObject({ op: z.literal('setWordControlPicture'), nodeId: IdentifierSchema, assetPath: PackagePathSchema, alt: z.string().optional(), geometry: DrawingGeometrySchema.optional() }),
    z.strictObject({ op: z.literal('removeWordControl'), nodeId: IdentifierSchema, preserveContent: z.boolean().optional() }),
]);
export type WordControlOperation = z.infer<typeof WordControlOperationSchema>;
