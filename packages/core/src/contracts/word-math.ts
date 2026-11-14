import { z } from 'zod';
import { IdentifierSchema, RevisionSchema } from './primitives.js';
export const WordMathStyleSchema = z.strictObject({ fonts: z.strictObject({ ascii: z.string().optional(), hAnsi: z.string().optional(), eastAsia: z.string().optional(), cs: z.string().optional() }).optional(), fontSize: z.number().positive().optional(), color: z.union([z.literal('auto'), z.string().regex(/^#[a-f0-9]{6}$/i)]).optional(), bold: z.boolean().optional(), italic: z.boolean().optional(), language: z.string().optional() });
const base = { mathId: IdentifierSchema, style: WordMathStyleSchema.optional() }, children = z.array(IdentifierSchema), one = children.length(1), two = children.length(2);
/** Flat typed math trees keep JSON Schema nonrecursive. Children are local IDs,
 * validated as one bounded tree; no XML, executable guides or second text cache. */
export const WordMathNodeSchema = z.discriminatedUnion('kind', [
    z.strictObject({ ...base, kind: z.literal('row'), children,argumentSize:z.number().int().min(-2).max(2).optional() }),
    z.strictObject({ ...base, kind: z.literal('text'), text: z.string(), variant: z.enum(['italic', 'normal', 'bold', 'boldItalic', 'doubleStruck', 'script', 'fraktur', 'sansSerif', 'monospace']).optional() }),
    z.strictObject({ ...base, kind: z.literal('fraction'), children: two, mode: z.enum(['bar', 'noBar', 'skewed', 'linear']).optional() }),
    z.strictObject({ ...base, kind: z.literal('radical'), children: children.min(1).max(2), hideDegree: z.boolean().optional() }),
    z.strictObject({ ...base, kind: z.literal('script'), mode: z.enum(['sub', 'sup', 'subsup', 'pre']), children: children.min(2).max(3) }),
    z.strictObject({ ...base, kind: z.literal('fence'), children: children.min(1), open: z.string().optional(), close: z.string().optional(), separator: z.string().optional(), grow: z.boolean().optional() }),
    z.strictObject({ ...base, kind: z.literal('matrix'), children: children.min(1), columnAlignments: z.array(z.enum(['left', 'center', 'right'])).optional(), verticalAlignment: z.enum(['top', 'center', 'bottom']).optional(), hidePlaceholders: z.boolean().optional() }),
    z.strictObject({ ...base, kind: z.literal('matrixRow'), children: children.min(1) }),
    z.strictObject({ ...base, kind: z.literal('nary'), children: children.length(3), operator: z.string(), limits: z.enum(['aboveBelow', 'subSup']).optional(), hideLower: z.boolean().optional(), hideUpper: z.boolean().optional() }),
    z.strictObject({ ...base, kind: z.literal('function'), children: two }),
    z.strictObject({ ...base, kind: z.literal('accent'), children: one, symbol: z.string() }),
    z.strictObject({ ...base, kind: z.literal('bar'), children: one, position: z.enum(['top', 'bottom']) }),
    z.strictObject({ ...base, kind: z.literal('limit'), children: two, position: z.enum(['lower', 'upper']) }),
    z.strictObject({ ...base, kind: z.literal('equationArray'), children: children.min(1) }),
    z.strictObject({ ...base, kind: z.literal('box'), children: one }),
]);
export type WordMathNode = z.infer<typeof WordMathNodeSchema>;
export const WordMathTreeSchema = z.strictObject({ rootMathId: IdentifierSchema, nodes: z.array(WordMathNodeSchema).min(1) });
export type WordMathTree = z.infer<typeof WordMathTreeSchema>;
export const WordEquationSchema = WordMathTreeSchema.extend({ nodeId: IdentifierSchema, display: z.enum(['inline', 'block']), alignment: z.enum(['left', 'right', 'center', 'centerGroup']).optional(), fontFamily: z.string().optional() });
export const WordMathOperationSchema = z.discriminatedUnion('op', [
    z.strictObject({ op: z.literal('insertWordEquation'), nodeId: IdentifierSchema, offset: RevisionSchema, math: WordMathTreeSchema, display: z.enum(['inline', 'block']).optional(), alignment: z.enum(['left', 'right', 'center', 'centerGroup']).optional(), fontFamily: z.string().optional() }),
    z.strictObject({ op: z.literal('setWordEquation'), nodeId: IdentifierSchema, math: WordMathTreeSchema, display: z.enum(['inline', 'block']).optional(), alignment: z.enum(['left', 'right', 'center', 'centerGroup']).optional(), fontFamily: z.string().optional() }),
    z.strictObject({ op: z.literal('removeWordEquation'), nodeId: IdentifierSchema }),
]);
export type WordMathOperation = z.infer<typeof WordMathOperationSchema>;
