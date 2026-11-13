import { z } from 'zod';
const index=z.number().int().nonnegative().safe(),count=z.number().int().positive().safe();
/** Shared HTML table changes; package budgets bound the actual occupied grid. */
export const TableChangeSchema=z.discriminatedUnion('kind',[
 z.strictObject({kind:z.enum(['insertRows','deleteRows','insertColumns','deleteColumns']),at:index,count}),
 z.strictObject({kind:z.literal('mergeCells'),row:index,column:index,rows:count,columns:count}),
 z.strictObject({kind:z.literal('splitCell'),row:index,column:index}),
 z.strictObject({kind:z.literal('resize'),grid:z.array(z.number().positive()).min(1).optional(),rowHeights:z.array(z.number().positive()).min(1).optional()}),
 z.strictObject({kind:z.literal('setCells'),row:index,column:index,values:z.array(z.array(z.string()).min(1)).min(1)}),
]);
export type TableChange=z.infer<typeof TableChangeSchema>;
