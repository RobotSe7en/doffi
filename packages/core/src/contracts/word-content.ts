import{z}from'zod';
import{WordMetadataSchema}from'./word.js';
/** Source identities are local transport references. The receiving transaction
 * allocates every actual node ID before attaching the semantic graph. */
export const WordInlineContentSemanticsSchema=WordMetadataSchema.pick({fields:true,hyperlinks:true,bookmarks:true,equations:true,controls:true}).partial();
export type WordInlineContentSemantics=z.infer<typeof WordInlineContentSemanticsSchema>;
