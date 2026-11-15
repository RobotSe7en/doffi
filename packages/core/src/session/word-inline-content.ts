import{DocumentError}from'../errors.js';
import type{WordMetadata}from'../contracts/word.js';
import type{WordOperation}from'../contracts/word-operations.js';
import type{DocumentEdit}from'../document/edit.js';
import type{ParsedHtml}from'../document/html.js';
import{isTextBoundary}from'../document/text.js';
import{trackWordInlineContent}from'./word-revisions.js';
import{attachWordContent}from'./word-content.js';
const invalid=(message:string):never=>{throw new DocumentError('InvalidOperation',message);};
/** Attach a clipboard graph only to newly allocated source nodes. Transport
 * identifiers never resolve against existing document nodes or resource paths. */
export function pasteWordInlineContent(word:WordMetadata,edit:DocumentEdit,parse:()=>ParsedHtml,operation:Extract<WordOperation,{op:'pasteWordInlineContent'}>):void{
 const source=parse(),target=source.getNode(operation.nodeId),{from,to,expectedText}=operation;
 if(!target||!/^p$|^h[1-6]$/.test(target.tagName)||from>to||to>target.text.length||target.text.slice(from,to)!==expectedText||!isTextBoundary(target.text,from)||!isTextBoundary(target.text,to))invalid('Inline Word content requires an exact paragraph range.');
 edit.assertInlineSource(operation.nodeId,operation.html);
 const start=source.textOffset(operation.nodeId)!+from,end=source.textOffset(operation.nodeId)!+to;
 const contains=(id:string):boolean=>{for(let current:string|null=id;current;current=source.peekNode(current)?.parentNodeId??null)if(current===operation.nodeId)return true;return false;};
 const protectedIds=[...(word.drawings??[]),...(word.equations??[]),...(word.noteReferences??[]),...(word.controls??[]),...(word.nativeContent?.fragments.filter(item=>item.kind!=='property')??[])].flatMap(item=>item.nodeId?[item.nodeId]:[]);
 if(from!==to)for(const id of protectedIds)if(contains(id)){const begin=source.textOffset(id)!,finish=begin+source.textLength(id)!;if(start<finish&&end>begin)throw new DocumentError('UnsupportedFeature','Replacing this object requires its structural Word operation.');}
 const semantics=operation.semantics??{},attach=(ids:ReadonlyMap<string,string>):void=>attachWordContent(word,edit,parse,ids,semantics);
 const keepEmpty=Object.values(semantics).some(items=>(items?.length??0)>0);
 if(word.trackChanges){trackWordInlineContent(word,edit,parse,{op:'trackWordInlineContent',nodeId:operation.nodeId,from,to,expectedText,html:operation.html},{inserted:attach,keepEmpty});return;}
 for(const id of edit.wrapSourceText(operation.nodeId,from,to,'span'))edit.removeSource(id);
 attach(edit.insertSourceAt(operation.nodeId,from,operation.html,undefined,true).identities);
}
