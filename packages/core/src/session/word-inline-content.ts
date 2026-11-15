import{DocumentError}from'../errors.js';
import type{WordMetadata}from'../contracts/word.js';
import type{WordOperation}from'../contracts/word-operations.js';
import type{DocumentEdit}from'../document/edit.js';
import type{ParsedHtml}from'../document/html.js';
import{isTextBoundary}from'../document/text.js';
import{trackWordInlineContent}from'./word-revisions.js';
import{createOpaqueId}from'../identifiers.js';
import{remapWordFieldBookmarks}from'./word-field-bookmarks.js';
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
 const semantics=operation.semantics??{},attach=(ids:ReadonlyMap<string,string>):void=>{
  const map=(id:string):string=>ids.get(id)??invalid('A clipboard semantic reference must belong to the supplied inline source.');
  const usedNames=new Set(word.bookmarks.map(bookmark=>bookmark.name)),names=new Map<string,string>(),newName=():string=>{let name:string;do{name='Copy_'+createOpaqueId().replaceAll('-','');}while(usedNames.has(name));usedNames.add(name);return name;};
  for(const bookmark of semantics.bookmarks??[]){if(names.has(bookmark.name))invalid('Copied bookmarks require unique source names.');const name=usedNames.has(bookmark.name)||!/^[_\p{L}][_\p{L}\p{N}]{0,39}$/u.test(bookmark.name)?newName():bookmark.name;usedNames.add(name);names.set(bookmark.name,name);word.bookmarks.push({...bookmark,bookmarkId:map(bookmark.startNodeId),name,startNodeId:map(bookmark.startNodeId),endNodeId:map(bookmark.endNodeId)});}
  const referenceName=(name:string):string=>{let result=names.get(name);if(!result){result=newName();names.set(name,result);}return result;};
  const inserted=parse();for(const after of ids.values()){const node=inserted.peekNode(after)!;for(const key of ['data-word-field-end','data-word-field-result']){const id=node.attributes[key];if(id)edit.attributes(after,{[key]:map(id)});}}
  word.fields.push(...(semantics.fields??[]).map(field=>{const copy={...field,nodeId:map(field.nodeId),...(field.endNodeId?{endNodeId:map(field.endNodeId)}:{}),...(field.resultNodeIds?{resultNodeIds:field.resultNodeIds.map(map)}:{}),...(field.instructionFields?{instructionFields:field.instructionFields.map(child=>({...child,fieldNodeId:map(child.fieldNodeId)}))}:{})};remapWordFieldBookmarks(copy,referenceName);return copy;}));
  word.hyperlinks.push(...(semantics.hyperlinks??[]).map(link=>({...link,nodeId:map(link.nodeId),...(link.endNodeId?{endNodeId:map(link.endNodeId)}:{}),...(link.bookmark&&!link.target?{bookmark:referenceName(link.bookmark)}:{})})));
  if(semantics.equations?.length)(word.equations??=[]).push(...semantics.equations.map(equation=>({...JSON.parse(JSON.stringify(equation))as typeof equation,nodeId:map(equation.nodeId)})));
  if(semantics.controls?.length)(word.controls??=[]).push(...semantics.controls.map(control=>{const copy={...control,nodeId:map(control.nodeId)};delete copy.sourceId;return copy;}));
 };
 const keepEmpty=Object.values(semantics).some(items=>(items?.length??0)>0);
 if(word.trackChanges){trackWordInlineContent(word,edit,parse,{op:'trackWordInlineContent',nodeId:operation.nodeId,from,to,expectedText,html:operation.html},{inserted:attach,keepEmpty});return;}
 for(const id of edit.wrapSourceText(operation.nodeId,from,to,'span'))edit.removeSource(id);
 attach(edit.insertSourceAt(operation.nodeId,from,operation.html,undefined,true).identities);
}
