import{DocumentError}from'../errors.js';
import type{WordMetadata}from'../contracts/word.js';
import type{WordInlineContentSemantics}from'../contracts/word-content.js';
import type{DocumentEdit}from'../document/edit.js';
import type{ParsedHtml}from'../document/html.js';
import{createOpaqueId}from'../identifiers.js';
import{remapWordFieldBookmarks}from'./word-field-bookmarks.js';
const invalid=(message:string):never=>{throw new DocumentError('InvalidOperation',message);};
export function attachWordContent(word:WordMetadata,edit:DocumentEdit,parse:()=>ParsedHtml,ids:ReadonlyMap<string,string>,semantics:WordInlineContentSemantics):void{
  const map=(id:string):string=>ids.get(id)??invalid('A clipboard semantic reference must belong to the supplied inline source.');
  const usedNames=new Set(word.bookmarks.map(bookmark=>bookmark.name)),names=new Map<string,string>(),newName=():string=>{let name:string;do{name='Copy_'+createOpaqueId().replaceAll('-','');}while(usedNames.has(name));usedNames.add(name);return name;};
  for(const bookmark of semantics.bookmarks??[]){if(names.has(bookmark.name))invalid('Copied bookmarks require unique source names.');const name=usedNames.has(bookmark.name)||!/^[_\p{L}][_\p{L}\p{N}]{0,39}$/u.test(bookmark.name)?newName():bookmark.name;usedNames.add(name);names.set(bookmark.name,name);word.bookmarks.push({...bookmark,bookmarkId:map(bookmark.startNodeId),name,startNodeId:map(bookmark.startNodeId),endNodeId:map(bookmark.endNodeId)});}
  const referenceName=(name:string):string=>{let result=names.get(name);if(!result){result=newName();names.set(name,result);}return result;};
  const inserted=parse();for(const after of ids.values()){const node=inserted.peekNode(after)!;for(const key of ['data-word-field-end','data-word-field-result']){const id=node.attributes[key];if(id)edit.attributes(after,{[key]:map(id)});}}
  word.fields.push(...(semantics.fields??[]).map(field=>{const copy={...field,nodeId:map(field.nodeId),...(field.endNodeId?{endNodeId:map(field.endNodeId)}:{}),...(field.resultNodeIds?{resultNodeIds:field.resultNodeIds.map(map)}:{}),...(field.instructionFields?{instructionFields:field.instructionFields.map(child=>({...child,fieldNodeId:map(child.fieldNodeId)}))}:{})};remapWordFieldBookmarks(copy,referenceName);return copy;}));
  word.hyperlinks.push(...(semantics.hyperlinks??[]).map(link=>({...link,nodeId:map(link.nodeId),...(link.endNodeId?{endNodeId:map(link.endNodeId)}:{}),...(link.bookmark&&!link.target?{bookmark:referenceName(link.bookmark)}:{})})));
  if(semantics.equations?.length)(word.equations??=[]).push(...semantics.equations.map(equation=>({...JSON.parse(JSON.stringify(equation))as typeof equation,nodeId:map(equation.nodeId)})));
  if(semantics.controls?.length)(word.controls??=[]).push(...semantics.controls.map(control=>{const copy={...control,nodeId:map(control.nodeId)};delete copy.sourceId;return copy;}));
}
