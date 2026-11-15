import{WordInlineContentSemanticsSchema}from'@doffi/core';
import type{ContentDocumentModel,WordInlineContentSemantics}from'@doffi/core';
const text=(node:Node):string=>node.nodeType===3?node.textContent??'':node.nodeType===1&&(node as Element).localName==='img'?'\ufffc':node.nodeType===1&&(node as Element).localName==='br'?'\n':[...node.childNodes].map(text).join('');
/** Only complete supported structures enter the semantic graph. Partial fields
 * and controls keep their selected text/format without copying outside text. */
export function copyWordClipboardSemantics(holder:HTMLElement,model:ContentDocumentModel,completeNodes?:ReadonlySet<string>):{semantics?:WordInlineContentSemantics;attributes:ReadonlyMap<string,Record<string,string>>}{
 const word=model.snapshot().metadata.word!,nodes=new Map([...holder.querySelectorAll<HTMLElement>('[data-node-id]')].map(node=>[node.dataset.nodeId!,node])),complete=(id:string):boolean=>{const node=nodes.get(id),source=model.getNode(id);return!!node&&!!source&&(completeNodes?completeNodes.has(id):text(node)===source.text);},attributes=new Map<string,Record<string,string>>();
 const fields=word.fields.filter(field=>complete(field.nodeId)&&(!field.endNodeId||complete(field.endNodeId))&&(field.instructionFields??[]).every(child=>nodes.has(child.fieldNodeId))).map(field=>({...field,...(field.resultNodeIds?{resultNodeIds:field.resultNodeIds.filter(id=>nodes.has(id))}:{})})),equations=word.equations?.filter(item=>complete(item.nodeId))??[],controls=word.controls?.filter(item=>complete(item.nodeId)&&!word.nativeContent?.fragments.some(fragment=>fragment.nodeId===item.nodeId))??[],hyperlinks=word.hyperlinks.filter(link=>nodes.has(link.nodeId)&&(!link.endNodeId||nodes.has(link.endNodeId))),bookmarks=word.bookmarks.filter(bookmark=>nodes.has(bookmark.startNodeId)&&nodes.has(bookmark.endNodeId));
 const bookmarkIds=new Set(bookmarks.flatMap(bookmark=>[bookmark.startNodeId,bookmark.endNodeId]));
 const controlIds=new Set(controls.map(item=>item.nodeId)),equationIds=new Set(equations.map(item=>item.nodeId)),linkIds=new Set(hyperlinks.flatMap(link=>[link.nodeId,...link.endNodeId?[link.endNodeId]:[]]));
 let selectedFields=fields;for(let changed=true;changed;){const ids=new Set(selectedFields.map(field=>field.nodeId)),next=selectedFields.filter(field=>!field.instructionFields?.some(child=>!ids.has(child.fieldNodeId)));changed=next.length!==selectedFields.length;selectedFields=next;}const fieldIds=new Set(selectedFields.map(field=>field.nodeId)),fieldMarkers=new Set(selectedFields.flatMap(field=>[field.nodeId,...field.endNodeId?[field.endNodeId]:[]]));
 for(const[id,node]of nodes){
  if(node.hasAttribute('data-word-field')&&!fieldMarkers.has(id))node.removeAttribute('data-word-field');
  for(const key of ['data-word-field-end','data-word-field-result'])if(node.hasAttribute(key)&&!fieldIds.has(node.getAttribute(key)!))node.removeAttribute(key);
  if(node.hasAttribute('data-word-content-control')&&!controlIds.has(id))for(const attribute of [...node.attributes])if(attribute.name==='data-word-content-control'||attribute.name.startsWith('data-word-control-'))node.removeAttribute(attribute.name);
  if(node.hasAttribute('data-word-equation')&&!equationIds.has(id))node.removeAttribute('data-word-equation');
  if(node.hasAttribute('data-word-link')&&!linkIds.has(id))node.removeAttribute('data-word-link');
  if(!bookmarkIds.has(id))node.removeAttribute('data-word-bookmark');
  const own=Object.fromEntries([...node.attributes].filter(attribute=>['data-word-field','data-word-field-end','data-word-field-result'].includes(attribute.name)).map(attribute=>[attribute.name,attribute.value]));if(Object.keys(own).length)attributes.set(id,own);
 }
 const semantics=WordInlineContentSemanticsSchema.parse({fields:selectedFields,equations,controls,hyperlinks,bookmarks});return{...(Object.values(semantics).some(items=>(items?.length??0)>0)?{semantics}:{}),attributes};
}
