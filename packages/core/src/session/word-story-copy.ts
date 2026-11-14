import type { WordMetadata } from '../contracts/word.js';
import type { DocumentEdit } from '../document/edit.js';
import type { ParsedHtml } from '../document/html.js';
/** Clone semantic ownership with the HTML IDs allocated by the shared editor. */
export function copyWordStorySemantics(word:WordMetadata,identities:ReadonlyMap<string,string>,edit:DocumentEdit,source:ParsedHtml,cloneId:string):void{
 const map=(id:string):string=>identities.get(id)??id;
 for(const [before,after]of identities){const node=source.peekNode(before)!;const attributes:Record<string,string>={};for(const key of ['data-word-field-end','data-word-field-result'])if(node.attributes[key])attributes[key]=map(node.attributes[key]!);if(Object.keys(attributes).length)edit.attributes(after,attributes);}
 word.fields.push(...word.fields.filter(item=>identities.has(item.nodeId)).map(item=>({...item,nodeId:map(item.nodeId),...(item.endNodeId?{endNodeId:map(item.endNodeId)}:{}),...(item.resultNodeIds?{resultNodeIds:item.resultNodeIds.map(map)}:{}),...(item.instructionFields?{instructionFields:item.instructionFields.map(field=>({...field,fieldNodeId:map(field.fieldNodeId)}))}:{})})));
 word.tabs.push(...word.tabs.filter(item=>identities.has(item.paragraphNodeId)).map(item=>({...item,paragraphNodeId:map(item.paragraphNodeId),stops:item.stops.map(stop=>({...stop}))})));
 word.tables.push(...word.tables.filter(item=>identities.has(item.nodeId)).map(item=>({...item,nodeId:map(item.nodeId),grid:[...item.grid]})));
 word.revisions.push(...word.revisions.filter(item=>identities.has(item.nodeId)).map(item=>({...item,nodeId:map(item.nodeId),sourceId:map(item.nodeId)})));
 if(word.equations)word.equations.push(...word.equations.filter(item=>identities.has(item.nodeId)).map(item=>({...JSON.parse(JSON.stringify(item)) as typeof item,nodeId:map(item.nodeId)})));
 if(word.controls)word.controls.push(...word.controls.filter(item=>identities.has(item.nodeId)).map(item=>({...JSON.parse(JSON.stringify(item)) as typeof item,nodeId:map(item.nodeId),sourceId:undefined})));
 for(const reference of word.noteReferences?.filter(item=>identities.has(item.nodeId))??[]){const note=word.notes!.find(note=>note.noteId===reference.noteId)!,story=word.stories.find(story=>story.storyId===note.storyId)!,article=source.nodeIds().find(id=>source.peekNode(id)!.parentNodeId===null)!,clone=edit.cloneSource(story.rootNodeId,article),id=clone.rootNodeIds[0]!;copyWordStorySemantics(word,clone.identities,edit,source,id);word.stories.push({...story,storyId:id,rootNodeId:id});word.notes!.push({...note,noteId:id,storyId:id});word.noteReferences!.push({...reference,nodeId:map(reference.nodeId),noteId:id});}
 if(word.drawings){
  const originals=word.drawings.filter(item=>identities.has(item.nodeId));for(const original of originals){const copy=JSON.parse(JSON.stringify(original)) as typeof original;copy.nodeId=map(original.nodeId);if(copy.kind==='connector'){if(copy.connector.startNodeId)copy.connector.startNodeId=map(copy.connector.startNodeId);if(copy.connector.endNodeId)copy.connector.endNodeId=map(copy.connector.endNodeId);}if('textStoryId'in copy&&copy.textStoryId){const story=word.stories.find(story=>story.storyId===copy.textStoryId)!;const article=source.nodeIds().find(id=>source.peekNode(id)!.parentNodeId===null)!,cloned=edit.cloneSource(story.rootNodeId,article),id=cloned.rootNodeIds[0]!;copyWordStorySemantics(word,cloned.identities,edit,source,id);copy.textStoryId=id;word.stories.push({storyId:id,rootNodeId:id,kind:'textBox'});}word.drawings.push(copy);}
 }
 const names=new Map<string,string>();
 word.bookmarks.push(...word.bookmarks.filter(item=>identities.has(item.startNodeId)&&identities.has(item.endNodeId)).map(item=>{const name=item.name+'_'+cloneId;names.set(item.name,name);return{...item,bookmarkId:map(item.startNodeId),name,startNodeId:map(item.startNodeId),endNodeId:map(item.endNodeId)};}));
 word.hyperlinks.push(...word.hyperlinks.filter(item=>identities.has(item.nodeId)).map(item=>({...item,nodeId:map(item.nodeId),...(item.endNodeId?{endNodeId:map(item.endNodeId)}:{}),...(item.bookmark&&names.has(item.bookmark)?{bookmark:names.get(item.bookmark)!}:{})})));
 if(word.nativeContent){const native=word.nativeContent,parts=new Map<string,string>();
  const fragments=native.fragments.filter(item=>item.nodeId&&identities.has(item.nodeId)).map((item,index)=>{let sourcePart=parts.get(item.sourcePart);if(!sourcePart){sourcePart='doffi-clone/'+cloneId+'/'+parts.size+'.xml';parts.set(item.sourcePart,sourcePart);}const nodeId=map(item.nodeId!),nativeId=cloneId+'-native-'+index;if(item.kind!=='property')edit.attributes(nodeId,{'data-word-native-object':nativeId});return{...item,nodeId,nativeId,sourcePart};});
  native.fragments.push(...fragments);native.relationships.push(...native.relationships.filter(item=>parts.has(item.sourcePart)).map(item=>({...item,sourcePart:parts.get(item.sourcePart)!,root:false})));
 }
}
