import type{ContentDocumentModel,DocumentNode,DocumentReadLimits}from'./runtime.js';
import{DocumentError}from'./errors.js';
import{parseHtmlTree}from'./document/html-parser.js';
import{supportedTags}from'./document/html.js';
import{readLimits}from'./document/validation.js';
import{textPositions}from'./document/text.js';
export interface WordTextQueryOptions{readonly revisionView?:'all'|'final'|'original';readonly limits?:Partial<DocumentReadLimits>}
export interface QueryFragment{readonly nodeId:string;readonly from:number;readonly to:number;readonly text:string}
export interface QueryParagraph{readonly node:DocumentNode;readonly text:string;readonly fragments:readonly QueryFragment[]}
/** Query projection retains exact source pieces; removing a revision from view
 * must never turn a discontiguous match into one destructive source interval. */
export function wordQueryParagraphs(model:ContentDocumentModel,options:WordTextQueryOptions={}):readonly QueryParagraph[]{
 const snapshot=model.snapshot();if(snapshot.manifest.kind!=='document'||snapshot.metadata.presentation!==undefined)throw new DocumentError('UnsupportedFormat','Word queries require a Word document.');
 const word=snapshot.metadata.word,sections=snapshot.metadata.sections,view=options.revisionView??'all',hideKinds=view==='final'?['delete','moveFrom','rowDelete','cellDelete']:view==='original'?['insert','moveTo','rowInsert','cellInsert']:[],hidden=new Set(word?.revisions.filter(item=>hideKinds.includes(item.kind)).map(item=>item.nodeId)),nodes=new Map<string,DocumentNode>();
 const get=(id:string):DocumentNode=>{let value=nodes.get(id);if(!value){value=model.getNode(id);if(!value)throw new DocumentError('AnchorNotFound','A Word query source is missing.');nodes.set(id,value);}return value;};
 const start=get(word?.stories.find(story=>story.kind==='main')?.rootNodeId??sections[0]!.rootNodeId),paragraphs:DocumentNode[]=[];
 const visit=(node:DocumentNode):void=>{if(hidden.has(node.nodeId)||node.attributes['data-word-native-object'])return;const before=paragraphs.length;if(/^p$|^h[1-6]$/.test(node.tagName)){paragraphs.push(node);return;}for(const id of node.childNodeIds)visit(get(id));if(['li','td','th'].includes(node.tagName)&&paragraphs.length===before&&node.text)paragraphs.push(node);};visit(start);
 const hiddenByParagraph=new Map<string,string[]>(),paragraphIds=new Set(paragraphs.map(node=>node.nodeId));
 for(const id of hidden){for(let current=get(id);current.parentNodeId;){current=get(current.parentNodeId);if(paragraphIds.has(current.nodeId)){const entries=hiddenByParagraph.get(current.nodeId)??[];entries.push(id);hiddenByParagraph.set(current.nodeId,entries);break;}}}
 const offsets=new Map<string,{from:number;to:number}>();if(hiddenByParagraph.size){const tree=parseHtmlTree(snapshot.html,readLimits(options.limits),supportedTags),positions=textPositions(tree);for(const[node,range]of positions)if('attrs'in node){const id=node.attrs.find(attr=>attr.name==='data-node-id')?.value;if(id)offsets.set(id,range);}}
 const mergedKinds=view==='final'?['paragraphDelete','paragraphMoveFrom']:view==='original'?['paragraphInsert','paragraphMoveTo']:[],merged=new Set(word?.revisions.filter(item=>mergedKinds.includes(item.kind)).map(item=>item.nodeId)),result:QueryParagraph[]=[];
 const adjacent=(first:DocumentNode,last:DocumentNode):boolean=>{if(!first.parentNodeId||first.parentNodeId!==last.parentNodeId)return false;const siblings=get(first.parentNodeId).childNodeIds,from=siblings.indexOf(first.nodeId),to=siblings.indexOf(last.nodeId);return to>from&&siblings.slice(from+1,to).every(id=>{const node=get(id);return ['div','span'].includes(node.tagName)&&!node.text&&/\bdata-word-(?:bookmark|move-range)=/.test(node.html);});};
 for(const[index,original]of paragraphs.entries()){
  const base=offsets.get(original.nodeId)?.from??0,intervals=(hiddenByParagraph.get(original.nodeId)??[]).map(id=>offsets.get(id)!).filter(Boolean).map(range=>({from:Math.max(0,range.from-base),to:Math.min(original.text.length,range.to-base)})).sort((a,b)=>a.from-b.from),fragments:QueryFragment[]=[];let at=0;
  for(const range of [...intervals,{from:original.text.length,to:original.text.length}]){if(range.from>at)fragments.push({nodeId:original.nodeId,from:at,to:range.from,text:original.text.slice(at,range.from)});at=Math.max(at,range.to);}
  let node=original;if(view==='original'){const before=word?.revisions.find(item=>item.nodeId===original.nodeId&&item.kind==='format'&&item.property==='paragraph')?.before;if(before){const attrs={...node.attributes};for(const[key,value]of Object.entries(before.attributes??{})){if(value===null)delete attrs[key];else attrs[key]=value;}node={...node,tagName:before.tagName??node.tagName,attributes:attrs};}}
  const text=fragments.map(item=>item.text).join(''),previous=result.at(-1),previousOriginal=paragraphs[index-1];
  if(previous&&previousOriginal&&merged.has(previousOriginal.nodeId)&&adjacent(previousOriginal,original))result[result.length-1]={node,text:previous.text+text,fragments:[...previous.fragments,...fragments]};else result.push({node,text,fragments});
 }
 return result;
}
