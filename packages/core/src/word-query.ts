import { SourceStyleResolver } from './document/source-css.js';
import { wordTextDefaults } from './document/word-format.js';
import { readLimits } from './document/validation.js';
import type { OpenDocumentOptions } from './runtime.js';
import { DocumentError } from './errors.js';
import type { ContentDocumentModel } from './runtime.js';
import type{DocumentOperation}from'./contracts/operations.js';
import{isTextBoundary}from'./document/text.js';
import{wordQueryParagraphs}from'./word-query-view.js';
import type{WordTextQueryOptions}from'./word-query-view.js';
export type{WordTextQueryOptions}from'./word-query-view.js';
export interface WordTextRange { readonly nodeId:string;readonly from:number;readonly to:number;readonly expectedText:string }
export interface WordTextMatch extends WordTextRange {readonly ranges?:readonly WordTextRange[]}
export interface WordTextFindOptions extends WordTextQueryOptions { readonly matchCase?:boolean;readonly wholeWord?:boolean }
export interface WordOutlineEntry { readonly nodeId:string;readonly level:number;readonly text:string }
export interface WordTextStatistics { readonly words:number;readonly characters:number;readonly charactersWithoutSpaces:number;readonly paragraphs:number }
/** Source offsets are UTF-16. A hidden revision or paragraph mark may divide
 * one visible match into several explicitly quoted source ranges. */
export function findWordText(model:ContentDocumentModel,query:string,options:WordTextFindOptions={}):readonly WordTextMatch[]{
 if(!query)return[];
 const regex=new RegExp(query.replace(/[.*+?^${}()|[\]\\]/g,'\\$&'),options.matchCase?'gu':'giu'),result:WordTextMatch[]=[];
 for(const paragraph of wordQueryParagraphs(model,options))for(const match of paragraph.text.matchAll(regex)){
  const from=match.index,to=from+match[0].length;
  if(options.wholeWord){const before=paragraph.text.slice(0,from).match(/.$/u)?.[0]??'',after=paragraph.text.slice(to).match(/^./u)?.[0]??'';if(/[\p{L}\p{N}_]/u.test(before)||/[\p{L}\p{N}_]/u.test(after))continue;}
  if(!isTextBoundary(paragraph.text,from)||!isTextBoundary(paragraph.text,to))continue;
  let visible=0;const ranges:WordTextRange[]=[];
  for(const fragment of paragraph.fragments){const begin=Math.max(from,visible),end=Math.min(to,visible+fragment.text.length);if(begin<end)ranges.push(Object.freeze({nodeId:fragment.nodeId,from:fragment.from+begin-visible,to:fragment.from+end-visible,expectedText:fragment.text.slice(begin-visible,end-visible)}));visible+=fragment.text.length;}
  const first=ranges[0]!,last=ranges.at(-1)!;result.push(Object.freeze({...first,to:first.nodeId===last.nodeId?last.to:first.to,expectedText:match[0],...(ranges.length>1?{ranges:Object.freeze(ranges)}:{})}));
 }
 return Object.freeze(result);
}
export function getWordOutline(model:ContentDocumentModel,options:WordTextQueryOptions={}):readonly WordOutlineEntry[]{
 const styles=new Map(model.snapshot().metadata.word?.styles.map(style=>[style.styleId,style]));
 return Object.freeze(wordQueryParagraphs(model,options).flatMap(({node:paragraph,text})=>{let level=paragraph.attributes['data-word-outline-level']!==undefined?Number(paragraph.attributes['data-word-outline-level']):undefined;for(let style=styles.get(paragraph.attributes['data-word-style']??'');level===undefined&&style;style=style.basedOn?styles.get(style.basedOn):undefined)level=style.outlineLevel;level??=/^h[1-6]$/.test(paragraph.tagName)?Number(paragraph.tagName[1])-1:undefined;return level!==undefined&&level<9?[Object.freeze({nodeId:paragraph.nodeId,level:level+1,text})]:[];}));
}
export function getWordTextStatistics(model:ContentDocumentModel,options:WordTextQueryOptions={}):WordTextStatistics{
 const nodes=wordQueryParagraphs(model,options),text=nodes.map(p=>p.text).join('\n').replaceAll('\ufffc',''),cjk=text.match(/\p{Script=Han}/gu)?.length??0,western=text.replace(/\p{Script=Han}/gu,' ').match(/[\p{L}\p{N}]+(?:['’\-][\p{L}\p{N}]+)*/gu)?.length??0;
 return Object.freeze({words:cjk+western,characters:Array.from(text.replaceAll('\n','')).length,charactersWithoutSpaces:Array.from(text.replace(/\s/gu,'')).length,paragraphs:nodes.length});
}

/** Build one normal transaction batch, validating every quoted piece before
 * returning any edit. Descending source order preserves offsets for all matches. */
export function wordTextReplacementOperations(model:ContentDocumentModel,matches:readonly WordTextMatch[],text:string,options:Pick<WordTextQueryOptions,'limits'>={}):readonly DocumentOperation[]{
 const nodes=new Map(wordQueryParagraphs(model,options).map((entry,index)=>[entry.node.nodeId,{node:entry.node,index}])),edits:{range:WordTextRange;text:string;index:number}[]=[];
 const editable=new Map<string,{from:number;to:number}[]>();for(const paragraph of wordQueryParagraphs(model,{...options,revisionView:'final'}))for(const fragment of paragraph.fragments){const ranges=editable.get(fragment.nodeId)??[];ranges.push(fragment);editable.set(fragment.nodeId,ranges);}
 for(const match of matches){
  const ranges=match.ranges??[match];
  if(!ranges.length||ranges.map(range=>range.expectedText).join('')!==match.expectedText)throw new DocumentError('InvalidOperation','The Word match has inconsistent source pieces.');
  let previous:{index:number;to:number}|undefined;
  for(const[index,range]of ranges.entries()){
   const entry=nodes.get(range.nodeId);
   if(!entry||range.from<0||range.to<=range.from||range.to>entry.node.text.length||entry.node.text.slice(range.from,range.to)!==range.expectedText||!isTextBoundary(entry.node.text,range.from)||!isTextBoundary(entry.node.text,range.to))throw new DocumentError('RevisionConflict','The matched Word source has changed.');
   if(previous&&(entry.index<previous.index||entry.index===previous.index&&range.from<previous.to))throw new DocumentError('InvalidOperation','A Word match must quote its source pieces in reading order.');
   previous={index:entry.index,to:range.to};edits.push({range,text:index===0?text:'',index:entry.index});
  }
 }
 edits.sort((a,b)=>b.index-a.index||b.range.from-a.range.from);for(let index=1;index<edits.length;index++){const left=edits[index]!,right=edits[index-1]!;if(left.range.nodeId===right.range.nodeId&&left.range.to>right.range.from)throw new DocumentError('InvalidOperation','Word replacement matches must not overlap.');}
 for(const {range}of edits){let at=range.from;for(const visible of editable.get(range.nodeId)??[])if(visible.from<=at&&visible.to>at)at=visible.to;if(at<range.to)throw new DocumentError('InvalidOperation','Deleted revision content must be reviewed before replacing it.');}
 return Object.freeze(edits.map(({range,text})=>({op:'replaceText' as const,nodeId:range.nodeId,from:range.from,to:range.to,expectedText:range.expectedText,text})));
}

/** Authored paragraph defaults for collapsed input and source formatting tools. */
export function getWordTextDefaults(model:ContentDocumentModel,nodeId:string,options:OpenDocumentOptions={}):Readonly<Record<string,string>>{
 const snapshot=model.snapshot();if(snapshot.metadata.presentation)throw new DocumentError('UnsupportedFormat','Word formatting defaults require a Word document.');let root=model.getNode(nodeId);if(!root)throw new DocumentError('AnchorNotFound','The Word text source is missing.');while(root.parentNodeId)root=model.getNode(root.parentNodeId)!;
 return Object.freeze(wordTextDefaults(snapshot.metadata.word??{stories:[{storyId:'main',kind:'main',rootNodeId:root.nodeId}],styles:[]},{peekNode:id=>model.getNode(id)},nodeId,new SourceStyleResolver(snapshot.styles.map(style=>style.css),readLimits(options.limits))));
}
