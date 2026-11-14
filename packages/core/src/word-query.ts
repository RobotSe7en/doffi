import { SourceStyleResolver } from './document/source-css.js';
import { wordTextDefaults } from './document/word-format.js';
import { readLimits } from './document/validation.js';
import type { OpenDocumentOptions } from './runtime.js';
import { DocumentError } from './errors.js';
import type { ContentDocumentModel,DocumentNode } from './runtime.js';
export interface WordTextMatch { readonly nodeId:string;readonly from:number;readonly to:number;readonly expectedText:string }
export interface WordTextFindOptions { readonly matchCase?:boolean;readonly wholeWord?:boolean }
export interface WordOutlineEntry { readonly nodeId:string;readonly level:number;readonly text:string }
export interface WordTextStatistics { readonly words:number;readonly characters:number;readonly charactersWithoutSpaces:number;readonly paragraphs:number }
function paragraphs(model:ContentDocumentModel):DocumentNode[]{
 const snapshot=model.snapshot();if(snapshot.manifest.kind!=='document'||snapshot.metadata.presentation!==undefined)throw new DocumentError('UnsupportedFormat','Word queries require a Word document.');
 const root=snapshot.metadata.word?.stories.find(s=>s.kind==='main')?.rootNodeId;
 const sections=snapshot.metadata.sections;
 if(!sections)throw new DocumentError('UnsupportedFormat','Word queries require Word page metadata.');
 const start=root?model.getNode(root):model.getNode(sections[0]!.rootNodeId);
 const result:DocumentNode[]=[];
 const visit=(node:DocumentNode):void=>{const before=result.length;if(node.attributes['data-word-native-object'])return;if(/^p$|^h[1-6]$/.test(node.tagName)){result.push(node);return;}for(const id of node.childNodeIds){const child=model.getNode(id);if(child)visit(child);}if(['li','td','th'].includes(node.tagName)&&result.length===before&&node.text)result.push(node);};
 if(start)visit(start);return result;
}
/** Source offsets are UTF-16 and never span a paragraph boundary. */
export function findWordText(model:ContentDocumentModel,query:string,options:WordTextFindOptions={}):readonly WordTextMatch[]{
 if(!query)return[];
 const regex=new RegExp(query.replace(/[.*+?^${}()|[\]\\]/g,'\\$&'),options.matchCase?'gu':'giu'),result:WordTextMatch[]=[];
 for(const paragraph of paragraphs(model))for(const match of paragraph.text.matchAll(regex)){
  const from=match.index,to=from+match[0].length;
  if(options.wholeWord){const before=paragraph.text.slice(0,from).match(/.$/u)?.[0]??'',after=paragraph.text.slice(to).match(/^./u)?.[0]??'';if(/[\p{L}\p{N}_]/u.test(before)||/[\p{L}\p{N}_]/u.test(after))continue;}
  result.push(Object.freeze({nodeId:paragraph.nodeId,from,to,expectedText:match[0]}));
 }
 return Object.freeze(result);
}
export function getWordOutline(model:ContentDocumentModel):readonly WordOutlineEntry[]{
 const styles=new Map(model.snapshot().metadata.word?.styles.map(style=>[style.styleId,style]));
 return Object.freeze(paragraphs(model).flatMap(paragraph=>{let level=paragraph.attributes['data-word-outline-level']!==undefined?Number(paragraph.attributes['data-word-outline-level']):undefined;for(let style=styles.get(paragraph.attributes['data-word-style']??'');level===undefined&&style;style=style.basedOn?styles.get(style.basedOn):undefined)level=style.outlineLevel;level??=/^h[1-6]$/.test(paragraph.tagName)?Number(paragraph.tagName[1])-1:undefined;return level!==undefined&&level<9?[Object.freeze({nodeId:paragraph.nodeId,level:level+1,text:paragraph.text})]:[];}));
}
export function getWordTextStatistics(model:ContentDocumentModel):WordTextStatistics{
 const nodes=paragraphs(model),text=nodes.map(p=>p.text).join('\n').replaceAll('\ufffc',''),cjk=text.match(/\p{Script=Han}/gu)?.length??0,western=text.replace(/\p{Script=Han}/gu,' ').match(/[\p{L}\p{N}]+(?:['’\-][\p{L}\p{N}]+)*/gu)?.length??0;
 return Object.freeze({words:cjk+western,characters:Array.from(text.replaceAll('\n','')).length,charactersWithoutSpaces:Array.from(text.replace(/\s/gu,'')).length,paragraphs:nodes.length});
}

/** Authored paragraph defaults for collapsed input and source formatting tools. */
export function getWordTextDefaults(model:ContentDocumentModel,nodeId:string,options:OpenDocumentOptions={}):Readonly<Record<string,string>>{
 const snapshot=model.snapshot();if(snapshot.metadata.presentation)throw new DocumentError('UnsupportedFormat','Word formatting defaults require a Word document.');let root=model.getNode(nodeId);if(!root)throw new DocumentError('AnchorNotFound','The Word text source is missing.');while(root.parentNodeId)root=model.getNode(root.parentNodeId)!;
 return Object.freeze(wordTextDefaults(snapshot.metadata.word??{stories:[{storyId:'main',kind:'main',rootNodeId:root.nodeId}],styles:[]},{peekNode:id=>model.getNode(id)},nodeId,new SourceStyleResolver(snapshot.styles.map(style=>style.css),readLimits(options.limits))));
}
