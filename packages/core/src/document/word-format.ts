import type { DeepReadonly } from '../runtime.js';
import type { WordMetadata } from '../contracts/word.js';
import type { ParsedHtml } from './html.js';
import type { DocumentEdit } from './edit.js';
import type { SourceStyleNode } from './source-css.js';
import type { SourceStyleResolver } from './source-css.js';
import{resolveWordStyleRun}from'../word-style-context.js';
import{wordTableRunDefaults}from'./word-table-run.js';
/** Resolve authored font defaults; browser substitutions never enter source. */
export function wordTextDefaults(word:DeepReadonly<Pick<WordMetadata,'stories'|'styles'|'defaultRunCssClass'>&Partial<Pick<WordMetadata,'tables'>>>,html:Pick<ParsedHtml,'peekNode'>,nodeId:string,resolver:SourceStyleResolver,maxWork=400000):Record<string,string>{
 let node=html.peekNode(nodeId)!;for(let current= node;current;){if(/^p$|^h[1-6]$/.test(current.tagName)){node=current;break;}if(!current.parentNodeId)break;current=html.peekNode(current.parentNodeId)!;}const project=(id:string):SourceStyleNode=>{const node=html.peekNode(id)!;return{tag:node.tagName,attrs:node.attributes,parent:node.parentNodeId?project(node.parentNodeId):undefined};};
 let root=node;while(root.parentNodeId&&!word.stories.some(story=>story.rootNodeId===root.nodeId))root=html.peekNode(root.parentNodeId)!;
 let css:Record<string,string>={...resolver.get(project(root.nodeId)),'background-color':'transparent','text-decoration-line':'none'};if(word.defaultRunCssClass)css=resolver.classStyle(word.defaultRunCssClass,css);
 const definition=word.styles.find(style=>style.styleId===node.attributes['data-word-style']&&style.kind==='paragraph')??(!node.attributes['data-word-style']?word.styles.find(style=>style.kind==='paragraph'&&style.default)??word.styles.find(style=>style.kind==='paragraph'&&style.styleId==='Normal'):undefined);
 const raw=resolveWordStyleRun(word,name=>resolver.specifiedClass(name),definition?.styleId,{...css,...wordTableRunDefaults(word,html,node.nodeId,resolver,maxWork)});css=resolver.get({tag:'span',attrs:{style:Object.entries(raw).map(([name,value])=>name+':'+value).join(';')},parent:project(root.nodeId)});
 const properties:Record<string,string>={'font-family':css['font-family']??'serif','font-size':css['font-size']??'12pt','font-weight':css['font-weight']??'400','font-style':css['font-style']??'normal','font-variant':css['font-variant']??'normal','text-transform':css['text-transform']??'none',color:css.color??'#000000','letter-spacing':css['letter-spacing']??'normal','vertical-align':'baseline','background-color':raw['background-color']??'transparent','text-decoration':css['text-decoration-line']??'none'};
 for(const name of ['text-decoration-style','text-decoration-color'])if(css[name]!==undefined)properties[name]=css[name]!;
 for(const role of ['ascii','hAnsi','eastAsia','cs'])properties['--word-font-'+role]=css['--word-font-'+role]??properties['font-family']!;
 properties['--word-font-size-cs']=css['--word-font-size-cs']??properties['font-size']!;properties['--word-bold-cs']=css['--word-bold-cs']??(css['font-weight']==='bold'||Number(css['font-weight'])>=600?'1':'0');properties['--word-italic-cs']=css['--word-italic-cs']??(css['font-style']==='italic'?'1':'0');properties['--word-kerning-threshold']=css['--word-kerning-threshold']??'0pt';
 for(const role of ['val','eastAsia','bidi'])if(css['--word-language-'+role]!==undefined)properties['--word-language-'+role]=css['--word-language-'+role]!;
 return properties;
}
export function clearCoveredWordRuns(word:WordMetadata,html:ParsedHtml,edit:DocumentEdit,nodeId:string,from:number,to:number):void{
 const base=html.textOffset(nodeId)!;
 const visit=(id:string):void=>{const node=html.peekNode(id)!,start=html.textOffset(id)!-base,end=start+html.getNode(id)!.text.length;for(const child of node.childNodeIds)visit(child);if(id===nodeId||start<from||end>to)return;
  const style=word.styles.find(style=>style.styleId===node.attributes['data-word-style']&&style.kind==='character');if(style)edit.attributes(id,{'data-word-style':undefined,class:(node.attributes.class??'').split(/\s+/).filter(name=>name!==style.cssClass&&name!==style.runCssClass).join(' ')});
  if(['strong','em','u','s','sub','sup'].includes(node.tagName))edit.renameSource(id,'span');
 };visit(nodeId);
}
