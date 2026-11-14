import type { DeepReadonly } from '../runtime.js';
import type { WordMetadata } from '../contracts/word.js';
import type { ParsedHtml } from './html.js';
import type { DocumentEdit } from './edit.js';
import type { SourceStyleNode } from './source-css.js';
import type { SourceStyleResolver } from './source-css.js';
import { DocumentError } from '../errors.js';
/** Resolve authored font defaults; browser substitutions never enter source. */
export function wordTextDefaults(word:DeepReadonly<Pick<WordMetadata,'stories'|'styles'|'defaultRunCssClass'>>,html:Pick<ParsedHtml,'peekNode'>,nodeId:string,resolver:SourceStyleResolver):Record<string,string>{
 const node=html.peekNode(nodeId)!,project=(id:string):SourceStyleNode=>{const node=html.peekNode(id)!;return{tag:node.tagName,attrs:node.attributes,parent:node.parentNodeId?project(node.parentNodeId):undefined};};
 let root=node;while(root.parentNodeId&&!word.stories.some(story=>story.rootNodeId===root.nodeId))root=html.peekNode(root.parentNodeId)!;
 let css=resolver.get(project(root.nodeId));if(word.defaultRunCssClass)css=resolver.classStyle(word.defaultRunCssClass,css);
 const definition=word.styles.find(style=>style.styleId===node.attributes['data-word-style']&&style.kind==='paragraph');const chain:(typeof word.styles)[number][]=[],seen=new Set<string>();for(let current=definition;current;current=current.basedOn?word.styles.find(style=>style.styleId===current!.basedOn):undefined){if(seen.has(current.styleId))throw new DocumentError('InvalidOperation','Word style inheritance is cyclic.');seen.add(current.styleId);chain.unshift(current);}
 for(const style of chain){css=resolver.classStyle(style.cssClass,css);if(style.runCssClass)css=resolver.classStyle(style.runCssClass,css);}
 const properties:Record<string,string>={'font-family':css['font-family']??'serif','font-size':css['font-size']??'12pt','font-weight':css['font-weight']??'400','font-style':css['font-style']??'normal','font-variant':css['font-variant']??'normal','text-transform':css['text-transform']??'none',color:css.color??'#000000','letter-spacing':css['letter-spacing']??'normal','vertical-align':'baseline','background-color':'transparent','text-decoration':'none'};
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
