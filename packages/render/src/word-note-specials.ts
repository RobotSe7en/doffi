import type{DeepReadonly,WordMetadata}from'@doffi/core';
import{DocumentError}from'@doffi/core';
import type{RenderLimits}from'./contracts.js';
import{cloneViewTree}from'./view-clone.js';
import{markViewNode,ownViewStyle}from'./view-nodes.js';
type Special='separator'|'continuationSeparator'|'continuationNotice';
export function wordNoteSpecials(document:Document,word:DeepReadonly<WordMetadata>|undefined,templates:ReadonlyMap<string,HTMLElement|undefined>,limits:Readonly<RenderLimits>):{
 create(kind:'footnote'|'endnote',type:Special):HTMLElement|undefined;
 height(kind:'footnote'|'endnote',type:Special,width:number,layer:HTMLElement):number;
}{
 for(const template of templates.values())for(const node of template?.querySelectorAll<HTMLElement>('[data-word-note-separator]')??[]){ownViewStyle(node,'display','width','max-width','border-bottom','height','vertical-align');node.style.cssText+=';display:inline-block;width:'+(node.dataset.wordNoteSeparator==='separator'?'144pt':'100%')+';max-width:100%;border-bottom:.5pt solid currentColor;height:.4em;vertical-align:middle';}
 const create=(kind:'footnote'|'endnote',type:Special):HTMLElement|undefined=>{
  const note=word?.notes?.find(note=>note.kind===kind&&note.type===type),template=note&&templates.get(note.storyId);
  if(template){const copy=cloneViewTree(template,limits.maxNodes);copy.dataset.sdkWordNoteStory=note!.storyId;copy.dataset.sdkWordNoteSpecial=type;copy.classList.remove('sdk-auxiliary');copy.style.cssText+=';display:flow-root;position:relative;border:0;padding:0;margin:0';return copy;}
  if(type==='continuationNotice')return;
  const line=markViewNode(document.createElement('div'));line.dataset.sdkWordNoteSpecial=type;line.style.cssText='height:12px;max-width:100%;border-top:.5pt solid #777;width:'+(type==='separator'?'144pt':'100%');return line;
 };
 const heights=new Map<string,number>();
 return{create,height(kind,type,width,layer){const key=kind+' '+type+' '+width;if(heights.has(key))return heights.get(key)!;const copy=create(kind,type);let value=0;if(copy){copy.style.cssText+=';position:absolute;visibility:hidden;width:'+width+'pt';layer.append(copy);try{if(document.getElementsByTagName('*').length>limits.maxNodes)throw new DocumentError('ResourceLimitExceeded','Note separator measurement exceeds the render node budget.');value=copy.getBoundingClientRect().height;}finally{copy.remove();}}heights.set(key,value);return value;}};
}
