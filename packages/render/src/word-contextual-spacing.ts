import type{ContentDocumentModel}from'@doffi/core';
import{ownViewStyle,viewNodeRole}from'./view-nodes.js';
const originals=new WeakMap<HTMLElement,{top:string;bottom:string;topPriority:string;bottomPriority:string}>();
/** Native contextual spacing suppresses this paragraph's own before/after
 * spacing only beside the identical paragraph style. Tables separate flows;
 * each cell and auxiliary story owns an independent paragraph sequence. */
export function composeWordContextualSpacing(document:Document,model:ContentDocumentModel):void{
 if(!model.snapshot().metadata.word)return;const root=document.querySelector('[data-office-kind="document"]');if(!root)return;
 const previous=new Map<Element,HTMLElement>(),before=new Map<HTMLElement,HTMLElement>(),after=new Map<HTMLElement,HTMLElement>(),paragraphs:HTMLElement[]=[],view=document.defaultView!;
 const group=(node:Element):Element=>node.closest('td,th,[data-office-story]')??root;
 for(const node of root.querySelectorAll<HTMLElement>('*')){
  if(!node.dataset.nodeId||viewNodeRole(node)==='remove'||node.closest('.sdk-repeat-header,.sdk-word-drawing-text'))continue;
  if(node.localName==='table'){previous.delete(group(node));continue;}
  if(!/^p$|^h[1-6]$/.test(node.localName))continue;
  if(node.hasAttribute('data-word-native-object')){previous.delete(group(node));continue;}
  if(node.closest('[data-word-native-object],[data-word-drawing]')||view.getComputedStyle(node).display==='none')continue;
  const flow=group(node),last=previous.get(flow);if(last){before.set(node,last);after.set(last,node);}previous.set(flow,node);paragraphs.push(node);
 }
 const style=(node:HTMLElement):string=>node.dataset.wordStyle??(/^h[1-6]$/.test(node.localName)?node.localName:'');
 for(const paragraph of paragraphs){const active=view.getComputedStyle(paragraph).getPropertyValue('--word-contextual-spacing').trim()==='1';let original=originals.get(paragraph);if(!active&&!original)continue;if(!original){original={top:paragraph.style.marginTop,bottom:paragraph.style.marginBottom,topPriority:paragraph.style.getPropertyPriority('margin-top'),bottomPriority:paragraph.style.getPropertyPriority('margin-bottom')};originals.set(paragraph,original);}
  ownViewStyle(paragraph,'margin-top','margin-bottom');
  for(const[side,neighbor,raw,priority]of [['top',before.get(paragraph),original.top,original.topPriority],['bottom',after.get(paragraph),original.bottom,original.bottomPriority]]as const){const property='margin-'+side;if(active&&neighbor&&style(neighbor)===style(paragraph))paragraph.style.setProperty(property,'0pt');else if(raw)paragraph.style.setProperty(property,raw,priority);else paragraph.style.removeProperty(property);}
 }
}
