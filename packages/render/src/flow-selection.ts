import { viewNodeRole } from './view-nodes.js';

/** Layout may unwrap spans, split text and normalize adjacent text nodes. Native
 * live ranges do not keep their character offset across those operations in all
 * engines. Retain offsets in the surviving source element for one measurement. */
export function retainFlowSelection(root:HTMLElement,ignored:(node:Node)=>boolean=node=>viewNodeRole(node)==='remove'):()=>void{
 const selection=root.ownerDocument.getSelection();
 const length=(node:Node):number=>ignored(node)?0:node.nodeType===3?node.textContent!.length:node.nodeType===1&&(node as Element).matches('br,img,[data-word-tab]')?1:[...node.childNodes].reduce((sum,child)=>sum+length(child),0);
 const capture=(node:Node|null,offset:number):{owner:HTMLElement;offset:number;forward:boolean}|undefined=>{
  if(!node||!root.contains(node))return;
  for(let parent:Node|null=node;parent&&parent!==root;parent=parent.parentNode)if(ignored(parent))return;
  const owner=(node.nodeType===1?node as Element:node.parentElement)?.closest<HTMLElement>('[data-node-id]')??root;
  let at=0,found=false;
  const visit=(item:Node):void=>{if(found||ignored(item))return;if(item===node){at+=item.nodeType===3?offset:[...item.childNodes].slice(0,offset).reduce((sum,child)=>sum+length(child),0);found=true;}else if(item.contains(node)){for(const child of item.childNodes)visit(child);}else at+=length(item);};visit(owner);
  return found?{owner,offset:at,forward:node.nodeType===3&&offset===0}:undefined;
 };
 const anchor=capture(selection?.anchorNode??null,selection?.anchorOffset??0),focus=capture(selection?.focusNode??null,selection?.focusOffset??0);
 return()=>{
  if(!anchor||!focus||!anchor.owner.isConnected||!focus.owner.isConnected)return;
  const point=(entry:typeof anchor):[Node,number]=>{
   let left=entry.offset,fallback:[Node,number]|undefined;
   const visit=(node:Node):[Node,number]|undefined=>{if(ignored(node))return;
    if(node.nodeType===3){const size=node.textContent!.length;if(left<size||left===size&&!entry.forward)return[node,left];if(left===size)fallback=[node,left];left-=size;return;}
    if(node.nodeType===1&&(node as Element).matches('br,img,[data-word-tab]')){if(left===0)return[node.parentNode!,[...node.parentNode!.childNodes].indexOf(node as ChildNode)];left--;return;}
    for(const child of node.childNodes){const found=visit(child);if(found)return found;}
    return;
   };
   return visit(entry.owner)??fallback??[entry.owner,entry.owner.childNodes.length];
  };
  selection!.setBaseAndExtent(...point(anchor),...point(focus));
 };
}
