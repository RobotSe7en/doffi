import{SourceStyleResolver,DEFAULT_DOCUMENT_READ_LIMITS,DocumentError}from'@doffi/core';
import type{ContentDocumentModel,SourceStyleNode}from'@doffi/core';
import type{RenderLimits}from'./contracts.js';
import{markViewNode,ownViewStyle}from'./view-nodes.js';
export interface WordRowSpacingView{readonly base:number;readonly rows:readonly number[]}
/** Row spacing consumes space inside the existing table bounds. Reversible
 * content scopes and physical-cell paint keep the source grid unchanged. */
export function composeWordRowSpacing(document:Document,model:ContentDocumentModel,limits:Readonly<RenderLimits>):ReadonlyMap<string,WordRowSpacingView>{
 const snapshot=model.snapshot(),known=new Set(snapshot.metadata.word?.tables.map(table=>table.nodeId)),result=new Map<string,WordRowSpacingView>();if(!known.size)return result;
 const css=new SourceStyleResolver(snapshot.styles.map(sheet=>sheet.css),DEFAULT_DOCUMENT_READ_LIMITS),view=document.defaultView!,project=(id:string):SourceStyleNode=>{const node=model.getNode(id)!;return{tag:node.tagName,attrs:node.attributes,parent:node.parentNodeId?project(node.parentNodeId):undefined};};let count=document.querySelectorAll('*').length;
 const pixels=(raw:string,row:HTMLElement):number=>{const match=/^([\d.]+)(pt|px|in|cm|mm|pc|pi|em|rem)?$/.exec(raw);if(match)return Number(match[1])*({pt:4/3,px:1,in:96,cm:96/2.54,mm:96/25.4,pc:16,pi:16,em:parseFloat(view.getComputedStyle(row).fontSize),rem:parseFloat(view.getComputedStyle(document.documentElement).fontSize)}[match[2]??'px']??1);const probe=document.createElement('span');probe.style.cssText='position:absolute;height:0;visibility:hidden;font-size:'+view.getComputedStyle(row).fontSize;probe.style.width=raw;document.body.append(probe);const value=probe.getBoundingClientRect().width;probe.remove();return value;};
 for(const table of document.querySelectorAll<HTMLTableElement>('table')){const id=table.dataset.nodeId;if(!id||!known.has(id))continue;const rows=[...table.rows],values=rows.map(row=>{if(!row.dataset.nodeId)return;const resolved=css.get(project(row.dataset.nodeId));return resolved['--word-row-cell-spacing']??resolved['--word-row-exception-cell-spacing'];});if(!values.some(value=>value!==undefined))continue;
  const base=parseFloat(view.getComputedStyle(table).borderSpacing)||0,width=table.getBoundingClientRect().width,spacing=values.map((value,index)=>value===undefined?base:pixels(value,rows[index]!));result.set(id,{base,rows:spacing});
  ownViewStyle(table,'width','table-layout','border-collapse','border-spacing');table.style.width=width+'px';table.style.tableLayout='fixed';table.style.borderCollapse='separate';table.style.borderSpacing=base+'px';
  const occupied:number[]=Array(snapshot.metadata.word!.tables.find(info=>info.nodeId===id)!.grid.length).fill(0);
  rows.forEach((row,index)=>{let column=0;for(const cell of row.cells){while(occupied[column]!>index)column++;const start=column,span=cell.colSpan,height=Math.min(cell.rowSpan||rows.length-index,rows.length-index);for(let c=column;c<column+span;c++)occupied[c]=index+height;column+=span;if(cell.dataset.wordGridGap)continue;
   const left=(base-spacing[index]!)*(start===0?1:.5),right=(base-spacing[index]!)*(column===occupied.length?1:.5),top=(base-spacing[index]!)*(index===0?1:.5),bottom=(base-spacing[index+height-1]!)*(index+height===rows.length?1:.5);
   if(!left&&!right&&!top&&!bottom)continue;if(++count>limits.maxNodes)throw new DocumentError('ResourceLimitExceeded','Row spacing exceeds the render node budget.');
   const content=markViewNode(document.createElement('div'),'unwrap');content.className='sdk-word-row-spacing-content';content.style.cssText=`width:calc(100% + ${left+right}px);margin-left:${-left}px;margin-right:${-right}px;margin-top:${-top}px;margin-bottom:${-bottom}px;min-width:0`;content.append(...cell.childNodes);cell.append(content);
  }});
 }
 return result;
}
