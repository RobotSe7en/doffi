import {SourceStyleResolver,DEFAULT_DOCUMENT_READ_LIMITS,DocumentError,wordTableCellProperties} from '@doffi/core';
import type {ContentDocumentModel,SourceStyleNode} from '@doffi/core';
import type {RenderLimits} from './contracts.js';
import {markViewNode,ownViewStyle} from './view-nodes.js';

/** Empty physical continuation cells own paint, never a second text/source
 * tree. Recalculate row bounds after pagination and font/layout changes. */
export function composeWordContinuationCells(document:Document,model:ContentDocumentModel,limits:Readonly<RenderLimits>):()=>void{
 const snapshot=model.snapshot(),word=snapshot.metadata.word,tables=word?.tables.filter(table=>table.continuationFormats?.some(entry=>entry.cellCssClass))??[];
 if(!tables.length)return()=>{};
 const css=new SourceStyleResolver(snapshot.styles.map(sheet=>sheet.css),DEFAULT_DOCUMENT_READ_LIMITS),projected=new Map<string,SourceStyleNode>(),sides=['top','right','bottom','left']as const;
 const project=(id:string):SourceStyleNode=>{let result=projected.get(id);if(result)return result;const node=model.getNode(id)!;result={tag:node.tagName,attrs:node.attributes,parent:node.parentNodeId?project(node.parentNodeId):undefined};projected.set(id,result);return result;};
 const layers=new WeakMap<HTMLTableCellElement,HTMLElement[]>();
 return()=>{
  let count=document.querySelectorAll('*').length;
  for(const info of tables)for(const table of document.querySelectorAll<HTMLTableElement>('table[data-node-id="'+CSS.escape(info.nodeId)+'"]')){
   const rows=[...table.rows],occupied:number[]=Array(info.grid.length).fill(0),entries=new Map(info.continuationFormats?.map(entry=>[entry.row+'\n'+entry.column,entry]));
   rows.forEach((row,rowIndex)=>{let column=0;for(const cell of row.cells){while(occupied[column]!>rowIndex)column++;const start=column,span=cell.colSpan,height=Math.min(cell.rowSpan||rows.length-rowIndex,rows.length-rowIndex);for(let c=column;c<column+span;c++)occupied[c]=rowIndex+height;column+=span;
    if(height<2||!Array.from({length:height-1},(_,i)=>entries.get((rowIndex+i+1)+'\n'+start)).some(entry=>entry?.cellCssClass))continue;
    const id=cell.dataset.nodeId;if(!id||!model.getNode(id))continue;
    let paints=layers.get(cell);if(!paints){paints=[];layers.set(cell,paints);ownViewStyle(cell,'position','isolation','background-color',...sides.map(side=>'border-'+side+'-color'));cell.style.position='relative';cell.style.isolation='isolate';cell.style.setProperty('background-color','transparent','important');for(const side of sides)cell.style.setProperty('border-'+side+'-color','transparent','important');}
    const box=cell.getBoundingClientRect(),computed=document.defaultView!.getComputedStyle(cell),left=(parseFloat(computed.borderLeftWidth)||0)/2,right=(parseFloat(computed.borderRightWidth)||0)/2,top=(parseFloat(computed.borderTopWidth)||0)/2;
    for(let index=0;index<height;index++){
     const physicalRow=rowIndex+index,entry=entries.get(physicalRow+'\n'+start),own=entry?.cellCssClass,base=project(id),context=own?wordTableCellProperties(word!,name=>css.specifiedClass(name),table.dataset.wordStyle,info,physicalRow,start,rows.length,span):undefined;
     const appearance=own?css.get({tag:base.tag,attrs:{style:Object.entries({...context,...css.specifiedClass(own)}).map(([key,value])=>key+':'+value).join(';')},parent:project(row.dataset.nodeId!)}):css.get(base);
     let paint=paints[index];if(!paint){if(++count>limits.maxNodes)throw new DocumentError('ResourceLimitExceeded','Continuation cell paint exceeds the render node budget.');paint=markViewNode(document.createElement('div'));paint.className='sdk-word-continuation-segment';paint.dataset.sdkWordPhysicalRow=String(physicalRow);paint.setAttribute('aria-hidden','true');cell.append(paint);paints[index]=paint;}
     const physical=rows[physicalRow]!.getBoundingClientRect(),bottom=rows[rowIndex+height-1]!.getBoundingClientRect().bottom;
     paint.style.cssText=`position:absolute;pointer-events:none;user-select:none;box-sizing:border-box;z-index:-1;left:${-left}px;right:${-right}px;top:${physical.top-box.top-top}px;height:${(index===height-1?bottom:physical.bottom)-physical.top}px;background-color:${appearance['background-color']}`;
     for(const side of sides){if(side==='top'&&index!==0||side==='bottom'&&index!==height-1)continue;for(const key of ['width','style','color'])paint.style.setProperty('border-'+side+'-'+key,appearance['border-'+side+'-'+key]??(key==='color'?appearance.color!:'initial'));}
    }
   }});
  }
 };
}
