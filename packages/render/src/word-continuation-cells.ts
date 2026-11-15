import {SourceStyleResolver,DEFAULT_DOCUMENT_READ_LIMITS,DocumentError,wordTableCellProperties} from '@doffi/core';
import type {ContentDocumentModel,SourceStyleNode} from '@doffi/core';
import type {RenderLimits} from './contracts.js';
import {markViewNode,ownViewStyle} from './view-nodes.js';
import type{WordRowSpacingView}from'./word-row-spacing.js';

/** Empty physical continuation cells own paint, never a second text/source
 * tree. Recalculate row bounds after pagination and font/layout changes. */
export function composeWordContinuationCells(document:Document,model:ContentDocumentModel,limits:Readonly<RenderLimits>,rowSpacing:ReadonlyMap<string,WordRowSpacingView>=new Map()):()=>void{
 const snapshot=model.snapshot(),word=snapshot.metadata.word,tables=word?.tables.filter(table=>rowSpacing.has(table.nodeId)||table.continuationFormats?.some(entry=>entry.cellCssClass))??[];
 if(!tables.length)return()=>{};
 const css=new SourceStyleResolver(snapshot.styles.map(sheet=>sheet.css),DEFAULT_DOCUMENT_READ_LIMITS),projected=new Map<string,SourceStyleNode>(),sides=['top','right','bottom','left']as const;
 const project=(id:string):SourceStyleNode=>{let result=projected.get(id);if(result)return result;const node=model.getNode(id)!;result={tag:node.tagName,attrs:node.attributes,parent:node.parentNodeId?project(node.parentNodeId):undefined};projected.set(id,result);return result;};
 const layers=new WeakMap<HTMLTableCellElement,HTMLElement[]>();
 return()=>{
  let count=document.querySelectorAll('*').length;
  for(const info of tables)for(const table of document.querySelectorAll<HTMLTableElement>('table[data-node-id="'+CSS.escape(info.nodeId)+'"]')){
   const rows=[...table.rows],occupied:number[]=Array(info.grid.length).fill(0),entries=new Map(info.continuationFormats?.map(entry=>[entry.row+'\n'+entry.column,entry])),spacing=rowSpacing.get(info.nodeId);
   rows.forEach((row,rowIndex)=>{let column=0;for(const cell of row.cells){while(occupied[column]!>rowIndex)column++;const start=column,span=cell.colSpan,height=Math.min(cell.rowSpan||rows.length-rowIndex,rows.length-rowIndex);for(let c=column;c<column+span;c++)occupied[c]=rowIndex+height;column+=span;
    if(cell.dataset.wordGridGap||!spacing&&(height<2||!Array.from({length:height-1},(_,i)=>entries.get((rowIndex+i+1)+'\n'+start)).some(entry=>entry?.cellCssClass)))continue;
    const id=cell.dataset.nodeId;if(!id||!model.getNode(id))continue;
    let paints=layers.get(cell);if(!paints){paints=[];layers.set(cell,paints);ownViewStyle(cell,'position','isolation','background-color',...sides.map(side=>'border-'+side+'-color'));cell.style.position='relative';cell.style.isolation='isolate';cell.style.setProperty('background-color','transparent','important');for(const side of sides)cell.style.setProperty('border-'+side+'-color','transparent','important');}
    const box=cell.getBoundingClientRect(),computed=document.defaultView!.getComputedStyle(cell),factor=document.defaultView!.getComputedStyle(table).borderCollapse==='collapse'?.5:1,left=(parseFloat(computed.borderLeftWidth)||0)*factor,right=(parseFloat(computed.borderRightWidth)||0)*factor,top=(parseFloat(computed.borderTopWidth)||0)*factor;
    for(let index=0;index<height;index++){
     const physicalRow=rowIndex+index,entry=entries.get(physicalRow+'\n'+start),own=entry?.cellCssClass,base=project(id),context=own?wordTableCellProperties(word!,name=>css.specifiedClass(name),table.dataset.wordStyle,info,physicalRow,start,rows.length,span):undefined;
     const appearance=own?css.get({tag:base.tag,attrs:{style:Object.entries({...context,...css.specifiedClass(own)}).map(([key,value])=>key+':'+value).join(';')},parent:project(row.dataset.nodeId!)}):css.get(base);
     let paint=paints[index];if(!paint){if(++count>limits.maxNodes)throw new DocumentError('ResourceLimitExceeded','Continuation cell paint exceeds the render node budget.');paint=markViewNode(document.createElement('div'));paint.className=height>1?'sdk-word-continuation-segment':'sdk-word-cell-spacing-paint';paint.dataset.sdkWordPhysicalRow=String(physicalRow);paint.setAttribute('aria-hidden','true');cell.append(paint);paints[index]=paint;}
     const physical=rows[physicalRow]!.getBoundingClientRect(),bottom=rows[rowIndex+height-1]!.getBoundingClientRect().bottom;
     const delta=spacing?(spacing.rows[physicalRow]??spacing.base)-spacing.base:0,logicalLeft=delta*(start===0?1:.5),logicalRight=delta*(start+span===info.grid.length?1:.5),bidi=table.dataset.sdkWordTableBidi==='true',verticalTop=index===0?delta*(physicalRow===0?1:.5):0,verticalBottom=index===height-1?delta*(physicalRow+1===rows.length?1:.5):0,segmentTop=physical.top-(index>0?(spacing?.base??0)/2:0),segmentBottom=(index===height-1?bottom:physical.bottom+(spacing?.base??0)/2);
     paint.style.cssText=`position:absolute;pointer-events:none;user-select:none;box-sizing:border-box;z-index:-1;left:${-left+(bidi?logicalRight:logicalLeft)}px;right:${-right+(bidi?logicalLeft:logicalRight)}px;top:${segmentTop-box.top-top+verticalTop}px;height:${Math.max(0,segmentBottom-segmentTop-verticalTop-verticalBottom)}px;background-color:${appearance['background-color']}`;
     for(const side of sides){if(side==='top'&&index!==0||side==='bottom'&&index!==height-1)continue;const visual=table.dataset.sdkWordTableBidi==='true'?(side==='left'?'right':side==='right'?'left':side):side;for(const key of ['width','style','color'])paint.style.setProperty('border-'+visual+'-'+key,appearance['border-'+side+'-'+key]??(key==='color'?appearance.color!:'initial'));}
    }
   }});
  }
 };
}
