import{wordTableRepeatHeader}from'./word-table-layout.js';
import { DocumentError,isDocumentTextBoundary } from '@doffi/core';
import type { RenderLimits } from './contracts.js';
import { pageFlow } from './flow.js';
import type { PageFlow,PageFlowConstraints } from './flow.js';
import { markViewNode, ownViewStyle, viewNodeRole } from './view-nodes.js';
import type { PreviewSource } from './source-index.js';
import{wordRowVisibleRect}from'./word-table-layout.js';

/** Measure one narrow, reversible source flow, then project logical columns onto
 * physical pages. Text and source elements retain identity, order and UTF-16 offsets. */
export function columnFlow(document:Document,flow:HTMLElement,area:number,pitch:number,width:number,count:number,gap:number,separator:boolean,limits:Readonly<RenderLimits>,source:PreviewSource,constraints?:PageFlowConstraints,pageOffset?:(page:number)=>number,widths?:readonly number[],spaces?:readonly number[]):PageFlow{
 if(count===1&&!pageOffset&&!widths)return pageFlow(document,flow,area,pitch,limits,1,constraints);
 if(count>limits.maxNodes)throw new DocumentError('ResourceLimitExceeded','Column count exceeds the rendering budget.');
 const equalWidth=(width-gap*(count-1))/count,columnWidths=widths??Array.from({length:count},()=>equalWidth),columnWidth=Math.max(...columnWidths),columnOffsets:number[]=[];let offsetX=0;for(let index=0;index<count;index++){columnOffsets.push(offsetX);offsetX+=columnWidths[index]!+(spaces?.[index]??gap);}if(columnWidths.some(width=>width<=0))throw new DocumentError('InvalidDocument','Column gaps leave no printable width.');
 flow.style.width=columnWidth+'px';
 const minimumWidth=Math.min(...columnWidths);if(minimumWidth<columnWidth-.01){const units=limits.maxPages*count;if(units*4>limits.maxNodes)throw new DocumentError('ResourceLimitExceeded','Unequal-column contours exceed the rendering budget.');const points:string[]=[];for(let column=0;column<units;column++){const x=columnWidths[column%count]!-minimumWidth;points.push(x+'px '+column*area+'px',x+'px '+(column+1)*area+'px');}const contour=markViewNode(document.createElement('div'));contour.className='sdk-column-exclusion';contour.style.cssText=`float:right;width:${columnWidth-minimumWidth}px;height:${units*area}px;pointer-events:none;shape-outside:polygon(${columnWidth-minimumWidth}px 0,${points.join(',')},${columnWidth-minimumWidth}px ${units*area}px)`;flow.prepend(contour);}
 const paginator=pageFlow(document,flow,area,area,{...limits,maxPages:limits.maxPages*count},count,constraints?{reserved:index=>constraints.reserved(Math.floor(index/count)),minimumPage:node=>{const minimum=constraints.minimumPage(node);return minimum===undefined?undefined:minimum*count;}}:undefined),wrappers:HTMLElement[]=[],positions=new Map<HTMLElement,string>(),backgrounds:HTMLElement[]=[],starts=new Map<HTMLElement,number>();
 const view=document.defaultView!,origin=()=>flow.getBoundingClientRect(),columnAt=(y:number)=>Math.max(0,Math.floor((y-origin().top+.25)/area));
 let projected=false;
 const offset=(column:number):{x:number;y:number}=>({x:columnOffsets[column%count]!+(pageOffset?.(Math.floor(column/count))??0),y:Math.floor(column/count)*pitch-column*area});
 const move=(element:HTMLElement,column:number):void=>{const delta=offset(column);if(!delta.x&&!delta.y)return;if(!positions.has(element))positions.set(element,element.style.cssText);ownViewStyle(element,'position','left','top');const style=view.getComputedStyle(element);element.style.position='relative';element.style.left=(delta.x+(parseFloat(style.left)||0))+'px';element.style.top=(delta.y+(parseFloat(style.top)||0))+'px';};
 function project():void{
  const base=origin(),nodes=new Map<Node,{rect:DOMRect;from:number;to:number}>();let measured=0,renderNodes=document.getElementsByTagName('*').length;
  const reserve=(count=1):void=>{renderNodes+=count;if(renderNodes>limits.maxNodes)throw new DocumentError('ResourceLimitExceeded','Column projection exceeds the rendering node budget.');};
  const measure=(node:Node):{from:number;to:number}|undefined=>{
   if(++measured>limits.maxNodes*4)throw new DocumentError('ResourceLimitExceeded','Column projection exceeds the render node budget.');
   if(node.nodeType!==1&&node.nodeType!==3)return;
   if(node.nodeType===1&&viewNodeRole(node)==='remove'&&((node as Element).classList.contains('sdk-page-gap')||(node as Element).classList.contains('sdk-column-exclusion')))return;
   let rect:DOMRect,textLines:DOMRect[]|undefined;if(node.nodeType===3){const range=document.createRange();range.selectNodeContents(node);rect=range.getBoundingClientRect();textLines=[...range.getClientRects()].flatMap(rect=>{const visible=wordRowVisibleRect(node.parentElement!,rect);return visible?[visible]:[];});}else rect=(node as Element).getBoundingClientRect();const clipped=wordRowVisibleRect(node.nodeType===1?node as Element:node.parentElement!,rect);if(!clipped)return;rect=clipped;
   if(rect.height<.1&&rect.width<.01)return;
   let from=textLines?.length?textLines.reduce((first,rect)=>Math.min(first,columnAt(rect.top+rect.height/2)),Infinity):columnAt(rect.top),to=textLines?.length?textLines.reduce((last,rect)=>Math.max(last,columnAt(rect.top+rect.height/2)),0):columnAt(Math.max(rect.top,rect.bottom-.5));
   if(node.nodeType===3&&from!==to){const parts:Text[]=[node as Text];let current=node as Text;for(let column=from+1;column<=to;column++){let low=0,high=current.length;while(low<high){const middle=Math.floor((low+high)/2),probe=document.createRange();probe.setStart(current,middle);probe.setEnd(current,Math.min(current.length,middle+1));const rect=probe.getBoundingClientRect();if(columnAt(rect.top+rect.height/2)<column)low=middle+1;else high=middle;}while(low<current.length&&!isDocumentTextBoundary(current.data,low))low++;if(low>0&&low<current.length){reserve();current=current.splitText(low);parts.push(current);}}if(parts.length===1)throw new DocumentError('UnsupportedFeature','A text fragment has no safe column boundary.');for(const part of parts)measure(part);return{from,to};}
   for(const child of [...node.childNodes]){const bounds=measure(child);if(bounds){from=Math.min(from,bounds.from);to=Math.max(to,bounds.to);}}
   nodes.set(node,{rect,from,to});return{from,to};
  };
  measure(flow);
  for(const table of flow.querySelectorAll<HTMLTableElement>('table')){
   const header:HTMLTableRowElement[]=[];for(const row of table.rows){if(viewNodeRole(row)==='remove')continue;if(!wordTableRepeatHeader(row))break;header.push(row);}if(!header.length)continue;
   const bounds=nodes.get(table)!;
   for(let column=bounds.from+1;column<=bounds.to;column++){
    reserve(header.reduce((total,row)=>total+row.querySelectorAll('*').length+1,0)+(table.querySelector(':scope > colgroup')?.querySelectorAll('*').length??0)+3);
    const clone=markViewNode(table.cloneNode(false) as HTMLTableElement),delta=offset(column);clone.contentEditable='false';clone.classList.add('sdk-repeat-header');clone.setAttribute('aria-hidden','true');clone.style.cssText+=`;position:absolute;left:${bounds.rect.left-base.left+delta.x}px;top:${Math.floor(column/count)*pitch}px;width:${bounds.rect.width}px;margin:0;pointer-events:none;user-select:none;z-index:2`;
    const columns=table.querySelector(':scope > colgroup');if(columns)clone.append(columns.cloneNode(true));const body=document.createElement('tbody');for(const row of header)body.append(row.cloneNode(true));clone.append(body);measured+=clone.querySelectorAll('*').length+1;if(measured>limits.maxNodes)throw new DocumentError('ResourceLimitExceeded','Repeated column headers exceed the rendering budget.');flow.append(clone);backgrounds.push(clone);source.registerInstances(clone);
   }
  }
  const paint=(element:HTMLElement,rect:DOMRect,from:number,to:number):void=>{
   const style=view.getComputedStyle(element);if((style.backgroundColor==='rgba(0, 0, 0, 0)'||style.backgroundColor==='transparent')&&['Top','Right','Bottom','Left'].every(side=>parseFloat(style.getPropertyValue('border-'+side.toLowerCase()+'-width'))===0))return;
   const color=style.backgroundColor,borders=Object.fromEntries(['top','right','bottom','left'].map(side=>[side,style.getPropertyValue('border-'+side)]));
   if(!positions.has(element))positions.set(element,element.style.cssText);ownViewStyle(element,'background-color','border-color');element.style.backgroundColor='transparent';element.style.borderColor='transparent';
   for(let column=from;column<=to;column++){const top=Math.max(rect.top-base.top,column*area),bottom=Math.min(rect.bottom-base.top,(column+1)*area);if(bottom<=top)continue;reserve();const delta=offset(column),fragment=markViewNode(document.createElement('div'));fragment.style.cssText=`position:absolute;pointer-events:none;box-sizing:border-box;left:${rect.left-base.left+delta.x}px;top:${top+delta.y}px;width:${Math.min(rect.width,columnWidths[column%count]!)}px;height:${bottom-top}px;background-color:${color};border-left:${borders.left};border-right:${borders.right};border-top:${column===from?borders.top:'0'};border-bottom:${column===to?borders.bottom:'0'};z-index:-1`;flow.append(fragment);backgrounds.push(fragment);}
  };
  const walk=(node:Node):void=>{
   const bounds=nodes.get(node);if(!bounds)return;
   if(node!==flow&&bounds.from===bounds.to){if(node.nodeType===1){if(widths&&bounds.rect.width>columnWidths[bounds.from%count]!+.1)paint(node as HTMLElement,bounds.rect,bounds.from,bounds.to);move(node as HTMLElement,bounds.from);}else if(bounds.rect.height>.1&&bounds.rect.width>.01){reserve();const wrapper=markViewNode(document.createElement('span'),'unwrap');node.parentNode!.insertBefore(wrapper,node);wrapper.append(node);wrappers.push(wrapper);move(wrapper,bounds.from);}return;}
   if(node.nodeType===3)throw new DocumentError('UnsupportedFeature','A text fragment crosses a measured column boundary.');
   if(node!==flow)paint(node as HTMLElement,bounds.rect,bounds.from,bounds.to);
   for(const child of [...node.childNodes])walk(child);
  };
  walk(flow);
  const pages=Math.ceil(paginator.count/count);flow.style.height=(Math.max(0,pages-1)*pitch+area)+'px';flow.style.minHeight=flow.style.height;flow.style.isolation='isolate';
  if(separator)for(let page=0;page<pages;page++)for(let column=1;column<count;column++){reserve();const line=markViewNode(document.createElement('div'));line.style.cssText=`position:absolute;pointer-events:none;left:${columnOffsets[column]!-(spaces?.[column-1]??gap)/2+(pageOffset?.(page)??0)}px;top:${page*pitch}px;height:${area}px;border-left:.5pt solid #555`;flow.append(line);backgrounds.push(line);}
  flow.dataset.sdkColumnCount=String(count);projected=true;
 }
 return{reset(){for(const node of backgrounds.splice(0))node.remove();for(const[element,style]of positions)element.style.cssText=style;positions.clear();for(const wrapper of wrappers.splice(0))wrapper.replaceWith(...wrapper.childNodes);flow.style.removeProperty('height');flow.style.removeProperty('min-height');projected=false;starts.clear();paginator.reset();},layout(element,forcePage,parity){if(widths){for(const section of [element,...element.querySelectorAll<HTMLElement>(".sdk-section")])if(section!==flow){ownViewStyle(section,"display");section.style.display="block";}}paginator.layout(element,forcePage,parity);starts.set(element,Math.floor(paginator.pageAt(element)/count));},project,get count(){return Math.ceil(paginator.count/count);},get blanks(){return new Set([...paginator.blanks].map(value=>Math.floor(value/count)));},pageAt(element){return starts.get(element)??(projected?Math.max(0,Math.floor((element.getBoundingClientRect().top-origin().top+.25)/pitch)):Math.floor(paginator.pageAt(element)/count));},headerHeight:paginator.headerHeight};
}
