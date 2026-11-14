import type {DeepReadonly,DocumentSection} from '@doffi/core';
import {DocumentError}from'@doffi/core';
import type {RenderLimits} from './contracts.js';
import type {PageFlow,PageFlowConstraints}from'./flow.js';
import{columnFlow}from'./column-flow.js';
import type{PreviewSource}from'./source-index.js';
import{markViewNode}from'./view-nodes.js';
export interface FlowSection{readonly element:HTMLElement;readonly definition:DeepReadonly<DocumentSection>}
interface Band{section:FlowSection;flow:HTMLElement;paginator:PageFlow;start:number;inset:number;firstColumn:number}
/** Source sections remain in reading order; owned bands place differing column
 * sets on shared physical pages without creating an editable DOM per page. */
export function sectionFlow(document:Document,flow:HTMLElement,area:number,pitch:number,width:number,limits:Readonly<RenderLimits>,source:PreviewSource,sections:()=>readonly FlowSection[],balanceEnd:(section:FlowSection)=>boolean,constraints?:PageFlowConstraints,pageOffset?:(page:number)=>number):PageFlow{
 const bands=new Map<HTMLElement,Band>(),starts=new Map<HTMLElement,number>(),blankPages=new Set<number>();
 let count=1,cursorPage=0,cursorInset=0,previous:Band|undefined;
 const size=()=>{flow.style.height=((count-1)*pitch+area)+'px';flow.style.minHeight=flow.style.height;};
 function bandFor(element:HTMLElement):Band{
  const known=bands.get(element);if(known)return known;
  const section=sections().find(section=>section.element===element);if(!section)throw new DocumentError('InvalidDocument','A section flow requires its source definition.');
  const wrapper=markViewNode(document.createElement('div'),'unwrap');wrapper.className='sdk-page-flow sdk-section-band';wrapper.style.cssText='position:absolute;left:0;top:0;display:flow-root';element.before(wrapper);wrapper.append(element);
  const columns=section.definition.columns;
  let band:Band;
  const local:PageFlowConstraints={reserved:index=>constraints?.reserved(band.start+index)??0,minimumPage:node=>{const minimum=constraints?.minimumPage(node);return minimum===undefined?undefined:Math.max(0,minimum-band.start);}};
  const paginator=columnFlow(document,wrapper,area,pitch,width,columns?.count??1,(columns?.gap??0)*4/3,columns?.separator??false,limits,source,local,pageOffset?(index=>pageOffset(band.start+index)):undefined,columns?.widths?.map(value=>value*4/3),columns?.spaces?.map(value=>value*4/3),{firstPageInset:()=>band.inset,firstColumn:()=>band.firstColumn,balanceEnd:()=>balanceEnd(section),compact:true});
  band={section,flow:wrapper,paginator,start:0,inset:0,firstColumn:0};bands.set(element,band);return band;
 }
 return{
  reset(){for(const band of bands.values())band.paginator.reset();starts.clear();blankPages.clear();count=1;cursorPage=0;cursorInset=0;previous=undefined;size();},
  layout(element,forcePage=false,parity){
   const band=bandFor(element);let page=cursorPage,inset=cursorInset;
   band.firstColumn=0;
   if(previous&&band.section.definition.breakType==='nextColumn'){
    page=previous.start+previous.paginator.count-1;
    const next=(previous.paginator.endColumn??0)+1,columns=band.section.definition.columns?.count??1;
    if(next<columns&&JSON.stringify(previous.section.definition.columns)===JSON.stringify(band.section.definition.columns)){band.firstColumn=next;inset=page===previous.start?previous.inset:0;}
    else{page++;inset=0;}
   }
   if(forcePage&&inset>.5){page++;inset=0;}
   if(area-(constraints?.reserved(page)??0)-inset<24){page++;inset=0;}
   if(parity!==undefined&&page%2!==parity){blankPages.add(page);page++;inset=0;}
   if(page>=limits.maxPages)throw new DocumentError('ResourceLimitExceeded','Document exceeds the render page budget.');
   band.start=page;band.inset=inset;band.flow.style.top=(page*pitch+inset)+'px';
   band.paginator.layout(element);starts.set(element,page);for(const blank of band.paginator.blanks)blankPages.add(page+blank);
   const last=band.paginator.count-1;count=Math.max(count,page+band.paginator.count);cursorPage=page+last;cursorInset=band.paginator.bottom(last)+(last===0?inset:0);
   for(const earlier of bands.values())if(earlier!==band&&starts.has(earlier.section.element)&&cursorPage>=earlier.start&&cursorPage<earlier.start+earlier.paginator.count)cursorInset=Math.max(cursorInset,earlier.paginator.bottom(cursorPage-earlier.start)+(cursorPage===earlier.start?earlier.inset:0));
   if(cursorInset>=area-(constraints?.reserved(cursorPage)??0)-.5){cursorPage++;cursorInset=0;}
   previous=band;size();
  },
  project(){for(const band of bands.values())band.paginator.project?.();size();},
  get count(){return count;},get blanks(){return blankPages;},
  pageAt(element){const exact=starts.get(element);if(exact!==undefined)return exact;for(const band of bands.values())if(band.flow.contains(element))return band.start+band.paginator.pageAt(element);return 0;},
  headerHeight(table){for(const band of bands.values())if(band.flow.contains(table))return band.paginator.headerHeight(table);return 0;},
  bottom(index){let bottom=0;for(const band of bands.values())if(index>=band.start&&index<band.start+band.paginator.count)bottom=Math.max(bottom,band.paginator.bottom(index-band.start)+(index===band.start?band.inset:0));return bottom;}
 };
}
