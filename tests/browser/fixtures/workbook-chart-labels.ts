import { createWorkbookNumberFormatter } from '@doffi/core';
import type { WorkbookSnapshot } from '@doffi/core';
import { createWorkbookSurface } from '../../../packages/render/src/workbook/surface.js';
import { WorkbookContentSources } from '../../../packages/render/src/workbook/content-source.js';
import { layoutWorkbookChart } from '../../../packages/render/src/workbook/chart-layout.js';
import type { WorkbookChartBinding } from '../../../packages/render/src/workbook/chart-values.js';

export async function measureNativeChartLabels(snapshot: WorkbookSnapshot, bindings: readonly WorkbookChartBinding[]) {
  const container=document.createElement('div');document.body.append(container);
  const surface=await createWorkbookSurface(container,snapshot,{widthPt:520,heightPt:340}),sources=new WorkbookContentSources(surface.document,snapshot),formatter=createWorkbookNumberFormatter({dateSystem:snapshot.metadata.workbook.dateSystem});
  const charts=snapshot.metadata.workbook.charts!,projection=sources.project(surface.source,charts.map(chart=>chart.nodeId)),before=surface.source.innerHTML;
  try {
    const results=[];
    for(const chart of charts){
      const binding=bindings.find(binding=>binding.chartId===chart.chartId)!;if(binding.kind!=='available')throw Error('Native labels need current values');
      const result=layoutWorkbookChart(surface.view,projection.get(chart.nodeId)!,chart,binding,480,300,'labels-'+chart.chartId,formatter);surface.view.append(result.element);
      const walker=surface.document.createTreeWalker(result.element,0xffff_ffff);let actual=1;while(walker.nextNode())actual++;
      const svg=result.element.querySelector('svg')!;
      const labels=[...svg.querySelectorAll<SVGTextElement>('[data-sdk-chart-label=data-label]')].map(node=>{const b=node.getBBox(),matrix=node.transform.baseVal.consolidate()?.matrix;return {text:node.textContent,index:Number(node.dataset.pointIndex),x:b.x+(matrix?.e??0),y:b.y+(matrix?.f??0),width:b.width,height:b.height};});
      const bars=[...svg.querySelectorAll<SVGRectElement>('[data-sdk-chart-point=column]')].map(node=>({y:node.y.baseVal.value,height:node.height.baseVal.value,fill:node.getAttribute('fill'),stroke:node.getAttribute('stroke')}));
      let negativePixel:number[]|undefined;
      if(chart.kind==='column'&&chart.minimum!<0&&binding.data.series[0]!.values[0]!<0){
        const bar=svg.querySelector<SVGRectElement>('[data-sdk-chart-point=column]')!,clone=svg.cloneNode(true) as SVGSVGElement;clone.setAttribute('width','480');clone.setAttribute('height','300');
        const url=URL.createObjectURL(new Blob([new XMLSerializer().serializeToString(clone)],{type:'image/svg+xml'})),image=new Image();
        try{image.src=url;await image.decode();const canvas=document.createElement('canvas');canvas.width=480;canvas.height=300;const context=canvas.getContext('2d')!;context.drawImage(image,0,0);negativePixel=[...context.getImageData(Math.floor(bar.x.baseVal.value+bar.width.baseVal.value/2),Math.floor(bar.y.baseVal.value+bar.height.baseVal.value/2),1,1).data];}
        finally{image.removeAttribute('src');URL.revokeObjectURL(url);}
      }
      results.push({labels,bars,negativePixel,values:binding.data.series[0]!.values,leaders:svg.querySelectorAll('[data-sdk-chart-leader-line]').length,invalid:/NaN|Infinity/.test(svg.outerHTML),nodes:result.nodes,actual,plot:result.layout.plot,over:chart.showLabelsOverMaximum,kind:chart.kind});
      result.element.remove();
    }
    return {results,unchanged:surface.source.innerHTML===before,cleaned:!surface.source.querySelector('[data-office-view=workbook-chart-label-measurement]')};
  } finally{projection.dispose();sources.dispose();formatter.dispose();surface.dispose();container.remove();}
}

export async function measureControlledChartLabels(snapshot: WorkbookSnapshot, bindings: readonly WorkbookChartBinding[], decoration?: 'solid' | 'double' | 'dotted' | 'dashed' | 'wavy') {
  const container=document.createElement('div');document.body.append(container);
  const css='.measured-label{font-family:"Liberation Sans";font-size:18pt;line-height:30pt;letter-spacing:2pt;--office-text-font-scale:0.5;'+(decoration?'text-decoration:underline line-through '+decoration:'')+'}.align-left{--office-axis-label-align:l}.align-right{--office-axis-label-align:r;--office-axis-label-skip:1}.keep-negative{--office-invert-negative:false}';
  const surface=await createWorkbookSurface(container,{...snapshot,styles:[...snapshot.styles,{path:'measured-label.css',css}]},{widthPt:520,heightPt:340}),sources=new WorkbookContentSources(surface.document,snapshot),formatter=createWorkbookNumberFormatter({dateSystem:snapshot.metadata.workbook.dateSystem});
  const charts=snapshot.metadata.workbook.charts!,projection=sources.project(surface.source,charts.map(chart=>chart.nodeId)),before=surface.source.innerHTML;
  try{
    const results=[];
    for(const alignment of ['left','right']){
      const original=charts[0]!,old=bindings.find(binding=>binding.chartId===original.chartId)!;if(old.kind!=='available')throw Error('No controlled chart');
      const chart={...original,minimum:-4,maximum:8,labelCssClass:'measured-label',categoryAxisCssClass:'align-'+alignment,textCssClass:'measured-label',legend:'none' as const};
      const binding={...old,data:{...old.data,categories:Array(3).fill('A long mixed category 中😀\nSecond line')}};
      const result=layoutWorkbookChart(surface.view,projection.get(chart.nodeId)!,chart,binding,300,220,'controlled-'+alignment,formatter);surface.view.append(result.element);
      const labels=[...result.element.querySelectorAll<SVGTextElement>('[data-sdk-chart-label=category-axis]')];
      const walker=surface.document.createTreeWalker(result.element,0xffff_ffff);let actual=1;while(walker.nextNode())actual++;
      const decorationPaint=labels.map(node=>{const carrier=node.closest('[data-sdk-chart-text-decoration]');return {text:node.textContent,accessible:!carrier?.hasAttribute('aria-hidden'),hidden:node.getAttribute('visibility')===carrier?.getAttribute('visibility'),strokes:carrier?.querySelectorAll('[data-sdk-chart-decoration-paint] line,[data-sdk-chart-decoration-paint] path').length??0};});
      results.push({alignment,decorationPaint,nodes:result.nodes,actual,anchors:labels.map(node=>node.getAttribute('text-anchor')),skipped:labels.filter(node=>node.dataset.sdkChartAutomaticSkip).length,size:labels[0]!.getAttribute('font-size'),spacing:labels[0]!.getAttribute('letter-spacing'),lineOffsets:[...labels[0]!.querySelectorAll('[data-sdk-chart-text-line]')].map(node=>Number(node.getAttribute('dy')))});
      result.element.remove();
    }
    const original=charts[8]!,binding=bindings.find(binding=>binding.chartId===original.chartId)!;if(binding.kind!=='available')throw Error('No negative source');
    const overridden={...original,pointStyles:[{seriesId:original.series[0]!.seriesId,index:0,cssClass:'keep-negative'}]};
    const negative=layoutWorkbookChart(surface.view,projection.get(original.nodeId)!,overridden,binding,480,300,'point-negative',formatter),overrideFill=negative.element.querySelector('[data-sdk-chart-point=column]')!.getAttribute('fill');negative.element.replaceChildren();
    const pie=charts[10]!,old=bindings.find(binding=>binding.chartId===pie.chartId)!;if(old.kind!=='available')throw Error('No pie source');
    const names=['Northern region quarterly\nActual delivery','Southern region quarterly\nActual delivery','Eastern region quarterly\nActual delivery'];
    const legend=layoutWorkbookChart(surface.view,projection.get(pie.nodeId)!,{...pie,showLabels:false,legend:'bottom',legendCssClass:'measured-label'},{...old,data:{...old.data,categories:names}},360,260,'wrapped-legend',formatter);surface.view.append(legend.element);
    const groups=[...legend.element.querySelectorAll<SVGGElement>('[data-sdk-chart-legend-entry]')].map(node=>{const b=node.getBBox(),m=node.transform.baseVal.consolidate()!.matrix;return{x:b.x+m.e,y:b.y+m.f,width:b.width,height:b.height};});
    legend.element.remove();
    const line=charts[0]!,lineBinding=bindings.find(binding=>binding.chartId===line.chartId)!;if(lineBinding.kind!=='available')throw Error('No line source');
    const large=layoutWorkbookChart(surface.view,projection.get(line.nodeId)!,{...line,showLabels:false,legend:'right',seriesStyles:line.series.map(series=>({seriesId:series.seriesId,cssClass:line.seriesStyles![0]!.cssClass,marker:'circle',markerSize:72}))},lineBinding,480,300,'large-marker',formatter);surface.view.append(large.element);
    const marker=large.element.querySelector<SVGGElement>('[data-sdk-chart-marker-role=legend]')!.getBBox(),walker=surface.document.createTreeWalker(large.element,0xffff_ffff);let actual=1;while(walker.nextNode())actual++;
    const result={results,overrideFill,largeMarker:{height:marker.height,rect:large.layout.legend!,nodes:large.nodes,actual},legend:{rect:legend.layout.legend!,groups},unchanged:surface.source.innerHTML===before,cleaned:!surface.source.querySelector('[data-office-view=workbook-chart-label-measurement]')};large.element.remove();return result;
  }finally{projection.dispose();sources.dispose();formatter.dispose();surface.dispose();container.remove();}
}

export async function measureChartNumberFields(snapshot: WorkbookSnapshot, bindings: readonly WorkbookChartBinding[]) {
  const container=document.createElement('div');document.body.append(container);
  const formats=['0_)','_(0_)','??/??','*x0','0*x*y','0_中','0_😀'];
  const css=formats.map((code,index)=>`.numeric-field-${index}{font-family:"Liberation Sans";font-size:11pt;--office-label-number-format:${JSON.stringify(code)};--office-label-number-format-source-linked:false;--office-axis-number-format:${JSON.stringify(code)};--office-axis-number-format-source-linked:false}`).join('');
  const surface=await createWorkbookSurface(container,{...snapshot,styles:[...snapshot.styles,{path:'chart-number-fields.css',css}]},{widthPt:520,heightPt:340}),sources=new WorkbookContentSources(surface.document,snapshot),formatter=createWorkbookNumberFormatter({dateSystem:snapshot.metadata.workbook.dateSystem});
  const original=snapshot.metadata.workbook.charts![0]!,old=bindings.find(binding=>binding.chartId===original.chartId)!;if(old.kind!=='available')throw Error('No format source');
  const projection=sources.project(surface.source,[original.nodeId]),source=projection.get(original.nodeId)!,before=surface.source.innerHTML;
  try{
    const results=[];
    for(const [index,format] of formats.entries()){
      const chart={...original,minimum:0,maximum:4,majorUnit:1,showLabels:true,labelContents:['value'] as const,titleVisible:false,legend:'none' as const,series:[original.series[0]!],labelCssClass:'numeric-field-'+index,valueAxisCssClass:'numeric-field-'+index};
      const binding={...old,data:{categories:['A','B','C'],series:[{...old.data.series[0]!,values:[1,2,3]}]}};
      const result=layoutWorkbookChart(surface.view,source,chart,binding,480,300,'number-'+index,formatter);surface.view.append(result.element);
      const spaces=[...result.element.querySelectorAll<SVGTSpanElement>('[data-sdk-chart-number-space]')].map(node=>({text:node.textContent,declared:Number(node.dataset.sdkChartNumberSpace),actual:node.previousElementSibling && (node.previousElementSibling as SVGTextContentElement).getNumberOfChars() ? node.getEndPositionOfChar(0).x-(node.previousElementSibling as SVGTextContentElement).getEndPositionOfChar((node.previousElementSibling as SVGTextContentElement).getNumberOfChars()-1).x : node.getComputedTextLength()+Number(node.getAttribute('dx')),hidden:node.getAttribute('visibility')}));
      const fills=[...result.element.querySelectorAll<SVGTSpanElement>('[data-sdk-chart-number-fill]')].map(node=>({text:node.textContent,width:node.getComputedTextLength()}));
      const walker=surface.document.createTreeWalker(result.element,0xffff_ffff);let actual=1;while(walker.nextNode())actual++;
      results.push({format,spaces,fills,labels:[...result.element.querySelectorAll('[data-sdk-chart-label=data-label]')].map(node=>node.textContent),nodes:result.nodes,actual});
      result.element.remove();
    }
    return {results,unchanged:surface.source.innerHTML===before,cleaned:!surface.source.querySelector('[data-office-view=workbook-chart-label-measurement]')};
  }finally{projection.dispose();sources.dispose();formatter.dispose();surface.dispose();container.remove();}
}
