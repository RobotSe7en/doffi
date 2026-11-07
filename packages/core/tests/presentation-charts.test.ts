import { openContentFixture as openDocument } from "../../../tests/content-fixture.js";
import {expect,test} from "vitest";
import { createDocumentSession, createPresentation } from "@doffi/core";
import type {DocumentOperation,DocumentSession} from "@doffi/core";
const p=(s:DocumentSession)=>s.snapshot().metadata.presentation!;
function prepare(s:DocumentSession,operations:DocumentOperation[]){const{documentId,revision}=s.snapshot().manifest;return s.prepare({protocolVersion:1,documentId,baseRevision:revision,operations});}
const commit=(s:DocumentSession,operations:DocumentOperation[])=>s.commit(prepare(s,operations));
const blank=()=>createDocumentSession(openDocument(createPresentation({documentId:"charts"})));
const geometry={x:20,y:20,width:400,height:200,rotation:0,flipH:false,flipV:false};
const data={categories:["甲","乙"],series:[{name:"系列",values:[10,null]}]};
test("chart creation, constants, nulls, type and title edits are one source transaction and reopen without caches",()=>{
  const s=blank();commit(s,[{op:"insertPresentationChart",parentNodeId:"slide-root-1",geometry,kind:"column",title:"标题",data}]);const chart=p(s).charts[0]!,series=p(s).dataSets[0]!.series[0]!;
  commit(s,[{op:"setPresentationChartData",nodeId:chart.nodeId,data:{categories:["一","二","三"],series:[{seriesId:series.seriesId,name:"中文系列",values:[0,20,null]}]}},{op:"setPresentationChart",nodeId:chart.nodeId,changes:{kind:"line",showLabels:true,legend:"right",minimum:0,maximum:30},title:"新标题"}]);
  expect(p(s).dataSets[0]!.series[0]).toMatchObject({seriesId:series.seriesId,values:[0,20,null]});expect(s.getNode(chart.titleNodeId!)!.text).toBe("新标题");expect(openDocument(s.serialize()).snapshot()).toEqual(s.snapshot());s.undo();expect(p(s).charts[0]!.kind).toBe("column");expect(p(s).dataSets[0]!.categories).toEqual(data.categories);
  const before=s.snapshot();expect(()=>commit(s,[{op:"setPresentationChart",nodeId:chart.nodeId,changes:{kind:"pie"}},{op:"setPresentationChartData",nodeId:chart.nodeId,data:{categories:["负数"],series:[{name:"坏数据",values:[-1]}]}}])).toThrow();expect(s.snapshot()).toEqual(before);
});
test("chart appearance detaches shared style classes and only invalidates pages that use the edited chart",()=>{
  const s=blank();commit(s,[{op:"insertPresentationChart",parentNodeId:"slide-root-1",geometry,kind:"column",data},{op:"insertSlide",layoutId:"layout-blank"}]);const chart=p(s).charts[0]!,series=p(s).dataSets[0]!.series[0]!;
  let candidate=prepare(s,[{op:"setPresentationChartStyle",nodeId:chart.nodeId,target:"series",seriesId:series.seriesId,properties:{"--office-fill-kind":"solid","--office-fill-color":"#123456"}}]);expect(candidate.affectedSlideIds).toEqual(["slide-1"]);s.commit(candidate);
  const firstClass=p(s).charts[0]!.seriesStyles![0]!.cssClass;commit(s,[{op:"duplicatePresentationObjects",nodeIds:[chart.nodeId],offset:{x:0,y:100}}]);const second=p(s).charts[1]!;expect(second.seriesStyles![0]!.cssClass).toBe(firstClass);
  candidate=prepare(s,[{op:"setPresentationChartStyle",nodeId:chart.nodeId,target:"series",seriesId:series.seriesId,properties:{"--office-fill-kind":"solid","--office-fill-color":"#654321"}}]);s.commit(candidate);expect(p(s).charts[0]!.seriesStyles![0]!.cssClass).not.toBe(firstClass);expect(p(s).charts[1]!.seriesStyles![0]!.cssClass).toBe(firstClass);
  expect(commit(s,[{op:"setPresentationChartStyle",nodeId:chart.nodeId,target:"series",seriesId:series.seriesId,properties:{"--office-fill-kind":"solid","--office-fill-color":"#654321"}}]).changed).toBe(false);
});
