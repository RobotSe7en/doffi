import{openDocument,SourceStyleResolver,DEFAULT_DOCUMENT_READ_LIMITS,getWordTextDefaults}from'@doffi/core';
import type{WordContentSemantics}from'@doffi/core';
import{importDocx}from'@doffi/office';
import{continuationCellFixture}from'./continuation-cell-fixture.js';
export function wordTableContentFixture():{html:string;semantics:WordContentSemantics}{
 const model=openDocument(importDocx(continuationCellFixture(),{documentId:'table-content-source',unsupportedContent:'reject'}),{kind:'document'}),snapshot=model.snapshot(),word=snapshot.metadata.word!,table=word.tables[0]!,css=new SourceStyleResolver(snapshot.styles.map(sheet=>sheet.css),DEFAULT_DOCUMENT_READ_LIMITS),names=new Set<string>();
 const visit=(id:string):void=>{const node=model.getNode(id)!;for(const name of node.attributes.class?.split(/\s+/)??[])if(name)names.add(name);for(const child of node.childNodeIds)visit(child);};visit(table.nodeId);for(const entry of table.continuationFormats??[]){if(entry.cellCssClass)names.add(entry.cellCssClass);for(const paragraph of entry.emptyParagraphs??[]){if(paragraph.paragraphCssClass)names.add(paragraph.paragraphCssClass);if(paragraph.runCssClass)names.add(paragraph.runCssClass);for(const run of paragraph.runs)if(run.cssClass)names.add(run.cssClass);}}
 const defaults=Object.entries(getWordTextDefaults(model,table.nodeId)).map(([key,value])=>key+':'+value).join(';').replaceAll('&','&amp;').replaceAll('"','&quot;');return{html:'<div style="'+defaults+'">'+model.getNode(table.nodeId)!.html+'</div>',semantics:JSON.parse(JSON.stringify({tables:[table],tabs:word.tabs,conditionalFormats:word.conditionalFormats,extraClasses:Object.fromEntries([...names].map(name=>[name,css.specifiedClass(name)]))}))as WordContentSemantics};
}
