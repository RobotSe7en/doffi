import type{WordMetadata}from'../contracts/word.js';import type{WordRevisionOperation}from'../contracts/word-revision.js';import type{ParsedHtml}from'../document/html.js';import type{DocumentEdit}from'../document/edit.js';import{DocumentError}from'../errors.js';
type Operation=Extract<WordRevisionOperation,{op:'reconcileWordTracked'}>;
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
function paragraphs(source:ParsedHtml,root:string):string[]{return source.nodeIds().filter(id=>{if(!/^p$|^h[1-6]$/.test(source.peekNode(id)!.tagName))return false;for(let current:string|null=id;current;current=source.peekNode(current)!.parentNodeId)if(current===root)return true;return false;});}
function same(a:readonly string[],b:readonly string[]):boolean{return a.length===b.length&&a.every((id,index)=>id===b[index]);}
/** Paragraph marks have their own revisions; unchanged moved text is retained. */
export function reconcileTrackedWord(word:WordMetadata,edit:DocumentEdit,parse:()=>ParsedHtml,operation:Operation):void{
 const input={op:'reconcileFragment'as const,nodeId:operation.nodeId,html:operation.html,textEdits:operation.textEdits},before=parse(),preview=edit.previewReconciliation(input),old=paragraphs(before,operation.nodeId),next=paragraphs(preview,operation.nodeId),oldIds=new Set(old),author=operation.author??word.trackingAuthor??'',date=operation.date??new Date().toISOString();
 const unchanged=(pairs:readonly [string,string][]):boolean=>pairs.every(([old,next])=>before.getNode(old)!.html===preview.getNode(next)!.html);
 if(next.length===old.length+1)for(let index=0;index<old.length;index++){
  const first=old[index]!,tail=next[index+1]!;if(next[index]!==first||oldIds.has(tail)||!same(old.slice(0,index),next.slice(0,index))||!same(old.slice(index+1),next.slice(index+2))||before.peekNode(first)!.parentNodeId!==preview.peekNode(first)!.parentNodeId||preview.peekNode(tail)!.parentNodeId!==preview.peekNode(first)!.parentNodeId)continue;
  if(before.getNode(first)!.text!==preview.getNode(first)!.text+preview.getNode(tail)!.text||!unchanged([...old.slice(0,index),...old.slice(index+1)].map(id=>[id,id])))continue;
  edit.apply(input);const current=parse(),actual=paragraphs(current,operation.nodeId),newTail=actual[index+1]!,oldAttributes=before.peekNode(first)!.attributes,newAttributes=current.peekNode(newTail)!.attributes;edit.attributes(newTail,Object.fromEntries(['style','class','lang','data-word-style','data-word-line-rule','data-word-numbering','data-word-level','data-word-outline-level'].flatMap(name=>oldAttributes[name]!==undefined&&newAttributes[name]===undefined?[[name,oldAttributes[name]]]:[])));for(const revision of word.revisions)if(revision.nodeId===first&&revision.kind.startsWith('paragraph'))revision.nodeId=newTail;
  for(const fragment of word.nativeContent?.fragments??[])if(fragment.nodeId===first&&fragment.kind==='property'&&fragment.propertyPath[0]===W+' pPr'&&(fragment.propertyPath[1]===W+' rPr'||fragment.propertyPath.length===1&&fragment.propertyName===W+' rPr'))fragment.nodeId=newTail;
  const tabs=word.tabs.find(tab=>tab.paragraphNodeId===first);if(tabs)word.tabs.push({...tabs,paragraphNodeId:newTail,stops:tabs.stops.map(stop=>({...stop}))});
  word.revisions.push({nodeId:first,sourceId:newTail+'-break',revisionId:newTail+'-break',kind:'paragraphInsert',author,date});return;
 }
 if(next.length===old.length-1)for(let index=0;index<next.length;index++){
  const first=old[index]!,second=old[index+1]!;if(next[index]!==first||!same(old.slice(0,index),next.slice(0,index))||!same(old.slice(index+2),next.slice(index+1))||before.peekNode(first)!.parentNodeId!==before.peekNode(second)!.parentNodeId)continue;
  if(preview.getNode(first)!.text!==before.getNode(first)!.text+before.getNode(second)!.text||!unchanged([...old.slice(0,index),...old.slice(index+2)].map(id=>[id,id])))continue;
  const inserted=word.revisions.find(revision=>revision.nodeId===first&&revision.kind==='paragraphInsert'&&revision.author===author);
  if(inserted){edit.mergeSourceParagraphMark(first);word.revisions=word.revisions.filter(revision=>revision!==inserted);return;}
  if(!word.revisions.some(revision=>revision.nodeId===first&&revision.kind==='paragraphDelete'))word.revisions.push({nodeId:first,sourceId:second+'-join',revisionId:second+'-join',kind:'paragraphDelete',author,date});return;
 }
 throw new DocumentError('UnsupportedFeature','This structural edit requires a complete tracked paragraph or table change.');
}
