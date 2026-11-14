import type{WordMetadata,WordRevision}from'@doffi/core';import{W,attr,fail,number}from'./ooxml/xml.js';import type{XmlNode}from'./ooxml/xml.js';
type Html={tag:string;attrs:Record<string,string>;children:(Html|string)[]};
interface ActiveMove{moveId:string;kind:'from'|'to';sourceId:string;startNodeId:string;author:string;date:string;columnFirst?:number;columnLast?:number;startDisplaced?:'prev'|'next';}
/** Range state is story-local; Word links source and destination by move name. */
export class WordMoveReader{
 private readonly active=new Map<string,ActiveMove>();
 private readonly tails:Partial<Record<'from'|'to',ActiveMove>>={};
 private readonly links=new WeakMap<ActiveMove,{previous:ActiveMove|undefined;next:ActiveMove|undefined}>();
 constructor(private readonly word:WordMetadata){}
 get pending():boolean{return this.active.size>0;}
 current(kind:'from'|'to'):string|undefined{return this.tails[kind]?.moveId;}
 marker(source:XmlNode,make:(tag:string,attrs:Record<string,string>)=>Html):Html|undefined{
  const match=/^move(From|To)Range(Start|End)$/.exec(source.name);if(source.uri!==W||!match)return;const kind=match[1]==='From'?'from':'to',sourceId=attr(source,'id');if(sourceId===undefined||!/^\d+$/.test(sourceId))fail(source,'A move range identity is invalid.');const key=kind+'\n'+sourceId,node=make('span',{'data-word-move-range':kind+'-'+match[2]!.toLowerCase()});
  if(match[2]==='Start'){const moveId=attr(source,'name'),author=attr(source,'author'),date=attr(source,'date');if(!moveId||author===undefined||date===undefined||!Number.isFinite(Date.parse(date))||this.active.has(key))fail(source,'A move range requires a unique identity, name, author and timestamp.');const entry:ActiveMove={moveId,kind,sourceId,startNodeId:node.attrs['data-node-id']!,author,date,...(attr(source,'colFirst')!==undefined?{columnFirst:number(attr(source,'colFirst'),source)}:{}),...(attr(source,'colLast')!==undefined?{columnLast:number(attr(source,'colLast'),source)}:{}),...(attr(source,'displacedByCustomXml')?{startDisplaced:attr(source,'displacedByCustomXml')as 'prev'|'next'}:{})},previous=this.tails[kind];this.links.set(entry,{previous,next:undefined});if(previous)this.links.get(previous)!.next=entry;this.tails[kind]=entry;this.active.set(key,entry);}
  else{const previous=this.active.get(key);if(!previous)fail(source,'A move range end has no start.');(this.word.moves??=[]).push({...previous,endNodeId:node.attrs['data-node-id']!,...(attr(source,'displacedByCustomXml')?{endDisplaced:attr(source,'displacedByCustomXml')as 'prev'|'next'}:{})});const links=this.links.get(previous)!;if(links.previous)this.links.get(links.previous)!.next=links.next;if(links.next)this.links.get(links.next)!.previous=links.previous;else if(links.previous)this.tails[kind]=links.previous;else delete this.tails[kind];this.active.delete(key);}return node;
 }
 mark(source:XmlNode,owner:Html):void{
  const kind=source.name==='ins'?'paragraphInsert':source.name==='del'?'paragraphDelete':source.name==='moveFrom'?'paragraphMoveFrom':'paragraphMoveTo',sourceId=attr(source,'id');if(sourceId===undefined)fail(source,'A paragraph revision has no identity.');const moveId=source.name==='moveFrom'?this.current('from'):source.name==='moveTo'?this.current('to'):undefined;if(source.name.startsWith('move')&&!moveId)fail(source,'A moved paragraph mark has no source range.');
  this.word.revisions.push({nodeId:owner.attrs['data-node-id']!,revisionId:owner.attrs['data-node-id']+'-'+kind,sourceId,kind,...(moveId?{moveId}:{}),...(attr(source,'author')!==undefined?{author:attr(source,'author')!}:{}),...(attr(source,'date')!==undefined?{date:attr(source,'date')!}:{})});
 }
 revision(source:XmlNode):Pick<WordRevision,'kind'|'moveId'>{const kind=source.name==='moveFrom'?'moveFrom':'moveTo',moveId=this.current(kind==='moveFrom'?'from':'to');if(!moveId)fail(source,'Moved run contents require their named move range.');return{kind,moveId};}
}
