import type {WordMetadata} from '../contracts/word.js';
import type {DocumentEdit} from '../document/edit.js';
import type {ParsedHtml} from '../document/html.js';
import type {SourceStyleEdit} from './source-styles.js';
import {copyWordSplitRun} from './word-format-revisions.js';

/** Retained deletions precede replacement text in canonical source. Clone only
 * the first replaced run's current formatting, without its old review history. */
export function insertWordReplacement(word:WordMetadata,edit:DocumentEdit,source:ParsedHtml,styles:SourceStyleEdit,nodeId:string,from:number,to:number,html:string):string[]{
    const point=source.textOffset(nodeId)!+from,path:string[]=[];let context=nodeId;
    for(;;){const child=source.peekNode(context)!.childNodeIds.find(id=>{const start=source.textOffset(id)!,length=source.textLength(id)!;return length>0&&start<=point&&start+length>point;});if(!child)break;const node=source.peekNode(child)!;if(!['span','a','strong','b','em','i','u','s','sub','sup','code','ins'].includes(node.tagName)||node.attributes['data-word-field']||node.attributes['data-word-content-control'])break;path.push(child);context=child;}
    const escape=(value:string):string=>value.replaceAll('&','&amp;').replaceAll('"','&quot;').replaceAll('<','&lt;');
    let content=html;
    for(const id of [...path].reverse()){
        const node=source.peekNode(id)!,attributes=Object.entries(node.attributes).filter(([key])=>['data-node-id','class','style','lang','dir','data-word-style','data-word-run','data-word-script','data-word-run-overlay','data-word-current-format','data-word-reset-text'].includes(key));
        // Links/revision ownership are not formatting. The source class and
        // direct declarations remain linked to their original named styles.
        const tag=['a','ins'].includes(node.tagName)?'span':node.tagName;
        content='<'+tag+attributes.map(([key,value])=>' '+key+'="'+escape(value)+'"').join('')+'>'+content+'</'+tag+'>';
    }
    content='<ins data-word-run-overlay="true" style="text-decoration:none;color:inherit">'+content+'</ins>';
    const result=edit.insertSourceInContainer(nodeId,to,content);copyWordSplitRun(word,result.splitIdentities,styles);styles.aliasNodeIds(result.identities);
    const native=word.nativeContent;
    if(native)native.fragments.push(...native.fragments.filter(fragment=>fragment.nodeId!==undefined&&result.identities.has(fragment.nodeId)&&fragment.kind==='property'&&!fragment.revisionPrevious&&!fragment.propertyPath.some(part=>part.endsWith('Change'))&&!fragment.propertyName?.endsWith('Change')).map((fragment,index)=>{
        const copy={...JSON.parse(JSON.stringify(fragment))as typeof fragment,nodeId:result.identities.get(fragment.nodeId!)!,nativeId:result.rootNodeIds[0]+'-replacement-'+index,...(fragment.propertyName?.endsWith(' rPr')?{currentRunPropertiesOnly:true}:{})};delete copy.revisionSourceId;delete copy.revisionPrevious;return copy;
    }));
    return result.rootNodeIds;
}
