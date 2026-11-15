import {DocumentError} from '../errors.js';
import type {WordMetadata} from '../contracts/word.js';
import type {WordRevision} from '../contracts/word-revision.js';
import type {DocumentEdit} from '../document/edit.js';
import type {ParsedHtml} from '../document/html.js';
import type {SourceStyleEdit} from './source-styles.js';
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
/** Paragraph marks are empty source runs; their history has no body-text range. */
export function formatWordParagraphMark(word:WordMetadata,edit:DocumentEdit,parse:()=>ParsedHtml,paragraphId:string,properties:Readonly<Record<string,string>>):void{
    const source=parse(),paragraph=source.peekNode(paragraphId)!;
    if(!/^p$|^h[1-6]$/.test(paragraph.tagName))throw new DocumentError('InvalidOperation','Paragraph mark formatting requires a paragraph.');
    if(!Object.keys(properties).length)return;
    const old=paragraph.childNodeIds.find(id=>source.peekNode(id)!.attributes['data-word-paragraph-mark']==='true'),id=old??edit.insertSource(paragraphId,'<span data-word-paragraph-mark="true"></span>').rootNodeIds[0]!;
    // Old packages attached mark properties to p/pPr. Move only that exact
    // property path to the mark's own native rPr boundary.
    for(const fragment of word.nativeContent?.fragments??[])if(fragment.nodeId===paragraphId&&fragment.kind==='property'&&fragment.propertyPath[0]===W+' pPr'){
        const count=fragment.propertyPath[1]===W+' rPr'?2:fragment.propertyPath.length===1&&fragment.propertyName===W+' rPr'?1:0;if(!count)continue;
        fragment.nodeId=id;fragment.ownerName='rPr';fragment.propertyPath=fragment.propertyPath.slice(count);if(fragment.propertyIndices)fragment.propertyIndices=fragment.propertyIndices.slice(count);
    }
    if(word.trackChanges){
        const existing=word.revisions.find(revision=>revision.nodeId===id&&revision.kind==='format'&&revision.property==='run');
        if(existing){const attributes=existing.before!.attributes??={};if(attributes['data-word-current-format']===undefined)attributes['data-word-current-format']=null;existing.author=word.trackingAuthor??'';existing.date=new Date().toISOString();}
        else{const attributes:NonNullable<NonNullable<WordRevision['before']>['attributes']>={};for(const key of ['style','class','lang','data-word-style','data-word-current-format']as const)attributes[key]=old?source.peekNode(old)!.attributes[key]??null:null;const revisionId=id+'-mark-format';word.revisions.push({nodeId:id,sourceId:revisionId,revisionId,kind:'format',property:'run',before:{attributes,inheritNativeProperties:true},author:word.trackingAuthor??'',date:new Date().toISOString()});}
    }
    const values={...properties};if(values['font-weight']!==undefined&&values['--word-bold-cs']===undefined)values['--word-bold-cs']=values['font-weight']==='bold'||Number(values['font-weight'])>=600?'1':'0';if(values['font-style']!==undefined&&values['--word-italic-cs']===undefined)values['--word-italic-cs']=values['font-style']==='italic'?'1':'0';
    for(const[key,value]of Object.entries(values))edit.sourceStyle(id,key,value);edit.attributes(id,{'data-word-current-format':'true'});
}

/** The same inheritance applies to typed text, browser reconciliation and AI
 * text operations; installed browser faces never become authored source. */
export function inheritWordEmptyParagraphText(word:WordMetadata,before:ParsedHtml,edit:DocumentEdit,parse:()=>ParsedHtml,styles:SourceStyleEdit):void{
    const current=parse();
    for(const markerId of current.nodeIds()){
        const marker=current.peekNode(markerId)!;if(marker.attributes['data-word-paragraph-mark']!=='true'||!marker.parentNodeId)continue;
        const id=marker.parentNodeId,previous=before.peekNode(id),paragraph=current.peekNode(id)!;if(!previous||!/^p$|^h[1-6]$/.test(previous.tagName)||paragraph.tagName!==previous.tagName||before.getNode(id)!.text||!current.getNode(id)!.text)continue;
        const oldRuns=previous.childNodeIds.filter(id=>before.peekNode(id)!.attributes['data-word-run']==='true');if(oldRuns.some(run=>current.getNode(run)?.text===current.getNode(id)!.text))continue;
        const attributes=Object.fromEntries(['class','style','lang','data-word-style','data-word-current-format','data-word-reset-text'].flatMap(key=>marker.attributes[key]!==undefined?[[key,marker.attributes[key]!]]:[]));
        for(const wrapper of edit.wrapWordParagraphText(id,attributes)){
        styles.aliasNodeIds(new Map([[marker.nodeId,wrapper]]));
        const native=word.nativeContent;if(native){const originals=native.fragments.filter(fragment=>fragment.nodeId===marker.nodeId&&fragment.kind==='property'&&!fragment.revisionPrevious&&!fragment.propertyPath.some(path=>path.endsWith('Change'))&&!fragment.propertyName?.endsWith('Change'));native.fragments.push(...originals.map((fragment,index)=>{const copy={...JSON.parse(JSON.stringify(fragment))as typeof fragment,nativeId:wrapper+'-mark-native-'+index,nodeId:wrapper,ownerName:'r',propertyPath:[W+' rPr',...fragment.propertyPath],...(fragment.propertyIndices?{propertyIndices:[0,...fragment.propertyIndices]}:{}),...(fragment.propertyName===W+' rPr'?{currentRunPropertiesOnly:true}:{})};delete copy.revisionSourceId;delete copy.revisionPrevious;return copy;}));}
        }
    }
}
