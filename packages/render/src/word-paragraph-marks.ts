import type {ContentDocumentModel,SourceStyleNode} from '@doffi/core';
import {createSourceStyleResolver} from './source-style-probe.js';
import {ownViewAttribute,ownViewStyle} from './view-nodes.js';
/** Empty lines and formatting symbols use their authored paragraph-mark run.
 * The paragraph's source paragraph formatting remains unchanged. */
export function composeWordParagraphMarks(document:Document,model:ContentDocumentModel):void{
    const markers=[...document.querySelectorAll<HTMLElement>('[data-word-paragraph-mark="true"]')];if(!markers.length)return;
    const resolver=createSourceStyleResolver(document),cache=new Map<string,SourceStyleNode>(),project=(id:string):SourceStyleNode=>{let value=cache.get(id);if(!value){const node=model.getNode(id)!;value={tag:node.tagName,attrs:node.attributes,parent:node.parentNodeId?project(node.parentNodeId):undefined};cache.set(id,value);}return value;};
    for(const marker of markers){
        const id=marker.dataset.nodeId,paragraph=marker.parentElement;if(!id||!paragraph?.dataset.nodeId)continue;
        const css=resolver.get(project(id));ownViewAttribute(marker,'contenteditable');marker.contentEditable='false';
        for(const key of ['font-family','font-size','font-weight','font-style','color','text-decoration-line','text-decoration-style','vertical-align'])if(css[key]!==undefined){const variable='--sdk-word-mark-'+key;ownViewStyle(paragraph,variable);paragraph.style.setProperty(variable,css[key]!);}
        if(!model.getNode(paragraph.dataset.nodeId)!.text)for(const key of ['font-family','font-size','font-weight','font-style','color','line-height'])if(css[key]!==undefined){ownViewStyle(paragraph,key);paragraph.style.setProperty(key,css[key]!);}
    }
}
