import type { ContentDocumentModel } from '@doffi/core';
import { DocumentError } from '@doffi/core';
import type { RenderLimits } from './contracts.js';
import { markViewNode, ownViewAttribute, ownViewStyle } from './view-nodes.js';
/** A native object's raster cache is a view of its preserved XML, never an edit target. */
export function composeWordNative(document: Document, model: ContentDocumentModel, urls: ReadonlyMap<string,string>, limits: Readonly<RenderLimits>): void {
    const nodes=new Map([...document.querySelectorAll<HTMLElement>('[data-node-id]')].map(node=>[node.dataset.nodeId!,node]));
    let count=document.querySelectorAll('*').length;
    for(const fragment of model.snapshot().metadata.word?.nativeContent?.fragments??[]){
        if(!fragment.nodeId||!fragment.preview)continue;
        const node=nodes.get(fragment.nodeId),url=urls.get(fragment.preview.assetPath);if(!node||!url)continue;
        if(++count>limits.maxNodes)throw new DocumentError('ResourceLimitExceeded','Native caches exceed the render node budget.');
        const image=markViewNode(document.createElement('img'));image.src=url;image.alt=node.title||'保留的原生内容';image.draggable=false;
        image.style.cssText='display:inline-block;max-width:100%;vertical-align:middle;pointer-events:none';
        if(fragment.preview.width)image.style.width=fragment.preview.width+'pt';
        if(fragment.preview.height)image.style.height=fragment.preview.height+'pt';
        ownViewAttribute(node,'data-sdk-native-preview');node.dataset.sdkNativePreview='true';
        ownViewStyle(node,'min-width','min-height','border','background','font-size','line-height');
        node.style.minWidth='0';node.style.minHeight='0';node.style.border='0';node.style.background='transparent';node.style.fontSize='0';node.style.lineHeight='0';
        node.append(image);
    }
}
