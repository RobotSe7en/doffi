import {createAssetDescriptor,DocumentError} from '@doffi/core';
import type {ContentDocumentModel,TransactionPrepareOptions} from '@doffi/core';

export const WORD_CLIPBOARD_BYTES:number=64*1024*1024;
const raster=/^image\/(png|jpeg|webp)$/;
/** Inline only the bytes already authorized by this model; never read a URL. */
export function copyClipboardImages(root:ParentNode,model:ContentDocumentModel):void{
    const snapshot=model.snapshot(),byNode=new Map(snapshot.metadata.word?.drawings?.map(item=>[item.nodeId,item]));let files:ReturnType<ContentDocumentModel['serialize']>|undefined,total=0;
    for(const image of root.querySelectorAll<HTMLImageElement>('img')){
        const url=new URL(image.getAttribute('src')!, 'https://doffi.invalid/'+snapshot.manifest.entrypoint),path=decodeURIComponent(url.pathname.slice(1)),asset=snapshot.manifest.assets.find(asset=>asset.path===path);
        if(url.origin!=='https://doffi.invalid'||!asset||!raster.test(asset.mimeType))throw new DocumentError('AssetUnavailable','复制图片需要当前文档已授权的栅格图片资源。');
        const bytes=model.getAsset?model.getAsset(path,WORD_CLIPBOARD_BYTES):(files??=model.serialize()).get(path);if(!bytes)throw new DocumentError('AssetUnavailable','复制图片的资源字节不可用。');
        total+=bytes.byteLength;if(total>WORD_CLIPBOARD_BYTES)throw new DocumentError('ResourceLimitExceeded','剪贴板图片超过 64 MiB。');
        let binary='';for(let offset=0;offset<bytes.length;offset+=32768)binary+=String.fromCharCode(...bytes.subarray(offset,offset+32768));image.setAttribute('src','data:'+asset.mimeType+';base64,'+btoa(binary));
        const drawing=byNode.get(image.dataset.nodeId??'');if(drawing){image.style.width=drawing.geometry.width+'pt';image.style.height=drawing.geometry.height+'pt';}
        image.removeAttribute('data-word-drawing');
    }
}
/** Decode local clipboard bytes before the transaction validates their MIME,
 * digest and final references. Text and every image commit in one undo step. */
export function pasteClipboardImages(root:ParentNode,model:ContentDocumentModel):TransactionPrepareOptions|undefined{
    const assets:NonNullable<TransactionPrepareOptions['assets']>[number][]=[],seen=new Set<string>(),entrypoint=model.snapshot().manifest.entrypoint;
    let total=0;
    for(const image of root.querySelectorAll<HTMLImageElement>('img')){
        const source=image.getAttribute('src')??'';if(!source.startsWith('data:'))continue;
        const match=/^data:(image\/(?:png|jpeg|webp));base64,([\s\S]*)$/i.exec(source);if(!match)throw new DocumentError('UnsupportedFeature','粘贴图片需要 PNG、JPEG 或 WebP 字节。');
        const encoded=match[2]!.replace(/[\t\r\n ]/g,'');if(encoded.length>Math.ceil(WORD_CLIPBOARD_BYTES/3)*4)throw new DocumentError('ResourceLimitExceeded','剪贴板图片超过 64 MiB。');
        let binary:string;try{binary=atob(encoded);}catch{throw new DocumentError('InvalidOperation','剪贴板图片的 Base64 数据无效。');}
        const bytes=Uint8Array.from(binary,character=>character.charCodeAt(0));total+=bytes.byteLength;if(total>WORD_CLIPBOARD_BYTES)throw new DocumentError('ResourceLimitExceeded','剪贴板图片超过 64 MiB。');
        const mime=match[1]!.toLowerCase(),extension=mime==='image/jpeg'?'jpg':mime.split('/')[1]!,asset=createAssetDescriptor('assets/clipboard.'+extension,mime,bytes),path='assets/clipboard-'+asset.sha256+'.'+extension,descriptor={...asset,path};
        if(!seen.has(path)){assets.push({descriptor,bytes});seen.add(path);}
        const base=entrypoint.split('/').slice(0,-1);image.setAttribute('src',[...base.map(()=>'..'),path].join('/'));image.removeAttribute('data-word-drawing');
    }
    return assets.length?{assets}:undefined;
}
