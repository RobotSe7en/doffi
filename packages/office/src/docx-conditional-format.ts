import type {DeepReadonly,WordConditionalFormatValue,WordMetadata} from '@doffi/core';
import {WordConditionalFormatValueSchema} from '@doffi/core';
import {child,check,checkAttributes,attr,fail,W} from './ooxml/xml.js';
import type {XmlNode} from './ooxml/xml.js';
import {empty} from './export/xml.js';
const names=Object.keys(WordConditionalFormatValueSchema.shape);
export function readWordConditionalFormat(properties:XmlNode|undefined):WordConditionalFormatValue|undefined {
 const source=child(properties,'cnfStyle');if(!source)return undefined;
 check(source,[]);checkAttributes(source,names,W);const result:WordConditionalFormatValue={};
 const val=attr(source,'val');if(val!==undefined){if(!/^[01]{12}$/.test(val))fail(source,'A conditional format bit mask requires twelve binary digits.');result.val=val;}
 for(const name of names){if(name==='val')continue;const raw=attr(source,name);if(raw===undefined)continue;if(!['0','1','true','false','on','off'].includes(raw))fail(source,'A conditional format flag is not Boolean.');Reflect.set(result,name,['1','true','on'].includes(raw));}
 return result;
}
export function bindWordConditionalFormat(word:WordMetadata,properties:XmlNode|undefined,nodeId:string):void {
 const format=readWordConditionalFormat(properties);if(format){word.conditionalFormats??=[];word.conditionalFormats.push({nodeId,format});}
}
export function writeWordConditionalFormat(format:DeepReadonly<WordConditionalFormatValue>|undefined):string {
 if(!format)return'';return empty('w:cnfStyle',Object.fromEntries(Object.entries(format).map(([key,value])=>['w:'+key,typeof value==='boolean'?Number(value):value])));
}
