import { DocumentError } from '@doffi/core';
import { attr,child,number,fail,unsupported } from './ooxml/xml.js';
import type { XmlNode } from './ooxml/xml.js';
import { empty } from './export/xml.js';
const width=/^[+-]?(?:\d+(?:\.\d+)?|\.\d+)(?:mm|cm|in|pt|pc|pi|%)?$/;
export const rowGapWidthPrefix=(side:'before'|'after'):string=>'data-word-'+side+'-width-';
export function readWordRowGapWidths(properties:XmlNode|undefined):Record<string,string>{
    const result:Record<string,string>={};
    for(const side of ['before','after']as const){const node=child(properties,side==='before'?'wBefore':'wAfter');if(!node)continue;const prefix=rowGapWidthPrefix(side),raw=attr(node,'w'),type=attr(node,'type');
        if(raw!==undefined&&!width.test(raw))fail(node,'A row gap width is malformed.');
        if(type!==undefined&&!['auto','nil','dxa','pct'].includes(type))unsupported(node,'This row gap width unit remains native.');
        result[prefix+'present']='true';if(raw!==undefined)result[prefix+'value']=raw;if(type!==undefined)result[prefix+'type']=type;
    }
    return result;
}
export function writeWordRowGapWidths(attributes:Readonly<Record<string,string>>):string{
    let result='';for(const side of ['before','after']as const){const prefix=rowGapWidthPrefix(side);if(attributes[prefix+'present']!=='true')continue;const raw=attributes[prefix+'value'],type=attributes[prefix+'type'];if(raw!==undefined&&!width.test(raw)||type!==undefined&&!['auto','nil','dxa','pct'].includes(type))throw new DocumentError('UnsupportedFeature','Invalid native row gap width.');result+=empty('w:'+(side==='before'?'wBefore':'wAfter'),{'w:w':raw,'w:type':type});}return result;
}
export function readWordRowGapCount(properties:XmlNode|undefined,side:'before'|'after'):number{
    const node=child(properties,side==='before'?'gridBefore':'gridAfter');if(!node)return 0;const result=number(attr(node,'val'),node);if(!Number.isSafeInteger(result)||result<0)fail(node,'A row gap requires a nonnegative integer grid count.');return result;
}
