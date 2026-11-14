import { WordNoteOptionsSchema } from '@doffi/core';
import type { DeepReadonly,WordMetadata } from '@doffi/core';
import { attr,child,check,unsupported,number } from './ooxml/xml.js';
import type { XmlNode } from './ooxml/xml.js';
import { empty } from './export/xml.js';
type Options=NonNullable<WordMetadata['footnoteOptions']>;
export function readNoteOptions(source:XmlNode|undefined):Options|undefined{
 if(!source)return;check(source,['pos','numFmt','numStart','numRestart','footnote','endnote']);const position=attr(child(source,'pos'),'val'),format=attr(child(source,'numFmt'),'val'),start=attr(child(source,'numStart'),'val'),restart=attr(child(source,'numRestart'),'val');const parsed=WordNoteOptionsSchema.safeParse({...(position?{position}:{}),...(format?{format}:{}),...(start?{start:number(start,source)}:{}),...(restart?{restart}:{})});if(!parsed.success)unsupported(source,'This note numbering or placement is preserved natively.');return parsed.data;
}
export function writeNoteOptions(kind:'footnote'|'endnote',options:DeepReadonly<Options>|undefined,specialIds:readonly number[]=[]):string{
 if(!options&&!specialIds.length)return'';const val=(name:string,value:string|number|undefined):string=>value===undefined?'':empty('w:'+name,{'w:val':value});return'<w:'+kind+'Pr>'+val('pos',options?.position)+val('numFmt',options?.format)+val('numStart',options?.start)+val('numRestart',options?.restart)+specialIds.map(id=>empty('w:'+kind,{'w:id':id})).join('')+'</w:'+kind+'Pr>';
}
