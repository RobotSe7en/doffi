import{copyWordClipboardSemantics}from'./word-clipboard-semantics.js';
import type{WordInlineContentSemantics}from'@doffi/core';
import {DEFAULT_DOCUMENT_READ_LIMITS,getWordTextDefaults,SourceStyleResolver,WordStyleSchema,DocumentError} from '@doffi/core';
import type {ContentDocumentModel,SourceStyleNode,WordMetadata,DeepReadonly} from '@doffi/core';
import{WORD_CLIPBOARD_BYTES,WORD_CLIPBOARD_FORMAT_ATTRIBUTE as marker,WORD_CLIPBOARD_NODE_ATTRIBUTE as identity}from'./word-clipboard-constants.js';
type Style=WordMetadata['styles'][number];
export interface WordClipboardStyleDefinition{style:Style;p:Record<string,string>;r:Record<string,string>}
export interface WordClipboardFormat{format:'doffi-word-clipboard';version:1|2;semantics?:WordInlineContentSemantics;styles:WordClipboardStyleDefinition[];nodes:Record<string,{style:string|null;class:string|null;styleId:string|null;sourceId?:string;attributes?:Record<string,string>}>}
const runKey=(key:string):boolean=>/^(?:font-|color$|background-color$|text-(?:decoration|transform)|vertical-align$|letter-spacing$|word-spacing$|--word-(?:font|run|kerning|language|bold|italic))/.test(key);
const bools=['font-weight','font-style','font-variant','text-transform'];
const declarations=(properties:Readonly<Record<string,string>>):string=>Object.entries(properties).map(([key,value])=>key+':'+value).join(';');
export function wordClipboardStyleDefinition(style:DeepReadonly<Style>,css:SourceStyleResolver):WordClipboardStyleDefinition{
    const own={...css.specifiedClass(style.cssClass),...(style.runCssClass?css.specifiedClass(style.runCssClass):{})},p=style.components?.p?css.specifiedClass(style.components.p):Object.fromEntries(Object.entries(own).filter(([key])=>style.paragraphProperties?.includes(key)??!runKey(key))),r={...(style.components?.r?css.specifiedClass(style.components.r):Object.fromEntries(Object.entries(own).filter(([key])=>style.runProperties?.includes(key)??runKey(key))))},resolved=style.runCssClass?css.classStyle(style.runCssClass,css.classStyle(style.cssClass)):css.classStyle(style.cssClass);
    for(const key of bools)if(r[key]!==undefined&&resolved[key]!==undefined)r[key]=resolved[key]!;
    return{style:WordStyleSchema.parse(style),p:{...p},r};
}
const encode=(value:unknown):string=>{const bytes=new TextEncoder().encode(JSON.stringify(value));let text='';for(let offset=0;offset<bytes.length;offset+=32768)text+=String.fromCharCode(...bytes.subarray(offset,offset+32768));return btoa(text);};

/** The standard HTML has authored font values; SDK readers also recover the
 * named style definitions. No unselected source text or runtime fonts enter it. */
export function copyWordClipboardFormat(holder:HTMLElement,model:ContentDocumentModel,contextId:string,completeNodes?:ReadonlySet<string>):void{
    const snapshot=model.snapshot(),word=snapshot.metadata.word;if(!word||holder.querySelector('table,[data-word-native-object]'))return;
    const block=[...holder.children].some(element=>element.matches('p,h1,h2,h3,h4,h5,h6,div,section,article,ul,ol'));
    const semantic=copyWordClipboardSemantics(holder,model,completeNodes);
    if(!block)for(const mark of holder.querySelectorAll('[data-word-paragraph-mark]'))mark.remove();
    const css=new SourceStyleResolver(snapshot.styles.map(sheet=>sheet.css),DEFAULT_DOCUMENT_READ_LIMITS,new Set(word.revisions.filter(item=>['insert','delete','moveFrom','moveTo'].includes(item.kind)).map(item=>item.nodeId))),cache=new Map<string,SourceStyleNode>();
    const project=(id:string):SourceStyleNode=>{let value=cache.get(id);if(!value){const node=model.getNode(id)!;value={tag:node.tagName,attrs:node.attributes,parent:node.parentNodeId?project(node.parentNodeId):undefined};cache.set(id,value);}return value;};
    const elements=[...holder.querySelectorAll<HTMLElement>('[data-node-id]')],classes=new Set(elements.flatMap(element=>[...element.classList])),used=new Set<string>(elements.flatMap(element=>element.dataset.wordStyle?[element.dataset.wordStyle]:[])),byId=new Map(word.styles.map(style=>[style.styleId,style]));
    for(const style of word.styles)if(classes.has(style.cssClass)||style.runCssClass&&classes.has(style.runCssClass))used.add(style.styleId);
    const pending=[...used];for(let index=0;index<pending.length;index++){const style=byId.get(pending[index]!);if(!style||!['paragraph','character'].includes(style.kind))return;for(const id of [style.basedOn,style.nextStyleId,style.linkedStyleId])if(id&&!used.has(id)){used.add(id);pending.push(id);}}
    const definitions:WordClipboardStyleDefinition[]=pending.map(id=>wordClipboardStyleDefinition(byId.get(id)!,css));
    const knownClasses=new Set(definitions.flatMap(item=>[item.style.cssClass,item.style.runCssClass,...Object.values(item.style.components??{})].filter((name):name is string=>!!name)));
    if(word.defaultRunCssClass)knownClasses.add(word.defaultRunCssClass);
    const defaults=word.defaultRunCssClass?css.specifiedClass(word.defaultRunCssClass):{},nodes:WordClipboardFormat['nodes']={};
    for(const[index,element]of elements.entries()){
        const id=element.dataset.nodeId!,key=String(index),sourceNode=project(id),remaining={...sourceNode,attrs:{...sourceNode.attrs,class:(sourceNode.attrs.class??'').split(/\s+/).filter(name=>!knownClasses.has(name)).join(' ')}},authored=[declarations(css.specified(remaining)),element.getAttribute('style')??''].filter(Boolean).join(';');
        nodes[key]={...(semantic.semantics?{sourceId:id,...(semantic.attributes.has(id)?{attributes:semantic.attributes.get(id)!}:{})}:{}),style:authored||null,class:[...element.classList].filter(name=>knownClasses.has(name)&&name!==word.defaultRunCssClass).join(' ')||null,styleId:element.getAttribute('data-word-style')};element.setAttribute(identity,key);
        const source=css.get(project(id));element.setAttribute('style',[element.getAttribute('style')??'',declarations(Object.fromEntries(Object.entries(source).filter(([property])=>runKey(property))))].filter(Boolean).join(';'));
        if(word.defaultRunCssClass&&element.classList.contains(word.defaultRunCssClass)&&!element.hasAttribute('data-word-paragraph-mark')){const parent=holder.ownerDocument.createElement('span');parent.dataset.wordRunOverlay='true';element.dataset.wordRunOverlay='true';parent.setAttribute('style',declarations(Object.fromEntries(Object.entries(defaults).filter(([property])=>property==='background-color'||property.startsWith('text-decoration')||property==='vertical-align'||property==='--word-run-shading'))));element.before(parent);parent.append(element);}
    }
    const wrapper=holder.ownerDocument.createElement(block?'div':'span');
    if(!block)wrapper.dataset.wordRunOverlay='true';
    wrapper.setAttribute('style',declarations(getWordTextDefaults(model,contextId)));
    wrapper.append(...holder.childNodes);holder.append(wrapper);
    const format:WordClipboardFormat={format:'doffi-word-clipboard',version:semantic.semantics?2:1,...(semantic.semantics?{semantics:semantic.semantics}:{}),styles:definitions,nodes};wrapper.setAttribute(marker,encode(format));
    if(holder.innerHTML.length>WORD_CLIPBOARD_BYTES)throw new DocumentError('ResourceLimitExceeded','复制格式超过剪贴板大小限制。');
}

