import { createAssetDescriptor, DocumentError } from '@doffi/core';
import type { ContentDocumentManifest, WordMetadata } from '@doffi/core';
import type { OpcPackage } from './ooxml/package.js';
import { W, R, MC, WP, unsupportedSource } from './ooxml/xml.js';
import type { XmlNode,UnsupportedXmlSource } from './ooxml/xml.js';
import { serializeXml } from './ooxml/serialize.js';

type Native = NonNullable<WordMetadata['nativeContent']>;
interface Pending { indices?:number[]; propertyIndex?:number; revisionSourceId?:string;revisionPrevious?:boolean; owner?: XmlNode; original: XmlNode; kind: 'inline' | 'block' | 'property'; runChild?:boolean; path: string[]; label: string; id: string }
const owners = new Set(['p', 'r', 'sdt', 'tbl', 'tr', 'tc', 'sectPr', 'style', 'settings', 'numbering', 'fonts', 'document', 'commentEx', 'commentsEx', 'commentsExtensible', 'commentsIds','commentId','commentExtensible','people','person','footnotes','endnotes','footnote','endnote']);
const propertyRoots = new Set(['sdtPr','sdtEndPr','fldChar','pPr', 'rPr', 'tblPr', 'trPr', 'tcPr', 'sectPr', 'settings', 'style', 'numbering', 'fonts', 'commentEx', 'commentsEx', 'commentsExtensible', 'commentsIds','commentId','commentExtensible','people','person','footnotes','endnotes','footnote','endnote']);
const generatedRelationships = new Set(['styles', 'numbering', 'settings', 'comments', 'commentsExtended','commentsIds','commentsExtensible','people', 'header', 'footer', 'footnotes', 'endnotes', 'image', 'chart', 'hyperlink']);
export class DocxNativeImport {
  private readonly pending: Pending[] = [];
  private readonly bindings = new Map<XmlNode, string>();
  private readonly placeholders = new Map<XmlNode, Pending>();
  private readonly parents = new Map<XmlNode, XmlNode>();
  private readonly roots = new Map<string, XmlNode>();
  private readonly suppressed = new Set<string>();
  private readonly recovered = new WeakMap<XmlNode, Set<string>>();
  private serial = 0;
  private readonly detached=new WeakSet<XmlNode>();private readonly batched=new Map<string,WeakSet<XmlNode>>();
  private readonly identities=new Map<string,NonNullable<Native['identities']>[number]>();
  private readonly ordinals=new WeakMap<XmlNode,number>();
  constructor(readonly original: OpcPackage, readonly main: string) {}
  initialize(): OpcPackage {
    const view: OpcPackage = {
      ...this.original,
      read: part => {
        const root = this.original.read(part);
        if (!this.roots.has(part)) { this.roots.set(part, root); this.index(root); }
        return root;
      },
      related: (part, type) => this.suppressed.has(part + '\n' + type) ? undefined : this.original.related(part, type),
    };
    return view;
  }
  private index(root: XmlNode): void {
    const counts=new Map<string,number>();for (const item of root.content) if (typeof item !== 'string') {const key=item.uri+' '+item.name,ordinal=counts.get(key)??0;this.ordinals.set(item,ordinal);counts.set(key,ordinal+1);this.parents.set(item, root); this.index(item); }
  }
  bind(source: XmlNode, attrs: Record<string, string>): void {
    this.bindings.set(source, attrs['data-node-id']!);
    const placeholder = this.placeholders.get(source);
    if (placeholder) { attrs['data-word-native-object'] = placeholder.id; attrs.title = placeholder.label; }
  }
  identity(source:XmlNode,nodeId:string,sourceId:string):void{const kind=source.uri===WP?'frame':'object';this.identities.set(nodeId+'\n'+kind,{nodeId,sourcePart:source.part,sourceId,kind});}
  beginAttempt(): void { this.bindings.clear();this.identities.clear(); }
  isPlaceholder(node:XmlNode):boolean{return this.placeholders.has(node);}
  recover(error: unknown): boolean {
    const info = unsupportedSource(error);
    if (!info) return false;
    const targetParent=this.parents.get(info.node),alternate=info.node.uri===MC&&info.node.name==='AlternateContent';
    if(info.elementType||alternate){const supported=new Set([W,R,WP,'http://schemas.openxmlformats.org/drawingml/2006/main','http://schemas.openxmlformats.org/drawingml/2006/picture','http://schemas.microsoft.com/office/word/2010/wordprocessingShape','http://schemas.microsoft.com/office/word/2010/wordprocessingGroup']);const unknownAlternative=(node:XmlNode):boolean=>node.uri===MC&&node.name==='AlternateContent'&&node.content.some(child=>typeof child!=='string'&&child.uri===MC&&child.name==='Choice'&&(child.attrs.get(' Requires')??'').split(/\s+/).filter(Boolean).some(prefix=>!supported.has(child.namespaces[prefix]??'')));
      if(info.elementType||unknownAlternative(info.node)){const key=['element',info.node.uri,info.node.name,targetParent?.uri??'',targetParent?.name??''].join('\n'),seen=this.batched.get(key)??new WeakSet<XmlNode>();this.batched.set(key,seen);const targets:XmlNode[]=[];const scan=(node:XmlNode):void=>{const parent=this.parents.get(node);if(node.uri===info.node.uri&&node.name===info.node.name&&parent?.uri===targetParent?.uri&&parent?.name===targetParent?.name&&(!alternate||unknownAlternative(node)))targets.push(node);for(const child of node.content)if(typeof child!=='string')scan(child);};for(const root of this.roots.values())if(!seen.has(root)){seen.add(root);scan(root);}if(targets.length){for(const node of targets)if(!this.isDetached(node)&&!this.recoverSource({node},error))return false;return true;}}
    }
    const collection=info.collection,parent=collection?this.parents.get(collection.owner):undefined;
    if(collection&&collection.owner.uri===W&&propertyRoots.has(collection.owner.name)){
      const key=[collection.owner.uri,collection.owner.name,parent?.uri??'',parent?.name??'',collection.uri,...collection.names].join('\n'),seen=this.batched.get(key)??new WeakSet<XmlNode>();this.batched.set(key,seen);const targets:XmlNode[]=[];
      const scan=(node:XmlNode):void=>{const owner=this.parents.get(node);if(node.uri===collection.owner.uri&&node.name===collection.owner.name&&owner?.uri===parent?.uri&&owner?.name===parent?.name)for(const child of node.content)if(typeof child!=='string'&&(child.uri!==collection.uri||!collection.names.includes(child.name)))targets.push(child);for(const child of node.content)if(typeof child!=='string')scan(child);};
      for(const root of this.roots.values())if(!seen.has(root)){seen.add(root);scan(root);}
      if(targets.length){for(const node of targets)if(!this.isDetached(node)&&!this.recoverSource({node},error))return false;return true;}
    }
    return this.recoverSource(info,error);
  }
  private isDetached(node:XmlNode):boolean{for(let current:XmlNode|undefined=node;current;current=this.parents.get(current))if(this.detached.has(current))return true;return false;}
  private recoverSource(info:UnsupportedXmlSource,error:unknown):boolean{
    const key = info.attribute ?? "#node";
    const recovered = this.recovered.get(info.node) ?? new Set<string>();
    if (recovered.has(key)) throw new DocumentError('UnsupportedFeature', 'A native feature could not be isolated without changing its surrounding content.', [{ code:'NativeRecoveryStalled',path:[info.node.part,info.node.name],message:'Native compatibility recovery made no progress.' }]);
    recovered.add(key); this.recovered.set(info.node, recovered);
    let target = info.node;
    const chain: XmlNode[] = [target];
    for (let parent = this.parents.get(target); parent; parent = this.parents.get(parent)) chain.push(parent);
    const change=chain.find(node=>node.uri===W&&/^(?:rPr|pPr|tblPr|trPr|tcPr|sectPr|tblGrid)Change$/.test(node.name)),revisionInfo=change?.attrs.get(W+' id')?{revisionSourceId:change.attrs.get(W+' id')!,revisionPrevious:chain.indexOf(change)>0}:{};
    const mergeContinuation = error instanceof Error && /vertical merge continuation/i.test(error.message);
    const root = chain.at(-1)!;
    if (root.uri !== W && !['commentsEx','commentsExtensible','commentsIds','people'].includes(root.name)) {
      const relation = [...this.original.relationships(this.main).values()].find(rel => !rel.external && rel.target === root.part);
      if (relation) { const key=this.main+'\n'+relation.type.slice(relation.type.lastIndexOf('/')+1); if(!this.suppressed.has(key)){this.suppressed.add(key);return true;} }
    }
    const alternate=chain.find(node=>node.uri===MC&&node.name==='AlternateContent');
    const propertyIndex = mergeContinuation ? -1 : chain.findIndex((n,index) => n.uri === W && (propertyRoots.has(n.name) || index === 1 && n.name === 'document') || ['commentEx','commentsEx','commentsExtensible','commentsIds','commentId','commentExtensible','people','person'].includes(n.name));
    // A drawing is indivisible until it has an editable geometry implementation.
    const atomicIndex = chain.findIndex(n => n.uri === W && ['drawing','pict','object','sdt'].includes(n.name) || n.uri==='http://schemas.openxmlformats.org/officeDocument/2006/math'&&['oMath','oMathPara'].includes(n.name));
    const drawing=atomicIndex<0?undefined:chain[atomicIndex];
    if (propertyIndex >= 0 && (!alternate||propertyIndex>chain.indexOf(alternate)) && (!drawing || drawing.name==='sdt'&&propertyIndex<atomicIndex)) {
      const owner = chain.slice(propertyIndex).find(n => (n.uri === W || ['commentEx','commentsEx','commentsExtensible','commentsIds','commentId','commentExtensible','people','person'].includes(n.name)) && owners.has(n.name)) ?? chain.at(-1)!;
      const path: string[] = [],indices:number[]=[];
      for (let current = target === owner ? undefined : this.parents.get(target); current && current !== owner; current = this.parents.get(current)) {path.unshift(current.uri + ' ' + current.name);indices.unshift(this.ordinals.get(current)??0);}
      if (info.attribute) {
        const attrs = new Map<string, string>([[info.attribute, target.attrs.get(info.attribute)!]]);
        this.pending.push({ ...revisionInfo,owner, original: { ...target, attrs, content: [] }, kind: 'property', path,indices,propertyIndex:this.ordinals.get(target)??0, label: '原生属性 ' + target.name, id: 'native-' + (++this.serial) });
        (target.attrs as Map<string, string>).delete(info.attribute);
      } else {
        const parent = this.parents.get(target);
        if (!parent || parent.content.indexOf(target) < 0) return false;
        this.pending.push({ ...revisionInfo,owner, original: target, kind: 'property', path,indices,propertyIndex:this.ordinals.get(target)??0, label: '原生属性 ' + target.name, id: 'native-' + (++this.serial) });
        this.detached.add(target);parent.content.splice(parent.content.indexOf(target), 1);
      }
      return true;
    }
    if(alternate)target=alternate;
    else if (mergeContinuation) target = chain.find(n => n.uri === W && n.name === 'tbl') ?? target;
    else if (drawing) target = drawing;
    else if(chain.some(n=>n.uri===W&&n.name==='r'))target=info.node;
    else if (chain.some(n => n.uri === W && n.name === 'tbl') && !['p','r'].includes(target.name)) target = chain.find(n => n.uri === W && n.name === 'tbl')!;
    const parent = this.parents.get(target);
    if (!parent || parent.content.indexOf(target) < 0 || this.placeholders.has(target)) return false;
    const inline = parent.uri === W && !['body','tc','sdtContent','hdr','ftr','comment'].includes(parent.name);
    // A larger indivisible object absorbs properties already recovered inside it.
    for (let index = this.pending.length - 1; index >= 0; index--) {
      const previous = this.pending[index]!; let inside = false;
      for (let current = previous.owner; current; current = this.parents.get(current)) if (current === target) inside = true;
      if (!inside || previous.kind !== 'property' || !previous.owner) continue;
      let owner = previous.owner;
      for (const path of previous.path) { const found = owner.content.find((n): n is XmlNode => typeof n !== 'string' && n.uri + ' ' + n.name === path); if (found) owner = found; }
      owner.content.push(previous.original); this.pending.splice(index, 1);
    }
    const pending: Pending = { original: target, kind: inline ? 'inline' : 'block', ...(parent.uri===W&&parent.name==='r'?{runChild:true}:{}), path: [], label: '原生内容：' + target.name, id: 'native-' + (++this.serial) };
    const leaf: XmlNode = { name: 't', uri: W, attrs: new Map(), namespaces: {}, content: ['\ufffc'], part: target.part };
    const run: XmlNode = { ...leaf, name: 'r', content: [leaf] };
    const replacement: XmlNode = inline ? run : { ...leaf, name: 'p', content: [run] };
    pending.owner = replacement; this.pending.push(pending); this.placeholders.set(replacement, pending);
    this.detached.add(target);parent.content.splice(parent.content.indexOf(target), 1, replacement); this.parents.set(replacement, parent); this.index(replacement);
    return true;
  }
  finish(files: Map<string, Uint8Array>, manifest: ContentDocumentManifest, word: WordMetadata): void {
    // Settings outside the editable model still belong to the native package.
    // Capture them directly once, rather than requiring one recovery pass per
    // harmless producer preference or compatibility flag.
    // Unmapped extension attributes on supported Word nodes remain native data.
    const XMLNS='http://www.w3.org/2000/xmlns/',XML='http://www.w3.org/XML/1998/namespace',W14='http://schemas.microsoft.com/office/word/2010/wordml';
    const capture=(node:XmlNode):void=>{
      if(node.uri===W){const attrs=new Map([...node.attrs].filter(([key])=>{const split=key.lastIndexOf(' '),uri=key.slice(0,split),name=key.slice(split+1);if(uri===MC&&name==='Ignorable'&&(node.attrs.get(key)??'').split(/\s+/).filter(Boolean).every(prefix=>[W,R,WP,MC,W14,'http://schemas.openxmlformats.org/drawingml/2006/main','http://schemas.openxmlformats.org/drawingml/2006/picture','http://schemas.microsoft.com/office/word/2010/wordprocessingShape','http://schemas.microsoft.com/office/word/2010/wordprocessingGroup'].includes(node.namespaces[prefix]??'')))return false;return uri&&![W,R,XMLNS,XML].includes(uri)&&!(uri===W14&&['paraId','textId'].includes(name))||uri===W&&this.bindings.has(node)&&['p','r','tbl','tr','tc','sectPr','sdt'].includes(node.name);}));
      if(attrs.size){let owner:XmlNode|undefined=node;while(owner&&!this.bindings.has(owner)&&!['style','styles','settings','numbering','fonts','document','comments','footnotes','endnotes'].includes(owner.name))owner=this.parents.get(owner);if(owner){const path:string[]=[],indices:number[]=[];for(let current=node===owner?undefined:this.parents.get(node);current&&current!==owner;current=this.parents.get(current)){path.unshift(current.uri+' '+current.name);indices.unshift(this.ordinals.get(current)??0);}let change:XmlNode|undefined;for(let current:XmlNode|undefined=node;current&&current!==owner;current=this.parents.get(current))if(current.uri===W&&current.name.endsWith('PrChange')){change=current;break;}this.pending.push({owner,original:{...node,attrs,content:[]},kind:'property',path,indices,propertyIndex:this.ordinals.get(node)??0,label:'原生扩展属性 '+node.name,id:'native-'+(++this.serial),...(change?.attrs.get(W+' id')?{revisionSourceId:change.attrs.get(W+' id')!,revisionPrevious:change!==node}:{})});}}}
      for(const child of node.content)if(typeof child!=='string')capture(child);
    };for(const root of this.roots.values())capture(root);
    const writtenSettings=new Set(['bordersDoNotSurroundHeader','bordersDoNotSurroundFooter','mirrorMargins','evenAndOddHeaders','defaultTabStop','autoHyphenation','gutterAtTop','trackRevisions','footnotePr','endnotePr']);
    for(const root of this.roots.values())if(root.uri===W&&root.name==='settings')for(const item of root.content)if(typeof item!=='string'&&!(item.uri===W&&(writtenSettings.has(item.name)||item.name==='updateFields'&&['0','false','off'].includes(item.attrs.get(W+' val')??''))))this.pending.push({owner:root,original:item,kind:'property',path:[],label:'原生设置 '+item.name,id:'native-'+(++this.serial)});
    const native: Native = { mainPart: this.main, fragments: [], parts: [], relationships: [] };
    const encoder = new (globalThis as unknown as { TextEncoder: new () => { encode(text: string): Uint8Array } }).TextEncoder(), parts = new Set<string>();
    const add = (bytes: Uint8Array, mime: string): string => {
      const path = 'assets/native-' + (manifest.assets.length + 1) + (/(?:\/|\+)xml(?:;|$)/i.test(mime) ? '.xml' : '.bin');
      files.set(path, bytes); manifest.assets.push(createAssetDescriptor(path, mime, bytes)); return path;
    };
    const relationships = (sourcePart: string): void => {
      for (const rel of this.original.relationships(sourcePart).values()) {
        if (!native.relationships.some(r => r.sourcePart === sourcePart && r.id === rel.id)) native.relationships.push({ sourcePart, ...rel });
        if (!rel.external) part(rel.target);
      }
    };
    const part = (sourcePart: string): void => {
      const root=this.roots.get(sourcePart),generated=root?.uri===W&&['document','styles','settings','numbering','comments','hdr','ftr','footnotes','endnotes'].includes(root.name)||['commentsEx','commentsExtensible','commentsIds','people'].includes(root?.name??'');
      if (parts.has(sourcePart) || sourcePart === this.main || generated) return;
      parts.add(sourcePart);
      const bytes = this.original.files.get(sourcePart);
      if (!bytes) throw new DocumentError('InvalidDocument', 'A preserved native dependency is missing.');
      const type = this.original.contentType(sourcePart) ?? 'application/octet-stream';
      if (/(?:\/|\+)xml(?:;|$)/i.test(type) || sourcePart.endsWith('.xml')) this.original.read(sourcePart);
      native.parts.push({ sourcePart, assetPath: add(bytes, type), contentType: type }); relationships(sourcePart);
    };
    const previews=new Map<string,string>();
    const preview=(root:XmlNode):Native['fragments'][number]['preview']=>{
      const fallback=root.uri===MC&&root.name==='AlternateContent'?root.content.find((node):node is XmlNode=>typeof node!=='string'&&node.uri===MC&&node.name==='Fallback'):undefined;
      const nodes:XmlNode[]=[],scan=(node:XmlNode):void=>{nodes.push(node);for(const child of node.content)if(typeof child!=='string')scan(child);};scan(fallback??root);
      for(const node of nodes){const id=node.attrs.get(R+' embed')??(node.name==='imagedata'?node.attrs.get(R+' id'):undefined);if(!id)continue;const relation=this.original.relationships(node.part).get(id);if(!relation||relation.external)continue;const mime=this.original.contentType(relation.target),bytes=this.original.files.get(relation.target);if(!mime||!bytes||!['image/png','image/jpeg','image/webp'].includes(mime))continue;let assetPath=previews.get(relation.target);if(!assetPath){assetPath=add(bytes,mime);previews.set(relation.target,assetPath);}const extent=nodes.find(item=>item.uri===WP&&item.name==='extent'),width=Number(extent?.attrs.get(' cx'))/12700,height=Number(extent?.attrs.get(' cy'))/12700;return{assetPath,...(Number.isFinite(width)&&width>0?{width}:{}),...(Number.isFinite(height)&&height>0?{height}:{})};}
      return undefined;
    };
    const fragmentXml: [string,string][] = [];
    const fragmentAssetPath = 'assets/native-fragments.json';
    for (const item of this.pending) {
      const nodeId = item.owner ? this.bindings.get(item.owner) : undefined;
      if (!nodeId && item.kind !== 'property') throw new DocumentError('InvalidDocument', 'A native content placeholder is not bound to the document.');
      const assetPath = fragmentAssetPath,cached=item.kind!=='property'?preview(item.original):undefined;
      fragmentXml.push([item.id,serializeXml(item.original)]);
      native.fragments.push({ nativeId: item.id,...(cached?{preview:cached}:{}), ...(nodeId ? { nodeId } : {}), kind: item.kind, ...(item.runChild?{runChild:true}:{}), sourcePart: item.original.part, assetPath, fragmentKey: item.id, propertyPath: item.path,...(item.indices?{propertyIndices:item.indices,propertyIndex:item.propertyIndex}:{}), ownerName: item.owner?.name ?? "document", ...(item.owner?.name === "style" && item.owner.attrs.get(W + " styleId") ? { styleId: item.owner.attrs.get(W + " styleId")! } : {}), label: item.label,propertyName:item.original.uri+' '+item.original.name,...(item.revisionSourceId?{revisionSourceId:item.revisionSourceId,revisionPrevious:item.revisionPrevious}: {}) });
      const binding=native.fragments[native.fragments.length-1]!;binding.fragmentFormat='json';binding.relationshipIds=[];
      // Relationships are package-local data; preserving never follows external links.
      const visit = (node: XmlNode): void => {
        for (const [key, id] of node.attrs) if (key.startsWith(R + ' ')) {
          const rel = this.original.relationships(node.part).get(id);
          if (!rel) throw new DocumentError('InvalidDocument', 'A native fragment refers to a missing relationship.');
          if(!binding.relationshipIds!.includes(id))binding.relationshipIds!.push(id);
          if (!native.relationships.some(r => r.sourcePart === node.part && r.id === rel.id)) native.relationships.push({ sourcePart: node.part, ...rel });
          if (!rel.external) part(rel.target);
        }
        for (const child of node.content) if (typeof child !== 'string') visit(child);
      };
      visit(item.original);
    }
    if (fragmentXml.length) {
      const bytes = encoder.encode(JSON.stringify({format:'doffi-word-native-fragments',version:1,fragments:Object.fromEntries(fragmentXml)}));
      files.set(fragmentAssetPath, bytes); manifest.assets.push(createAssetDescriptor(fragmentAssetPath, 'application/json', bytes));
    }
    for (const rel of this.original.relationships(this.main).values()) {
      const short = rel.type.slice(rel.type.lastIndexOf('/') + 1);
      if (!generatedRelationships.has(short)) {
        if (!native.relationships.some(r => r.sourcePart === this.main && r.id === rel.id)) native.relationships.push({ sourcePart: this.main, ...rel, root: true });
        if (!rel.external) part(rel.target);
      }
    }
    if (native.fragments.length || native.parts.length||native.relationships.length){native.identities=[...this.identities.values()];word.nativeContent = native;}
  }
}
