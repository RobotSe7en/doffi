import { DocumentError } from '@doffi/core';
import type { ContentDocumentModel, DocumentReadLimits, DeepReadonly, WordMetadata } from '@doffi/core';
import { W, R, PKG, xml } from './ooxml/xml.js';
import type { XmlNode } from './ooxml/xml.js';
import { serializeXml } from './ooxml/serialize.js';
import { declaration, escapeXml } from './export/xml.js';
export const NATIVE_MARKER = 'urn:doffi:transient:native';
export const nativeMarker = (nodeId: string | undefined): string => nodeId ? ' xmlns:dn="' + NATIVE_MARKER + '" dn:node="' + escapeXml(nodeId) + '"' : '';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const codecs = globalThis as unknown as { TextEncoder: new () => { encode(text: string): Uint8Array } };
function relPath(part: string): string { const slash = part.lastIndexOf('/'); return part ? part.slice(0, slash + 1) + '_rels/' + part.slice(slash + 1) + '.rels' : '_rels/.rels'; }
function name(node: XmlNode): string { return node.uri + ' ' + node.name; }
function merge(target: XmlNode, source: XmlNode): void {
  for (const [key, value] of source.attrs) (target.attrs as Map<string, string>).set(key, value);
  if (source.content.length) target.content.splice(0, target.content.length, ...source.content);
}
/** Reattach residual native data to the just-compiled current source. */
export function restoreDocxNative(model: ContentDocumentModel, generated: ReadonlyMap<string, Uint8Array>, limits: DocumentReadLimits): ReadonlyMap<string, Uint8Array> {
  const native = model.snapshot().metadata.word?.nativeContent;
  if (!native) return generated;
  const files = new Map(generated), sourceFiles = model.serialize(), encoder = new codecs.TextEncoder();
  const budget = { nodes: 0 }, parsed = new Map<string, XmlNode>();
  const read = (bytes: Uint8Array, path: string): XmlNode => xml(bytes, path, limits, budget);
  const fragmentAssets = new Map<string, XmlNode>();
  const fragmentEntries = new Map<string, Map<string, XmlNode>>();
  const readFragment = (fragment: typeof native.fragments[number]): XmlNode => {
    let root = fragmentAssets.get(fragment.assetPath);
    if (!root) { const bytes=sourceFiles.get(fragment.assetPath);if(!bytes)throw new DocumentError('AssetUnavailable','A native fragment is missing.');root=read(bytes,fragment.assetPath);fragmentAssets.set(fragment.assetPath,root); }
    if (!fragment.fragmentKey) return root;
    let entries = fragmentEntries.get(fragment.assetPath);
    if (!entries) { entries = new Map(); for (const item of root.content) if(typeof item!=='string' && item.uri==='urn:doffi:word:native-fragments' && item.name==='item') { const key=item.attrs.get(' id'),child=item.content.find((node):node is XmlNode=>typeof node!=='string'); if(!key||!child||entries.has(key))throw new DocumentError('InvalidDocument','The native fragment collection is invalid.');entries.set(key,child); } fragmentEntries.set(fragment.assetPath,entries); }
    const node=entries.get(fragment.fragmentKey);if(!node)throw new DocumentError('InvalidDocument','A native fragment collection entry is missing.');return node;
  };
  const nodeParts = new Map<string, {path:string;name:string}[]>();
  for (const [path, bytes] of files) if (path.endsWith('.xml') && path !== '[Content_Types].xml') {
    const root = read(bytes, path); parsed.set(path, root);
    const scan = (node: XmlNode): void => {
      for(const id of [node.attrs.get(NATIVE_MARKER+' node'),node.attrs.get(NATIVE_MARKER+' container')])if(id){const locations=nodeParts.get(id)??[];locations.push({path,name:node.name});nodeParts.set(id,locations);}
      for (const item of node.content) if (typeof item !== 'string') scan(item);
    };
    scan(root);
  }
  const partMap = new Map<string, string>([[native.mainPart, 'word/document.xml']]);
  for (const fragment of native.fragments) {
    const locations=fragment.nodeId?nodeParts.get(fragment.nodeId):undefined;
    const mapped = locations?.find(location=>location.name===fragment.ownerName)?.path??locations?.[0]?.path;
    if (mapped) partMap.set(fragment.sourcePart, mapped);
    else if (fragment.kind === 'property' && !fragment.nodeId) {
      const rootNames: Record<string,string> = { settings: 'settings', style: 'styles', styles: 'styles', numbering: 'numbering', fonts: 'fonts', commentsEx: 'commentsEx', commentsExtensible: 'commentsExtensible', commentsIds: 'commentsIds',people:'people' };
      const rootName = rootNames[fragment.ownerName];
      const entry = rootName ? [...parsed].find(([,root]) => root.name === rootName) : undefined;
      if (entry) partMap.set(fragment.sourcePart, entry[0]);
    }
  }
  const requiredParts = new Set<string>(), requiredRelations = new Set<string>();
  const requireRelation = (part: string, id: string): void => {
    const key=part+'\n'+id;if(requiredRelations.has(key))return;
    const relation=native.relationships.find(r=>r.sourcePart===part&&r.id===id);if(!relation)return;
    requiredRelations.add(key);
    if(!relation.external){requiredParts.add(relation.target);for(const child of native.relationships)if(child.sourcePart===relation.target)requireRelation(child.sourcePart,child.id);}
  };
  for(const relation of native.relationships)if(relation.root)requireRelation(relation.sourcePart,relation.id);
  for(const fragment of native.fragments){
    const root=readFragment(fragment);const collect=(node:XmlNode):void=>{for(const[key,value]of node.attrs)if(key.startsWith(R+' '))requireRelation(fragment.sourcePart,value);for(const child of node.content)if(typeof child!=='string')collect(child);};collect(root);
  }
  const retained = native.parts.filter(part => requiredParts.has(part.sourcePart) && !partMap.has(part.sourcePart));
  retained.forEach((part,index) => partMap.set(part.sourcePart, 'word/preserved/part-' + (index + 1) + (part.sourcePart.match(/\.[^.\/]+$/)?.[0] ?? '.bin')));
  const relationMaps = new Map<string, Map<string, string>>();
  const relationRoots = new Map<string, XmlNode>();
  const relationshipRoot = (part: string): XmlNode => {
    let root = relationRoots.get(part);
    if (!root) { const bytes = files.get(relPath(part)); root = bytes ? read(bytes, relPath(part)) : { name:'Relationships', uri:PKG, attrs:new Map(), namespaces:{}, content:[], part:relPath(part) }; relationRoots.set(part,root); }
    return root;
  };
  for (const rel of native.relationships) {
    if(!requiredRelations.has(rel.sourcePart+'\n'+rel.id))continue;
    const outputPart = partMap.get(rel.sourcePart); if (!outputPart) continue;
    const target = rel.external ? rel.target : partMap.get(rel.target);
    if (!target) continue;
    const root = relationshipRoot(outputPart), relations = root.content.filter((n):n is XmlNode => typeof n !== 'string');
    const resolved = rel.external ? target : '/' + target;
    const same = relations.find(n => n.attrs.get(' Type') === rel.type && n.attrs.get(' Target') === resolved && (n.attrs.get(' TargetMode') === 'External') === rel.external);
    let id = same?.attrs.get(' Id');
    if (!id) {
      const used = new Set(relations.map(n=>n.attrs.get(' Id'))); id=rel.id;
      for(let serial=1;used.has(id);serial++) id='nativeRel'+serial;
      root.content.push({name:'Relationship',uri:PKG,attrs:new Map([[' Id',id],[' Type',rel.type],[' Target',resolved],...(rel.external ? [[' TargetMode','External']] as [string,string][] : [])]),namespaces:{},content:[],part:relPath(outputPart)});
    }
    let mapping = relationMaps.get(rel.sourcePart); if(!mapping){mapping=new Map();relationMaps.set(rel.sourcePart,mapping);} mapping.set(rel.id,id);
  }
  const fragmentXml = (fragment: DeepReadonly<NonNullable<WordMetadata['nativeContent']>['fragments'][number]>): XmlNode => {
    const copy=(node:XmlNode):XmlNode=>({...node,attrs:new Map(node.attrs),content:node.content.map(item=>typeof item==='string'?item:copy(item))});
    const root=copy(readFragment(fragment)), mapping=relationMaps.get(fragment.sourcePart);
    const rewrite=(node:XmlNode):void=>{for(const [key,value]of node.attrs)if(key.startsWith(R+' ')&&mapping?.has(value))(node.attrs as Map<string,string>).set(key,mapping.get(value)!);for(const item of node.content)if(typeof item!=='string')rewrite(item);};rewrite(root);return root;
  };
  const restored=new Set<string>();
  const byNode=new Map<string,typeof native.fragments[number][]>();
  for(const fragment of native.fragments)if(fragment.nodeId){const list=byNode.get(fragment.nodeId)??[];list.push(fragment);byNode.set(fragment.nodeId,list);}
  const applyProperty=(root:XmlNode,fragment:typeof native.fragments[number]):void=>{
    restored.add(fragment.nativeId);
    let target=root;
    for(const path of fragment.propertyPath){let next=target.content.find((n):n is XmlNode=>typeof n!=='string'&&name(n)===path);if(!next){const split=path.lastIndexOf(' ');next={uri:path.slice(0,split),name:path.slice(split+1),attrs:new Map(),namespaces:{},content:[],part:root.part};target.content.unshift(next);}target=next;}
    const original=fragmentXml(fragment);
    if(name(original)===name(target))merge(target,original);
    else{const current=target.content.find((n):n is XmlNode=>typeof n!=='string'&&name(n)===name(original));if(current)merge(current,original);else target.content.push(original);}
  };
  const apply=(root:XmlNode):XmlNode=>{
    const id=root.attrs.get(NATIVE_MARKER+' node'), object=root.attrs.get(NATIVE_MARKER+' object');
    if(object){const fragment=native.fragments.find(f=>f.nativeId===object&&f.kind!=='property');if(!fragment)throw new DocumentError('InvalidDocument','A native object descriptor is missing.');restored.add(fragment.nativeId);if(!fragment.runChild)return fragmentXml(fragment);const container=root.attrs.get(NATIVE_MARKER+' container');for(const key of ['node','object','container'])(root.attrs as Map<string,string>).delete(NATIVE_MARKER+' '+key);root.content.push(fragmentXml(fragment));for(const residual of container?byNode.get(container)??[]:[])if(residual.kind==='property')applyProperty(root,residual);return root;}
    const container=root.attrs.get(NATIVE_MARKER+' container');(root.attrs as Map<string,string>).delete(NATIVE_MARKER+' node');(root.attrs as Map<string,string>).delete(NATIVE_MARKER+' container');
    for(let index=0;index<root.content.length;index++){const child=root.content[index]!;if(typeof child!=='string')root.content[index]=apply(child);}
    for(const fragment of [...(id?byNode.get(id)??[]:[]),...(container?byNode.get(container)??[]:[])])if(fragment.kind==='property'&&fragment.ownerName===root.name)applyProperty(root,fragment);
    return root;
  };
  for(const [path,root]of parsed){
    apply(root);
    for(const fragment of native.fragments)if(fragment.kind==='property'&&!fragment.nodeId&&partMap.get(fragment.sourcePart)===path){const target=fragment.styleId?root.content.find((n):n is XmlNode=>typeof n!=='string'&&n.name==='style'&&n.attrs.get(W+' styleId')===fragment.styleId):root;if(!target)throw new DocumentError('ExportFailed','A preserved native style owner is missing.');applyProperty(target,fragment);}
    files.set(path,encoder.encode(declaration+serializeXml(root)));
  }
  for(const part of retained){const bytes=sourceFiles.get(part.assetPath);if(!bytes)throw new DocumentError('AssetUnavailable','A retained OPC part is missing.');files.set(partMap.get(part.sourcePart)!,bytes);}
  for(const [part,root]of relationRoots)files.set(relPath(part),encoder.encode(declaration+serializeXml(root)));
  for(const fragment of native.fragments)if(!restored.has(fragment.nativeId)&&!retained.some(part=>part.sourcePart===fragment.sourcePart))throw new DocumentError('ExportFailed','A preserved native fragment could not be attached to its current source owner.',[{code:'UnattachedNativeContent',path:[fragment.sourcePart,fragment.nativeId],message:fragment.label}]);
  const types=read(files.get('[Content_Types].xml')!,'[Content_Types].xml');
  for(const part of retained)types.content.push({name:'Override',uri:CT,attrs:new Map([[' PartName','/'+partMap.get(part.sourcePart)!],[' ContentType',part.contentType]]),namespaces:{},content:[],part:'[Content_Types].xml'});
  files.set('[Content_Types].xml',encoder.encode(declaration+serializeXml(types)));
  let total=0;for(const bytes of files.values()){total+=bytes.length;if(bytes.length>limits.maxFileBytes)throw new DocumentError('ResourceLimitExceeded','A preserved Office part exceeds its byte budget.');}
  if(files.size>limits.maxFiles||total>limits.maxPackageBytes)throw new DocumentError('ResourceLimitExceeded','Preserved Office output exceeds its package budget.');
  return files;
}
/** Markers exist only between compilation and native reattachment. */
export function nativeObjectMarker(id:string,block:boolean,nodeId?:string,containerId?:string):string{return '<w:'+ (block?'p':'r')+' xmlns:w="'+W+'" xmlns:dn="'+NATIVE_MARKER+'" dn:object="'+escapeXml(id)+'"'+(nodeId?' dn:node="'+escapeXml(nodeId)+'"':'')+(containerId?' dn:container="'+escapeXml(containerId)+'"':'')+'/>';}
