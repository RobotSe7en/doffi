import type { XmlNode } from './xml.js';
import { escapeXml } from '../export/xml.js';
const XML = 'http://www.w3.org/XML/1998/namespace';
const XMLNS = 'http://www.w3.org/2000/xmlns/';
/** Serialize expanded names with explicit bindings, including detached native fragments. */
export function serializeXml(root: XmlNode, mapAttribute: (uri: string, name: string, value: string) => string = (_u, _n, v) => v): string {
  const prefixes = new Map<string, string>([[XML, 'xml']]), reserved = new Set<string>(['xml']),bindings=new Map<string,Set<string>>(),aliases=new Map<string,string[]>();
  const reserve = (node: XmlNode): void => {for(const[name,uri]of Object.entries(node.namespaces)){reserved.add(name);const values=bindings.get(name)??new Set<string>();values.add(uri);bindings.set(name,values);const names=aliases.get(uri)??[];if(name&&!names.includes(name))names.push(name);aliases.set(uri,names);}for(const child of node.content)if(typeof child!=='string')reserve(child);};
  reserve(root); let serial = 0;
  const prefix = (uri: string): string => {
    if (!uri) return '';
    let value = prefixes.get(uri);
    if(!value){
      // A globally unambiguous existing alias is safe for expanded names and
      // QName-valued attributes. Reusing it also keeps repeated native round
      // trips from accumulating a fresh namespace declaration on every pass.
      value=aliases.get(uri)?.find(name=>name!=='xml'&&bindings.get(name)?.size===1);
      if(!value){do{value='doffiXml'+(++serial);}while(reserved.has(value));reserved.add(value);}prefixes.set(uri,value);
    }
    return value + ':';
  };
  const collect = (node: XmlNode): void => {
    prefix(node.uri);
    for (const key of node.attrs.keys()) { const split = key.lastIndexOf(' '); const uri = key.slice(0, split); if (uri !== XMLNS) prefix(uri); }
    for (const item of node.content) if (typeof item !== 'string') collect(item);
  };
  collect(root);
  const write = (node: XmlNode, top: boolean, inherited: ReadonlyMap<string,string> = new Map()): string => {
    const scope = new Map(inherited);
    const name = prefix(node.uri) + node.name;
    let out = '<' + name;
    if (top) for (const [uri, p] of prefixes) if (uri !== XML) { out += ' xmlns:' + p + '="' + escapeXml(uri) + '"'; scope.set(p,uri); }
    const defaultUri=node.namespaces[''];if(defaultUri!==undefined&&scope.get('')!==defaultUri){out+=' xmlns="'+escapeXml(defaultUri)+'"';scope.set('',defaultUri);}
    // Preserve aliases used in mc:Ignorable and other QName-valued attributes.
    for (const [p, uri] of Object.entries(node.namespaces)) if (p && p !== 'xml' && uri !== 'urn:doffi:transient:native' && ![...prefixes.values()].includes(p) && scope.get(p) !== uri) { out += ' xmlns:' + p + '="' + escapeXml(uri) + '"'; scope.set(p,uri); }
    for (const [key, value] of node.attrs) {
      const split = key.lastIndexOf(' '), uri = key.slice(0, split), local = key.slice(split + 1);
      if (uri === XMLNS) continue;
      out += ' ' + prefix(uri) + local + '="' + escapeXml(mapAttribute(uri, local, value)).replace(/\n/g, '&#10;').replace(/\t/g, '&#9;') + '"';
    }
    return out + '>' + node.content.map(item => typeof item === 'string' ? escapeXml(item) : write(item, false, scope)).join('') + '</' + name + '>';
  };
  return write(root, true);
}
