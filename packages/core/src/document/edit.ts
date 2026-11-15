import { defaultTreeAdapter as adapter } from "parse5";
import type { DefaultTreeAdapterTypes as Tree } from "parse5";
import type { DocumentContentOperation } from "../contracts/operations.js";
import type { DocumentReadLimits, TextOffsetMapping } from "../runtime.js";
import { DocumentError } from "../errors.js";
import { normalizeCss, setInlineStyle } from "./css.js";
import { setDrawingInlineStyle } from "./drawing-css.js";
import { styleSourceText } from "./text-style.js";
import type { SourceTextRange,SourceTextStyleChange } from "./text-style.js";
import { bodyTags, parseHtml, supportedTags } from "./html.js";
import { parseHtmlFragment, parseHtmlTree, serializeHtmlOuter, serializeHtmlTree } from "./html-parser.js";
import { decodeText, encode, resolveReference } from "./package.js";
import type { HtmlSourceDescriptor, ParsedHtml } from "./html.js";
import type { DocumentPackage } from "../runtime.js";
import { checkLimit } from "./validation.js";
import { isTextBoundary, textPositions } from "./text.js";
import { wordScriptRole } from '../word-script.js';

type Element = Tree.Element;
type Node = Tree.Node;
type Child = Tree.ChildNode;
const blocks = new Set(["article", "section", "div", "p", "h1", "h2", "h3", "h4", "h5", "h6", "li", "td", "th"]);
const isElement = (node: Node): node is Element => "tagName" in node;
const attr = (node: Element, name: string) => node.attrs.find((item) => item.name === name)?.value;
function fail(message: string): never { throw new DocumentError("InvalidOperation", message); }

/** Format adapters share the same bounded HTML editor and source validation. */
export interface EditableHtmlSource {
  readonly files: DocumentPackage;
  readonly snapshot: { readonly html: string; readonly manifest: HtmlSourceDescriptor;
    readonly metadata: { readonly presentation?: { readonly hyperlinks: readonly { readonly nodeId: string }[] };
      readonly workbook?: { readonly objectHyperlinks?: readonly { readonly nodeId: string }[] | undefined } } };
  readonly html: ParsedHtml;
}

function setAttribute(node: Element, name: string, value: string): void {
  const existing = node.attrs.find((item) => item.name === name);
  if (existing) existing.value = value;
  else node.attrs.push({ name, value });
}

function visit(root: Node, callback: (node: Node) => void): void {
  const stack = [root];
  while (stack.length) {
    const node = stack.pop()!;
    callback(node);
    if ("childNodes" in node) {
      for (let index = node.childNodes.length - 1; index >= 0; index--) stack.push(node.childNodes[index]!);
    }
  }
}

function ancestors(node: Node): Element[] {
  const result: Element[] = [];
  for (let current: Node | null = node; current; current = "parentNode" in current ? current.parentNode : null) {
    if (isElement(current)) result.push(current);
  }
  return result;
}

interface TextLeaf { node: Tree.TextNode | Element; start: number; end: number; value: string }
function textLeaves(root: Node): TextLeaf[] {
  const leaves: TextLeaf[] = [];
  let offset = 0;
  visit(root, (node) => {
    const value = node.nodeName === "#text" && "value" in node ? node.value
      : isElement(node) && node.tagName === "br" ? "\n"
        : isElement(node) && node.tagName === "img" ? "\ufffc" : undefined;
    if (value !== undefined) {
      leaves.push({ node: node as TextLeaf["node"], start: offset, end: offset + value.length, value });
      offset += value.length;
    }
  });
  return leaves;
}

/** Match import's source font roles, keeping marks and variation selectors with their base. */
type FontScript = "ascii" | "eastAsia" | "hAnsi" | "cs";
function fontScript(value: string): FontScript | undefined {
  return ["ascii", "eastAsia", "hAnsi", "cs"].includes(value) ? value as FontScript : undefined;
}
function scriptParts(value: string, initialScript?: FontScript): { script: FontScript; text: string }[] {
  const parts: ReturnType<typeof scriptParts> = [];
  for (const character of value) {
    const previous = parts.at(-1);
    const script = wordScriptRole(character, previous?.script ?? initialScript);
    if (previous?.script === script) previous.text += character;
    else parts.push({ script, text: character });
  }
  return parts;
}

/** Private mutable tree for one preparation. No parser objects escape the core. */
export class DocumentEdit {
  private readonly tree: Tree.Document;
  private readonly baseIds: ReadonlySet<string>;
  private createdNodes = 0;
  readonly textEdits: TextOffsetMapping[] = [];
  readonly replacedTextNodes = new Set<string>();
  readonly structuralTextNodes = new Set<string>();

  constructor(private readonly state: EditableHtmlSource, private readonly limits: DocumentReadLimits, private readonly allocate: () => string) {
    this.tree = parseHtmlTree(state.snapshot.html, limits, supportedTags);
    this.baseIds = new Set(state.html.nodeIds());
  }

  private index(): Map<string, Element> {
    const index = new Map<string, Element>();
    visit(this.tree, (node) => {
      if (isElement(node)) {
        const id = attr(node, "data-node-id");
        if (id !== undefined) index.set(id, node);
      }
    });
    return index;
  }

  /** Private format adapters write typed source attributes, never renderer DOM. */
  attributes(nodeId: string, values: Readonly<Record<string, string | undefined>>): void {
    const target = this.index().get(nodeId);
    if (!target) fail("The format operation targets a missing source node.");
    for (const [name, value] of Object.entries(values)) {
      if (value === undefined) target.attrs = target.attrs.filter((item) => item.name !== name);
      else setAttribute(target, name, value);
    }
  }

  renameSource(nodeId: string, tag: string): void {
    const node = this.index().get(nodeId); if (!node || !bodyTags.has(tag)) fail("A source rename requires a supported existing node.");
    node.tagName = tag; node.nodeName = tag;
  }
  shellAncestors(): readonly { tagName: string; attributes: Readonly<Record<string,string>> }[] {
    const html=this.tree.childNodes.find(node=>isElement(node)&&node.tagName==="html");
    if(!html || !isElement(html)) return [];
    const body=html.childNodes.find(node=>isElement(node)&&node.tagName==="body");
    return [body,html].filter((node):node is Element=>!!node&&isElement(node)).map(node=>({tagName:node.tagName,attributes:Object.fromEntries(node.attrs.map(attr=>[attr.name,attr.value]))}));
  }

  orderAttributes(nodeId:string,names:readonly string[]):void {
    const target=this.index().get(nodeId);if(!target)fail("The format source is missing.");
    const original=new Map(names.map((name,index)=>[name,index])),current=new Map(target.attrs.map((attr,index)=>[attr.name,index]));
    target.attrs.sort((a,b)=>(original.get(a.name)??names.length+current.get(a.name)!)-(original.get(b.name)??names.length+current.get(b.name)!));
  }

  /** Core-owned format insertion; caller HTML is still parsed, bounded and validated. */
  insertSource(parentId: string, source: string, beforeId?: string,copyIdentities=false,insertionOffset?:number): { rootNodeIds: string[]; identities: ReadonlyMap<string, string> } {
    const index = this.index(), container = index.get(parentId);let parent=container,before:Tree.ChildNode|undefined=beforeId === undefined ? undefined : index.get(beforeId);
    if (!parent || beforeId !== undefined && (!before || before.parentNode !== parent)) fail("A format insertion requires an existing parent and sibling.");
    const fragment = parseHtmlFragment(source, parent, this.limits, bodyTags), identities = new Map<string, string>();
    visit(fragment, node => {
      if (!isElement(node)) return;
      const previous = attr(node, "data-node-id"), id = this.allocate();
      if (previous !== undefined) {
        if(!copyIdentities)fail("New format content cannot supply source identities.");
        if (identities.has(previous)) fail("Copied source identities must be unique.");
        identities.set(previous, id);
      }
      setAttribute(node, "data-node-id", id);
      if (attr(node, "id") !== undefined) fail("Copying custom HTML IDs requires an explicit style mapping.");
      node.attrs = node.attrs.filter(item => item.name !== "data-office-field-id");
      const style = attr(node, "style");
      if (style !== undefined) setAttribute(node, "style", normalizeCss(style, "declarationList", this.limits, { nodes: 0 }));
    });
    const positions = textPositions(this.tree), range = positions.get(parent)!;let from=before ? positions.get(before)!.from : range.to;
    if(insertionOffset!==undefined){
      const leaves:Tree.ChildNode[]=[];visit(parent,node=>{if(node.nodeName==='#text'||isElement(node)&&['img','br'].includes(node.tagName))leaves.push(node as Tree.ChildNode);});
      const text=leaves.map(node=>node.nodeName==='#text'&&'value'in node?node.value:isElement(node)&&node.tagName==='img'?'\ufffc':'\n').join('');if(!Number.isSafeInteger(insertionOffset)||insertionOffset<0||insertionOffset>text.length||!isTextBoundary(text,insertionOffset))fail('A format insertion requires an exact UTF-16 source boundary.');
      from=range.from+insertionOffset;const leaf=leaves.find(node=>positions.get(node)!.to>from)||leaves.at(-1);
      if(leaf){
        let atomic:Element|undefined;for(let node:Tree.Node|null=leaf.parentNode;node&&node!==container;node='parentNode'in node?node.parentNode:null)if(isElement(node)&&(node.tagName==='del'||attr(node,'data-word-native-object')||attr(node,'data-word-drawing')||attr(node,'data-word-note-reference')||attr(node,'data-word-equation')||['checkBox','dropDown','date','picture'].includes(attr(node,'data-word-content-control')??'')))atomic=node;
        if(atomic){const bounds=positions.get(atomic)!;if(from!==bounds.from&&from!==bounds.to)fail('An atomic object cannot contain a partial insertion.');parent=atomic.parentNode as Element;before=from===bounds.from?atomic:parent.childNodes[parent.childNodes.indexOf(atomic)+1];}
        else{parent=leaf.parentNode as Element;const local=from-positions.get(leaf)!.from;
          if(leaf.nodeName==='#text'&&'value'in leaf&&local>0&&local<leaf.value.length){const tail=this.text(leaf.value.slice(local));leaf.value=leaf.value.slice(0,local);const next=parent.childNodes[parent.childNodes.indexOf(leaf)+1];if(next)adapter.insertBefore(parent,tail,next);else adapter.appendChild(parent,tail);before=tail;}
          else before=local<=0?leaf:parent.childNodes[parent.childNodes.indexOf(leaf)+1];
        }
      }
    }
    this.recordSplice(positions, from, from, container!, textPositions(fragment).get(fragment)!.to, true);
    const rootNodeIds = fragment.childNodes.filter(isElement).map(node => attr(node, "data-node-id")!);
    for (const child of [...fragment.childNodes]) { adapter.detachNode(child); if (before) adapter.insertBefore(parent, child, before); else adapter.appendChild(parent, child); }
    return { rootNodeIds, identities };
  }
  insertSourceAt(nodeId:string,offset:number,source:string):{rootNodeIds:string[];identities:ReadonlyMap<string,string>}{return this.insertSource(nodeId,source,undefined,false,offset);}
  resourceReference(assetPath:string):string{
    if(!this.state.snapshot.manifest.assets.some(asset=>asset.path===assetPath)||!this.state.files.has(assetPath))throw new DocumentError('AssetUnavailable','The drawing resource is not authorized in this transaction.');
    const directory=this.state.snapshot.manifest.entrypoint.split('/').slice(0,-1),parts=assetPath.split('/');while(directory.length&&directory[0]===parts[0]){directory.shift();parts.shift();}return[...directory.map(()=>'..'),...parts.map(encodeURIComponent)].join('/');
  }

  cloneSource(nodeId: string, parentId: string, beforeId?: string): { rootNodeIds: string[]; identities: ReadonlyMap<string, string> } {
    const target = this.index().get(nodeId);
    if (!target) fail("The copied source is missing.");
    return this.insertSource(parentId, serializeHtmlOuter(target), beforeId,true);
  }

  /** Internal table adapter may address parser-created row groups by core identity. */
  sourceParent(nodeId:string):string {
    const target=this.index().get(nodeId),parent=target?.parentNode;
    if(!parent||!isElement(parent)||parent.tagName==="body")fail("The source parent is unavailable.");
    let id=attr(parent,"data-node-id");if(!id){id=this.allocate();setAttribute(parent,"data-node-id",id);}return id;
  }
  sourceStyle(nodeId:string,property:string,value:string):void {
    const target=this.index().get(nodeId);if(!target)fail("The format style target is unavailable.");setAttribute(target,"style",setInlineStyle(attr(target,"style")??"",property,value,this.limits));
  }
  sourceDrawingStyle(nodeId: string, property: string, value: string): void {
    const target = this.index().get(nodeId); if (!target) fail("The drawing style target is unavailable.");
    setAttribute(target, "style", setDrawingInlineStyle(attr(target, "style") ?? "", property, value, this.limits));
  }
  formatSourceText(nodeId: string, range: SourceTextRange, properties: Readonly<Record<string, string>>,onChange?:(change:SourceTextStyleChange)=>void,preferred?:ReadonlySet<string>,force=false,nativeRun=false): void {
    const target = this.index().get(nodeId); if (!target) fail("The text format target is unavailable.");
    const styles = this.state.snapshot.manifest.styles.map(path => decodeText(this.state.files.get(path)!));
    styleSourceText(target, range, properties, this.limits, this.allocate, styles,undefined,onChange,preferred,force,nativeRun);
  }
  /** Split an inline property owner without changing text or moving atomic
   * content out of its owner. Only the split ancestor path receives new IDs. */
  splitSourceInline(nodeId:string,offset:number):ReadonlyMap<string,string>{
    const target=this.index().get(nodeId);if(!target||!target.parentNode||!isElement(target.parentNode))fail('The inline split target is unavailable.');
    const positions=textPositions(this.tree),area=positions.get(target)!,text=textLeaves(target).map(leaf=>leaf.value).join('');
    if(offset<=0||offset>=text.length||!isTextBoundary(text,offset))fail('An inline split requires an interior UTF-16 boundary.');
    const boundary=area.from+offset,identities=new Map<string,string>();
    const split=(node:Element):Element=>{
      if(!['span','a','strong','b','em','i','u','s','sub','sup','code'].includes(node.tagName)||node.attrs.some(attribute=>['data-word-field','data-word-content-control','data-word-native-object','data-word-drawing','data-word-equation','data-word-note-reference'].includes(attribute.name)))fail('A formatting boundary cannot split an atomic Word object.');
      const right=adapter.createElement(node.tagName,node.namespaceURI,node.attrs.map(attribute=>({...attribute}))),old=attr(node,'data-node-id'),id=this.allocate();setAttribute(right,'data-node-id',id);if(old)identities.set(old,id);
      for(const child of [...node.childNodes]){const span=positions.get(child);if(!span)continue;if(span.from>=boundary){adapter.detachNode(child);adapter.appendChild(right,child);}else if(span.to>boundary){if(child.nodeName==='#text'&&'value'in child){const at=boundary-span.from;adapter.appendChild(right,this.text(child.value.slice(at)));child.value=child.value.slice(0,at);}else if(isElement(child))adapter.appendChild(right,split(child));}}
      return right;
    };
    const right=split(target),parent=target.parentNode,sibling=parent.childNodes[parent.childNodes.indexOf(target)+1];if(sibling)adapter.insertBefore(parent,right,sibling);else adapter.appendChild(parent,right);return identities;
  }
  /** Revision containers retain source leaves and object ownership; wrapping
   * changes no UTF-16 text offsets. A wrapper is allocated for each text leaf. */
  wrapSourceText(nodeId:string,from:number,to:number,tag:'ins'|'del'|'span'):string[]{
    const target=this.index().get(nodeId);if(!target)fail('The revision source is missing.');
    const positions=textPositions(this.tree),bounds=positions.get(target)!,leaves=textLeaves(target),value=leaves.map(leaf=>leaf.value).join('');
    if(from>to||!isTextBoundary(value,from)||!isTextBoundary(value,to))fail('A revision range requires exact UTF-16 boundaries.');
    if(from===to)return[];
    const ids:string[]=[],wrapped=new Set<Node>();
    for(const leaf of leaves){if(leaf.end<=from||leaf.start>=to)continue;let node:Tree.ChildNode=leaf.node as Tree.ChildNode;
      for(let parent:Node|null=node;parent&&parent!==target;parent='parentNode'in parent?parent.parentNode:null)if(isElement(parent)&&(attr(parent,'data-word-native-object')||attr(parent,'data-word-drawing')||attr(parent,'data-word-equation')||attr(parent,'data-word-note-reference')||attr(parent,'data-word-field')||attr(parent,'data-word-content-control')))node=parent;
      if(wrapped.has(node))continue;wrapped.add(node);const span=positions.get(node)!,begin=Math.max(bounds.from+from,span.from),end=Math.min(bounds.from+to,span.to);
      if(isElement(node)&&(begin!==span.from||end!==span.to))fail('A revision cannot split an atomic Word object.');
      const parent=node.parentNode;if(!parent||!isElement(parent))fail('A revision requires a source parent.');const wrapper=this.make(tag);ids.push(attr(wrapper,'data-node-id')!);
      if(node.nodeName==='#text'&&'value'in node){const localFrom=begin-span.from,localTo=end-span.from;if(localFrom)adapter.insertBefore(parent,this.text(node.value.slice(0,localFrom)),node);adapter.insertBefore(parent,wrapper,node);adapter.appendChild(wrapper,this.text(node.value.slice(localFrom,localTo)));if(localTo<node.value.length)adapter.insertBefore(parent,this.text(node.value.slice(localTo)),node);adapter.detachNode(node);}
      else{adapter.insertBefore(parent,wrapper,node);adapter.detachNode(node);adapter.appendChild(wrapper,node);}
    }
    return ids;
  }
  /** New text in an empty Word paragraph inherits its mark's source run.
   * Existing inline overrides remain inside this default scope. */
  wrapWordParagraphText(nodeId:string,attributes:Readonly<Record<string,string>>):string[]{
    const target=this.index().get(nodeId);if(!target||!/^p$|^h[1-6]$/.test(target.tagName))fail('Run defaults require a paragraph.');
    const result:string[]=[];let wrapper:Element|undefined;
    for(const child of [...target.childNodes]){
      if(isElement(child)&&(attr(child,'data-word-paragraph-mark')==='true'||!textLeaves(child).some(leaf=>leaf.value))){wrapper=undefined;continue;}
      if(!wrapper){wrapper=this.make('span',Object.entries({...attributes,'data-word-run':'true','data-word-run-overlay':'true'}).map(([name,value])=>({name,value})));adapter.insertBefore(target,wrapper,child);result.push(attr(wrapper,'data-node-id')!);}
      adapter.detachNode(child);adapter.appendChild(wrapper,child);visit(child,node=>{if(isElement(node)&&node.tagName==='span'&&!this.baseIds.has(attr(node,'data-node-id')??'')&&!attr(node,'data-word-run-overlay'))setAttribute(node,'data-word-run-overlay','true');});
    }
    return result;
  }
  previewReconciliation(operation:Extract<DocumentContentOperation,{op:'reconcileFragment'}>):ParsedHtml{
    const html=this.html(),parsed=parseHtml(html,this.state.snapshot.manifest,this.state.files,this.limits,{nodes:0}),preview=new DocumentEdit({...this.state,html:parsed,snapshot:{...this.state.snapshot,html}},this.limits,this.allocate);preview.apply(operation);return parseHtml(preview.html(),this.state.snapshot.manifest,this.state.files,this.limits,{nodes:0});
  }
  /** A format adapter has already checked the owning source identity. */
  replaceSourceText(nodeId: string, edit: { readonly from: number; readonly to: number; readonly expectedText: string; readonly text: string }): void {
    const target = this.index().get(nodeId); if (!target) fail("The source text target is unavailable.");
    this.replaceText(target, { op: "replaceText", nodeId, from: edit.from, to: edit.to, expectedText: edit.expectedText, text: edit.text });
  }
  removeSource(nodeId:string,contents=false):void {
    const target=this.index().get(nodeId);if(!target||!target.parentNode||!isElement(target.parentNode)||target.parentNode.tagName==="body")fail("The document root cannot be removed.");
    const positions=textPositions(this.tree),range=positions.get(target)!;
    this.recordSplice(positions,range.from,range.to,contents?target:target.parentNode,0,true);
    if(contents)for(const child of [...target.childNodes])adapter.detachNode(child);else adapter.detachNode(target);
  }
  unwrapSource(nodeId:string):void{const target=this.index().get(nodeId);if(!target||!target.parentNode||!isElement(target.parentNode)||target.parentNode.tagName==='body')fail('The document root cannot be unwrapped.');const parent=target.parentNode;for(const child of [...target.childNodes]){adapter.detachNode(child);adapter.insertBefore(parent,child,target);}adapter.detachNode(target);}
  mergeSourceParagraphMark(nodeId:string):void{
    const node=this.index().get(nodeId);if(!node||!/^p$|^h[1-6]$/.test(node.tagName)||!node.parentNode||!isElement(node.parentNode))fail('A paragraph mark requires its paragraph.');
    const parent=node.parentNode,siblings=parent.childNodes,markers:Element[]=[];let next:Element|undefined;
    for(let index=siblings.indexOf(node)+1;index<siblings.length;index++){const candidate=siblings[index]!;if(candidate.nodeName==='#text'&&'value'in candidate&&!candidate.value.trim())continue;if(isElement(candidate)&&['div','span'].includes(candidate.tagName)&&textLeaves(candidate).every(leaf=>!leaf.value)){let marker=false;visit(candidate,item=>{if(isElement(item)&&(attr(item,'data-word-move-range')||attr(item,'data-word-bookmark')))marker=true;});if(marker){markers.push(candidate);continue;}}if(isElement(candidate)&&/^p$|^h[1-6]$/.test(candidate.tagName))next=candidate;break;}
    if(!next)return;const before=next.childNodes[0],append=(child:Tree.ChildNode):void=>{adapter.detachNode(child);if(before)adapter.insertBefore(next!,child,before);else adapter.appendChild(next!,child);};for(const child of [...node.childNodes]){if(isElement(child)&&attr(child,'data-word-paragraph-mark')){adapter.detachNode(child);continue;}append(child);}for(const marker of markers){if(marker.tagName==='span')append(marker);else{for(const child of [...marker.childNodes])append(child);adapter.detachNode(marker);}}adapter.detachNode(node);
    for(const ancestor of ancestors(next)){const id=attr(ancestor,'data-node-id');if(id)this.structuralTextNodes.add(id);}
  }


  /** Format-owned bodies accept new fragments without exposing writable source
   * identities. Equal canonical content retains the original nodes and bytes. */
  replaceSourceContent(nodeId: string, source: string): boolean {
    const target = this.index().get(nodeId); if (!target) fail("The source body is missing.");
    const fragment = parseHtmlFragment(source, target, this.limits, bodyTags);
    visit(fragment, node => { if (isElement(node) && (attr(node, "data-node-id") !== undefined || attr(node, "id") !== undefined)) fail("New format content cannot supply source identities."); });
    const canonical = (fragment: Tree.DocumentFragment): string => {
      visit(fragment, node => { if (isElement(node)) { node.attrs = node.attrs.filter(item => item.name !== "data-node-id"); const style = attr(node, "style"); if (style !== undefined) setAttribute(node, "style", normalizeCss(style, "declarationList", this.limits, { nodes: 0 })); } });
      return serializeHtmlTree(fragment);
    };
    const old = parseHtmlFragment(target.childNodes.map(serializeHtmlOuter).join(""), target, this.limits, bodyTags);
    if (canonical(fragment) === canonical(old)) return false;
    this.removeSource(nodeId, true); this.insertSource(nodeId, source); return true;
  }

  /** Structural movement preserves node identity; ancestor text anchors need exact recovery. */
  moveSource(nodeIds: readonly string[], parentId: string, beforeId?: string): void {
    const index = this.index(), parent = index.get(parentId), before = beforeId === undefined ? undefined : index.get(beforeId);
    const targets = nodeIds.map(id => index.get(id));
    if (!parent || beforeId !== undefined && (!before || before.parentNode !== parent) || targets.some(target => !target)) fail("A format move requires existing source nodes and a sibling.");
    if (new Set(nodeIds).size !== nodeIds.length || targets.some(target => target === before || ancestors(parent).includes(target!) || targets.some(other => other !== target && ancestors(target!).includes(other!)))) fail("A format move cannot contain cycles or nested selections.");
    for (const target of targets) for (const ancestor of [...ancestors(target!).slice(1), ...ancestors(parent)]) { const id = attr(ancestor, "data-node-id"); if (id) this.replacedTextNodes.add(id); }
    for (const target of targets) { adapter.detachNode(target!); if (before) adapter.insertBefore(parent, target!, before); else adapter.appendChild(parent, target!); }
  }

  private make(tag: string, attributes: { name: string; value: string }[] = []): Element {
    checkLimit(++this.createdNodes, this.limits.maxNodes);
    return adapter.createElement(tag, "http://www.w3.org/1999/xhtml" as Element["namespaceURI"], [
      { name: "data-node-id", value: this.allocate() }, ...attributes,
    ]);
  }

  private text(value: string): Tree.TextNode {
    checkLimit(++this.createdNodes, this.limits.maxNodes);
    return { nodeName: "#text", value, parentNode: null };
  }

  private insertedText(value: string, word: boolean, inheritedScript: string): Child[] {
    const result: Child[] = [];
    for (const [index, line] of value.split("\n").entries()) {
      if (index) result.push(this.make("br"));
      if (!line) continue;
      if (!word) result.push(this.text(line));
      else for (const part of scriptParts(line, fontScript(inheritedScript))) {
        if (part.script === inheritedScript) {
          result.push(this.text(part.text));
          continue;
        }
        const span = this.make("span", [
          { name: "data-word-script", value: part.script },
          { name: "style", value: "font-family:var(--word-font-" + part.script + ",var(--word-font-ascii,var(--word-font-hAnsi,inherit)))" },
        ]);
        adapter.appendChild(span, this.text(part.text));
        result.push(span);
      }
    }
    return result;
  }

  /** Record one splice for each surviving base node whose normalized text changes. */
  private recordSplice(positions: ReturnType<typeof textPositions>, from: number, to: number, parent: Element, insertedLength: number, structural = false): void {
    const insertionAncestors = new Set(ancestors(parent));
    for (const [id, element] of this.index()) {
      if (!this.baseIds.has(id)) continue;
      const range = positions.get(element)!;
      const begin = Math.max(range.from, from);
      const end = Math.min(range.to, to);
      const length = insertionAncestors.has(element) ? insertedLength : 0;
      if (begin < end || length) {
        this.textEdits.push({ nodeId: id, from: Math.max(0, begin - range.from), to: Math.max(0, end - range.from), insertedLength: length });
        if (structural) this.structuralTextNodes.add(id);
      }
    }
  }

  private replaceText(target: Element, operation: Extract<DocumentContentOperation, { op: "replaceText" }>): void {
    if (["img", "br"].includes(target.tagName)) fail("Text operations require a content element, not an inline object.");
    const { from, to, text, expectedText } = operation;
    const leaves = textLeaves(target);
    const original = leaves.map((leaf) => leaf.value).join("");
    if (from > to || to > original.length || original.slice(from, to) !== expectedText
      || !isTextBoundary(original, from) || !isTextBoundary(original, to)) fail("The text range or expected source text does not match.");
    if ([...text].some((character) => character.length === 1 && character.charCodeAt(0) >= 0xd800 && character.charCodeAt(0) <= 0xdfff)
      || /[\u0000\r]/u.test(text)) fail("Replacement text must be well-formed Unicode with LF line breaks.");
    if (expectedText === text) return;
    const selected = leaves.filter((leaf) => leaf.end > from && leaf.start < to);
    const start = leaves.find((leaf) => leaf.start <= from && leaf.end > from) ?? leaves.at(-1);
    if (selected.some((leaf) => isElement(leaf.node) && leaf.node.tagName === "img")) {
      fail("Inline objects must be changed through structural or asset operations.");
    }
    if (selected.some((leaf) => ancestors(leaf.node).some((node) => attr(node, "data-word-tab") !== undefined
      || attr(node, "data-word-break") !== undefined))) fail("Native tabs and explicit page breaks require structural operations.");
    const context = (node: Node) => ancestors(node).find((item) => blocks.has(item.tagName));
    const nativeLinks = new Set([...(this.state.snapshot.metadata.presentation?.hyperlinks ?? []), ...(this.state.snapshot.metadata.workbook?.objectHyperlinks ?? [])].map(link => link.nodeId));
    const semantic = (node: Node) => ancestors(node).find((item) => ["ins", "del"].includes(item.tagName)
      || attr(item, "data-word-link") !== undefined || attr(item, "data-word-field") !== undefined
      || attr(item, "data-word-field-result") !== undefined || attr(item, "data-office-field") !== undefined
      || attr(item, "data-office-link") !== undefined || nativeLinks.has(attr(item, "data-node-id")!));
    if (start && selected.some((leaf) => context(leaf.node) !== context(start.node) || semantic(leaf.node) !== semantic(start.node))) {
      fail("Text replacement cannot cross paragraphs, cells, stories or native semantic boundaries.");
    }
    // Empty markers are source boundaries too; text replacement must not move them silently.
    let offset = 0;
    visit(target, (node) => {
      if (isElement(node) && node.attrs.some((item) => ["data-word-field", "data-word-field-end", "data-word-bookmark", "data-word-comment-mark", "data-word-tab", "data-word-break"].includes(item.name))
        && offset > from && offset < to) fail("Text replacement crosses a native Word marker.");
      if (node.nodeName === "#text" && "value" in node) offset += node.value.length;
      else if (isElement(node) && ["br", "img"].includes(node.tagName)) offset++;
    });
    const parent = start ? start.node.parentNode : target;
    if (!parent || !isElement(parent)) fail("The text insertion point is not an editable element.");
    const inherited = ancestors(parent);
    const inheritedScript = inherited.map((item) => attr(item, "data-word-script")).find((script) => script !== undefined) ?? "ascii";
    const inserted = this.insertedText(text, inherited.some((item) => attr(item, "data-word-run") !== undefined || attr(item, "data-word-script") !== undefined), inheritedScript);
    const positions = textPositions(this.tree);
    const globalFrom = start ? positions.get(start.node)!.from + from - start.start : positions.get(target)!.from;
    this.recordSplice(positions, globalFrom, globalFrom + to - from, parent, text.length);
    if (!start) {
      for (const node of inserted) adapter.appendChild(target, node);
      return;
    }
    const reference = start.node;
    if (isElement(reference) && ["img", "br"].includes(reference.tagName) && from === to) {
      const next = from >= start.end ? parent.childNodes[parent.childNodes.indexOf(reference) + 1] : reference;
      for (const node of inserted) {
        if (next) adapter.insertBefore(parent, node, next);
        else adapter.appendChild(parent, node);
      }
      return;
    }
    if (reference.nodeName === "#text" && "value" in reference) {
      const local = Math.max(0, from - start.start);
      if (local) adapter.insertBefore(parent, this.text(reference.value.slice(0, local)), reference);
      reference.value = reference.value.slice(local);
    }
    for (const node of inserted) adapter.insertBefore(parent, node, reference);
    for (const leaf of leaves) {
      const begin = Math.max(from, leaf.start);
      const end = Math.min(to, leaf.end);
      if (begin >= end) continue;
      if (leaf.node.nodeName === "#text" && "value" in leaf.node) {
        // The starting leaf's prefix was already split off above.
        leaf.node.value = leaf.node === reference
          ? leaf.value.slice(end - leaf.start)
          : leaf.value.slice(0, begin - leaf.start) + leaf.value.slice(end - leaf.start);
      } else adapter.detachNode(leaf.node);
    }
  }

  /** Reconcile a validated editor view while retaining only identities inside this source subtree. */
  private reconcile(target: Element, operation: Extract<DocumentContentOperation, { op: "reconcileFragment" }>): void {
    const before = parseHtml(this.html(), this.state.snapshot.manifest, this.state.files, this.limits, { nodes: 0 });
    const allowed = new Map<string, Element>();
    visit(target, (node) => { if (isElement(node)) { const id = attr(node, "data-node-id"); if (id) allowed.set(id, node); } });
    const fragment = parseHtmlFragment(operation.html, target, this.limits, bodyTags), seen = new Set<string>();
    const cssBudget = { nodes: 0 };
    visit(fragment, (node) => {
      if (!isElement(node)) return;
      const id = attr(node, "data-node-id");
      if (id !== undefined) {
        if (!this.baseIds.has(id) || !allowed.has(id) || node === target || id === operation.nodeId || seen.has(id))
          fail("Reconciled identities must be unique descendants of the base target subtree.");
        seen.add(id);
        if (this.state.snapshot.metadata.presentation && attr(allowed.get(id)!, "data-office-kind") !== undefined) {
          const original = allowed.get(id)!;
          const protectedName = (name: string) => /^data-office-(?:kind$|story$|x$|y$|width$|height$|rotation$|flip-|child-|crop-|adjust-|connector-preset$)/.test(name);
          const names = new Set([...original.attrs, ...node.attrs].map(item => item.name).filter(protectedName));
          if ([...names].some(name => attr(original, name) !== attr(node, name))) fail("Reconciliation cannot change typed presentation geometry or source roles.");
        }
      } else setAttribute(node, "data-node-id", this.allocate());
      const style = attr(node, "style");
      if (style !== undefined) setAttribute(node, "style", normalizeCss(style, "declarationList", this.limits, cssBudget));
    });
    const proofs = new Map<string, { value: string; edits: typeof operation.textEdits; inserted: { from: number; to: number }[] }>();
    for (const edit of operation.textEdits) {
      if (!allowed.has(edit.nodeId) || !this.baseIds.has(edit.nodeId)) fail("Text mapping proof targets must belong to the base target subtree.");
      const proof = proofs.get(edit.nodeId) ?? { value: before.getNode(edit.nodeId)!.text, edits: [], inserted: [] };
      if (edit.from > edit.to || edit.to > proof.value.length || proof.value.slice(edit.from, edit.to) !== edit.expectedText
        || !isTextBoundary(proof.value, edit.from) || !isTextBoundary(proof.value, edit.to)
        || /[\u0000\r]/u.test(edit.text)
        || [...edit.text].some((character) => character.length === 1 && character.charCodeAt(0) >= 0xd800 && character.charCodeAt(0) <= 0xdfff))
        fail("Reconciled text mapping proof does not match the source text or Unicode boundaries.");
      proof.value = proof.value.slice(0, edit.from) + edit.text + proof.value.slice(edit.to);
      if (edit.expectedText !== edit.text) {
        proof.edits.push(edit);
        const delta = edit.text.length - edit.to + edit.from, inserted: typeof proof.inserted = [];
        for (const range of proof.inserted) {
          if (range.to <= edit.from) inserted.push(range);
          else if (range.from >= edit.to) inserted.push({ from: range.from + delta, to: range.to + delta });
          else {
            if (range.from < edit.from) inserted.push({ from: range.from, to: edit.from });
            if (range.to > edit.to) inserted.push({ from: edit.from + edit.text.length, to: range.to + delta });
          }
        }
        if (edit.text.length) inserted.push({ from: edit.from, to: edit.from + edit.text.length });
        const merged: typeof inserted = [];
        for (const range of inserted.sort((left, right) => left.from - right.from)) {
          const last = merged.at(-1);
          if (last && range.from <= last.to) last.to = Math.max(last.to, range.to);
          else merged.push({ ...range });
        }
        proof.inserted = merged;
      }
      proofs.set(edit.nodeId, proof);
    }
    for (const child of [...target.childNodes]) adapter.detachNode(child);
    for (const child of [...fragment.childNodes]) { adapter.detachNode(child); adapter.appendChild(target, child); }
    const textNodes: Tree.TextNode[] = [];
    visit(target, (node) => { if (node.nodeName === "#text" && "value" in node) textNodes.push(node); });
    const positions = textPositions(this.tree);
    for (const node of textNodes) {
      const inherited = ancestors(node), word = inherited.some((item) => attr(item, "data-word-run") !== undefined || attr(item, "data-word-script") !== undefined);
      if (!word || !node.parentNode || !isElement(node.parentNode)) continue;
      const existing = inherited.find((item) => { const id = attr(item, "data-node-id"); return id !== undefined && proofs.has(id); });
      if (!existing) continue;
      const proof = proofs.get(attr(existing, "data-node-id")!)!, offset = positions.get(node)!.from - positions.get(existing)!.from;
      const ranges = proof.inserted.map((range) => ({ from: Math.max(0, range.from - offset), to: Math.min(node.value.length, range.to - offset) }))
        .filter((range) => range.from < range.to);
      if (!ranges.length) continue;
      const script = inherited.map((item) => attr(item, "data-word-script")).find((value) => value !== undefined) ?? "ascii";
      if (ranges.every((range) => scriptParts(node.value.slice(range.from, range.to), fontScript(script)).every((part) => part.script === script))) continue;
      let cursor = 0;
      for (const range of ranges) {
        if (range.from > cursor) adapter.insertBefore(node.parentNode, this.text(node.value.slice(cursor, range.from)), node);
        for (const part of this.insertedText(node.value.slice(range.from, range.to), true, script)) adapter.insertBefore(node.parentNode, part, node);
        cursor = range.to;
      }
      if (cursor < node.value.length) adapter.insertBefore(node.parentNode, this.text(node.value.slice(cursor)), node);
      adapter.detachNode(node);
    }
    const after = parseHtml(this.html(), this.state.snapshot.manifest, this.state.files, this.limits, { nodes: 0 });
    for (const [id] of allowed) {
      if (!after.hasNode(id)) continue;
      const previous = before.getNode(id)!.text, next = after.getNode(id)!.text, proof = proofs.get(id);
      if (proof && proof.value !== next) fail("Reconciled text mapping proof differs from the final source text.");
      if (proof) {
        for (const edit of proof.edits) this.textEdits.push({ nodeId: id, from: edit.from, to: edit.to, insertedLength: edit.text.length });
        if (proof.edits.length) this.structuralTextNodes.add(id);
      } else if (previous !== next) this.replacedTextNodes.add(id);
    }
    for (const parent of ancestors(target).slice(1)) {
      const id = attr(parent, "data-node-id");
      if (id && before.getNode(id)?.text !== after.getNode(id)?.text) this.replacedTextNodes.add(id);
    }
  }

  apply(operation: Exclude<DocumentContentOperation, import("../contracts/word-operations.js").WordOperation>): void {
    this.createdNodes = 0;
    if (!this.baseIds.has(operation.nodeId)) fail("Operation targets must belong to the transaction's base snapshot.");
    const target = this.index().get(operation.nodeId);
    if (!target) fail("An earlier operation removed the target node.");
    switch (operation.op) {
      case "replaceText": this.replaceText(target, operation); break;
      case "reconcileFragment": this.reconcile(target, operation); break;
      case "setStyle": {
        if(attr(target,'data-word-drawing')&&/^(?:width|height|position|left|right|top|bottom|transform)$/.test(operation.property))fail('Use Word drawing geometry operations to change drawing dimensions and position.');
        const style = (this.state.snapshot.metadata.presentation ? setDrawingInlineStyle : setInlineStyle)(attr(target, "style") ?? "", operation.property, operation.value, this.limits);
        setAttribute(target, "style", style);
        break;
      }
      case "replaceAsset": {
        if (target.tagName !== "img") fail("Asset replacement requires an image node.");
        const asset = this.state.snapshot.manifest.assets.find((item) => item.path === operation.assetPath);
        if (!asset || !this.state.files.has(operation.assetPath)) throw new DocumentError("AssetUnavailable", "The replacement asset is not declared and available in this session.");
        const directory = this.state.snapshot.manifest.entrypoint.split("/").slice(0, -1);
        const parts = operation.assetPath.split("/");
        while (directory.length && directory[0] === parts[0]) { directory.shift(); parts.shift(); }
        const relative = [...directory.map(() => ".."), ...parts.map(encodeURIComponent)].join("/");
        if (resolveReference(relative, this.state.snapshot.manifest.entrypoint) !== asset.path) fail("The image path cannot be resolved.");
        setAttribute(target, "src", relative);
        break;
      }
      case "removeNode": {
        if (!target.parentNode || !isElement(target.parentNode) || target.parentNode.tagName === "body") fail("The document root cannot be removed.");
        const positions = textPositions(this.tree);
        const range = positions.get(target)!;
        this.recordSplice(positions, range.from, range.to, target.parentNode, 0, true);
        adapter.detachNode(target);
        break;
      }
      case "replaceFragment":
      case "insertFragment": {
        const sibling = operation.op === "insertFragment" && ["before", "after"].includes(operation.position);
        const parent = sibling ? target.parentNode : target;
        if (!parent || !isElement(parent) || parent.tagName === "body") fail("Fragments must remain inside the document root.");
        const fragment = parseHtmlFragment(operation.html, parent, this.limits, bodyTags);
        const cssBudget = { nodes: 0 };
        visit(fragment, (node) => {
          if (isElement(node)) {
            if (attr(node, "data-node-id") !== undefined) fail("Fragment node identities must be allocated by the core.");
            setAttribute(node, "data-node-id", this.allocate());
            const style = attr(node, "style");
            if (style !== undefined) setAttribute(node, "style", normalizeCss(style, "declarationList", this.limits, cssBudget));
          }
        });
        const reference = operation.op === "replaceFragment" || operation.position === "end" ? undefined
          : operation.position === "start" ? target.childNodes[0]
            : operation.position === "before" ? target
              : parent.childNodes[parent.childNodes.indexOf(target) + 1];
        const positions = textPositions(this.tree);
        const range = positions.get(parent)!;
        const from = operation.op === "replaceFragment" ? range.from : reference ? positions.get(reference)!.from : range.to;
        const to = operation.op === "replaceFragment" ? range.to : from;
        this.recordSplice(positions, from, to, parent, textPositions(fragment).get(fragment)!.to, true);
        if (operation.op === "replaceFragment") {
          visit(target, (node) => {
            if (isElement(node)) {
              const id = attr(node, "data-node-id");
              if (id !== undefined && this.baseIds.has(id)) this.replacedTextNodes.add(id);
            }
          });
          for (const child of [...target.childNodes]) adapter.detachNode(child);
        }
        for (const child of [...fragment.childNodes]) {
          adapter.detachNode(child);
          if (reference) adapter.insertBefore(parent, child, reference);
          else adapter.appendChild(parent, child);
        }
        break;
      }
    }
    const source = serializeHtmlTree(this.tree);
    checkLimit(source.length, this.limits.maxFileBytes);
    checkLimit(encode(source).byteLength, this.limits.maxFileBytes);
    // Catch browser tree correction instead of silently moving or dropping inserted content.
    if (parseHtml(source, this.state.snapshot.manifest, this.state.files, this.limits, { nodes: 0 }).html !== source) {
      fail("The operation produces an unstable HTML structure.");
    }
  }

  html(): string { return serializeHtmlTree(this.tree); }
}
