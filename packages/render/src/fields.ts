import { DocumentError } from "@doffi/core";
import type { DeepReadonly, ContentDocumentModel, DocumentTransaction, WordMetadata } from "@doffi/core";
import type { DocumentRenderReport, LayoutDiagnostic, LayoutFieldResult, RenderedPage } from "./contracts.js";
import { pageLabel } from "./pagination.js";
import { hideFromViewSelection, markViewNode, measureViewContent } from "./view-nodes.js";
import type { PreviewSource } from "./source-index.js";

interface FieldView {
  readonly id: string;
  readonly kind: LayoutFieldResult['kind'];
  readonly instruction: string;
  readonly results: readonly string[];
  readonly display: HTMLElement;
}
export interface LayoutFields {
  readonly changed: boolean;
  evaluate(pages: readonly RenderedPage[], nodePages: Readonly<Record<string, readonly string[]>>): readonly LayoutFieldResult[];
  repeat(pages: readonly RenderedPage[]): void;
  readonly diagnostics: readonly LayoutDiagnostic[];
}
function displayText(document: Document, target: HTMLElement, value: string): void {
  target.replaceChildren();
  let role = "ascii", part = "";
  const append = (): void => {
    if (!part) return;
    const span = document.createElement("span"); span.dataset.wordScript = role;
    span.style.fontFamily = `var(--word-font-${role},var(--word-font-ascii,inherit))`; span.textContent = part; target.append(span); part = "";
  };
  for (const character of value) {
    const next = /\p{Mark}/u.test(character) ? role : /[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Hangul}]/u.test(character) ? "eastAsia"
      : /[\p{Script=Arabic}\p{Script=Hebrew}]/u.test(character) ? "cs" : character.codePointAt(0)! <= 0x7f ? "ascii" : "hAnsi";
    if (next !== role) { append(); role = next; } part += character;
  }
  append();
}

/** Results are a reversible display projection. Explicit source cache updates are
 * returned as ordinary core operations by createLayoutFieldTransaction below. */
export function layoutFields(document: Document, source: PreviewSource, word: DeepReadonly<WordMetadata> | undefined): LayoutFields {
  const views: FieldView[] = [], diagnostics: LayoutDiagnostic[] = [];
  let changed = false;
  const all = word?.fields ?? [];
  const instructionOnly = new Set(all.flatMap((field) => field.instructionFields?.map((item) => item.fieldNodeId) ?? []));
  const styles=new Map(word?.styles.map(style=>[style.styleId,style]));
  const headings=[...source.root.querySelectorAll<HTMLElement>('p,h1,h2,h3,h4,h5,h6')].flatMap(element=>{let level=element.dataset.wordOutlineLevel!==undefined?Number(element.dataset.wordOutlineLevel):undefined;for(let style=styles.get(element.dataset.wordStyle??'');level===undefined&&style;style=style.basedOn?styles.get(style.basedOn):undefined)level=style.outlineLevel;level??=/^h[1-6]$/.test(element.localName)?Number(element.localName[1])-1:undefined;return level!==undefined&&level<9?[{id:element.dataset.nodeId!,level:level+1,text:source.elements.get(element.dataset.nodeId!)?.element.textContent?.trim()??''}]:[];});

  for (const field of all) {
    const kind = field.instruction.trim().split(/\s+/)[0]!.toUpperCase();
    if (field.locked) {
      diagnostics.push({ code: "UnsupportedField", nodeId: field.nodeId, message: "Locked field result remains cached." }); continue;
    }
    if (!["PAGE", "NUMPAGES", "TOC",'REF','PAGEREF','NOTEREF','SEQ'].includes(kind) || instructionOnly.has(field.nodeId)) {
      if (!instructionOnly.has(field.nodeId)) diagnostics.push({ code: "UnsupportedField", nodeId: field.nodeId, message: "Field result remains cached: " + kind });
      continue;
    }
    const results = field.resultNodeIds ?? [], first = results[0] ? source.elements.get(results[0])?.element : undefined;
    if (!first) { diagnostics.push({ code: "UnsupportedField", nodeId: field.nodeId, message: "Field has no editable cached result scope." }); continue; }
    if (kind === "TOC" && /\\(?:f|a|b|c|s|t)\b/i.test(field.instruction)) {
      diagnostics.push({ code: "UnsupportedField", nodeId: field.nodeId, message: "This TOC switch combination remains cached." }); continue;
    }
    if((kind==='REF'&&/\\(?:n|r|w|p|d|f)\b/i.test(field.instruction)||kind==='PAGEREF'&&/\\p\b/i.test(field.instruction)||kind==='NOTEREF'&&/\\p\b/i.test(field.instruction))){diagnostics.push({code:'UnsupportedField',nodeId:field.nodeId,message:'This reference switch combination remains cached.'});continue;}
    const display = markViewNode(document.createElement("span"));
    measureViewContent(display);
    display.className = "sdk-field-display"; display.dataset.sdkFieldNodeId = field.nodeId; display.style.whiteSpace = "pre-wrap";
    if (kind === "TOC") display.style.display = "block";
    for (const id of results) {
      const result = source.elements.get(id)?.element; if (!result) continue;
      const cache = markViewNode(document.createElement("span"), "unwrap");
      hideFromViewSelection(cache);
      cache.style.cssText = "position:absolute;visibility:hidden;pointer-events:none";
      cache.append(...result.childNodes); result.append(cache); result.contentEditable = "false";
      source.setFieldView(id, display);
    }
    first.append(display);
    views.push({ id: field.nodeId, kind: kind as FieldView["kind"], instruction: field.instruction, results, display });
  }
  function value(view: FieldView, pages: readonly RenderedPage[], nodePages: Readonly<Record<string, readonly string[]>>, page?: RenderedPage): string {
    if(['REF','PAGEREF','NOTEREF'].includes(view.kind)){
      const match=/^\s*(?:REF|PAGEREF|NOTEREF)\s+(?:"([^"]+)"|([^\s\\]+))/i.exec(view.instruction),bookmark=word?.bookmarks.find(bookmark=>bookmark.name===(match?.[1]??match?.[2])),start=bookmark?source.elements.get(bookmark.startNodeId):undefined,end=bookmark?source.elements.get(bookmark.endNodeId):undefined;if(!start||!end)return'错误！未找到引用源。';
      if(view.kind==='PAGEREF')return pages.find(page=>nodePages[bookmark!.startNodeId]?.includes(page.pageId))?.label??'1';
      if(view.kind==='NOTEREF'){const reference=word?.noteReferences?.find(reference=>{const node=source.elements.get(reference.nodeId);return node&&node.from>=start.from&&node.to<=end.from;});if(!reference)return'错误！未找到注释。';const marker=source.elements.get(reference.nodeId)?.element;return marker?.lastElementChild?.textContent??reference.customMark??'1';}
      return source.tokens.filter(token=>token.to>start.from&&token.from<end.from).map(token=>token.text.slice(Math.max(0,start.from-token.from),Math.min(token.text.length,end.from-token.from))).join('');
    }
    if(view.kind==='SEQ'){
      const match=/^\s*SEQ\s+(?:"([^"]+)"|([^\s\\]+))/i.exec(view.instruction),name=match?.[1]??match?.[2],fields=views.filter(field=>field.kind==='SEQ').sort((a,b)=>(source.elements.get(a.id)?.from??0)-(source.elements.get(b.id)?.from??0));let number=0;for(const field of fields){const key=/^\s*SEQ\s+(?:"([^"]+)"|([^\s\\]+))/i.exec(field.instruction);if((key?.[1]??key?.[2])!==name)continue;const restart=/\\r\s+(\d+)/i.exec(field.instruction);number=restart?Number(restart[1]):/\\c\b/i.test(field.instruction)?number:number+1;if(field===view)break;}return String(number);
    }
    if (view.kind === "NUMPAGES") return String(pages.length);
    if (view.kind === "PAGE") {
      const current = page ?? pages.find((item) => nodePages[view.id]?.includes(item.pageId)) ?? pages[0];
      const format = /\\\*\s+(ROMAN|roman|ALPHABETIC|alphabetic|Arabic)\b/.exec(view.instruction)?.[1];
      const selected = format === "ROMAN" ? "upperRoman" : format === "roman" ? "lowerRoman" : format === "ALPHABETIC" ? "upperLetter" : format === "alphabetic" ? "lowerLetter" : "decimal";
      return current ? format ? pageLabel(current.number, selected) : current.label : "1";
    }
    const levels = /\\o\s+"([1-6])-([1-6])"/i.exec(view.instruction), min = Number(levels?.[1] ?? 1), max = Number(levels?.[2] ?? 3);
    return headings.filter((item) => item.level >= min && item.level <= max).map((heading) => {
      const current = pages.find((item) => nodePages[heading.id]?.includes(item.pageId));
      return heading.text + "\t" + (current?.label ?? "1");
    }).join("\n");
  }
  return { diagnostics, get changed() { return changed; }, evaluate(pages, nodePages) {
    changed = false;
    return Object.freeze(views.map((view) => {
      const result = value(view, pages, nodePages);
      if (view.display.textContent !== result) { displayText(document, view.display, result); changed = true; }
      return Object.freeze({ fieldNodeId: view.id, resultNodeIds: view.results, kind: view.kind, value: result });
    }));
  }, repeat(pages) {
    for (const page of pages) {
      const paper = document.querySelector<HTMLElement>('.sdk-paper[data-page-id="' + CSS.escape(page.pageId) + '"]');
      const repeated = document.querySelectorAll<HTMLElement>('.sdk-repeat-header[data-sdk-repeat-page-id="' + CSS.escape(page.pageId) + '"] .sdk-field-display[data-sdk-field-node-id]');
      for (const node of [...paper?.querySelectorAll<HTMLElement>(".sdk-field-display[data-sdk-field-node-id]") ?? [], ...repeated]) {
        const view = views.find((item) => item.id === node.dataset.sdkFieldNodeId); if (!view) continue;
        displayText(document, node, value(view, pages, {}, page));
      }
    }
  } };
}

export function createLayoutFieldTransaction(model: ContentDocumentModel, report: DocumentRenderReport): DocumentTransaction {
  const snapshot = model.snapshot();
  if (report.layout !== "pages" || !report.pages.length) throw new DocumentError("InvalidOperation", "Layout fields require a completed pages layout.");
  if (report.documentId !== snapshot.manifest.documentId || report.revision !== snapshot.manifest.revision)
    throw new DocumentError("RevisionConflict", "Layout fields belong to a different document revision.");
  const operations: DocumentTransaction["operations"] = [];
  const seen = new Set<string>();
  for (const field of report.fields) {
    const definition = snapshot.metadata.word?.fields.find((item) => item.nodeId === field.fieldNodeId);
    if (!definition || JSON.stringify(definition.resultNodeIds ?? []) !== JSON.stringify(field.resultNodeIds)) throw new DocumentError("InvalidOperation", "Layout field scopes differ from the document source.");
    for (const [index, nodeId] of field.resultNodeIds.entries()) {
      if (seen.has(nodeId)) throw new DocumentError("InvalidOperation", "Layout field result scopes overlap."); seen.add(nodeId);
      const node = model.getNode(nodeId); if (!node) throw new DocumentError("InvalidOperation", "Layout field result scope is missing.");
      const text = index === 0 ? field.value : "";
      if (node.text !== text) operations.push({ op: "replaceText", nodeId, from: 0, to: node.text.length, expectedText: node.text, text });
    }
  }
  return { protocolVersion: 1, documentId: snapshot.manifest.documentId, baseRevision: snapshot.manifest.revision, operations };
}
