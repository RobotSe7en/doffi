import { DocumentError } from "@doffi/core";
import type { DeepReadonly, ContentDocumentMetadata, DocumentSection, PageSetup } from "@doffi/core";
import type { LayoutDiagnostic, RenderedPage, RenderLimits } from "./contracts.js";
import { markViewNode } from "./view-nodes.js";
import { repeatTableHeaders } from "./table-headers.js";
import { pageFlow } from "./flow.js";
import { columnFlow } from './column-flow.js';
import{composeWordPageBorders}from'./word-page-borders.js';
import type { PageFlow } from "./flow.js";
import { ownViewStyle } from "./view-nodes.js";
import type { PreviewSource } from "./source-index.js";
import { applyWordGrid, wordPageSetup,facingWordPageSetup } from './word-grid.js';
import type { WordNoteLayout } from './word-notes.js';

const px = (value: number): number => value * 96 / 72;
const pt = (value: number): number => value * 72 / 96;
const gutter = 18;
export function pageLabel(value: number, format = "decimal"): string {
  if (format === "upperLetter" || format === "lowerLetter") {
    let label = "";
    for (let n = value; n > 0; n = Math.floor((n - 1) / 26)) label = String.fromCharCode(65 + (n - 1) % 26) + label;
    return format === "lowerLetter" ? label.toLowerCase() : label;
  }
  if (format === "upperRoman" || format === "lowerRoman") {
    if (value > 1_000_000) throw new DocumentError("ResourceLimitExceeded", "Roman page numbering exceeds the rendering budget.");
    let label = "", n = value;
    for (const [amount, letters] of [[1000,"M"],[900,"CM"],[500,"D"],[400,"CD"],[100,"C"],[90,"XC"],[50,"L"],[40,"XL"],[10,"X"],[9,"IX"],[5,"V"],[4,"IV"],[1,"I"]] as const)
      while (n >= amount) { label += letters; n -= amount; }
    return format === "lowerRoman" ? label.toLowerCase() : label;
  }
  return String(value);
}
interface SectionView { readonly definition: DeepReadonly<DocumentSection>; readonly element: HTMLElement }
interface Group {
  startIndex:number;
  readonly columns:string;
  readonly setup: DeepReadonly<PageSetup>;
  readonly box: HTMLElement;
  readonly flow: HTMLElement;
  readonly sections: SectionView[];
  readonly paginator: PageFlow;
}
export interface Pagination {
  measure(): readonly RenderedPage[];
  nodePages(pages: readonly RenderedPage[]): Readonly<Record<string, readonly string[]>>;
  readonly diagnostics: readonly LayoutDiagnostic[];
  dispose(): void;
}

/** Browser-independent measured pages over a single, reversible source DOM. */
export function paginate(document: Document, source: PreviewSource, metadata: DeepReadonly<ContentDocumentMetadata>, limits: Readonly<RenderLimits>,notes?:WordNoteLayout): Pagination {
  if (metadata.presentation !== undefined) throw new DocumentError("UnsupportedFeature", "Presentation surfaces do not use Word pagination.");
  const root = source.root;
  const diagnostics: LayoutDiagnostic[] = [], groups: Group[] = [];
  const sourceElements = new Map([root, ...root.querySelectorAll<HTMLElement>("[data-node-id]")].map((node) => [node.dataset.nodeId!, node]));
  if (document.querySelectorAll("*").length > limits.maxNodes) throw new DocumentError("ResourceLimitExceeded", "Document exceeds the render node budget.");
  const mainSections = metadata.sections.map((definition) => ({ definition, element: sourceElements.get(definition.rootNodeId) }))
    .filter((item): item is SectionView => Boolean(item.element && root.contains(item.element)))
    .sort((a, b) => a.element === b.element ? 0 : a.element.compareDocumentPosition(b.element) & 4 ? -1 : 1);
  if (!mainSections.length) throw new DocumentError("InvalidDocument", "The main story has no page section.");
  if (mainSections.some(({ element }, index) => element === root ? mainSections.length !== 1 : element.parentElement !== root
    || mainSections.some((other, ordinal) => ordinal !== index && (element.contains(other.element) || other.element.contains(element)))))
    throw new DocumentError("UnsupportedFeature", "Pagination requires non-overlapping section roots directly in the main story.");
  if (mainSections[0]!.element !== root) {
    // Core fragment operations may place valid story content between section
    // roots. Keep its exact source order and inherit the preceding page setup
    // (or the first setup for leading content), without moving it into a source
    // section or inventing persisted metadata.
    const wrappers = mainSections.map((section) => {
      const wrapper = markViewNode(document.createElement("div"), "unwrap"); wrapper.style.display = "flow-root";
      section.element.style.width = "auto"; section.element.style.padding = "0"; section.element.style.margin = "0";
      return { definition: section.definition, element: wrapper };
    });
    const ordinals = new Map(mainSections.map((section, ordinal) => [section.element as Node, ordinal]));
    let index = 0;
    for (const node of [...root.childNodes]) {
      const ordinal = ordinals.get(node); if (ordinal !== undefined) index = ordinal;
      wrappers[index]!.element.append(node);
    }
    root.append(...wrappers.map((section) => section.element));
    mainSections.splice(0, mainSections.length, ...wrappers);
  }
  const maximumWidth = Math.max(...mainSections.map((item) => (item.definition.pageSetup ?? metadata.pageSetup).width));
  root.style.width = maximumWidth + "pt"; root.style.padding = "0"; root.style.margin = "0 auto"; root.style.background = "transparent"; root.style.position = "relative";
  const article = root.closest<HTMLElement>('[data-office-kind="document"]')!;
  article.style.width = maximumWidth + "pt"; article.style.padding = "0"; article.style.background = "transparent"; article.style.position = "relative"; article.style.zIndex = "1";
  const layoutStyle = document.createElement("style");
  layoutStyle.textContent = ".sdk-page-group{position:relative;box-sizing:border-box;margin:0 auto 18pt}.sdk-page-group:last-child{margin-bottom:0}.sdk-page-flow{position:relative;display:flow-root}.sdk-page-flow .sdk-section{display:flow-root}.sdk-page-layer{position:absolute;pointer-events:none;z-index:0}.sdk-paper{position:absolute;box-sizing:border-box;background:white;box-shadow:0 1px 5px #0003}.sdk-page-story{position:absolute;box-sizing:border-box;overflow:hidden}.sdk-page-story p{min-height:0}.sdk-page-story .sdk-auxiliary{display:block;border:0;padding:0}.sdk-story-label,.sdk-auxiliary{display:none}.sdk-page-story{display:block}.sdk-page-number{position:absolute;right:0;bottom:-14pt;color:#68717b;font:9px system-ui}[data-word-break]{border:0!important;margin:0!important;height:0;line-height:0;display:block}";
  document.head.append(layoutStyle);
  for (const section of mainSections) {
    applyWordGrid(document,section.element,section.definition,limits);
    const setup = wordPageSetup(section.definition,metadata.pageSetup), { width, height, margins } = setup;
    if (width - margins.left - margins.right <= 0 || height - margins.top - margins.bottom <= 0 || width > 20_000 || height > 20_000)
      throw new DocumentError("UnsupportedFeature", "Page dimensions must provide a positive printable area within the render size limit.");
    let group = groups.at(-1);
    const columns=JSON.stringify(section.definition.columns??{count:1,gap:0});
    if (!group || JSON.stringify(group.setup) !== JSON.stringify(setup) || group.columns!==columns) {
      if (group && section.definition.breakType === "continuous") diagnostics.push({ code: "SectionGeometryChange", nodeId: section.definition.rootNodeId, message: "A continuous section with different page geometry starts on a new page." });
      const box = markViewNode(document.createElement("div"), "unwrap"), flow = markViewNode(document.createElement("div"), "unwrap");
      box.className = "sdk-page-group"; flow.className = "sdk-page-flow";
      box.style.cssText = `width:${width}pt;padding:${margins.top}pt 0 ${margins.bottom}pt;overflow:clip`;
      flow.style.cssText = `width:${width - margins.left - margins.right}pt;margin-left:${margins.left}pt`;
      box.append(flow);
      const definition=section.definition.columns;
      const constraints=notes?{reserved:(index:number)=>notes.reserved(group!.startIndex+index),minimumPage:(node:Node)=>{const minimum=notes.minimumPage(node);return minimum===undefined?undefined:Math.max(0,minimum-group!.startIndex);}}:undefined;
      group = { startIndex:0,setup, columns, box, flow, sections: [], paginator: definition&&definition.count>1||metadata.word?.mirrorMargins ? columnFlow(document,flow,px(height-margins.top-margins.bottom),px(height+gutter),px(width-margins.left-margins.right),definition?.count??1,px(definition?.gap??0),definition?.separator??false,limits,source,constraints,metadata.word?.mirrorMargins?(index=>(group!.startIndex+index)%2===1?px(setup.margins.right-setup.margins.left):0):undefined) : pageFlow(document, flow, px(height - margins.top - margins.bottom), px(height + gutter), limits,1,constraints) }; groups.push(group);
      if (section.element === root) { flow.append(...root.childNodes); root.append(box); }
      else root.insertBefore(box, section.element);
    }
    group.sections.push(section);
    if (section.element !== root) {
      const node = section.element;
      node.style.width = "auto"; node.style.padding = "0"; node.style.margin = "0"; node.style.boxSizing = "border-box";
      node.style.breakBefore = "auto";
      group.flow.append(node);
    }
  }
  const layer = markViewNode(document.createElement("div")); layer.className = "sdk-page-layer"; article.before(layer);
  const front=markViewNode(document.createElement('div'));front.className='sdk-page-border-layer';front.style.cssText='position:absolute;pointer-events:none;z-index:3';article.after(front);
  const storyTemplates = new Map((metadata.word?.stories ?? []).map((story) => [story.rootNodeId, document.querySelector<HTMLElement>('[data-node-id="' + CSS.escape(story.rootNodeId) + '"]')]));
  const repeated = repeatTableHeaders(document, root, source, limits);
  let pages: readonly RenderedPage[] = [], disposed = false;
  const sectionStarts = (group: Group): number[] => group.sections.map(({ element }) => element === root ? 0 : group.paginator.pageAt(element));
  function measure(): readonly RenderedPage[] {
    if (disposed) return pages;
    repeated.clear();
    for (const group of groups) group.paginator.reset();
    let physical = 0;
    for (const group of groups) {
      group.startIndex=physical;
      for (const section of group.sections) {
        const kind = section.definition.breakType, first = section === mainSections[0];
        const parity = !first && (kind === "oddPage" || kind === "evenPage") ? ((kind === "oddPage" ? 0 : 1) + physical) % 2 : undefined;
        group.paginator.layout(section.element === root ? group.flow : section.element, !first && kind !== "continuous", parity);
      }
      physical += group.paginator.count;
      if (physical > limits.maxPages) throw new DocumentError("ResourceLimitExceeded", "Document exceeds the render page budget.");
      group.paginator.project?.();
    }
    source.reindex();
    const result: RenderedPage[] = [], surface = root.getBoundingClientRect();
    let nextNumber = 1;
    const startsBySection = new Map<string, number>(), previousRefs: { headers: Record<string,string>; footers: Record<string,string> } = { headers: {}, footers: {} };
    const refs = new Map<string, { headers: Record<string,string>; footers: Record<string,string> }>();
    for (const section of mainSections) {
      Object.assign(previousRefs.headers, section.definition.headers); Object.assign(previousRefs.footers, section.definition.footers);
      refs.set(section.definition.sectionId, { headers: { ...previousRefs.headers }, footers: { ...previousRefs.footers } });
    }
    layer.replaceChildren();front.replaceChildren(); layer.style.left = surface.left - document.body.getBoundingClientRect().left + "px"; layer.style.top = surface.top - document.body.getBoundingClientRect().top + "px"; layer.style.width = maximumWidth + "pt";
    front.style.left=layer.style.left;front.style.top=layer.style.top;front.style.width=layer.style.width;
    let renderNodes = document.querySelectorAll("*").length;
    for (const group of groups) {
      const { setup } = group, starts = sectionStarts(group), pitch = px(setup.height + gutter);
      const count = group.paginator.count;
      const box = group.box.getBoundingClientRect();
      for (let index = 0; index < count; index++) {
        if (result.length >= limits.maxPages) throw new DocumentError("ResourceLimitExceeded", "Document exceeds the render page budget.");
        let owner = 0;
        for (let candidate = 1; candidate < starts.length; candidate++) if (starts[candidate]! <= index) owner = candidate;
        let definition = group.sections[owner]!.definition;
        const blank = group.paginator.blanks.has(index) ? group.sections.find((_section, ordinal) => index === starts[ordinal]! - 1) : undefined;
        if (blank) definition = mainSections[Math.max(0, mainSections.indexOf(blank) - 1)]!.definition;
        const first = !blank && !startsBySection.has(definition.sectionId);
        if (first) { startsBySection.set(definition.sectionId, result.length); if (definition.pageNumberStart !== undefined) nextNumber = definition.pageNumberStart; }
        const number = nextNumber++, page: RenderedPage = Object.freeze({ pageId: "page-" + (result.length + 1), index: result.length, sectionId: definition.sectionId,
          number, label: pageLabel(number, definition.pageNumberFormat), rect: Object.freeze({ x: pt(box.left - surface.left), y: pt(box.top - surface.top + index * pitch), width: setup.width, height: setup.height }), setup:facingWordPageSetup(setup,result.length,metadata.word?.mirrorMargins??false) as PageSetup, blank: Boolean(blank) });
        result.push(page);
        const paper = document.createElement("div"); paper.className = "sdk-paper"; paper.dataset.pageId = page.pageId; paper.setAttribute("aria-hidden", "true");
        paper.style.cssText = `left:${page.rect.x}pt;top:${page.rect.y}pt;width:${setup.width}pt;height:${setup.height}pt`;
        layer.append(paper);
        for (const kind of ["headers", "footers"] as const) {
          if (blank) continue;
          const variant = first && definition.differentFirstPage ? "first" : metadata.word?.evenAndOddHeaders && number % 2 === 0 ? "even" : "default";
          const reference = refs.get(definition.sectionId)?.[kind][variant];
          const id = metadata.word?.stories.find((story) => story.storyId === reference)?.rootNodeId;
          const template = id ? storyTemplates.get(id) : null;
          if (!template) continue;
          renderNodes += template.querySelectorAll("*").length + 1;
          if (renderNodes > limits.maxNodes) throw new DocumentError("ResourceLimitExceeded", "Repeated stories exceed the render node budget.");
          const story = template.cloneNode(true) as HTMLElement; story.classList.remove("sdk-auxiliary"); story.classList.add("sdk-page-story"); story.dataset.sdkAuxiliary = "true";
          const distance = kind === "headers" ? definition.headerDistance ?? 36 : definition.footerDistance ?? 36;
          const available = Math.max(0, (kind === "headers" ? setup.margins.top : setup.margins.bottom) - distance);
          story.style.cssText += `;left:${page.setup.margins.left}pt;width:${setup.width - setup.margins.left - setup.margins.right}pt;max-height:${available}pt;${kind === "headers" ? "top:" : "bottom:"}${distance}pt`;
          paper.append(story);
          if (story.scrollHeight > px(available) + 1 && !diagnostics.some((item) => item.code === "ContentOverflow" && item.nodeId === id))
            diagnostics.push({ code: "ContentOverflow", nodeId: id!, message: "A header or footer exceeds the available margin area and is clipped." });
        }
        renderNodes+=composeWordPageBorders(document,article,paper,front,page,definition.pageBorders,first,metadata.word);if(renderNodes>limits.maxNodes)throw new DocumentError('ResourceLimitExceeded','Page borders exceed the rendering node budget.');
        const label = document.createElement("span"); label.className = "sdk-page-number"; label.textContent = String(page.index + 1); paper.append(label); layer.append(paper);
      }
    }
    pages = Object.freeze(result); repeated.project(pages);
    // Clip source painting to printable areas; borders and merged-cell backgrounds
    // cannot bleed into margins or the inter-page gutter. Stories are siblings.
    for (const group of groups) {
      const contours = Array.from({ length: group.paginator.count }, (_, index) => group.paginator.blanks.has(index) ? "" :
        `M ${-px(group.setup.margins.left)} ${px(index * (group.setup.height + gutter))} h ${px(group.setup.width)} v ${px(group.setup.height - group.setup.margins.top - group.setup.margins.bottom)-(notes?.reserved(group.startIndex+index)??0)} h ${-px(group.setup.width)} Z`).join(" ");
      ownViewStyle(group.flow, "clip-path"); group.flow.style.clipPath = `path("${contours}")`;
    }
    return pages;
  }
  return { measure, diagnostics, nodePages(current) {
    const surface = root.getBoundingClientRect(), result: Record<string, readonly string[]> = Object.create(null) as Record<string, readonly string[]>;
    for (const [id] of sourceElements) {
      const rectangles = source.elementRectangles(id);
      result[id] = Object.freeze(current.filter((page) => !page.blank && rectangles.some((rect) => rect.bottom > surface.top + px(page.rect.y + page.setup.margins.top) + 0.1 && rect.top < surface.top + px(page.rect.y + page.rect.height - page.setup.margins.bottom) - 0.1)).map((page) => page.pageId));
    }
    return Object.freeze(result);
  }, dispose() { disposed = true; repeated.dispose(); layer.remove();front.remove(); layoutStyle.remove(); } };
}
