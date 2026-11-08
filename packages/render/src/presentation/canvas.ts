import { DocumentError, evaluatePresentationPreset, readPresentationGeometry } from "@doffi/core";
import type { DeepReadonly, ContentDocumentModel, PresentationGeometry, PresentationMetadata, PresentationObject } from "@doffi/core";
import type { DocumentRenderReport, LayoutDiagnostic, LayoutFieldResult, RenderedPage, RenderLimits } from "../contracts.js";
import { PreviewSource } from "../source-index.js";
import { hideFromViewSelection, markViewNode, measureViewContent, ownViewAttribute, ownViewStyle } from "../view-nodes.js";
import { numeric, property, resolvedColor, svgFill } from "./appearance.js";
import { renderChart } from "./chart.js";
import type { BasicChartScene } from "../chart.js";
import { prepareChartTextDecorations } from "../chart-text-decoration.js";
import { PresentationInheritance } from "./inheritance.js";
import { svgElement } from "./svg.js";
import { decorateLists, textTransform } from "./text.js";
import { resolveConnector } from "./connectors.js";
import { officeFontSegments } from "../font-scripts.js";
import { createDrawingGeometryView } from "../drawing-geometry.js";

type ObjectDefinition = DeepReadonly<PresentationObject>;
const attributes = (node: Element): Record<string, string> => Object.fromEntries([...node.attributes].map(a => [a.name, a.value]));
const styleView = (node: HTMLElement, values: Readonly<Record<string, string>>): void => {
  ownViewStyle(node, ...Object.keys(values)); for (const [name, value] of Object.entries(values)) node.style.setProperty(name, value);
};
const frameCss = `html{background:#eceef1}body{margin:0;padding:24px;font-family:'Noto Sans CJK SC',sans-serif}article[data-office-kind=presentation]{margin:0 auto;position:relative}p,h1,h2,h3,h4,h5,h6{margin:0;line-height:1.2;font-weight:inherit}p{min-height:1em}table{border-spacing:0}td,th{padding:0;font-weight:inherit;text-align:inherit;vertical-align:top}body>article>[data-office-kind]:not([data-office-kind=slide]){display:none}body>article>[data-office-kind=slide]{position:relative;box-sizing:border-box;overflow:hidden;margin:0 0 18pt;isolation:isolate;background:white}div:where([data-office-kind=text],[data-office-kind=geometry]){padding:3.6pt 7.2pt}div:where([data-office-kind=chart]){font-size:10pt}:where([data-office-kind=text],[data-office-kind=geometry],[data-office-kind=image],[data-office-kind=group],[data-office-kind=connector],[data-office-kind=table],[data-office-kind=chart]){--office-fill-kind:initial;--office-fill-color:initial;--office-line-kind:initial;--office-line-color:initial;--office-shadow-kind:initial;--office-text-anchor:initial}.sdk-ppt-notes,.sdk-ppt-comments{font:14px sans-serif;white-space:pre-wrap;background:#fff;padding:12px;margin:12px 0}.sdk-ppt-notes summary{font-weight:bold;cursor:pointer}.sdk-ppt-text{box-sizing:border-box;display:flex;flex-direction:column;position:absolute;white-space:pre-wrap;overflow-wrap:normal}.sdk-ppt-text>p{flex:none}span:where([data-sdk-hyperlink]){color:var(--office-theme-hlink);text-decoration:underline}div:where([data-sdk-chart-title]){font-weight:bold}.sdk-ppt-marker{position:absolute;pointer-events:none;white-space:pre}.sdk-ppt-hidden-label{position:absolute;right:8pt;top:4pt;font:10px sans-serif;background:#fff;color:#666;z-index:2147483647}`;

export interface PresentationCanvas {
  readonly source: PreviewSource;
  readonly fields: readonly LayoutFieldResult[];
  finalizeCharts(): void;
  measure(): Pick<DocumentRenderReport, "pages" | "nodePages" | "diagnostics">;
}
/** Fixed slide composition. Every injected object is a reversible view node. */
export function composePresentation(document: Document, model: ContentDocumentModel, limits: Readonly<RenderLimits>,surfaceId?:string): PresentationCanvas {
  const snapshot = model.snapshot(), metadata = snapshot.metadata.presentation;
  if (snapshot.manifest.kind !== "presentation" || !metadata) throw new DocumentError("UnsupportedFormat", "A presentation canvas requires a presentation package.");
  const source = new PreviewSource(document, model), root = source.root;
  let count = document.querySelectorAll("*").length, serial = 0;
  const account = (amount: number): void => { count += amount; if (count > limits.maxNodes) throw new DocumentError("ResourceLimitExceeded", "Presentation rendering exceeds its node budget."); };
  if (metadata.slides.length > limits.maxPages) throw new DocumentError("ResourceLimitExceeded", "Presentation rendering exceeds its page budget.");
  account(0);
  const css = document.createElement("style"); css.textContent = frameCss; document.head.insertBefore(css, document.head.querySelector("[data-sdk-source-styles]"));
  styleView(root, { width: metadata.size.width + "pt" });
  const byId = new Map([...source.elements].map(([id, entry]) => [id, entry.element]));
  const objects = new Map(metadata.objects.map(o => [o.nodeId, o])), layouts = new Map(metadata.layouts.map(l => [l.layoutId, l])), masters = new Map(metadata.masters.map(m => [m.masterId, m]));
  const inheritance = new PresentationInheritance(document, snapshot.styles.map(s => s.css).join("\n"), limits.maxNodes);
  const fields: LayoutFieldResult[] = [], diagnostics: LayoutDiagnostic[] = [], textBoxes: { node: HTMLElement; content: HTMLElement; fit: string; scale: number; lineReduction: number; fonts: readonly { node: HTMLElement; size: number; line: number }[]; resize(width: number, height: number): void; width: number; height: number }[] = [];
  const connectors: { node: HTMLElement; definition: DeepReadonly<Extract<PresentationObject, { kind: "connector" }>> }[] = [];
  const chartScenes: BasicChartScene[] = [];
  const surfaceLayout=metadata.layouts.find(l=>l.rootNodeId===surfaceId),surfaceMaster=metadata.masters.find(m=>m.rootNodeId===surfaceId),surfaceNotesMaster=metadata.notesMasters.find(m=>m.rootNodeId===surfaceId);
  if(surfaceId&&!surfaceLayout&&!surfaceMaster&&!surfaceNotesMaster)throw new DocumentError("InvalidOperation","The requested presentation definition surface is missing.");
  const slideNodes = [...root.children].filter((node): node is HTMLElement => node instanceof document.defaultView!.HTMLElement && node.dataset.officeKind === "slide");
  const slides=surfaceId?[byId.get(surfaceId)!]:slideNodes;
  const size=surfaceNotesMaster?metadata.notesSize??metadata.size:metadata.size;
  if(surfaceId){for(const slide of slideNodes)styleView(slide,{display:"none"});styleView(slides[0]!,{display:"block",position:"relative",overflow:"hidden",isolation:"isolate","background-color":"white",margin:"0 0 18pt"});styleView(root,{width:size.width+"pt"});}
  const entryFor=(node:HTMLElement):DeepReadonly<PresentationMetadata>["slides"][number]=>metadata.slides.find(s=>s.rootNodeId===node.dataset.nodeId)??{slideId:surfaceLayout?.layoutId??surfaceMaster?.masterId??surfaceNotesMaster!.notesMasterId,rootNodeId:surfaceId!,layoutId:surfaceLayout?.layoutId??metadata.layouts[0]!.layoutId};
  const surfaceObjects = (surface: HTMLElement): HTMLElement[] => [...surface.children].filter((node): node is HTMLElement => objects.has(node.getAttribute("data-node-id") ?? ""));
  const placeholder = (surface: HTMLElement, object: ObjectDefinition): HTMLElement | undefined => surfaceObjects(surface).find(node=>{const candidate=objects.get(node.dataset.nodeId!)!;return candidate.placeholder&&object.placeholder&&candidate.placeholder.index===object.placeholder.index&&candidate.placeholder.type===object.placeholder.type;})??surfaceObjects(surface).find(node => {
    const candidate = objects.get(node.dataset.nodeId!)!;
    return candidate.placeholder && object.placeholder && candidate.placeholder.index === object.placeholder.index && (candidate.placeholder.type === object.placeholder.type || [candidate.placeholder.type, object.placeholder.type].every(t => ["title", "ctrTitle"].includes(t)));
  });
  const geometry = (nodes: readonly HTMLElement[]): PresentationGeometry => {
    for (let i = nodes.length - 1; i >= 0; i--) { const result = readPresentationGeometry(attributes(nodes[i]!)); if (result) return result; }
    throw new DocumentError("InvalidDocument", "A presentation object has no resolved geometry.");
  };
  const instancesBySlide = new Map<string, Set<string>>();
  const process = (node: HTMLElement, definition: ObjectDefinition, prototypes: readonly HTMLElement[], slideIndex: number, master: DeepReadonly<PresentationMetadata>["masters"][number], instance = false): void => {
    const chain = [...prototypes, node], g = geometry(chain);
    const id = "ppt-" + ++serial;
    inheritance.apply(node, prototypes);
    const computed = document.defaultView!.getComputedStyle(node);
    const insets = { left: parseFloat(computed.paddingLeft) || 0, top: parseFloat(computed.paddingTop) || 0, right: parseFloat(computed.paddingRight) || 0, bottom: parseFloat(computed.paddingBottom) || 0 };
    const anchor = property(computed, "--office-text-anchor", "t"), shadow = property(computed, "--office-shadow-kind");
    const background = property(computed, "--office-fill-kind"), shadowColor = resolvedColor(document, computed, "--office-shadow", "#000000");
    if (definition.kind === "geometry" && background !== "linear" && background !== "none" && definition.preset !== "line" && !property(computed, "--office-text-color")) inheritance.apply(node, [], new Map([["color", "var(--office-theme-lt1,#ffffff)"]]));
    let shadowFilter = "none";
    if (shadow === "outer") {
      const angle = numeric(computed, "--office-shadow-direction", 0) * Math.PI / 180, distance = numeric(computed, "--office-shadow-distance", 0);
      shadowFilter = `drop-shadow(${distance * Math.cos(angle)}pt ${distance * Math.sin(angle)}pt ${numeric(computed, "--office-shadow-blur", 0) / 2}pt ${shadowColor})`;
    }
    styleView(node, { position: "absolute", left: g.x + "pt", top: g.y + "pt", width: g.width + "pt", height: g.height + "pt", "box-sizing": "border-box", padding: "0", margin: "0", "transform-origin": "50% 50%", transform: `rotate(${g.rotation}deg) scale(${g.flipH ? -1 : 1},${g.flipV ? -1 : 1})`, filter: shadowFilter });
    if (node.dataset.officeHidden === "true") styleView(node, { visibility: "hidden" });
    if (definition.kind === "group") {
      const wrapper = markViewNode(document.createElement("div"), "unwrap"), x = Number(node.dataset.officeChildX), y = Number(node.dataset.officeChildY), w = Number(node.dataset.officeChildWidth), h = Number(node.dataset.officeChildHeight);
      wrapper.style.cssText = `position:absolute;left:0;top:0;width:${w}pt;height:${h}pt;transform-origin:0 0;transform:scale(${g.width / w},${g.height / h}) translate(${-x}pt,${-y}pt)`;
      wrapper.append(...node.childNodes); node.append(wrapper);
      for (const child of [...wrapper.children] as HTMLElement[]) { const object = objects.get(child.dataset.nodeId!); if (object) process(child, object, [], slideIndex, master, instance); }
    } else if (definition.kind === "geometry" || definition.kind === "connector" || definition.kind === "text") {
      const adjustments: Record<string, number> = {};
      for (const prototype of chain) for (const attr of prototype.attributes) if (attr.name.startsWith("data-office-adjust-")) adjustments[attr.name.slice(19)] = Number(attr.value);
      const preset = definition.kind === "geometry" ? definition.preset : definition.kind === "connector" ? node.dataset.officeConnectorPreset ?? (definition.routing === "straight" ? "straightConnector1" : "bentConnector3") : "rect";
      const evaluated = evaluatePresentationPreset(preset, g.width, g.height, adjustments);
      const boxColor=computed.backgroundColor,boxPaint=definition.kind!=="connector"&&preset!=="line"&&boxColor!=="transparent"&&!/rgba\([^)]*,\s*0\)$/.test(boxColor)?boxColor:undefined;
      const svg = createDrawingGeometryView(document, computed, { preset, geometry: evaluated, widthPt: g.width, heightPt: g.height, id, orientation: { rotationDegrees: g.rotation, flipH: g.flipH, flipV: g.flipV },
        fillFallback: boxPaint ?? (definition.kind === "geometry" && preset !== "line" ? "#4472c4" : "none"),
        strokeFallback: definition.kind === "connector" || preset === "line" ? "#000000" : "none", maxNodes: limits.maxNodes - count });
      if(definition.kind!=="connector"&&preset!=="line")styleView(node,{"background-color":"transparent"});
      const children = [...node.childNodes]; node.append(svg);
      if (definition.kind === "connector") connectors.push({ node, definition });
      if (children.length || definition.kind==="text" || definition.kind==="geometry"&&evaluated.textRect) {
        const content = markViewNode(document.createElement("div"), "unwrap"); content.className = "sdk-ppt-text";
        const rect = evaluated.textRect;
        content.style.cssText += ";" + textTransform(node, rect ?? { x: 0, y: 0, width: g.width, height: g.height }) + `;overflow-wrap:anywhere;padding:${insets.top}px ${insets.right}px ${insets.bottom}px ${insets.left}px;justify-content:${anchor === "ctr" ? "center" : anchor === "b" ? "flex-end" : anchor === "dist" ? "space-around" : anchor === "just" ? "space-between" : "flex-start"}`;
        if (!rect) content.style.display = "none";
        content.append(...children); node.append(content);
        let text = {} as NonNullable<Extract<PresentationObject, { kind: "text" }>["text"]>;
        for (const prototype of chain) { const entry = objects.get(prototype.dataset.nodeId!); if (entry && "text" in entry) text = { ...text, ...entry.text }; }
        content.style.whiteSpace = text.wrap === false ? "pre" : "pre-wrap";
        ownViewAttribute(node,"data-sdk-text-auto-fit");node.setAttribute("data-sdk-text-auto-fit",text.autoFit??"none");
        const role = definition.placeholder?.type, textStyle = role === "title" || role === "ctrTitle" ? master.titleStyle : role === "body" ? master.bodyStyle : master.otherStyle;
        for (const p of content.querySelectorAll<HTMLElement>("p")) {
          const level = p.dataset.officeLevel ?? "0", base = document.createElement("p"); base.className = textStyle ? textStyle + " " + textStyle + "-level-" + level : "";
          const ancestors = prototypes.flatMap(prototype => [...prototype.querySelectorAll<HTMLElement>("p")].filter(candidate => (candidate.dataset.officeLevel ?? "0") === level).slice(0, 1));
          inheritance.apply(p, [base, ...ancestors]);
        }
        const fonts = [...content.querySelectorAll<HTMLElement>("p,span")].map(node => { const style = document.defaultView!.getComputedStyle(node); return { node, size: parseFloat(style.fontSize), line: parseFloat(style.lineHeight) }; });
        textBoxes.push({ node, content, fit: text.autoFit ?? "none", scale: text.fontScale ?? 1, lineReduction: text.lineSpacingReduction ?? 0, fonts, width: g.width, height: g.height,
          resize(width, height) {
            if (width > 1_000_000 || height > 1_000_000) throw new DocumentError("ResourceLimitExceeded", "Text auto-fit exceeds its geometry budget.");
            const resolved = evaluatePresentationPreset(preset, width, height, adjustments), rect = resolved.textRect;
            const previousNodes = svg.querySelectorAll("*").length + 1;
            const replacement = createDrawingGeometryView(document, computed, { preset, geometry: resolved, widthPt: width, heightPt: height, id, orientation: { rotationDegrees: g.rotation, flipH: g.flipH, flipV: g.flipV },
              fillFallback: boxPaint ?? (definition.kind === "geometry" && preset !== "line" ? "#4472c4" : "none"),
              strokeFallback: definition.kind === "connector" || preset === "line" ? "#000000" : "none", maxNodes: limits.maxNodes - count + previousNodes });
            account(replacement.querySelectorAll("*").length + 1 - previousNodes);
            styleView(node, { width: width + "pt", height: height + "pt" });
            svg.setAttribute("viewBox", `0 0 ${Math.max(width, 0.001)} ${Math.max(height, 0.001)}`); svg.style.width = width + "pt"; svg.style.height = height + "pt";
            svg.replaceChildren(...replacement.childNodes);
            if (rect) {
              const geometryStyle = document.createElement("div").style; geometryStyle.cssText = textTransform(node, rect);
              for (const name of geometryStyle) content.style.setProperty(name, geometryStyle.getPropertyValue(name));
            }
          },
        });
      }
    } else if (definition.kind === "image") {
      const image = node.tagName === "IMG" ? node as HTMLImageElement : node.querySelector("img")!;
      const l = Number(node.dataset.officeCropLeft ?? 0), r = Number(node.dataset.officeCropRight ?? 0), t = Number(node.dataset.officeCropTop ?? 0), b = Number(node.dataset.officeCropBottom ?? 0);
      if (image === node) {
        // Imported image objects use an outer div. This branch also supports the
        // contract's direct image node without introducing a new source identity.
        const wrapper = markViewNode(document.createElement("div"), "unwrap");
        node.before(wrapper); wrapper.append(node); wrapper.style.cssText = node.style.cssText; wrapper.style.overflow = "hidden";
        styleView(image, { left: -l / (1 - l - r) * 100 + "%", top: -t / (1 - t - b) * 100 + "%", width: 100 / (1 - l - r) + "%", height: 100 / (1 - t - b) + "%", transform: "none", filter: "none" });
      } else { styleView(node, { overflow: "hidden" }); styleView(image, { position: "absolute", left: -l / (1 - l - r) * 100 + "%", top: -t / (1 - t - b) * 100 + "%", width: 100 / (1 - l - r) + "%", height: 100 / (1 - t - b) + "%" }); }
    } else if (definition.kind === "table") {
      const table = node.querySelector("table")!, columns = markViewNode(document.createElement("colgroup")), total = definition.grid.reduce((sum, n) => sum + n, 0);
      for (const width of definition.grid) { const col = document.createElement("col"); col.style.width = width / total * 100 + "%"; columns.append(col); }
      table.prepend(columns);
      styleView(table,{width:"100%",height:"100%","table-layout":"fixed"});
      for (const cell of table.querySelectorAll<HTMLElement>("td,th")) {
        const s = document.defaultView!.getComputedStyle(cell); if (property(s, "--office-cell-fill-kind") === "solid") styleView(cell, { "background-color": resolvedColor(document, s, "--office-cell-fill") });
        for (const side of ["top", "right", "bottom", "left"]) if (property(s, "--office-cell-border-" + side + "-color")) styleView(cell, { ["border-" + side + "-color"]: resolvedColor(document, s, "--office-cell-border-" + side) });
      }
    } else if (definition.kind === "chart") {
      const chart = metadata.charts.find(c => c.chartId === definition.chartId)!, data = metadata.dataSets.find(d => d.dataSetId === chart.dataSetId)!;
      chartScenes.push(renderChart(node, chart, data, g.width, g.height, id, account, { rotationDegrees: g.rotation, flipH: g.flipH, flipV: g.flipV }));
    }
    for (const field of node.querySelectorAll<HTMLElement>('[data-office-field="slidenum"]')) {
      if (field.querySelector(".sdk-field-display")) continue;
      const cache = markViewNode(document.createElement("span"), "unwrap"); cache.append(...field.childNodes); cache.style.display = "none"; hideFromViewSelection(cache);
      const display = markViewNode(document.createElement("span")); display.className = "sdk-field-display"; display.textContent = String(slideIndex + 1); measureViewContent(display); field.append(cache, display);
      source.setFieldView(field.dataset.nodeId!, display);
      if (!fields.some(f => f.fieldNodeId === field.dataset.nodeId)) fields.push({ fieldNodeId: field.dataset.nodeId!, resultNodeIds: [field.dataset.nodeId!], value: String(slideIndex + 1), kind: "PAGE" });
    }
    if (definition.kind !== "chart" && definition.kind !== "group") account(node.querySelectorAll("svg *,svg,.sdk-ppt-text").length);
  };
  for (const [index, slide] of slides.entries()) {
    const entry=entryFor(slide),layout=surfaceId?surfaceLayout:layouts.get(entry.layoutId)!,master=surfaceMaster??(surfaceNotesMaster?{...metadata.masters[0]!,rootNodeId:surfaceNotesMaster.rootNodeId,themeId:surfaceNotesMaster.themeId,otherStyle:surfaceNotesMaster.textStyle}:masters.get(layout!.masterId)!);
    const masterRoot = byId.get(master.rootNodeId)!, layoutRoot = layout?byId.get(layout.rootNodeId)!:masterRoot, theme = metadata.themes.find(t => t.themeId === master.themeId)!;
    ownViewAttribute(masterRoot, "class");for(const t of metadata.themes)masterRoot.classList.remove(t.cssClass);masterRoot.classList.add(theme.cssClass);
    // Base theme classes on surfaces are references. Compose them in order so a
    // master color-map override is not reset by the same base theme on each slide.
    ownViewAttribute(layoutRoot, "class"); ownViewAttribute(slide, "class");for(const t of metadata.themes){if(layoutRoot!==masterRoot)layoutRoot.classList.remove(t.cssClass);if(slide!==masterRoot)slide.classList.remove(t.cssClass);}
    if(layoutRoot!==masterRoot)inheritance.apply(layoutRoot, [masterRoot]);if(slide!==masterRoot)inheritance.apply(slide, layoutRoot===slide?[masterRoot]:[masterRoot,layoutRoot]);
    styleView(slide, { width: size.width + "pt", height: size.height + "pt" });
    const s = document.defaultView!.getComputedStyle(slide), fill = property(s, "--office-fill-kind");
    if (fill === "solid") styleView(slide, { "background-color": resolvedColor(document, s, "--office-fill") });
    else if (fill === "linear") {
      const background = markViewNode(svgElement(document, "svg", { width: "100%", height: "100%", viewBox: `0 0 ${size.width} ${size.height}` })); background.style.cssText = "position:absolute;inset:0;pointer-events:none";
      background.append(svgElement(document, "rect", { width: size.width, height: size.height, fill: svgFill(document, s, background, "slide-bg-" + index, { x: 0, y: 0, width: size.width, height: size.height }) })); slide.prepend(background);
    }
    const instanceIds = new Set<string>(); instancesBySlide.set(slide.dataset.nodeId!, instanceIds);
    const inherited: HTMLElement[] = [];
    if (entry.showMasterShapes !== false&&(!surfaceId||surfaceLayout)) {
      for (const template of [...(layout?.showMasterShapes === false ? [] : surfaceObjects(masterRoot)), ...(surfaceId?[]:surfaceObjects(layoutRoot))]) {
        const definition = objects.get(template.dataset.nodeId!)!, ph = definition.placeholder;
        if (ph && !["ftr", "hdr", "dt", "sldNum"].includes(ph.type)) continue;
        if (ph) {
          const keys: Readonly<Record<string, string>> = { ftr: "footer", hdr: "header", dt: "date", sldNum: "slide-number" };
          const key = keys[ph.type]!;
          const setting = slide.getAttribute("data-office-show-" + key) ?? layoutRoot.getAttribute("data-office-show-" + key) ?? masterRoot.getAttribute("data-office-show-" + key);
          if (setting === "false" || surfaceObjects(slide).some(o => objects.get(o.dataset.nodeId!)?.placeholder?.type === ph.type)) continue;
          if (template.parentElement === masterRoot && surfaceObjects(layoutRoot).some(o => objects.get(o.dataset.nodeId!)?.placeholder?.type === ph.type)) continue;
        }
        account(template.querySelectorAll("*").length + 1);
        const clone = markViewNode(template.cloneNode(true) as HTMLElement); inherited.push(clone);
        for (const element of [clone, ...clone.querySelectorAll<HTMLElement>("[data-node-id]")]) instanceIds.add(element.dataset.nodeId!);
      }
    }
    const first = surfaceObjects(slide)[0]; for (const clone of inherited) { slide.insertBefore(clone, first ?? null); const definition = objects.get(clone.dataset.nodeId!)!; process(clone, definition, [], index, master, true); source.registerInstances(clone); }
    for (const objectNode of surfaceObjects(slide).filter(node => !inherited.includes(node))) {
      const definition = objects.get(objectNode.dataset.nodeId!)!;
      const prototypes = definition.placeholder ? [masterRoot!==objectNode.parentElement?placeholder(masterRoot, definition):undefined,layoutRoot!==objectNode.parentElement?placeholder(layoutRoot, definition):undefined].filter((n): n is HTMLElement => !!n) : [];
      process(objectNode, definition, prototypes, index, master);
    }
    if (entry.hidden) { const label = markViewNode(document.createElement("span")); label.className = "sdk-ppt-hidden-label"; label.textContent = "隐藏幻灯片"; slide.append(label); }
    const note = !surfaceId&&metadata.stories.find(story => story.kind === "notes" && story.slideId === entry.slideId);
    if (note && note.kind === "notes") {
      const body = byId.get(note.bodyNodeId ?? note.rootNodeId)!, panel = markViewNode(document.createElement("details")), heading = document.createElement("summary");
      panel.className = "sdk-ppt-notes"; heading.textContent = "演讲者备注"; panel.append(heading);
      const noteRoot = byId.get(note.rootNodeId)!, noteMaster = metadata.notesMasters.find(m => m.notesMasterId === note.notesMasterId), noteMasterRoot = noteMaster && byId.get(noteMaster.rootNodeId);
      if (noteMaster && noteMasterRoot) {
        const theme = metadata.themes.find(theme => theme.themeId === noteMaster.themeId)!;
        ownViewAttribute(noteMasterRoot, "class");for(const t of metadata.themes)noteMasterRoot.classList.remove(t.cssClass);noteMasterRoot.classList.add(theme.cssClass);
        ownViewAttribute(noteRoot, "class");for(const t of metadata.themes)noteRoot.classList.remove(t.cssClass);
      }
      const content = document.createElement("div"), clone = body.cloneNode(true) as HTMLElement;
      content.className = "sdk-ppt-note-content";
      clone.classList.add("sdk-ppt-rich-story");
      clone.removeAttribute("data-office-kind"); clone.removeAttribute("data-office-story"); content.append(clone); panel.append(content); slide.after(panel);
      inheritance.apply(content, [...(noteMasterRoot ? [noteMasterRoot] : []), noteRoot]);
      const appearance = document.defaultView!.getComputedStyle(content);
      if (property(appearance, "--office-fill-kind") === "solid") content.style.backgroundColor = resolvedColor(document, appearance, "--office-fill");
      if (noteMaster?.textStyle) for (const p of clone.querySelectorAll<HTMLElement>("p")) {
        const prototype = document.createElement("p"); prototype.className = noteMaster.textStyle + " " + noteMaster.textStyle + "-level-" + (p.dataset.officeLevel ?? "0"); inheritance.apply(p, [prototype]);
      }
      account(panel.querySelectorAll("*").length + 1); source.registerInstances(clone);
    }
  }
  for (const { node, definition } of connectors) resolveConnector(node, definition, objects);
  const comments = snapshot.annotations;
  if (comments.length&&!surfaceId) {
    const panel = markViewNode(document.createElement("section")); panel.className = "sdk-ppt-comments"; panel.setAttribute("aria-label", "批注");
    const statuses = { open: "待处理", resolved: "已解决", closed: "已关闭", needs_reanchor: "待重新定位" };
    const appendBody = (parent: HTMLElement, id: string | undefined, plain: string | undefined, modern: boolean): void => {
      const body = id && byId.get(id);
      if (!body) { parent.append(document.createTextNode(plain ?? "")); return; }
      const clone = body.cloneNode(true) as HTMLElement; clone.classList.add("sdk-ppt-rich-story"); parent.append(clone);source.registerInstances(clone);
      const story = metadata.stories.find(story => story.rootNodeId === id), slide = metadata.slides.find(slide => slide.slideId === story?.slideId);
      if (modern && slide) inheritance.apply(clone, [byId.get(masters.get(layouts.get(slide.layoutId)!.masterId)!.rootNodeId)!]);
    };
    root.append(panel);
    for (const comment of comments) {
      const item = document.createElement("div"); item.textContent = (comment.author?.name ?? "") + (comment.date ? " · " + comment.date : "") + " · " + statuses[comment.status] + "\n";
      const info = metadata.comments.find(c => c.annotationId === comment.annotationId);
      if (info) item.dataset.commentFormat = info.format;
      panel.append(item);
      appendBody(item, comment.bodyNodeId, comment.instruction, info?.format !== "classic");
      for (const reply of comment.replies ?? []) {
        const child = document.createElement("div"); child.textContent = (reply.author?.name ?? "") + (reply.date ? " · " + reply.date : "") + "：";
        item.append(child); appendBody(child, reply.bodyNodeId, reply.instruction, info?.format !== "classic");
      }
    }
    account(panel.querySelectorAll("*").length + 1);
  }
  for (const link of metadata.hyperlinks) {
    const node = byId.get(link.nodeId); if (node) { ownViewAttribute(node, "data-sdk-hyperlink"); node.setAttribute("data-sdk-hyperlink", ""); }
  }
  for (const node of root.querySelectorAll<HTMLElement>("[data-node-id]")) {
    const style = document.defaultView!.getComputedStyle(node);
    if (property(style, "--office-text-color")) styleView(node, { color: resolvedColor(document, style, "--office-text") });
  }
  // Script wrappers retain exact UTF-16 text while resolving native font roles.
  const texts: Text[] = [], walker = document.createTreeWalker(root, document.defaultView!.NodeFilter.SHOW_TEXT);
  for (let n = walker.nextNode(); n; n = walker.nextNode()) if (!n.parentElement?.closest("svg,.sdk-field-display")) texts.push(n as Text);
  for (const text of texts) {
    const parent = text.parentElement!; if (!text.textContent) continue;
    const chunks = officeFontSegments(text.data);
    account(chunks.length); const fragment = document.createDocumentFragment();
    for (const chunk of chunks) {
      const span = markViewNode(document.createElement("span"), "unwrap"); span.dataset.officeScript = chunk.role; span.style.fontFamily = `var(--office-font-${chunk.role},var(--office-minor-${chunk.role},${document.defaultView!.getComputedStyle(parent).fontFamily}))`;
      span.textContent = chunk.text; fragment.append(span);
    }
    text.replaceWith(fragment);
  }
  decorateLists(root, account);
  source.reindex(); account(0);
  return { source, fields, finalizeCharts() { for (const scene of chartScenes) prepareChartTextDecorations(scene, account); chartScenes.length = 0; }, measure() {
    for (const box of textBoxes) {
      const apply = (scale: number): void => { for (const { node, size, line } of box.fonts) styleView(node, { "font-size": size * scale + "px", ...(Number.isFinite(line) ? { "line-height": line * scale * (1 - box.lineReduction) + "px" } : {}) }); };
      if (box.fit === "normal") {
        let scale = box.scale; apply(scale);
        for (let pass = 0; pass < limits.maxLayoutPasses && box.content.clientHeight > 0 && (box.content.scrollHeight > box.content.clientHeight + 1 || box.content.scrollWidth > box.content.clientWidth + 1); pass++) {
          scale *= Math.min(0.95, Math.sqrt(box.content.clientHeight / box.content.scrollHeight), box.content.clientWidth / box.content.scrollWidth); scale = Math.max(0.01, scale); apply(scale);
        }
      } else if (box.scale !== 1) apply(box.scale);
      if (box.fit === "shape" && box.content.clientHeight > 0) {
        for (let pass = 0; pass < limits.maxLayoutPasses; pass++) {
          const extraHeight = Math.max(0, box.content.scrollHeight - box.content.clientHeight), extraWidth = Math.max(0, box.content.scrollWidth - box.content.clientWidth);
          if (extraHeight <= 1 && extraWidth <= 1) break;
          box.width += extraWidth * 0.75; box.height += extraHeight * 0.75; box.resize(box.width, box.height);
        }
      }
    }
    for (const { node, definition } of connectors) resolveConnector(node, definition, objects);
    const origin = root.getBoundingClientRect(), pages: RenderedPage[] = [], nodePages: Record<string, string[]> = {};
    for (const [index, node] of slides.entries()) {
      const slide=entryFor(node),rect = node.getBoundingClientRect(), pageId = slide.slideId;
      pages.push({ pageId, index: index + 1, sectionId: slide.slideId, number: index + 1, label: String(index + 1), rect: { x: (rect.left - origin.left) * 0.75, y: (rect.top - origin.top) * 0.75, width: size.width, height: size.height }, setup: { unit: "pt", width: size.width, height: size.height, margins: { top: 0, right: 0, bottom: 0, left: 0 } }, blank: false });
      for (const id of [node.dataset.nodeId!, ...[...node.querySelectorAll<HTMLElement>("[data-node-id]")].map(n => n.dataset.nodeId!), ...instancesBySlide.get(node.dataset.nodeId!) ?? []]) { const entries = nodePages[id] ??= []; if (!entries.includes(pageId)) entries.push(pageId); }
    }
    const overflow = textBoxes.filter(({ content }) => content.offsetWidth && (content.scrollHeight > content.clientHeight + 1 || content.scrollWidth > content.clientWidth + 1)).map(({ node }) => ({ code: "ContentOverflow" as const, nodeId: node.dataset.nodeId!, message: "Slide text exceeds its resolved text rectangle with the selected fonts." }));
    source.setPages(pages); return { pages, nodePages, diagnostics: [...diagnostics, ...overflow] };
  } };
}
