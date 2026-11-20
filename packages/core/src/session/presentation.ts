import type { Annotation } from "../contracts/annotations.js";
import type { DocumentContentOperation } from "../contracts/operations.js";
import type { ContentDocumentMetadata } from "../contracts/document.js";
import type { DocumentState } from "../document/state.js";
import type { ParsedHtml } from "../document/html.js";
import type { DocumentEdit } from "../document/edit.js";
import { assetDigest } from "../document/assets.js";
import { encode } from "../document/package.js";
import { invalid } from "../document/validation.js";
import { presentationStyleImpact } from "./presentation-style-impact.js";

export function guardPresentationContent(base: DocumentState, op: DocumentContentOperation): void {
  const p = base.snapshot.metadata.presentation;
  if (!p) return;
  const object = p.objects.find(o => o.nodeId === op.nodeId), node = base.html.peekNode(op.nodeId)!;
  if (op.op === "setStyle" && node.attributes["data-office-kind"] !== undefined && /^(?:(?:(?:min|max)-)?(?:width|height)|margin(?:-(?:top|right|bottom|left))?)$/i.test(op.property)) invalid("Presentation object coordinates and dimensions require a typed geometry operation.");
  if (op.op === "removeNode" && ["presentation", "slide", "master", "layout", "notes-master", "notes", "comment"].includes(node.attributes["data-office-kind"] ?? "")) invalid("Presentation surfaces require a format management operation.");
  if (["replaceText", "replaceFragment", "insertFragment", "reconcileFragment"].includes(op.op) && node.attributes["data-office-kind"] !== undefined
    && (!object || !["text", "geometry", "table", "chart"].includes(object.kind)) && !["notes", "comment"].includes(node.attributes["data-office-kind"]!)) invalid("Content operations require a presentation text content domain.");
}

/** Only complete deleted objects are reconciled; damaged required references still fail. */
export function updatePresentationMetadata(metadata: ContentDocumentMetadata, before: DocumentState, after: ParsedHtml, annotations: Annotation[]): void {
  const p = metadata.presentation;
  if (!p) return;
  p.objects = p.objects.filter(item => after.hasNode(item.nodeId));
  if (p.nativePreview) p.nativePreview = p.nativePreview.filter(item => item.feature !== "tags" || after.hasNode(item.ownerId));
  const objects = new Set(p.objects.map(o => o.nodeId));
  for (const object of p.objects) if (object.kind === "connector") {
    if (object.startNodeId && !objects.has(object.startNodeId)) { delete object.startNodeId; delete object.startSite; }
    if (object.endNodeId && !objects.has(object.endNodeId)) { delete object.endNodeId; delete object.endSite; }
  }
  p.charts = p.charts.filter(item => after.hasNode(item.nodeId));
  const usedData = new Set(p.charts.map(c => c.dataSetId)), previousData = new Set(before.snapshot.metadata.presentation?.charts.map(c => c.dataSetId));
  p.dataSets = p.dataSets.filter(item => usedData.has(item.dataSetId) || !previousData.has(item.dataSetId));
  p.hyperlinks = p.hyperlinks.filter(item => after.hasNode(item.nodeId));
  p.stories = p.stories.filter(item => item.kind !== "comment" || after.hasNode(item.rootNodeId));
  const byId = new Map(annotations.map(a => [a.annotationId, a]));
  p.comments = p.comments.filter(item => byId.has(item.annotationId));
  const existing = new Set(p.comments.map(c => c.annotationId));
  const previousAnnotations = new Set(before.snapshot.annotations.map(a => a.annotationId));
  for (const annotation of annotations) if (!existing.has(annotation.annotationId) && !previousAnnotations.has(annotation.annotationId)) p.comments.push({ annotationId: annotation.annotationId, format: "modern" });
}

export function removeDeletedCommentBodies(before: DocumentState, annotations: readonly Annotation[], edit: DocumentEdit,metadata:ContentDocumentMetadata): void {
  if (!before.snapshot.metadata.presentation) return;
  const kept = new Set(annotations.flatMap(a => [a.bodyNodeId, ...(a.replies ?? []).map(r => r.bodyNodeId)]).filter(id => id !== undefined));
  const originals=new Set(before.snapshot.metadata.presentation.stories.map(s=>s.rootNodeId));
  const roots=new Set([...before.snapshot.annotations.flatMap(a=>[a,...a.replies??[]]).flatMap(b=>b.bodyNodeId?[b.bodyNodeId]:[]),...metadata.presentation!.stories.filter(s=>s.kind==="comment"&&!originals.has(s.rootNodeId)).map(s=>s.rootNodeId)]);
  for(const root of roots)if(!kept.has(root))edit.removeSource(root);
}

export function updatePresentationRecovery(metadata: ContentDocumentMetadata, before: DocumentState, annotations: readonly Annotation[],workflows:ReadonlyMap<string,Annotation["status"]>): void {
  const p = metadata.presentation;
  if (!p) return;
  const originals = new Map(before.snapshot.annotations.map(a => [a.annotationId, a]));
  for (const comment of p.comments) {
    const annotation = annotations.find(a => a.annotationId === comment.annotationId)!;
    const original = originals.get(comment.annotationId);
    if (annotation.status === "needs_reanchor") {
      const workflow=workflows.get(comment.annotationId);
      if(workflow&&workflow!=="needs_reanchor")comment.recoveryStatus=workflow;
      else if (original && original.status !== "needs_reanchor") comment.recoveryStatus = original.status;
    } else delete comment.recoveryStatus;
  }
}

/** A derived fingerprint of each slide and its transitive source/data dependencies. */
export function presentationSlideFingerprints(state: Pick<DocumentState, "snapshot" | "html">,layoutOnly=false): Map<string, string> {
  const { snapshot, html } = state, p = snapshot.metadata.presentation, result = new Map<string, string>();
  if (!p) return result;
  const digest = (value: unknown): string => assetDigest(encode(JSON.stringify(value)));
  const hashes = new Map<string, string>();
  const source = (id: string): string => {
    let hash = hashes.get(id); if (hash) return hash;
    const node = html.peekNode(id);
    if (!node) return "missing";
    hash = digest([html.signature(id), node.childNodeIds.map(source)]); hashes.set(id, hash); return hash;
  };
  const roots = new Set([...p.slides, ...p.masters, ...p.layouts, ...p.notesMasters, ...p.stories].map(s => s.rootNodeId));
  const bySurface = new Map<string, unknown[]>(), owners = new Map<string, string | undefined>();
  const owner = (id: string): string | undefined => {
    if (owners.has(id)) return owners.get(id);
    const node = html.peekNode(id), parent = node?.parentNodeId;
    const value = roots.has(id) ? id : parent ? owner(parent) : undefined; owners.set(id, value); return value;
  };
  const charts = new Map(p.charts.map(c => [c.chartId, c])), data = new Map(p.dataSets.map(d => [d.dataSetId, d]));
  for (const object of [...p.objects].sort((a,b) => a.nodeId.localeCompare(b.nodeId))) {
    const surface = owner(object.nodeId); if (!surface) continue;
    const items = bySurface.get(surface) ?? [], chart = object.kind === "chart" ? charts.get(object.chartId) : undefined;
    items.push([object, chart, chart ? data.get(chart.dataSetId) : undefined]); bySurface.set(surface, items);
  }
  for(const link of [...p.hyperlinks].sort((a,b)=>a.nodeId.localeCompare(b.nodeId))){const surface=owner(link.nodeId);if(surface){const items=bySurface.get(surface)??[];items.push({hyperlink:link});bySurface.set(surface,items);}}
  const surface = (id: string): string => digest([source(id), bySurface.get(id)]);
  const surfaces = new Map([...roots].map(id => [id, surface(id)]));
  const root=html.peekNode(html.nodeIds().find(id=>html.peekNode(id)?.attributes["data-office-kind"]==="presentation")!)!;
  const shared = digest([p.size,root.attributes]);
  const css=presentationStyleImpact(snapshot.styles),symbols=new Map<string,{classes:Set<string>;ids:Set<string>}>();
  const collect=(id:string):{classes:Set<string>;ids:Set<string>}=>{
    const previous=symbols.get(id);if(previous)return previous;const node=html.peekNode(id)!,value={classes:new Set((node.attributes.class??"").split(/\s+/).filter(Boolean)),ids:new Set(node.attributes.id?[node.attributes.id]:[])};
    for(const child of node.childNodeIds){const nested=collect(child);for(const name of nested.classes)value.classes.add(name);for(const name of nested.ids)value.ids.add(name);}symbols.set(id,value);return value;
  };
  const layouts = new Map(p.layouts.map(l => [l.layoutId, l])), masters = new Map(p.masters.map(m => [m.masterId, m])), themes = new Map(p.themes.map(t => [t.themeId, t]));
  for (const slide of p.slides) {
    const layout = layouts.get(slide.layoutId), master = layout ? masters.get(layout.masterId) : undefined;
    const classes=new Set((root.attributes.class??"").split(/\s+/)),ids=new Set(root.attributes.id?[root.attributes.id]:[]);
    for(const surface of [slide.rootNodeId,layout?.rootNodeId,master?.rootNodeId])if(surface){const value=collect(surface);for(const name of value.classes)classes.add(name);for(const name of value.ids)ids.add(name);}
    if(master)classes.add(themes.get(master.themeId)!.cssClass);
    for(const key of [master?.titleStyle,master?.bodyStyle,master?.otherStyle])if(key){classes.add(key);for(let level=0;level<9;level++)classes.add(key+"-level-"+level);}
    const dependencies=new Set([slide.rootNodeId,layout?.rootNodeId,master?.rootNodeId]);
    for(const object of p.objects)if(dependencies.has(owner(object.nodeId))&&object.kind==="chart"){const chart=charts.get(object.chartId)!;for(const[key,value]of Object.entries(chart))if(key.endsWith("CssClass")&&typeof value==="string")classes.add(value);for(const style of [...chart.seriesStyles??[],...chart.pointStyles??[]]){classes.add(style.cssClass);if(style.labelCssClass)classes.add(style.labelCssClass);}}
    const {name:_slideName,...slideLayout}=slide,{name:_layoutName,...layoutGeometry}=layout??{};
    result.set(slide.slideId, digest([shared,css(classes,ids),layoutOnly?slideLayout:slide, surfaces.get(slide.rootNodeId), layoutOnly?layoutGeometry:layout, layout ? surfaces.get(layout.rootNodeId) : null, master, master ? surfaces.get(master.rootNodeId) : null, master ? themes.get(master.themeId) : null]));
  }
  return result;
}

export function affectedPresentationSlides(before: Pick<DocumentState, "snapshot" | "html">, after: Pick<DocumentState, "snapshot" | "html">,layoutOnly=false): string[] {
  const a = presentationSlideFingerprints(before,layoutOnly), b = presentationSlideFingerprints(after,layoutOnly);
  return [...new Set([...a.keys(), ...b.keys()])].filter(id => a.get(id) !== b.get(id));
}
