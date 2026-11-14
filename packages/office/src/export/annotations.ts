import { DocumentError } from "@doffi/core";
import type { Annotation, DeepReadonly } from "@doffi/core";
import type { SourceElement, SourceNode, SourceText, SourceTree } from "./source.js";
import { element } from "./source.js";
import { empty } from "./xml.js";

export interface UnexportedAnnotation { readonly annotationId: string; readonly reason: "region" | "needs_reanchor" | "disjoint" | "cross_story" | "unrepresentable" }
export interface ExportComment { readonly id: string; readonly annotation: DeepReadonly<Annotation>; readonly paragraphId: string }
interface Endpoint { node: SourceNode; offset: number }
interface Range { start: Endpoint; end: Endpoint; story: SourceElement; first: number; last: number }
/** Ranges bind to concrete source leaves. Equal paragraph-end/start offsets do
 * not accidentally move a comment onto the neighbouring paragraph. */
export class ExportAnnotations {
  readonly comments: ExportComment[] = [];
  readonly omitted: UnexportedAnnotation[] = [];
  private readonly events = new WeakMap<SourceNode, Map<number, { starts: string[]; ends: string[] }>>();
  constructor(tree: SourceTree, annotations: readonly DeepReadonly<Annotation>[], stories: readonly SourceElement[], maximumWork: number, emitted: (node: SourceText) => boolean) {
    const leaves: SourceNode[] = [];
    let work = 0;
    const step = (): void => { if (++work > maximumWork) throw new DocumentError("ResourceLimitExceeded", "DOCX annotation mapping exceeds its work budget."); };
    const contains = (parent: SourceElement, child: SourceNode): boolean => {
      for (let current: SourceNode | undefined = child; current; current = current.parent) { step(); if (current === parent) return true; }
      return false;
    };
    const visit = (node: SourceNode): void => {
      if (!element(node)) { if (emitted(node)) leaves.push(node); }
      else if (["img", "br"].includes(node.tag)||node.attrs["data-word-drawing"]||node.attrs["data-word-native-object"]) leaves.push(node); else node.children.forEach(visit);
    };
    stories.forEach(visit);
    const range = (id: string, from?: number, to?: number): Range | undefined => {
      const node = tree.byId.get(id); if (!node) return;
      const story = stories.find((story) => contains(story, node)); if (!story || story.attrs["data-office-story"] === "comment") return;
      const begin = node.from + (from ?? 0), end = node.from + (to ?? node.to - node.from);
      const candidates = leaves.filter((leaf) => { step(); return contains(node, leaf) && leaf.to > begin && leaf.from < end; });
      if (!candidates.length) {
        if (begin !== end) return;
        const point = leaves.find((leaf) => { step(); return contains(node, leaf) && leaf.from <= begin && leaf.to >= begin; });
        if (point) return { start: { node: point, offset: begin - point.from }, end: { node: point, offset: begin - point.from }, story, first: begin, last: end };
        // Empty markers/paragraphs have an explicit emission site.
        if (node.to !== node.from || !["p", "h1", "h2", "h3", "h4", "h5", "h6", "span", "strong", "em", "u", "s", "sub", "sup", "ins", "del"].includes(node.tag)) return;
        let paragraph: SourceElement | undefined = node;
        while (paragraph && !/^p$|^h[1-6]$/.test(paragraph.tag)) paragraph = paragraph.parent;
        if (!paragraph) return;
        return { start: { node, offset: 0 }, end: { node, offset: 0 }, story, first: begin, last: end };
      }
      const first = candidates[0]!, last = candidates.at(-1)!;
      return { start: { node: first, offset: Math.max(begin, first.from) - first.from }, end: { node: last, offset: Math.min(end, last.to) - last.from }, story, first: begin, last: end };
    };
    for (const annotation of annotations) {
      let reason: UnexportedAnnotation["reason"] | undefined;
      const anchor = annotation.anchor;
      if (annotation.status === "needs_reanchor") reason = "needs_reanchor";
      else if (anchor.kind === "region") reason = "region";
      const ranges: (Range | undefined)[] = reason ? [] : anchor.kind === "text" ? [range(anchor.nodeId, anchor.from, anchor.to)] : anchor.kind === "text-set" ? anchor.ranges.map((item) => range(item.nodeId, item.from, item.to)) : anchor.kind === "node" ? [range(anchor.nodeId)] : anchor.kind === "node-set" ? anchor.nodeIds.map((id) => range(id)) : [];
      if (!reason && (!ranges.length || ranges.some((item) => !item))) reason = "unrepresentable";
      const sorted = (ranges.filter((item): item is Range => Boolean(item))).sort((left, right) => left.first - right.first);
      if (!reason && sorted.some((item) => item.story !== sorted[0]!.story)) reason = "cross_story";
      let last = sorted[0]?.last ?? 0;
      if (!reason) for (const item of sorted.slice(1)) {
        if (item.first > last && leaves.some((leaf) => { step(); return contains(item.story, leaf) && leaf.from < item.first && leaf.to > last; })) { reason = "disjoint"; break; }
        last = Math.max(last, item.last);
      }
      if (reason) { this.omitted.push({ annotationId: annotation.annotationId, reason }); continue; }
      const id = String(this.comments.length), paragraphId = (this.comments.length + 1).toString(16).padStart(8, "0").toUpperCase();
      this.comments.push({ id, annotation, paragraphId });
      const first = sorted[0]!, end = sorted.reduce((left, right) => left.last >= right.last ? left : right).end;
      this.add(first.start, id, "starts"); this.add(end, id, "ends");
    }
  }
  private add(endpoint: Endpoint, id: string, kind: "starts" | "ends"): void {
    let offsets = this.events.get(endpoint.node); if (!offsets) { offsets = new Map(); this.events.set(endpoint.node, offsets); }
    let event = offsets.get(endpoint.offset); if (!event) { event = { starts: [], ends: [] }; offsets.set(endpoint.offset, event); }
    event[kind].push(id);
  }
  offsets(node: SourceNode): number[] { return [...this.events.get(node)?.keys() ?? []].sort((a, b) => a - b); }
  at(node: SourceNode, offset: number): string {
    const event = this.events.get(node)?.get(offset); if (!event) return "";
    this.events.get(node)!.delete(offset);
    return event.starts.map((id) => empty("w:commentRangeStart", { "w:id": id })).join("") + event.ends.map((id) => empty("w:commentRangeEnd", { "w:id": id }) + '<w:r>' + empty("w:commentReference", { "w:id": id }) + '</w:r>').join("");
  }
}
