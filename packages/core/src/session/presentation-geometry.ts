import type { DeepReadonly } from "../runtime.js";
import type { PresentationGeometry, PresentationMetadata, PresentationObject } from "../contracts/presentation.js";
import type { PresentationOperation } from "../contracts/operations.js";
import type { ParsedHtml } from "../document/html.js";
import type { DocumentEdit } from "../document/edit.js";
import { readPresentationGeometry, matchPresentationPlaceholder } from "../document/presentation.js";
import { evaluatePresentationPreset } from "../presentation-geometry.js";
import { evaluatePresentationCustomGeometry } from "../presentation-custom-geometry.js";
import { invalid } from "../document/validation.js";

export type Matrix = readonly [number, number, number, number, number, number];
const identity: Matrix = [1, 0, 0, 1, 0, 0];
export function multiply(a: Matrix, b: Matrix): Matrix {
  return [a[0]*b[0]+a[2]*b[1], a[1]*b[0]+a[3]*b[1], a[0]*b[2]+a[2]*b[3], a[1]*b[2]+a[3]*b[3], a[0]*b[4]+a[2]*b[5]+a[4], a[1]*b[4]+a[3]*b[5]+a[5]];
}
export function point(m: Matrix, p: { x: number; y: number }): { x: number; y: number } { return { x: m[0]*p.x+m[2]*p.y+m[4], y: m[1]*p.x+m[3]*p.y+m[5] }; }
export function inverse(m: Matrix): Matrix {
  const determinant = m[0]*m[3]-m[1]*m[2];
  if (!Number.isFinite(determinant) || determinant === 0) invalid("A group transform is not invertible.");
  return [m[3]/determinant, -m[1]/determinant, -m[2]/determinant, m[0]/determinant, (m[2]*m[5]-m[3]*m[4])/determinant, (m[1]*m[4]-m[0]*m[5])/determinant];
}
export function local(g: PresentationGeometry): Matrix {
  const angle = g.rotation*Math.PI/180, c = Math.cos(angle), s = Math.sin(angle), x = g.flipH ? -1 : 1, y = g.flipV ? -1 : 1;
  return [c*x, s*x, -s*y, c*y, g.x+g.width/2-c*x*g.width/2+s*y*g.height/2, g.y+g.height/2-s*x*g.width/2-c*y*g.height/2];
}
const names = ["x", "y", "width", "height", "rotation", "flip-h", "flip-v"] as const;
export function geometryAttributes(g: PresentationGeometry): Record<string, string> {
  return Object.fromEntries(names.map(name => ["data-office-" + name, String(name === "flip-h" ? g.flipH : name === "flip-v" ? g.flipV : g[name])]));
}
const same = (a: PresentationGeometry | undefined, b: PresentationGeometry): boolean => !!a && a.x === b.x && a.y === b.y && a.width === b.width && a.height === b.height && a.rotation === b.rotation && a.flipH === b.flipH && a.flipV === b.flipV;

/** Private derived geometry; no mutable transform tree is retained between transactions. */
export function presentationTransforms(p: DeepReadonly<PresentationMetadata>, html: Pick<ParsedHtml,"peekNode">) {
  const objects = new Map(p.objects.map(o => [o.nodeId, o])), surfaces = new Set([...p.slides, ...p.layouts, ...p.masters, ...p.notesMasters, ...p.stories].map(s => s.rootNodeId));
  const owners = new Map<string, string>(), geometryCache = new Map<string, PresentationGeometry>(), worldCache = new Map<string, Matrix>();
  const owner = (id: string): string => {
    const previous = owners.get(id); if (previous) return previous;
    let current: string | null = id;
    while (current !== null && !surfaces.has(current)) current = html.peekNode(current)?.parentNodeId ?? null;
    if (current === null) return invalid("A presentation object has no source surface.");
    owners.set(id, current); return current;
  };
  const bySurface = new Map<string, Map<string, DeepReadonly<PresentationObject>>>();
  for (const object of p.objects) if (object.placeholder) {
    const surface = owner(object.nodeId), entries = bySurface.get(surface) ?? new Map();
    entries.set(object.placeholder.type + ":" + object.placeholder.index, object); bySurface.set(surface, entries);
  }
  const geometry = (id: string): PresentationGeometry => {
    const cached = geometryCache.get(id); if (cached) return cached;
    const source = html.peekNode(id); if (!source) return invalid("A geometry target is missing.");
    let result = readPresentationGeometry(source.attributes);
    if (!result) {
      const object = objects.get(id), surface = owner(id), slide = p.slides.find(s => s.rootNodeId === surface);
      const layout = slide ? p.layouts.find(l => l.layoutId === slide.layoutId) : p.layouts.find(l => l.rootNodeId === surface);
      const master = layout ? p.masters.find(m => m.masterId === layout.masterId) : undefined;
      const notes = p.stories.find(s => s.kind === "notes" && s.rootNodeId === surface);
      const notesMaster = notes?.kind === "notes" ? p.notesMasters.find(m => m.notesMasterId === notes.notesMasterId) : undefined;
      for (const parent of [layout?.rootNodeId, master?.rootNodeId, notesMaster?.rootNodeId]) {
        if (!parent || parent === surface || !object?.placeholder) continue;
        const inherited = matchPresentationPlaceholder(
          bySurface.get(parent)?.values() ?? [], object.placeholder, candidate => candidate.placeholder,
          parent === layout?.rootNodeId ? "layout" : "master",
        );
        if (inherited) { result = readPresentationGeometry(html.peekNode(inherited.nodeId)!.attributes); if (result) break; }
      }
    }
    if (!result) return invalid("A presentation transform cannot be resolved.");
    geometryCache.set(id, result); return result;
  };
  const parent = (id: string): Matrix => {
    const parentId = html.peekNode(id)?.parentNodeId, object = parentId ? objects.get(parentId) : undefined;
    if (object?.kind !== "group") return identity;
    const g = geometry(object.nodeId), a = html.peekNode(object.nodeId)!.attributes;
    const sx = g.width / Number(a["data-office-child-width"]), sy = g.height / Number(a["data-office-child-height"]);
    return multiply(world(object.nodeId), [sx, 0, 0, sy, -Number(a["data-office-child-x"])*sx, -Number(a["data-office-child-y"])*sy]);
  };
  const world = (id: string): Matrix => {
    let result = worldCache.get(id); if (result) return result;
    result = multiply(parent(id), local(geometry(id))); worldCache.set(id, result); return result;
  };
  const endpoint = (id: string, site: number): { x: number; y: number } => {
    const object = objects.get(id); if (!object) return invalid("A connector target is missing.");
    const g = geometry(id), adjustments = Object.fromEntries(Object.entries(html.peekNode(id)!.attributes).filter(([key]) => key.startsWith("data-office-adjust-")).map(([key, value]) => [key.slice(19), Number(value)]));
    const custom = object.kind === "geometry" || object.kind === "image" ? object.customGeometry : undefined;
    const connection = (custom ? evaluatePresentationCustomGeometry(custom, g.width, g.height) :
      evaluatePresentationPreset(object.kind === "geometry" ? object.preset : "rect", g.width, g.height, object.kind === "geometry" ? adjustments : {})).connections[site];
    if (!connection) return invalid("A connector site does not exist on its target preset.");
    return point(world(id), connection);
  };
  return { owner, geometry, world, parent, endpoint };
}

export function applyPresentationGeometry(operation: Extract<PresentationOperation, { op: "setPresentationGeometry" }>, p: PresentationMetadata, edit: DocumentEdit, before: ParsedHtml): void {
  const object = p.objects.find(item => item.nodeId === operation.nodeId);
  if (!object) invalid("Geometry operations require a registered presentation object.");
  if (operation.childCoordinates && object.kind !== "group") invalid("Only groups have child coordinates.");
  const source = before.peekNode(operation.nodeId)!.attributes;
  const attrs: Record<string, string | undefined> = same(readPresentationGeometry(source), operation.geometry)
    ? Object.fromEntries(names.map(name => ["data-office-" + name, source["data-office-" + name]])) : geometryAttributes(operation.geometry);
  if (operation.childCoordinates) for (const key of ["x", "y", "width", "height"] as const) {
    const name = "data-office-child-" + key, value = operation.childCoordinates[key];
    attrs[name] = Number(source[name]) === value ? source[name] : String(value);
  }
  edit.attributes(operation.nodeId, attrs);
}

/** Bound connector caches move with targets in the same candidate and history entry. */
export function updatePresentationConnectors(p: PresentationMetadata, beforeMetadata: DeepReadonly<PresentationMetadata>, before: ParsedHtml, after: ParsedHtml, edit: DocumentEdit): void {
  const previous = presentationTransforms(beforeMetadata, before), next = presentationTransforms(p, after);
  const originalObjects = new Map(beforeMetadata.objects.map(object => [object.nodeId, object]));
  for (const object of p.objects) {
    if (object.kind !== "connector" || !after.hasNode(object.nodeId) || !object.startNodeId && !object.endNodeId) continue;
    const targets = [object.startNodeId, object.endNodeId].filter((id): id is string => id !== undefined && before.hasNode(id) && after.hasNode(id));
    const original = originalObjects.get(object.nodeId);
    const bindingChanged = original?.kind !== "connector" || object.routing!==original.routing || object.startNodeId !== undefined && (object.startNodeId !== original.startNodeId || object.startSite !== original.startSite)
      || object.endNodeId !== undefined && (object.endNodeId !== original.endNodeId || object.endSite !== original.endSite);
    const changed = !before.hasNode(object.nodeId) || bindingChanged
      || targets.some(id => JSON.stringify(originalObjects.get(id)) !== JSON.stringify(p.objects.find(o => o.nodeId === id)) || JSON.stringify(previous.world(id)) !== JSON.stringify(next.world(id)) || JSON.stringify(previous.geometry(id)) !== JSON.stringify(next.geometry(id)))
      || JSON.stringify(previous.world(object.nodeId)) !== JSON.stringify(next.world(object.nodeId));
    if (!changed) continue;
    const g = next.geometry(object.nodeId), transform = next.world(object.nodeId), parentInverse = inverse(next.parent(object.nodeId));
    const a = point(parentInverse, object.startNodeId ? next.endpoint(object.startNodeId, object.startSite ?? 0) : point(transform, { x: 0, y: 0 }));
    const b = point(parentInverse, object.endNodeId ? next.endpoint(object.endNodeId, object.endSite ?? 0) : point(transform, { x: g.width, y: g.height }));
    const clean = (value: number): number => { if (!Number.isFinite(value)) return invalid("A connector transform exceeds finite geometry."); return Math.abs(value) < 1e-10 ? 0 : value; };
    const result: PresentationGeometry = { x: clean(Math.min(a.x, b.x)), y: clean(Math.min(a.y, b.y)), width: clean(Math.abs(b.x-a.x)), height: clean(Math.abs(b.y-a.y)), rotation: 0, flipH: b.x<a.x, flipV: b.y<a.y };
    edit.attributes(object.nodeId, geometryAttributes(result));
  }
}
